import fitz  # PyMuPDF
import cv2
import numpy as np
import os
import re
import uuid
import threading  # guard against duplicate processing requests for one file.
import ipaddress
from urllib.parse import urlsplit
from extraction_engine import PdfExtractionEngine, OcrEngine, UrlValidator

from flask import Flask, jsonify, render_template, request, send_from_directory, Response

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
QR_DIR = os.path.join(UPLOAD_DIR, "qr")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(QR_DIR, exist_ok=True)

app = Flask(__name__, template_folder=TEMPLATES_DIR)

EMAIL_REGEX = re.compile(
    r"(?i)\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,24}\b"
)

# One broad candidate matcher preserves protocol, query and fragment text.
URL_CANDIDATE_REGEX = re.compile(
    r'''(?ix)\b(?:
        (?:https?|ftp)://[^\s<>()\[\]{}"']+
        |mailto:[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,24}
        |tel:\+?[0-9][0-9().\s-]{5,}[0-9]
        |www\.[^\s<>()\[\]{}"']+
        |(?:\d{1,3}\.){3}\d{1,3}(?::\d{2,5})?(?:[/?#][^\s<>()\[\]{}"']*)?
        |(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,24}
         (?::\d{2,5})?(?:[/?#][^\s<>()\[\]{}"']*)?
    )'''
)

TRAILING_PUNCT = ").,;:'\"'!?]>"

# Single shared OCR engine: PaddleOCR when available, EasyOCR fallback
# otherwise (see OcrEngine in extraction_engine.py). Loaded once at import
# time, never per page or per request.
reader = OcrEngine(lang="en", gpu=False)
url_validator = UrlValidator()

app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024

# global progress: { file_id: {current_page, total_pages} }
progress = {}
processing_files = set()  # tracks PDFs currently being processed.
processing_lock = threading.Lock()  # protects processing_files across Flask threads.

# OpenCV detector objects are cached per worker thread.  This avoids repeated
# native allocations while also avoiding unsafe concurrent use of one object.
_qr_local = threading.local()


def get_qr_detector():
    detector = getattr(_qr_local, "detector", None)
    if detector is None:
        detector = cv2.QRCodeDetector()
        _qr_local.detector = detector
    return detector


def is_valid_url_candidate(token: str, strict_ocr: bool = False) -> bool:
    """Validate schemes, hosts, ports and OCR-specific artifact patterns."""
    if not token or any(ch.isspace() for ch in token if not token.lower().startswith("tel:")):
        return False

    lower = token.lower()
    if lower.startswith("mailto:"):
        return bool(EMAIL_REGEX.fullmatch(token[7:]))
    if lower.startswith("tel:"):
        return bool(re.fullmatch(r"tel:\+?[0-9][0-9(). -]{5,}[0-9]", token, re.I))

    explicit_scheme = bool(re.match(r"(?i)^(?:https?|ftp)://", token))
    parse_value = token if explicit_scheme else f"//{token}"
    try:
        parsed = urlsplit(parse_value)
        host = parsed.hostname or ""
        if not host or parsed.port is not None and not (1 <= parsed.port <= 65535):
            return False
    except ValueError:
        return False

    # Reject title-cased sentence fragments masquerading as bare
    # domains (for example, Code.These) without using domain blacklists.
    if not explicit_scheme and not token.lower().startswith("www.") and not parsed.path:
        original_host = token.split(":", 1)[0].rstrip(".")
        labels = original_host.split(".")
        if len(labels) >= 2 and all(re.fullmatch(r"[A-Z][a-z]+", label) for label in labels):
            return False

    try:
        ipaddress.ip_address(host)
        return True
    except ValueError:
        pass

    if len(host) > 253:
        return False
    if "." not in host:
        return explicit_scheme and bool(re.fullmatch(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?", host, re.I))
    labels = host.rstrip(".").split(".")
    if any(not re.fullmatch(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?", label, re.I) for label in labels):
        return False
    tld = labels[-1]
    if not (re.fullmatch(r"[a-z]{2,24}", tld, re.I) or tld.lower().startswith("xn--")):
        return False

    return True


def extract_links_exact_with_pos(text, strict_ocr=False):
    if not text:
        return []

    # Occurrence-safe: never join separate lines or construct a URL candidate.
    # Every result must exist as one exact contiguous string in the source text.
    reconstructed = text

    # Match the extractor's exact source text. OCR spacing/case is not
    # repaired because doing so would construct a URL not present in the PDF.
    cleaned = reconstructed
    if not strict_ocr:
        # Exclude bare email addresses but preserve explicit mailto: URLs exactly.
        cleaned = EMAIL_REGEX.sub(
            lambda match: match.group(0)
            if cleaned[max(0, match.start() - 7):match.start()].lower() == "mailto:"
            else "",
            cleaned
        )
    found = []

    for match in URL_CANDIDATE_REGEX.finditer(cleaned):
        token = strip_trailing(match.group(0))
        if token and is_valid_url_candidate(token, strict_ocr=strict_ocr):
            found.append((token, match.start()))

    return found


def strip_trailing(token: str) -> str:
    return (token or "").strip().rstrip(TRAILING_PUNCT)


def render_page_bgr(page: fitz.Page, zoom: float) -> np.ndarray:
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    rgb = np.frombuffer(
        pix.samples, dtype=np.uint8
    ).reshape(pix.height, pix.width, 3)
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def detect_qr_multi(img_bgr: np.ndarray, thorough=True):
    qr = get_qr_detector()
    out = []

    # Fast path: one native decode pass for digital text-heavy pages.
    if not thorough:
        if hasattr(qr, "detectAndDecodeMulti"):
            ok, data_list, points, _ = qr.detectAndDecodeMulti(img_bgr)
            if ok and data_list is not None:
                for index, data in enumerate(data_list):
                    if data:
                        pts = points[index] if points is not None else None
                        out.append((data, pts))
        if not out:
            data, points, _ = qr.detectAndDecode(img_bgr)
            if data:
                out.append((data, points))
        return out

    def decode_variant(image, point_mapper=None):
        decoded = []
        if hasattr(qr, "detectAndDecodeMulti"):
            ok, data_list, points, _ = qr.detectAndDecodeMulti(image)
            if ok and data_list is not None:
                for index, data in enumerate(data_list):
                    pts = points[index] if points is not None else None
                    if pts is not None and point_mapper:
                        pts = point_mapper(pts) if point_mapper else None
                    decoded.append((data, pts))

        data, points, _ = qr.detectAndDecode(image) if not decoded else ("", None, None)
        if data:
            if points is not None and point_mapper:
                points = point_mapper(points) if point_mapper else None
            decoded.append((data, points))

        for data, points in decoded:
            if data:
                out.append((data, points))

    # Strict staged fallback with early exit: 300 DPI, 600-equivalent, image
    # enhancement, then rotations.  Each expensive stage runs only on a miss.
    decode_variant(img_bgr)
    if out:
        return out
    base = cv2.resize(img_bgr, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    decode_variant(base, point_mapper=lambda points: np.asarray(points) / 2.0)
    if out:
        return out
    gray = cv2.cvtColor(base, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8)).apply(gray)
    sharpen = cv2.addWeighted(
        clahe, 1.8, cv2.GaussianBlur(clahe, (0, 0), 1.0), -0.8, 0
    )
    threshold = cv2.adaptiveThreshold(
        sharpen, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 31, 5
    )
    for prepared in (clahe, sharpen, threshold):
        decode_variant(prepared, point_mapper=lambda points: np.asarray(points) / 2.0)
        if out:
            return out

    height, width = base.shape[:2]
    for variant in (sharpen, threshold):
        for rotation in (cv2.ROTATE_90_CLOCKWISE, cv2.ROTATE_180,
                         cv2.ROTATE_90_COUNTERCLOCKWISE):
            candidate = cv2.rotate(variant, rotation)

            def restore(points, code=rotation):
                pts = np.array(points, dtype=np.float32).reshape(-1, 2).copy()
                xs, ys = pts[:, 0].copy(), pts[:, 1].copy()
                if code == cv2.ROTATE_90_CLOCKWISE:
                    pts[:, 0], pts[:, 1] = ys, height - 1 - xs
                elif code == cv2.ROTATE_180:
                    pts[:, 0], pts[:, 1] = width - 1 - xs, height - 1 - ys
                else:
                    pts[:, 0], pts[:, 1] = width - 1 - ys, xs
                return pts / 2.0

            decode_variant(candidate, point_mapper=restore)
            if out:
                return out

    return out


def crop_quad(img_bgr: np.ndarray, pts, pad=14):
    if pts is None:
        return None

    pts = np.array(pts, dtype=np.float32).reshape(-1, 2)
    x, y, w, h = cv2.boundingRect(pts.astype(np.int32))

    x = max(0, x - pad)
    y = max(0, y - pad)
    w = min(img_bgr.shape[1] - x, w + 2 * pad)
    h = min(img_bgr.shape[0] - y, h + 2 * pad)

    return img_bgr[y:y+h, x:x+w]


def extract_url_from_qr_payload_exact(payload: str):
    if not payload:
        return None
    matches = url_validator.candidates(payload.strip())
    return matches[0][0] if matches else None


def detect_qr_with_preview(
    page: fitz.Page,
    file_id: str,
    page_num: int,
    img_bgr: np.ndarray = None,
    thorough: bool = True,
    render_zoom: float = 4.1667
):
    # Reuse the cached 300-DPI render shared with conditional OCR.
    if img_bgr is None:
        img_bgr = render_page_bgr(page, 4.1667)

    items = []

    detections = detect_qr_multi(img_bgr, thorough=thorough)
    # Final cheap fallback: isolate embedded-image regions and retry at native
    # crop scale.  Cropping removes page clutter without rerendering.
    if not detections and page is not None:
        for image in page.get_images(full=True):
            for rect in page.get_image_rects(image[0]):
                x0, y0 = max(0, int(rect.x0 * render_zoom)), max(0, int(rect.y0 * render_zoom))
                x1 = min(img_bgr.shape[1], int(np.ceil(rect.x1 * render_zoom)))
                y1 = min(img_bgr.shape[0], int(np.ceil(rect.y1 * render_zoom)))
                crop = img_bgr[y0:y1, x0:x1]
                if crop.size == 0:
                    continue
                for payload, points in detect_qr_multi(crop, thorough=False):
                    if points is not None:
                        points = np.asarray(points, dtype=np.float32)
                        points[..., 0] += x0
                        points[..., 1] += y0
                    detections.append((payload, points))

    for d, pts in detections:
        pay = (d or "").strip()
        if not pay:
            continue

        preview_url = None
        bbox = None
        crop = crop_quad(img_bgr, pts)

        if pts is not None:
            point_array = np.array(pts, dtype=np.float32).reshape(-1, 2)
            bbox = [
                float(point_array[:, 0].min() / render_zoom),
                float(point_array[:, 1].min() / render_zoom),
                float(point_array[:, 0].max() / render_zoom),
                float(point_array[:, 1].max() / render_zoom),
            ]

        if crop is not None and crop.size > 0:
            name = f"{file_id}_p{page_num}_{len(items)}_{abs(hash(pay))}.png"
            cv2.imwrite(os.path.join(QR_DIR, name), crop)
            preview_url = f"/qr/{name}"

        items.append({
            "page": page_num,
            "payload": pay,
            "url": extract_url_from_qr_payload_exact(pay),
            "preview": preview_url,
            "bbox": bbox,
            "source": "QR"  # retain extractor metadata internally/API-side.
        })

    return items


# The Flask API delegates to the modular page-parallel extraction engine;
# routes and the response schema remain unchanged. Worker count is derived
# from available CPUs and page count (capped) rather than fixed.
def process_pdf(pdf_path: str, file_id: str) -> dict:
    engine = PdfExtractionEngine(
        ocr_reader=reader,
        qr_detector=detect_qr_with_preview,
        max_workers=min(os.cpu_count() or 1, 6),
    )
    return engine.extract(pdf_path, file_id, progress[file_id])


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/assets/<path:filename>")
def assets(filename):
    path = os.path.join(TEMPLATES_DIR, filename)

    if not os.path.isfile(path):
        return jsonify({"error": "Asset not found"}), 404

    if filename.endswith(".css"):
        return Response(
            open(path, "rb").read(),
            mimetype="text/css"
        )

    if filename.endswith(".js"):
        return Response(
            open(path, "rb").read(),
            mimetype="application/javascript"
        )

    return Response(
        open(path, "rb").read(),
        mimetype="application/octet-stream"
    )


@app.post("/api/upload")
def upload_pdf():
    if "file" not in request.files:
        return jsonify({"error": "Missing file"}), 400

    f = request.files["file"]

    if not f.filename:
        return jsonify({"error": "No selected file"}), 400

    if not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF allowed"}), 400

    file_id = str(uuid.uuid4())
    f.save(os.path.join(UPLOAD_DIR, f"{file_id}.pdf"))

    # init minimal progress
    progress[file_id] = {
        "current_page": 0,
        "total_pages": 0,
        "done": False,  # keep progress JSON consistent from upload onward.
        "links_count": 0,
        "qr_count": 0
    }

    return jsonify({"file_id": file_id}), 200


@app.get("/api/process/<file_id>")
def api_process(file_id):
    pdf_path = os.path.join(UPLOAD_DIR, f"{file_id}.pdf")

    if not os.path.exists(pdf_path):
        return jsonify({"error": "File not found"}), 404

    # prevent two requests from processing the same upload concurrently.
    with processing_lock:
        if file_id in processing_files:
            return jsonify({"error": "File is already being processed"}), 409
        processing_files.add(file_id)

    try:
        return jsonify(process_pdf(pdf_path, file_id=file_id)), 200
    finally:
        with processing_lock:
            processing_files.discard(file_id)


@app.get("/qr/<path:filename>")
def serve_qr(filename):
    return send_from_directory(QR_DIR, filename)


@app.get("/api/progress/<file_id>")
def get_progress(file_id):
    # always return finite, bounded and internally consistent progress.
    state = progress.get(file_id, {})
    total = max(0, int(state.get("total_pages", 0) or 0))
    current = max(0, int(state.get("current_page", 0) or 0))
    current = min(current, total) if total else 0
    done = bool(state.get("done", False))

    if done:
        current = total

    return jsonify({
        "current_page": current,
        "total_pages": total,
        "done": done,
        "links_count": max(0, int(state.get("links_count", 0) or 0)),
        "qr_count": max(0, int(state.get("qr_count", 0) or 0))
    })


if __name__ == "__main__":
    app.run(debug=False, use_reloader=False)

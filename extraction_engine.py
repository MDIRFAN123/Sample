"""Priority-ordered, page-parallel PDF URL/QR extraction engine.

Extraction priority (per PDF_EXTRACTION_PRIORITY):
    1. Native PDF text        (master source of truth)
    2. Hyperlinks (PDF /URI actions)
    3. Visible form-field values
    4. QR codes                (kept in a wholly separate result set)
    5. OCR                     (fallback only, region-scoped, never on
                                 already-searchable text)

Performance notes (see accompanying analysis):
  * The source PDF is opened exactly once per worker thread and reused for
    every page that thread handles -- never once per page.
  * Each page is rendered to a bitmap exactly once; that same bitmap is
    reused for QR detection and for any OCR fallback.
  * Each page's text is parsed exactly once via a single shared
    ``TextPage`` -- "words", "text" and "blocks" all reuse it instead of
    re-walking the page three times.
  * OCR is skipped completely on searchable pages that contain neither
    embedded images nor vector artwork, since there is nothing there an
    OCR pass could find that native extraction hasn't already found.
"""

import ipaddress
import os
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from urllib.parse import urlsplit

import cv2
import fitz
import idna
import numpy as np
import tldextract


URL_CANDIDATE = re.compile(
    r'''(?ix)(?<![@\w])(?:
        (?:https?|ftp)://[^\s<>()\[\]{}"']+
        |www\.[^\s<>()\[\]{}"']+
        |localhost(?::\d{1,5})?(?:[/\?#][^\s<>()\[\]{}"']*)?
        |(?:\d{1,3}\.){3}\d{1,3}(?::\d{1,5})?(?:[/\?#][^\s<>()\[\]{}"']*)?
        |(?:[^\W_](?:[^\W_]|-){0,62}\.)+[^\W_](?:[^\W_]|-){0,62}
         (?::\d{1,5})?(?:[/\?#][^\s<>()\[\]{}"']*)?
    )''',
    re.UNICODE,
)
EMAIL = re.compile(r"(?i)^[a-z0-9._%+-]+@[a-z0-9.-]+\.[^\W_]{2,63}$", re.UNICODE)
TRAILING = ").,;:'\"!?]>"
BOX_TOLERANCE = 3.0


@dataclass
class UrlOccurrence:
    token: str
    page: int
    bbox: list | None
    source: str
    occurrence_id: str
    sources: list = field(default_factory=list)

    def as_json(self):
        return {
            "token": self.token,
            "page": self.page,
            "source": self.source,
            "pages": [self.page],
            "sources": list(self.sources),
            "bbox": self.bbox,
        }


class UrlValidator:
    """Candidate detection plus RFC-style parsing and PSL domain validation.

    Fully dynamic: no hardcoded TLD list. Domain validity is decided by the
    Public Suffix List (via tldextract), so new/future TLDs, IDNs, bare
    domains, IPv4 literals, localhost, ports, paths, queries and fragments
    are all supported without special-casing.
    """

    def __init__(self):
        self.psl = tldextract.TLDExtract(
            suffix_list_urls=(), include_psl_private_domains=True, cache_dir=None
        )

    def candidates(self, text):
        if not text:
            return []
        results = []
        for match in URL_CANDIDATE.finditer(text):
            token = match.group(0).rstrip(TRAILING)
            if token and self.valid(token):
                results.append((token, match.start(), match.start() + len(token)))
        return results

    def valid(self, token):
        if not token or any(char.isspace() for char in token if not token.lower().startswith("tel:")):
            return False
        if re.search(r"%(?![0-9A-Fa-f]{2})|[\x00-\x1f\x7f]", token):
            return False
        lower = token.lower()
        # This engine reports web URLs, not e-mail addresses or telephone
        # identifiers.  Keeping these out at the validator boundary also
        # prevents OCR text such as an e-mail domain becoming a URL.
        if lower.startswith(("mailto:", "tel:")) or EMAIL.fullmatch(token):
            return False

        has_scheme = bool(re.match(r"(?i)^(?:https?|ftp)://", token))
        parsed = urlsplit(token if has_scheme else f"//{token}")
        try:
            host = parsed.hostname or ""
            port = parsed.port
        except ValueError:
            return False
        if not host or (port is not None and not 1 <= port <= 65535):
            return False
        if host.lower() == "localhost":
            return True
        try:
            ipaddress.ip_address(host)
            return True
        except ValueError:
            pass
        try:
            ascii_host = idna.encode(host.rstrip(".")).decode("ascii")
        except idna.IDNAError:
            return False
        labels = ascii_host.split(".")
        if any(not re.fullmatch(r"(?i)[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?", label) for label in labels):
            return False
        raw_authority = re.split(r"[/\?#]", token, maxsplit=1)[0]
        original_labels = raw_authority.rsplit("@", 1)[-1].split(":", 1)[0].rstrip(".").split(".")
        # Upper-case dotted initials followed by a title-cased word are a
        # prose abbreviation pattern, not a hostname.  Case is intentionally
        # considered only for bare candidates; explicit URLs remain valid.
        if (not has_scheme and len(original_labels) >= 3
                and all(len(label) == 1 and label.isupper()
                        for label in original_labels[:-1])
                and original_labels[-1][:1].isupper()):
            return False
        extracted = self.psl(ascii_host)
        # A syntactically valid label is not enough: an Internet hostname
        # must have both a PSL suffix and a registrable label.  The only
        # deliberate exceptions are IP literals and localhost above.
        return bool(extracted.suffix and extracted.domain)


class OcrEngine:
    """OCR backend wrapper: prefers PaddleOCR, falls back to EasyOCR.

    Exposes a single ``readtext(image, detail=1, paragraph=False)`` call
    that always returns ``[(points, text, confidence), ...]`` regardless of
    which backend actually loaded, so the rest of the pipeline never has to
    know or care which OCR engine is in use.

    Both backends' underlying inference is not guaranteed safe to call
    concurrently from multiple threads on one shared instance, so calls are
    serialized with a lock. This only serializes the (already minority)
    OCR-fallback path -- it does not affect the page-parallel PDF parsing.
    """

    def __init__(self, lang="en", gpu=False):
        self._lock = threading.Lock()
        self.kind = None
        self.backend = None
        try:
            from paddleocr import PaddleOCR

            self.backend = PaddleOCR(
                lang=lang, use_angle_cls=False, show_log=False, use_gpu=gpu
            )
            self.kind = "paddle"
        except Exception:
            import easyocr

            self.backend = easyocr.Reader([lang], gpu=gpu)
            self.kind = "easyocr"

    def readtext(self, image, detail=1, paragraph=False):
        with self._lock:
            if self.kind == "easyocr":
                return self.backend.readtext(image, detail=detail, paragraph=paragraph)
            result = self.backend.ocr(image, cls=False)
            rows = (result or [None])[0] or []
            return [(points, text, confidence) for points, (text, confidence) in rows]


class PdfExtractionEngine:
    """Priority-ordered, page-parallel PDF URL and QR extraction engine."""

    def __init__(self, ocr_reader, qr_detector, max_workers=6):
        self.ocr_reader = ocr_reader
        self.qr_detector = qr_detector
        self.max_workers = max_workers
        self.validator = UrlValidator()

    @staticmethod
    def boxes_match(first, second):
        if not first or not second:
            return False
        ax0, ay0, ax1, ay1 = map(float, first)
        bx0, by0, bx1, by1 = map(float, second)
        if min(ax1, bx1) >= max(ax0, bx0) and min(ay1, by1) >= max(ay0, by0):
            return True
        gap_x = max(ax0 - bx1, bx0 - ax1, 0.0)
        gap_y = max(ay0 - by1, by0 - ay1, 0.0)
        return (gap_x * gap_x + gap_y * gap_y) ** 0.5 <= BOX_TOLERANCE

    def add_occurrence(self, occurrences, token, page, bbox, source):
        existing = next((
            item for item in occurrences
            if item.token == token and item.page == page and self.boxes_match(item.bbox, bbox)
        ), None)
        if existing is None:
            existing = UrlOccurrence(
                token=token, page=page, bbox=bbox, source=source,
                occurrence_id=f"p{page}-o{len(occurrences) + 1}", sources=[]
            )
            occurrences.append(existing)
        if source not in existing.sources:
            existing.sources.append(source)
        return existing

    def native_occurrences(self, words, page_num):
        """Build the master occurrence list from a page's already-extracted
        word list (see ``process_page``) -- never re-parses the page."""
        lines = {}
        for word in words:
            if len(word) >= 8 and str(word[4]).strip():
                lines.setdefault((word[5], word[6]), []).append(word)
        found = []
        for line_words in lines.values():
            line_words.sort(key=lambda word: word[7])
            chunks, spans, offset = [], [], 0
            for word in line_words:
                value = str(word[4])
                if chunks:
                    chunks.append(" ")
                    offset += 1
                start = offset
                chunks.append(value)
                offset += len(value)
                spans.append((start, offset, word))
            line = "".join(chunks)
            for token, start, end in self.validator.candidates(line):
                selected = [word for left, right, word in spans if left < end and right > start]
                if selected:
                    bbox = [
                        min(word[0] for word in selected), min(word[1] for word in selected),
                        max(word[2] for word in selected), max(word[3] for word in selected),
                    ]
                    self.add_occurrence(found, token, page_num, bbox, "TEXT")
        return found

    def attach_hyperlinks(self, page, page_num, occurrences):
        for link in page.get_links():
            uri = link.get("uri")
            bbox = list(link.get("from")) if link.get("from") else None
            for token, _, _ in self.validator.candidates(uri or ""):
                same = [item for item in occurrences if item.page == page_num and item.token == token]
                target = next((item for item in same if self.boxes_match(item.bbox, bbox)), None)
                if target is None:
                    target = self.add_occurrence(occurrences, token, page_num, bbox, "PDF_LINK")
                elif "PDF_LINK" not in target.sources:
                    target.sources.append("PDF_LINK")

    def attach_visible_fields(self, page, page_num, occurrences):
        for widget in page.widgets() or []:
            bbox = list(widget.rect) if widget.rect else None
            for token, _, _ in self.validator.candidates(str(widget.field_value or "")):
                target = next((
                    item for item in occurrences
                    if item.page == page_num and item.token == token
                    and self.boxes_match(item.bbox, bbox)
                ), None)
                if target is None:
                    target = self.add_occurrence(
                        occurrences, token, page_num, bbox, "PDF_FORM"
                    )
                elif "PDF_FORM" not in target.sources:
                    target.sources.append("PDF_FORM")

    @staticmethod
    def render_page(page, dpi=300):
        zoom = dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
        return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR), zoom

    @staticmethod
    def preprocess(image):
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)
        blurred = cv2.GaussianBlur(gray, (0, 0), 1.0)
        return cv2.addWeighted(gray, 1.6, blurred, -0.6, 0)

    def ocr_region(self, page_image, page_rect, region_rect, scale, retry=True):
        region = fitz.Rect(region_rect) & fitz.Rect(page_rect)
        if region.is_empty:
            return []
        x0, y0 = max(0, int(region.x0 * scale)), max(0, int(region.y0 * scale))
        x1, y1 = min(page_image.shape[1], int(np.ceil(region.x1 * scale))), min(page_image.shape[0], int(np.ceil(region.y1 * scale)))
        crop = page_image[y0:y1, x0:x1]
        if crop.size == 0:
            return []
        attempts = [(crop, 1.0)]
        if retry:
            attempts.append((cv2.resize(crop, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC), 1.5))
        detections = []
        for image, retry_scale in attempts:
            prepared = self.preprocess(image)
            results = self.ocr_reader.readtext(prepared, detail=1, paragraph=False)
            best = 0.0
            for points, text, confidence in results:
                for token, _, _ in self.validator.candidates(text):
                    xs = [point[0] for point in points]
                    ys = [point[1] for point in points]
                    divisor = scale * retry_scale
                    bbox = [
                        region.x0 + min(xs) / divisor, region.y0 + min(ys) / divisor,
                        region.x0 + max(xs) / divisor, region.y0 + max(ys) / divisor,
                    ]
                    detections.append((token, bbox))
                    best = max(best, float(confidence or 0.0))
            if detections and best >= 0.65:
                break
        return detections

    def ocr_regions(self, page, page_image, scale, searchable, native_rects, page_images, vector_drawings):
        """Region-scoped OCR fallback.

        On a fully searchable page with no embedded images and no vector
        artwork, there is nothing left for OCR to usefully find (native
        text already covers everything visible), so this returns
        immediately without touching the page again -- no margin scan, no
        image-region scan, no OCR calls at all.
        """
        page_rect = page.rect
        # The PDF-wide searchable decision is made before workers start.  A
        # searchable document must never pay the OCR cost or acquire the OCR
        # inference lock.  Image-only documents use the already cached render.
        if searchable:
            return []
        return self.ocr_region(page_image, page_rect, page_rect, scale, retry=True)

    def process_page(self, get_document, page_index, file_id):
        document = get_document()
        page = document[page_index]
        page_num = page_index + 1

        # One shared TextPage backs every text extraction below, so the page
        # content stream is parsed exactly once regardless of how many
        # different "views" (words / plain text / blocks) we need of it.
        textpage = page.get_textpage()
        words = page.get_text("words", sort=True, textpage=textpage)
        native_text = page.get_text("text", textpage=textpage) or ""
        blocks = page.get_text("blocks", textpage=textpage)
        native_rects = [list(block[:4]) for block in blocks if (block[4] or "").strip()]

        occurrences = self.native_occurrences(words, page_num)
        self.attach_hyperlinks(page, page_num, occurrences)
        self.attach_visible_fields(page, page_num, occurrences)

        # Classify per page so mixed/hybrid documents OCR only their actual
        # image pages while searchable pages always remain on the fast path.
        searchable = bool(native_text.strip())
        page_image, scale = self.render_page(page, 300)

        qr_items = self.qr_detector(
            page, file_id=file_id, page_num=page_num, img_bgr=page_image,
            thorough=True, render_zoom=scale
        )

        page_images = page.get_images(full=True)
        vector_drawings = []
        if searchable:
            # Only worth inspecting vector artwork (logo/watermark shapes) as
            # a possible OCR trigger when the page is otherwise searchable;
            # non-searchable pages already get a full-page OCR pass below.
            try:
                vector_drawings = page.get_drawings()
            except Exception:
                vector_drawings = []

        for token, bbox in self.ocr_regions(page, page_image, scale, searchable, native_rects, page_images, vector_drawings):
            if searchable and any(self.boxes_match(bbox, native_bbox) for native_bbox in native_rects):
                continue
            self.add_occurrence(occurrences, token, page_num, bbox, "OCR")
        return page_num, occurrences, qr_items, bool(not searchable)

    def extract(self, pdf_path, file_id, progress_state):
        with fitz.open(pdf_path) as document:
            total_pages = len(document)
        progress_state.update({
            "current_page": 0, "pages_done": 0, "total_pages": total_pages,
            "done": False, "links_count": 0, "qr_count": 0,
        })
        pages = {}
        lock = threading.Lock()

        # The PDF is opened exactly once per worker thread (never once per
        # page) and that same handle is reused for every page the thread
        # is assigned; all opened handles are closed once processing ends.
        thread_local = threading.local()
        opened_documents = []
        documents_lock = threading.Lock()

        def get_document():
            document = getattr(thread_local, "document", None)
            if document is None:
                document = fitz.open(pdf_path)
                thread_local.document = document
                with documents_lock:
                    opened_documents.append(document)
            return document

        workers = min(
            max(1, os.cpu_count() or 1),
            max(1, total_pages),
            max(1, self.max_workers),
        )
        try:
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {
                    executor.submit(
                        self.process_page, get_document, index, file_id
                    ): index
                    for index in range(total_pages)
                }
                completed = 0
                for future in as_completed(futures):
                    page_num, occurrences, qr_items, full_page_ocr = future.result()
                    pages[page_num] = (occurrences, qr_items, full_page_ocr)
                    with lock:
                        completed += 1
                        progress_state["pages_done"] = completed
                        progress_state["current_page"] = completed
                        progress_state["links_count"] = sum(len(value[0]) for value in pages.values())
                        progress_state["qr_count"] = sum(len(value[1]) for value in pages.values())
        finally:
            for document in opened_documents:
                document.close()

        links, qr_items = [], []
        pages_scanned = 0
        for page_num in sorted(pages):
            occurrences, page_qrs, used_full_page_ocr = pages[page_num]
            links.extend(item.as_json() for item in occurrences)
            qr_items.extend(page_qrs)
            pages_scanned += int(used_full_page_ocr)
        for item in qr_items:
            item["pages"] = [item["page"]]
        progress_state.update({
            "current_page": total_pages, "pages_done": total_pages, "done": True,
            "links_count": len(links), "qr_count": len(qr_items),
        })
        return {
            "total_pages": total_pages,
            "pages_scanned": pages_scanned,
            "links_count": len(links),
            "links": links,
            "qr_count": len(qr_items),
            "qr_items": qr_items,
        }

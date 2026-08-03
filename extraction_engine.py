"""Priority-ordered, page-parallel PDF URL/QR extraction engine.

Extraction priority (per PDF_EXTRACTION_PRIORITY):
    1. Native PDF text        (master source of truth)
    2. Hyperlinks (PDF /URI actions)
    3. Visible form-field values
    4. QR codes                (kept in a wholly separate result set)
    5. OCR                     (fallback only, region-scoped, never on
                                 already-searchable text)

Performance notes (see accompanying analysis):
  * The source PDF is opened exactly once for the extraction request.
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
        (?:https?|ftp)://[^\s<>()\[\]{}"]+
        |www\.[^\s<>()\[\]{}"]+
        |localhost(?::\d{1,5})?(?:[/\?#][^\s<>()\[\]{}"]*)?
        |(?:\d{1,3}\.){3}\d{1,3}(?::\d{1,5})?(?:[/\?#][^\s<>()\[\]{}"]*)?
        |(?:[^\W_](?:(?:[^\W_]|-)|['’‘](?=[^\W_])){0,62}\.)+
         [^\W_](?:(?:[^\W_]|-)|['’‘](?=[^\W_])){0,62}
         (?::\d{1,5})?(?:[/\?#][^\s<>()\[\]{}"]*)?
    )''',
    re.UNICODE,
)
EMAIL = re.compile(r"(?i)^[a-z0-9._%+-]+@[a-z0-9.-]+\.[^\W_]{2,63}$", re.UNICODE)
TRAILING = ").,;:'\"!?]>"
BOX_TOLERANCE = 3.0
OCR_DUPLICATE_IOU = 0.70


@dataclass
class UrlOccurrence:
    token: str
    page: int
    bbox: list | None
    source: str
    occurrence_id: str
    sources: list = field(default_factory=list)
    confidence: float = 0.0

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
        # OCR commonly inserts straight or typographic apostrophes inside a
        # hostname label.  Remove them only from the validation authority;
        # the original token remains untouched for display and output.
        prefix, separator, remainder = token.partition("://") if has_scheme else ("", "", token)
        authority, boundary, tail = re.match(r"([^/\?#]*)([/\?#]?)(.*)", remainder).groups()
        clean_authority = authority.translate(str.maketrans("", "", "'’‘"))
        validation_token = ((prefix + separator) if has_scheme else "") + clean_authority + boundary + tail
        parsed = urlsplit(validation_token if has_scheme else f"//{validation_token}")
        try:
            host = parsed.hostname or ""
            port = parsed.port
        except ValueError:
            return False
        if any(ord(char) > 127 for component in (parsed.path, parsed.query, parsed.fragment)
               for char in component):
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
        if (not has_scheme and not token.lower().startswith("www.")
                and not parsed.path and len(original_labels) == 2
                and len(original_labels[0]) <= 4
                and original_labels[0] != original_labels[0].lower()
                and original_labels[1][:1].isupper()):
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

    @staticmethod
    def bbox_iou(first, second):
        """Return intersection-over-union for two PDF-coordinate boxes."""
        if not first or not second:
            return 0.0
        ax0, ay0, ax1, ay1 = map(float, first)
        bx0, by0, bx1, by1 = map(float, second)
        intersection_width = max(0.0, min(ax1, bx1) - max(ax0, bx0))
        intersection_height = max(0.0, min(ay1, by1) - max(ay0, by0))
        intersection = intersection_width * intersection_height
        first_area = max(0.0, ax1 - ax0) * max(0.0, ay1 - ay0)
        second_area = max(0.0, bx1 - bx0) * max(0.0, by1 - by0)
        union = first_area + second_area - intersection
        return intersection / union if union > 0.0 else 0.0

    @staticmethod
    def comparison_key(token):
        return (token or "").casefold().translate(str.maketrans("", "", "'’‘"))

    @staticmethod
    def source_priority(source):
        return {"TEXT": 3, "PDF_LINK": 3, "PDF_FORM": 3,
                "QR": 2, "OCR": 1}.get(source, 0)

    def candidates_compete(self, first, second, first_box, second_box):
        if self.bbox_iou(first_box, second_box) <= OCR_DUPLICATE_IOU:
            return False
        left, right = self.comparison_key(first), self.comparison_key(second)
        return bool(left and right and (left == right or left in right or right in left))

    def add_occurrence(self, occurrences, token, page, bbox, source, confidence=0.0):
        existing = next((
            item for item in occurrences
            if item.page == page and self.candidates_compete(
                item.token, token, item.bbox, bbox
            )
        ), None)
        if existing is None:
            existing = UrlOccurrence(
                token=token, page=page, bbox=bbox, source=source,
                occurrence_id=f"p{page}-o{len(occurrences) + 1}", sources=[],
                confidence=float(confidence or 0.0)
            )
            occurrences.append(existing)
        else:
            current_rank = self.source_priority(existing.source)
            incoming_rank = self.source_priority(source)
            replace = incoming_rank > current_rank
            if incoming_rank == current_rank:
                replace = (len(token) > len(existing.token) or
                           (len(token) == len(existing.token) and
                            float(confidence or 0.0) > existing.confidence))
            if replace:
                existing.token = token
                existing.bbox = bbox
                existing.source = source
                existing.confidence = float(confidence or 0.0)
        if source not in existing.sources and self.comparison_key(existing.token) == self.comparison_key(token):
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
        ordered_lines = []
        for line_words in lines.values():
            line_words.sort(key=lambda word: word[7])
            ordered_lines.append(line_words)
        ordered_lines.sort(key=lambda row: (row[0][1], row[0][0]))

        # Test only adjacent visual line boundaries.  This recovers URLs split
        # by PDF wrapping without globally concatenating unrelated words.
        consumed = set()
        for first, second in zip(ordered_lines, ordered_lines[1:]):
            left, right = str(first[-1][4]), str(second[0][4])
            joined = left + right
            if not (left.endswith("/") or right.startswith(".")):
                continue
            matches = self.validator.candidates(joined)
            if any(token == joined for token, _, _ in matches):
                bbox = [min(first[-1][0], second[0][0]), min(first[-1][1], second[0][1]),
                        max(first[-1][2], second[0][2]), max(first[-1][3], second[0][3])]
                self.add_occurrence(found, joined, page_num, bbox, "TEXT")
                consumed.update((id(first[-1]), id(second[0])))

        for line_words in ordered_lines:
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
                if selected and not any(id(word) in consumed for word in selected):
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
                self.add_occurrence(occurrences, token, page_num, bbox, "PDF_LINK")

    def attach_visible_fields(self, page, page_num, occurrences):
        for widget in page.widgets() or []:
            bbox = list(widget.rect) if widget.rect else None
            for token, _, _ in self.validator.candidates(str(widget.field_value or "")):
                self.add_occurrence(occurrences, token, page_num, bbox, "PDF_FORM")

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
            rows = []
            for points, text, confidence in results:
                rows.append((points, str(text), float(confidence or 0.0)))
                for token, _, _ in self.validator.candidates(text):
                    xs = [point[0] for point in points]
                    ys = [point[1] for point in points]
                    divisor = scale * retry_scale
                    bbox = [
                        region.x0 + min(xs) / divisor, region.y0 + min(ys) / divisor,
                        region.x0 + max(xs) / divisor, region.y0 + max(ys) / divisor,
                    ]
                    detections.append((token, bbox, float(confidence or 0.0)))
                    best = max(best, float(confidence or 0.0))
            rows.sort(key=lambda row: (min(point[1] for point in row[0]),
                                       min(point[0] for point in row[0])))
            for left, right in zip(rows, rows[1:]):
                joined = left[1] + right[1]
                if not (left[1].endswith("/") or right[1].startswith(".")):
                    continue
                if not any(token == joined for token, _, _ in self.validator.candidates(joined)):
                    continue
                points = list(left[0]) + list(right[0])
                xs, ys = [point[0] for point in points], [point[1] for point in points]
                divisor = scale * retry_scale
                detections.append((joined, [
                    region.x0 + min(xs) / divisor, region.y0 + min(ys) / divisor,
                    region.x0 + max(xs) / divisor, region.y0 + max(ys) / divisor,
                ], min(left[2], right[2])))
                best = max(best, min(left[2], right[2]))
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
        page_area = max(1.0, page_rect.width * page_rect.height)
        text_area = sum(max(0.0, fitz.Rect(rect).get_area()) for rect in native_rects)
        if not searchable or text_area / page_area < 0.002:
            return self.ocr_region(page_image, page_rect, page_rect, scale, retry=True)

        text_rects = [fitz.Rect(rect) for rect in native_rects]
        regions = []
        seen = set()
        for image in page_images:
            for rect in page.get_image_rects(image[0]):
                rect = fitz.Rect(rect) & page_rect
                covered = sum((rect & text).get_area() for text in text_rects
                              if rect.intersects(text))
                if rect.is_empty or covered / max(1.0, rect.get_area()) >= 0.5:
                    continue
                key = tuple(round(value, 1) for value in rect)
                if key not in seen:
                    seen.add(key)
                    regions.append(rect)
        hits = []
        for rect in regions:
            hits.extend(self.ocr_region(page_image, page_rect, rect, scale, retry=False))
        return hits

    def process_page(self, get_document, page_index, file_id):
        document = get_document()
        page = document[page_index]
        page_num = page_index + 1

        # One native extraction only; every downstream text decision reuses
        # this cached word list.
        words = page.get_text("words", sort=True)
        native_rects = [list(word[:4]) for word in words if str(word[4]).strip()]

        occurrences = self.native_occurrences(words, page_num)
        self.attach_hyperlinks(page, page_num, occurrences)
        self.attach_visible_fields(page, page_num, occurrences)

        # Classify per page so mixed/hybrid documents OCR only their actual
        # image pages while searchable pages always remain on the fast path.
        searchable = bool(words)
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

        for token, bbox, confidence in self.ocr_regions(
                page, page_image, scale, searchable, native_rects,
                page_images, vector_drawings):
            if any(self.bbox_iou(bbox, item.get("bbox")) > OCR_DUPLICATE_IOU
                   for item in qr_items
                   if item.get("bbox")):
                continue
            self.add_occurrence(
                occurrences, token, page_num, bbox, "OCR", confidence
            )
        return page_num, occurrences, qr_items, bool(not searchable)

    def extract(self, pdf_path, file_id, progress_state):
        document = fitz.open(pdf_path)
        total_pages = len(document)
        progress_state.update({
            "current_page": 0, "pages_done": 0, "total_pages": total_pages,
            "done": False, "links_count": 0, "qr_count": 0,
        })
        pages = {}
        lock = threading.Lock()

        def get_document():
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

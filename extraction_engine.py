"""Page-parallel QR code extraction for PDF documents."""

import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import cv2
import fitz
import numpy as np


class PdfExtractionEngine:
    """Render PDF pages in parallel and delegate QR decoding to the detector."""

    def __init__(self, qr_detector, max_workers=6):
        self.qr_detector = qr_detector
        self.max_workers = max_workers

    @staticmethod
    def render_page(page, dpi=300):
        zoom = dpi / 72.0
        pixmap = page.get_pixmap(
            matrix=fitz.Matrix(zoom, zoom),
            alpha=False
        )

        rgb = np.frombuffer(
            pixmap.samples,
            dtype=np.uint8
        ).reshape(
            pixmap.height,
            pixmap.width,
            3
        )

        return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR), zoom

    def process_page(self, pdf_path, page_index, file_id):
        """
        Process one page using its own PyMuPDF Document instance.

        This avoids sharing a fitz.Document/Page object between worker
        threads, which can cause intermittent PDF processing failures.
        """
        document = None

        try:
            document = fitz.open(pdf_path)

            page = document[page_index]
            page_num = page_index + 1

            image, scale = self.render_page(page)

            qr_items = self.qr_detector(
                page,
                file_id=file_id,
                page_num=page_num,
                img_bgr=image,
                thorough=True,
                render_zoom=scale,
            )

            return page_num, qr_items, None

        except Exception as exc:
            # Do not allow one problematic page to terminate the
            # complete PDF scan.
            return page_index + 1, [], str(exc)

        finally:
            if document is not None:
                document.close()

    def extract(self, pdf_path, file_id, progress_state):
        started_at = time.monotonic()

        # Open the PDF only to determine page count.
        # This document is immediately closed and is NOT shared
        # with worker threads.
        document = None

        try:
            document = fitz.open(pdf_path)
            total_pages = len(document)
        finally:
            if document is not None:
                document.close()

        progress_state.update({
            "current_page": 0,
            "pages_processed": 0,
            "total_pages": total_pages,
            "qr_count": 0,
            "elapsed_seconds": 0.0,
            "done": False,
        })

        if total_pages == 0:
            elapsed_seconds = time.monotonic() - started_at

            progress_state.update({
                "current_page": 0,
                "pages_processed": 0,
                "qr_count": 0,
                "elapsed_seconds": elapsed_seconds,
                "done": True,
            })

            return {
                "total_pages": 0,
                "qr_count": 0,
                "qr_items": [],
                "scan_time_seconds": elapsed_seconds,
            }

        page_results = {}
        page_errors = {}

        progress_lock = threading.Lock()

        workers = min(
            max(1, os.cpu_count() or 1),
            max(1, total_pages),
            self.max_workers,
        )

        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(
                    self.process_page,
                    pdf_path,
                    index,
                    file_id,
                ): index
                for index in range(total_pages)
            }

            for future in as_completed(futures):
                page_index = futures[future]

                try:
                    page_num, qr_items, page_error = future.result()

                except Exception as exc:
                    # Extra safety: even if process_page itself somehow
                    # raises unexpectedly, continue processing other pages.
                    page_num = page_index + 1
                    qr_items = []
                    page_error = str(exc)

                page_results[page_num] = qr_items

                if page_error:
                    page_errors[page_num] = page_error

                with progress_lock:
                    processed = progress_state["pages_processed"] + 1

                    total_qr_count = sum(
                        len(items)
                        for items in page_results.values()
                    )

                    progress_state.update({
                        "pages_processed": processed,
                        "current_page": processed,
                        "qr_count": total_qr_count,
                        "elapsed_seconds": (
                            time.monotonic() - started_at
                        ),
                    })

        # Preserve original page ordering.
        qr_items = []

        for page_num in sorted(page_results):
            qr_items.extend(page_results[page_num])

        # Determine payload type after all pages have completed.
        for item in qr_items:
            item["payload_type"] = self.payload_type(
                item.get("payload", "")
            )

        elapsed_seconds = time.monotonic() - started_at

        # Mark the scan complete only after all worker tasks have finished.
        progress_state.update({
            "current_page": total_pages,
            "pages_processed": total_pages,
            "qr_count": len(qr_items),
            "elapsed_seconds": elapsed_seconds,
            "done": True,
        })

        # Keep the existing API response structure unchanged.
        return {
            "total_pages": total_pages,
            "qr_count": len(qr_items),
            "qr_items": qr_items,
            "scan_time_seconds": elapsed_seconds,
        }

    @staticmethod
    def payload_type(payload):
        value = str(payload or "").strip()

        if value.upper().startswith("WIFI:"):
            return "Wi-Fi"

        if value.upper().startswith("BEGIN:VCARD"):
            return "vCard"

        if value.startswith(("{", "[")):
            try:
                json.loads(value)
                return "JSON"
            except (TypeError, ValueError):
                pass

        if value.lower().startswith("mailto:") or re.fullmatch(
            r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+"
            r"@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
            value,
        ):
            return "Email"

        if value.lower().startswith("tel:"):
            return "Phone"

        return "Text"
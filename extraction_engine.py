"""Page-parallel QR code extraction for PDF documents."""

import json
import logging
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import nullcontext

import cv2
import fitz
import numpy as np


logger = logging.getLogger(__name__)


class PdfExtractionEngine:
    """Render each PDF page once and delegate QR decoding to the detector."""

    def __init__(self, qr_detector, max_workers=6):
        self.qr_detector = qr_detector
        self.max_workers = max_workers

    @staticmethod
    def render_page(page, dpi=300):
        zoom = dpi / 72.0
        pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        rgb = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(
            pixmap.height, pixmap.width, 3
        )
        return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR), zoom

    def process_page(self, pdf_path, page_index, total_pages, file_id):
        page_num = page_index + 1
        logger.info("Starting page %s/%s", page_num, total_pages)
        try:
            with fitz.open(pdf_path) as document:
                page = document[page_index]
                image, scale = self.render_page(page)
                return page_num, self.qr_detector(
                    page, file_id=file_id, page_num=page_num, img_bgr=image,
                    thorough=True, render_zoom=scale
                )
        except Exception:
            logger.exception("PDF page %s failed; continuing scan", page_num)
            raise
        finally:
            logger.info("Finished page %s/%s", page_num, total_pages)

    def extract(self, pdf_path, file_id, progress_state, progress_lock=None):
        started_at = time.monotonic()
        with fitz.open(pdf_path) as document:
            total_pages = len(document)
        with progress_lock if progress_lock is not None else nullcontext():
            progress_state.update({
                "current_page": 0,
                "pages_processed": 0,
                "total_pages": total_pages,
                "qr_count": 0,
                "elapsed_seconds": 0.0,
                "done": False,
                "status": "PROCESSING",
                "error": None,
            })
        page_results = {}
        workers = min(
            max(1, os.cpu_count() or 1), max(1, total_pages), self.max_workers
        )

        try:
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {
                    executor.submit(
                        self.process_page, pdf_path, index, total_pages, file_id
                    ): index
                    for index in range(total_pages)
                }
                for future in as_completed(futures):
                    page_num = futures[future] + 1
                    try:
                        page_num, qr_items = future.result()
                    except Exception:
                        qr_items = []
                    page_results[page_num] = qr_items or []
                    with progress_lock if progress_lock is not None else nullcontext():
                        processed = progress_state["pages_processed"] + 1
                        progress_state.update({
                            "pages_processed": processed,
                            "current_page": processed,
                            "qr_count": sum(
                                len(items) for items in page_results.values()
                            ),
                            "elapsed_seconds": time.monotonic() - started_at,
                        })
        except Exception:
            with progress_lock if progress_lock is not None else nullcontext():
                progress_state.update({
                    "elapsed_seconds": time.monotonic() - started_at,
                })
            logger.exception("PDF scan failed")
            raise

        qr_items = []
        for page_num in sorted(page_results):
            qr_items.extend(page_results[page_num])
        for item in qr_items:
            item["payload_type"] = self.payload_type(item.get("payload", ""))

        elapsed_seconds = time.monotonic() - started_at
        with progress_lock if progress_lock is not None else nullcontext():
            progress_state.update({
                "current_page": total_pages,
                "pages_processed": total_pages,
                "qr_count": len(qr_items),
                "elapsed_seconds": elapsed_seconds,
            })
        logger.info(
            "PDF scan completed: %s pages processed, %s QR codes found",
            total_pages,
            len(qr_items),
        )
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
            r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
            value,
        ):
            return "Email"
        if value.lower().startswith("tel:"):
            return "Phone"
        return "Text"

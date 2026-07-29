"""Download, inspect, redact, and format Excel workbooks for the UI."""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

from openpyxl import load_workbook

from config import Settings


class WorkbookError(RuntimeError):
    pass


# Browser-based downloading is handled by downloader.py
# WorkbookProcessor only processes local workbook files.

class WorkbookProcessor:
    MAX_SAMPLE_ROWS = 5
    MAX_SAMPLE_COLUMNS = 20
    MAX_PREVIEW_ROWS = 5000
    MAX_PREVIEW_COLUMNS = 200
    PII_PATTERNS = (
        ("ssn_hyphen", re.compile(r"(?<!\d)\d{3}-\d{2}-\d{4}(?!\d)")),
        ("ssn_dot", re.compile(r"(?<!\d)\d{3}\.\d{2}\.\d{4}(?!\d)")),
        ("ssn_space", re.compile(r"(?<!\d)\d{3}\s\d{2}\s\d{4}(?!\d)")),
        ("card_plain", re.compile(r"(?<!\d)\d{16}(?!\d)")),
        ("card_hyphen", re.compile(r"(?<!\d)\d{4}-\d{4}-\d{4}-\d{4}(?!\d)")),
        ("card_dot", re.compile(r"(?<!\d)\d{4}\.\d{4}\.\d{4}\.\d{4}(?!\d)")),
        ("card_space", re.compile(r"(?<!\d)\d{4}\s\d{4}\s\d{4}\s\d{4}(?!\d)")),
    )
    RULES_DIR = Path(__file__).resolve().parents[1] / "data" / "sheet_rules"
    CONTEXT_DIR = Path(__file__).resolve().parents[1] / "data" / "sheet_contexts"

    def build_payload(self, workbook_path: Path, workfront_content: dict | str | None = None) -> dict:
        if isinstance(workfront_content, str):
            workfront_content = {"OFFER_DESCRIPTION": workfront_content}
        workfront_content = {
            column: self._json_safe_value(value)
            for column, value in (workfront_content or {}).items()
        }
        workfront_text = "\n".join(
            f"{column}: {'N/A' if value is None or str(value).strip() == '' else value}"
            for column, value in workfront_content.items()
        )
        try:
            workbook = load_workbook(workbook_path, read_only=True, data_only=False)
        except Exception as error:
            raise WorkbookError("The downloaded file is not a readable Excel workbook.") from error

        pii_count = 0
        sheet_data: dict[str, dict] = {}
        try:
            for sheet in workbook.worksheets:
                rendered, found = self._sheet_sample(sheet)
                pii_count += found
                preview_rows = self._preview_rows(sheet)
                dimensions = f"{sheet.max_row} rows × {sheet.max_column} columns"
                sheet_data[sheet.title] = {
                    "offer_description": workfront_text,
                    "workfront_content": workfront_content,
                    "workbook_context": self._sheet_context(
                        workbook_path.name, sheet.title, workbook.sheetnames
                    ),
                    "sheet_prompt": self._sheet_instructions(sheet.title, dimensions, rendered),
                    "sheet_sample": rendered,
                    "dimensions": dimensions,
                    "generated_prompt": self._generated_prompt(workfront_text, workbook_path.name, sheet.title, dimensions, rendered),
                    "preview_rows": preview_rows,
                    "preview_truncated": (
                        sheet.max_row > self.MAX_PREVIEW_ROWS
                        or sheet.max_column > self.MAX_PREVIEW_COLUMNS
                    ),
                }
        finally:
            workbook.close()

        return {
            "workbook_name": workbook_path.name,
            "workbook_type": workbook_path.suffix.removeprefix(".").upper() or "XLSX",
            "sheet_count": len(workbook.sheetnames),
            "sheets": workbook.sheetnames,
            "sheet_data": sheet_data,
            "workfront_content": workfront_content,
            "has_pii": pii_count > 0,
            "pii_count": pii_count,
            "pii_status": f"{pii_count} detected for masking" if pii_count else "None detected"
        }

    @staticmethod
    def _json_safe_value(value):
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        if hasattr(value, "isoformat"):
            return value.isoformat()
        if hasattr(value, "read"):
            return str(value.read())
        return str(value)

    def _sheet_sample(self, sheet) -> tuple[str, int]:
        rows: list[str] = []
        pii_count = 0
        for row in sheet.iter_rows(max_row=self.MAX_SAMPLE_ROWS, max_col=self.MAX_SAMPLE_COLUMNS, values_only=True):
            values = []
            for value in row:
                rendered_value = "" if value is None else str(value)
                pii_count += self._count_pii(rendered_value)
                values.append(rendered_value)
            rows.append(" | ".join(values).rstrip(" |"))
        return "\n".join(rows).strip() or "(The sheet is empty.)", pii_count

    def _preview_rows(self, sheet) -> list[list[str]]:
        return [
            ["" if value is None else str(value) for value in row]
            for row in sheet.iter_rows(
                max_row=min(sheet.max_row, self.MAX_PREVIEW_ROWS),
                max_col=min(sheet.max_column, self.MAX_PREVIEW_COLUMNS),
                values_only=True,
            )
        ]

    def _count_pii(self, text: str) -> int:
        return sum(
            len(pattern.findall(text))
            for _, pattern in self.PII_PATTERNS
        )

    def _mask_pii(self, text: str) -> tuple[str, int]:
        matches = 0
        for _, pattern in self.PII_PATTERNS:
            text, count = pattern.subn("[REDACTED]", text)
            matches += count
        return text, matches

    @staticmethod
    def _workbook_context(name: str, sheets: list[str]) -> str:
        return f"Workbook: {name}\nAvailable sheets: {', '.join(sheets)}\nCredit Card and SSN values will be masked when the prompt is generated."

    @staticmethod
    def _sheet_prompt(name: str, dimensions: str, sample: str) -> str:
        return (
            f"Validate the '{name}' sheet ({dimensions}). "
            "Review structure, required fields, data types, formulas, and values."
        )

    @classmethod
    def _sheet_instructions(cls, name: str, dimensions: str, sample: str) -> str:
        configured = cls._load_sheet_text(cls.RULES_DIR, name)
        return configured or cls._sheet_prompt(name, dimensions, sample)

    @classmethod
    def _sheet_context(cls, workbook_name: str, sheet_name: str, sheets: list[str]) -> str:
        configured = cls._load_sheet_text(cls.CONTEXT_DIR, sheet_name)
        if configured:
            return configured
        return (
            f"Workbook: {workbook_name}\n"
            f"Active worksheet: {sheet_name}\n"
            f"Available sheets: {', '.join(sheets)}\n"
            "Credit Card and SSN values will be masked when the prompt is generated."
        )

    @staticmethod
    def _load_sheet_text(directory: Path, sheet_name: str) -> str:
        filename = re.sub(r"[^a-z0-9]+", "_", sheet_name.lower()).strip("_") + ".txt"
        path = directory / filename
        if not path.is_file():
            return ""
        return path.read_text(encoding="utf-8").strip()

    @classmethod
    def _generated_prompt(cls, offer: str, name: str, sheet: str, dimensions: str, sample: str) -> str:
        prompt = (
            "# Workbook Validation Prompt\n\n"
            f"## Offer Description\n{offer or '(Not available)'}\n\n"
            f"## Workbook Context\nWorkbook: {name}\nSheet: {sheet} ({dimensions})\n\n"
            f"## Masked Sheet Sample\n{sample}\n\n"
            "## Expected Output\nList validation issues with row/column references, severity, and suggested remediation."
        )
        return cls()._mask_pii(prompt)[0]

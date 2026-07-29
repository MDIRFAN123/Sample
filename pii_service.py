"""Detect the approved PII and sensitive-data patterns in a prompt."""

from __future__ import annotations

import re
from collections import Counter


class PIIScanner:
    """Pattern-based scanner that returns types/counts, never matched values."""

    # This is the approved detection set. Do not add categories without review.
    PII_PATTERNS = {
        "ssn": (
            re.compile(r"(?<!\d)\d{3}-\d{2}-\d{4}(?!\d)"),
        ),
        "credit_card": (
            re.compile(r"(?<!\d)\d{16}(?!\d)"),
            re.compile(r"(?<!\d)\d{4}-\d{4}-\d{4}-\d{4}(?!\d)"),
            re.compile(r"(?<!\d)\d{4} \d{4} \d{4} \d{4}(?!\d)"),
        ),
    }

    SCAN_ORDER = ("ssn", "credit_card")

    SEVERITY = {
        "ssn": "High",
        "credit_card": "High",
    }

    def scan(self, text: str) -> dict:
        findings = Counter()
        occupied: list[tuple[int, int]] = []
        details: list[dict] = []

        for label in self.SCAN_ORDER:
            for pattern in self.PII_PATTERNS[label]:
                for match in pattern.finditer(text):
                    if not self._overlaps(match.span(), occupied):
                        findings[label] += 1
                        occupied.append(match.span())
                        details.append(
                            {
                                "type": label,
                                "start": match.start(),
                                "end": match.end(),
                                "masked_preview": self._masked_preview(match.group()),
                                "severity": self.SEVERITY.get(label, "Low"),
                            }
                        )

        types = dict(sorted(findings.items()))
        details.sort(key=lambda item: item["start"])
        return {
            "has_pii": bool(types),
            "count": sum(types.values()),
            "types": types,
            "findings": details,
        }

    def scan_sections(self, sections: dict[str, str], worksheet: str) -> dict:
        all_findings: list[dict] = []
        type_counts = Counter()

        for section, text in sections.items():
            finding_worksheet = worksheet
            finding_section = section
            if "::" in section:
                finding_worksheet, finding_section = section.split("::", 1)
            result = self.scan(text)
            type_counts.update(result["types"])
            for finding in result["findings"]:
                finding.update(
                    {
                        "section": finding_section,
                        "worksheet": finding_worksheet,
                    }
                )
                all_findings.append(finding)

        types = dict(sorted(type_counts.items()))
        return {
            "has_pii": bool(all_findings),
            "count": len(all_findings),
            "types": types,
            "findings": all_findings,
        }

    @staticmethod
    def _masked_preview(value: str) -> str:
        visible = "".join(character for character in value if character.isalnum())[-4:]
        return f"{'*' * max(4, len(value) - len(visible))}{visible}"

    @staticmethod
    def _overlaps(span: tuple[int, int], occupied: list[tuple[int, int]]) -> bool:
        return any(span[0] < end and start < span[1] for start, end in occupied)

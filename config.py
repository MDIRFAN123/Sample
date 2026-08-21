"""Application configuration loaded from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent

TEAM_CONFIG = {
    "COPS": {
        "dsr_enabled": True,
        "sheet_selection": False,
        "complete_workbook_preview": True,
        "two_chatbot_flow": True,
        "rules_folder": "data/sheet_rules/COPS",
    },
    "MQA": {
        "dsr_enabled": False,
        "sheet_selection": True,
        "complete_workbook_preview": False,
        "two_chatbot_flow": False,
        "rules_folder": "data/sheet_rules/MQA",
    },
    "Channel Marketing": {
        "dsr_enabled": False,
        "sheet_selection": True,
        "complete_workbook_preview": False,
        "two_chatbot_flow": False,
        "rules_folder": "data/sheet_rules/Channel_Marketing",
    },
}


@dataclass(frozen=True)
class Settings:
    download_dir: Path
    monitored_download_dir: Path
    download_monitor_timeout: int
    max_workbook_bytes: int
    ora_user: str | None
    ora_password: str | None
    ora_dsn: str | None
    workbook_base_url: str
    data_share_base_url: str
    dsr_control_workbook: Path = BASE_DIR / "data" / "dsr" / "Past_Mistakes_20250820.xlsx"
    cops_gpt_url: str = "https://chatgpt.com/"

    @classmethod
    def from_environment(cls) -> "Settings":
        return cls(
            download_dir=Path(os.getenv("WORKBOOK_DOWNLOAD_DIR", BASE_DIR / "data" / "downloads")),
            monitored_download_dir=Path(
                os.getenv("MONITORED_DOWNLOAD_DIR", Path.home() / "Downloads")
            ),
            download_monitor_timeout=int(os.getenv("DOWNLOAD_MONITOR_TIMEOUT", "180")),
            max_workbook_bytes=int(os.getenv("MAX_WORKBOOK_BYTES", 25 * 1024 * 1024)),
            ora_user=os.getenv("ORA_USER"),
            ora_password=os.getenv("ORA_PASS"),
            ora_dsn=os.getenv("ORA_DSN"),
            workbook_base_url=os.getenv(
                "WORKBOOK_BASE_URL", "https://workfront.adobe.com/internal/download?versionID="
            ),
            data_share_base_url=os.getenv("DATA_SHARE_BASE_URL", "https://workfront.adobe.com/"),
            dsr_control_workbook=Path(os.getenv("DSR_CONTROL_WORKBOOK", BASE_DIR / "data" / "dsr" / "Past_Mistakes_20250820.xlsx")),
            cops_gpt_url=os.getenv("COPS_GPT_URL", "https://chatgpt.com/"),
        )

"""Application configuration loaded from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent


def env_bool(name: str, default: bool) -> bool:
    return os.getenv(name, str(default)).strip().lower() in {"1", "true", "yes", "on"}

TEAM_CONFIG = {
    "COPS": {
        "dsr_enabled": True,
        "sheet_selection": False,
        "complete_workbook_preview": True,
        "two_chatbot_flow": True,
        "rules_folder": "data/sheet_rules/COPS",
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
    cops_rovo_url: str = "https://irfannmd11.atlassian.net/"
    cops_gpt_url: str = "https://chatgpt.com/"
    team_app: str = ""
    enterprise_mode: bool = True
    auto_download_enabled: bool = True
    auto_upload_enabled: bool = False

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
            cops_rovo_url=os.getenv("COPS_ROVO_URL", "https://irfannmd11.atlassian.net/"),
            cops_gpt_url=os.getenv("COPS_GPT_URL", "https://chatgpt.com/"),
            team_app=os.getenv("TEAM_APP", "COPS").strip(),
            enterprise_mode=env_bool("ENTERPRISE_MODE", True),
            auto_download_enabled=env_bool("AUTO_DOWNLOAD_ENABLED", True),
            auto_upload_enabled=env_bool("AUTO_UPLOAD_ENABLED", False),
        )

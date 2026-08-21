"""Create and serve the two-file COPS validation package."""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path


def safe_token(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_") or "unknown"


def package_directory(downloads: Path, project: str, task: str) -> Path:
    return downloads / f"COPS_Validation_{safe_token(project)}_{safe_token(task)}"


def retain_workbook(source: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / source.name
    shutil.copy2(source, target)
    return target


def create_zip(target_dir: Path) -> Path:
    archive = target_dir / "COPS_Validation_Package.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        for path in target_dir.iterdir():
            if path.is_file() and path != archive:
                bundle.write(path, path.name)
    return archive

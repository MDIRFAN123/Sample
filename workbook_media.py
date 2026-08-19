"""Offline normalization for workbook image formats unsupported by openpyxl."""
from __future__ import annotations

from io import BytesIO
import logging
from pathlib import Path
import tempfile
import zipfile

from PIL import Image

LOGGER = logging.getLogger(__name__)


def normalize_wmf_images(workbook_path: Path) -> int:
    """Convert embedded WMF media to PNG in-place, retaining the original on failure."""
    if not zipfile.is_zipfile(workbook_path):
        return 0
    converted: dict[str, tuple[str, bytes]] = {}
    with zipfile.ZipFile(workbook_path, "r") as archive:
        for name in archive.namelist():
            if not name.lower().startswith("xl/media/") or not name.lower().endswith(".wmf"):
                continue
            try:
                with Image.open(BytesIO(archive.read(name))) as image:
                    output = BytesIO()
                    image.convert("RGBA").save(output, "PNG", optimize=True)
                converted[name] = (f"{name[:-4]}.png", output.getvalue())
            except Exception as error:
                LOGGER.warning("Unable to convert embedded WMF image %s in %s; preserving original: %s", name, workbook_path, error)
        if not converted:
            return 0
        with tempfile.NamedTemporaryFile(suffix=workbook_path.suffix, dir=workbook_path.parent, delete=False) as handle:
            temporary = Path(handle.name)
        try:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
                for item in archive.infolist():
                    if item.filename in converted:
                        new_name, content = converted[item.filename]
                        target.writestr(new_name, content)
                        continue
                    content = archive.read(item.filename)
                    if item.filename.endswith(".rels"):
                        text = content.decode("utf-8")
                        for old, (new, _) in converted.items():
                            text = text.replace(Path(old).name, Path(new).name)
                        content = text.encode("utf-8")
                    elif item.filename == "[Content_Types].xml":
                        text = content.decode("utf-8")
                        if 'Extension="png"' not in text:
                            text = text.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')
                        content = text.encode("utf-8")
                    target.writestr(item, content)
            temporary.replace(workbook_path)
        except Exception:
            temporary.unlink(missing_ok=True)
            raise
    LOGGER.info("Converted %d WMF image(s) to PNG in %s", len(converted), workbook_path)
    return len(converted)

"""
Browser-based Workfront downloader.

Opens the authenticated Workfront download URL in the default browser,
waits until the workbook has finished downloading,
and returns the downloaded local file path.
"""

from pathlib import Path
import os
import re
import time
import webbrowser


PARTIAL_EXTS = (
    ".crdownload",
    ".part",
    ".tmp"
)

EXCEL_EXTS = (
    ".xlsx",
    ".xlsm",
    ".xls"
)


def downloads_dir():
    """
    Returns the user's Downloads folder.
    """
    return os.path.join(
        os.path.expanduser("~"),
        "Downloads"
    )


def is_partial(filename: str):
    return filename.lower().endswith(PARTIAL_EXTS)


def is_excel(filename: str):
    return filename.lower().endswith(EXCEL_EXTS)


def split_name(filename: str):
    base, ext = os.path.splitext(filename)
    return base, ext.lower()


def candidate_regex(base: str, ext: str):
    """
    Matches

    Workbook.xlsx
    Workbook (1).xlsx
    Workbook (2).xlsx
    """

    return re.compile(
        rf"^{re.escape(base)}(?:\s*\(\d+\))?{re.escape(ext)}$",
        re.IGNORECASE
    )


def pick_best_match(download_path: str, expected_filename: str):

    base, ext = split_name(expected_filename)

    regex = candidate_regex(base, ext)

    newest = None

    for file in os.listdir(download_path):

        if is_partial(file):
            continue

        if not is_excel(file):
            continue

        if not regex.match(file):
            continue

        full = os.path.join(download_path, file)

        try:
            modified = os.path.getmtime(full)
        except OSError:
            continue

        if newest is None or modified > newest[0]:
            newest = (modified, full)

    return newest[1] if newest else None


def is_download_complete(filepath: str):
    """
    Ensures Chrome has completely finished writing the file.
    """

    partial_files = [
        filepath + ".crdownload",
        filepath + ".part",
        filepath + ".tmp"
    ]

    for partial in partial_files:
        if os.path.exists(partial):
            return False

    try:
        size1 = os.path.getsize(filepath)
        time.sleep(1)
        size2 = os.path.getsize(filepath)

        return size1 == size2

    except Exception:
        return False


def wait_for_download(
    expected_filename: str,
    timeout: int = 180
):
    """
    Wait until the workbook download has completed.
    """

    download_path = downloads_dir()

    start = time.time()

    while time.time() - start < timeout:

        path = pick_best_match(
            download_path,
            expected_filename
        )

        if path:

            if is_download_complete(path):
                return path

        time.sleep(1)

    raise FileNotFoundError(
        f"Workbook '{expected_filename}' was not downloaded "
        f"within {timeout} seconds."
    )


def download_to_local(
    download_url: str,
    expected_filename: str,
    timeout: int = 180
):
    """
    Opens the Workfront download URL.

    Browser downloads the workbook.

    Waits until the download completes.

    Returns the downloaded workbook path.
    """

    if not download_url:
        raise ValueError("Download URL is empty.")

    if not expected_filename:
        raise ValueError("Expected filename is empty.")

    print("=" * 80)
    print("Opening Workfront download URL")
    print(download_url)
    print("=" * 80)

    opened = webbrowser.open(download_url)

    if not opened:
        raise RuntimeError(
            "Unable to open the default browser."
        )

    print("Browser opened successfully.")

    print(
        f"Waiting for workbook '{expected_filename}'..."
    )

    workbook = wait_for_download(
        expected_filename,
        timeout
    )

    print("=" * 80)
    print("Download completed.")
    print(workbook)
    print("=" * 80)

    return Path(workbook)

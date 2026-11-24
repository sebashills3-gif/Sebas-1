"""After running the Sanctions Check in Genie:

1. Wait for the "Sanctions Check_dd.mm.yyyy_hh.mm.ss" folder that Genie
   creates in C:\\IM Genie Bot\\Sanctions Check.
2. Take the PDF that was generated inside it.
3. Move it to the evidence folder (where the screenshots are).
"""

import re
import time
from datetime import datetime, timedelta
from pathlib import Path

from prestaging import move_files

SANCTIONS_FOLDER = Path(r"C:\IM Genie Bot\Sanctions Check")

FOLDER_PATTERN = re.compile(
    r"^Sanctions Check_(\d{2}\.\d{2}\.\d{4}_\d{2}\.\d{2}\.\d{2})$",
    re.IGNORECASE,
)
DATE_FORMAT = "%d.%m.%Y_%H.%M.%S"


def find_folder(since, base=SANCTIONS_FOLDER):
    """Most recent Sanctions Check folder created since `since`."""
    base = Path(base)
    if not base.exists():
        return None

    candidates = []
    for file_path in base.iterdir():
        matched = FOLDER_PATTERN.match(file_path.name)
        if not file_path.is_dir() or not matched:
            continue
        try:
            stamp = datetime.strptime(matched.group(1), DATE_FORMAT)
        except ValueError:
            continue
        if stamp >= since - timedelta(seconds=2):
            candidates.append((stamp, file_path))

    return max(candidates)[1] if candidates else None


def wait_for_pdfs(since, base=SANCTIONS_FOLDER, timeout_s=900, pause_s=3):
    """Wait for the new folder and for its PDF to be complete.

    Returns the list of PDFs (normally one) in the folder.
    """
    start_time = time.time()
    previous_sizes = None

    while time.time() - start_time < timeout_s:
        folder = find_folder(since, base)

        if folder is not None:
            pdfs = sorted(
                p for p in folder.iterdir()
                if p.is_file() and p.suffix.lower() == ".pdf"
            )
            sizes = {p: p.stat().st_size for p in pdfs}

            # The PDF must exist and must have stopped growing.
            if pdfs and sizes == previous_sizes:
                return pdfs
            previous_sizes = sizes

        time.sleep(pause_s)

    raise TimeoutError(
        f"The folder with the Sanctions Check PDF did not appear in {base} "
        f"after {timeout_s} seconds."
    )


def process_sanctions(since, evidence_folder, base=SANCTIONS_FOLDER):
    """Move the Sanctions Check PDF to the evidence folder and return its
    new paths."""
    print()
    print(f"Waiting for the Sanctions Check folder in {base}...", flush=True)
    pdfs = wait_for_pdfs(since, base)

    if len(pdfs) > 1:
        print(f"Warning: the folder has {len(pdfs)} PDFs; all are moved.")

    moved = move_files(pdfs, evidence_folder)
    for file_path in moved:
        print(f"PDF moved: {file_path}")

    return moved

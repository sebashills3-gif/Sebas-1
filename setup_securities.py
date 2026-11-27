"""Support for options 4 and 5 of the menu:

* Choose the base Excel with a file picker.
* Create the folder for the run that starts at BLOOMBOT.
* Rebuild, from an existing evidence folder, everything needed for the
  securities set up in IMPACT (BLOOMBOT XLS and Sanctions Check PDF). In
  this mode every security is assumed to belong to HBSI, so the filtered
  Excel is not needed.
"""

import shutil
from datetime import datetime
from pathlib import Path

from prestaging import read_xls_matrix
from sanctions_pdf import verify_securities
from skeletons import build_securities, process_matrix

FILTERED_NAME = "TUR_requests_filtered.xlsx"

# Entity assumed for every security in the "setup only" mode.
ASSUMED_ENTITY = "HBSI"


# ============================================================
# FILE PICKER
# ============================================================

def latest_evidence_folder(base=None):
    """The most recent 'TUR Evidence_*' folder (by name), or `base`."""
    base = Path(base or Path.cwd())
    folders = sorted(
        (p for p in base.glob("TUR Evidence_*") if p.is_dir()),
        key=lambda p: p.name,
    )
    return folders[-1] if folders else base


def choose_file(title_text, initial_folder=None):
    """Open the Windows file picker. If it is not available, ask for the
    path on the console. Returns a Path, or None if cancelled."""
    initial = Path(initial_folder or latest_evidence_folder())
    file_types = [("Excel", "*.xlsx *.xls"), ("All files", "*.*")]

    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        file_path = filedialog.askopenfilename(
            title=title_text,
            initialdir=str(initial),
            filetypes=file_types,
            parent=root,
        )
        root.destroy()

    except Exception:
        file_path = input(
            f"{title_text}\nType the file path: "
        ).strip().strip('"')

    return Path(file_path) if file_path else None


# ============================================================
# OPTION 4: new folder for the run from BLOOMBOT
# ============================================================

def new_bloombot_folder(filtered_excel, base=None):
    """Create 'TUR Evidence_<date>_BLOOMBOT' and copy the filtered Excel
    there (so option 5 can use that folder later). Returns (folder,
    path of the copy)."""
    base = Path(base or Path.cwd())
    dt = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    folder = base / f"TUR Evidence_{dt}_BLOOMBOT"
    folder.mkdir(parents=True, exist_ok=True)

    copy_path = folder / FILTERED_NAME
    shutil.copy2(filtered_excel, copy_path)
    return folder, copy_path


# ============================================================
# OPTION 5: set up from an existing folder
# ============================================================

def _most_recent(paths):
    paths = list(paths)
    return max(paths, key=lambda p: p.stat().st_mtime) if paths else None


def locate_files(file_path):
    """From the chosen file, look in its folder for what is needed.

    You can choose the BLOOMBOT XLS (GIMP*.xls) or any other Excel of the
    folder (then the most recent GIMP*.xls of that folder is used). Returns
    {"folder", "xls", "pdfs"}; whatever is not found is None (or an empty
    list).
    """
    file_path = Path(file_path)
    folder = file_path.parent

    if file_path.suffix.lower() == ".xls":
        xls = file_path
    else:
        xls = _most_recent(
            p for p in folder.iterdir()
            if p.suffix.lower() == ".xls" and p.name.upper().startswith("GIMP")
        )

    pdf = _most_recent(
        p for p in folder.iterdir() if p.suffix.lower() == ".pdf"
    )

    return {
        "folder": folder,
        "xls": xls,
        "pdfs": [pdf] if pdf else [],
    }


def prepare_setup(file_path):
    """Rebuild the securities ready for IMPACT from an evidence folder: read
    the BLOOMBOT XLS, classify, assume every security is HBSI and check the
    Sanctions Check PDF.

    Returns (securities, folder). Raises RuntimeError if the BLOOMBOT XLS
    cannot be found.
    """
    files = locate_files(file_path)
    folder = files["folder"]

    if files["xls"] is None:
        raise RuntimeError(
            f"Could not find the BLOOMBOT XLS (GIMP*.xls) in {folder}."
        )

    print()
    print(f"Base folder: {folder}")
    print(f"BLOOMBOT XLS: {files['xls'].name}")
    print(
        "Sanctions PDF: "
        + (files["pdfs"][0].name if files["pdfs"] else "(none)")
    )
    print(f"Every security is assumed to belong to {ASSUMED_ENTITY}.")

    matrix = read_xls_matrix(files["xls"])
    entities = [{ASSUMED_ENTITY} for _ in build_securities(matrix)]

    securities = process_matrix(matrix, folder, entities)
    verify_securities(files["pdfs"], matrix, securities, folder)

    return securities, folder

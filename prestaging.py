"""After running the BLOOMBOT macro:

1. Wait for the GIMP*.xls and GIMP*.xps files generated in the DUMP folder
   and move them to the evidence folder.
2. Read the identifiers of each security from the XLS.
3. Paste them into the open "PreStaging File" workbook and save it.

Requires Windows with Excel and: pip install xlwings
"""

import re
import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path

DUMP_FOLDER = Path(r"C:\Users\45502358\Desktop\DUMP")

# XLS row (1 = first) where each identifier is. Column A holds the
# mnemonics and from column B onward there is one security per column.
XLS_ROWS = {
    "ISIN": 21,
    "CUSIP": 22,
    "COMMON": 23,
    "FIGI_ID": 91,
}
# The SEDOL can appear in any of these rows (there can be several).
SEDOL_ROWS = range(82, 87)
# Values Bloomberg returns when there is no data; they do not count as SEDOL.
NO_DATA = ("#N/A", "N.A.", "N/A", "#NAME", "#VALUE", "#REF")
LAST_XLS_ROW = 102

# PreStaging File column where each identifier is pasted.
PRESTAGING_COLUMNS = {
    "ISIN": "C",
    "SEDOL": "G",
    "CUSIP": "K",
    "COMMON": "O",
    "FIGI_ID": "S",
}
PRESTAGING_FIRST_ROW = 5
PRESTAGING_NAME = "PreStaging File"


# ============================================================
# 1. DUMP FILES
# ============================================================

def find_files(quantity, since, dump_folder=DUMP_FOLDER):
    """Look for GIMPddmmyyhhmmss_<quantity>.xls/.xps created since `since`.

    Returns {extension: path} of the most recent group that has an XLS.
    """
    pattern = re.compile(rf"^GIMP(\d{{12}})_{quantity}$", re.IGNORECASE)
    bundles = {}

    if not Path(dump_folder).exists():
        return {}

    for file_path in Path(dump_folder).iterdir():
        extension = file_path.suffix.lower()
        matched = pattern.match(file_path.stem)
        if extension not in (".xls", ".xps") or not matched:
            continue
        try:
            stamp = datetime.strptime(matched.group(1), "%d%m%y%H%M%S")
        except ValueError:
            continue
        if stamp >= since - timedelta(seconds=2):
            bundles.setdefault((stamp, file_path.stem), {})[extension] = file_path

    with_xls = [key for key, bundle in bundles.items() if ".xls" in bundle]
    if not with_xls:
        return {}
    return bundles[max(with_xls)]


def _sizes(files):
    return {
        extension: file_path.stat().st_size
        for extension, file_path in files.items()
    }


def wait_for_files(quantity, since, dump_folder=DUMP_FOLDER,
                   timeout_s=300, xps_grace_s=30, pause_s=2):
    """Wait for the XLS and the XPS to appear (and stop growing).

    If the XLS appears but the XPS does not arrive after `xps_grace_s`, it
    continues with the XLS only. Returns (xls, xps_or_None).
    """
    start_time = time.time()
    xls_seen_at = None
    previous_sizes = None

    while time.time() - start_time < timeout_s:
        files = find_files(quantity, since, dump_folder)

        if ".xls" in files:
            if xls_seen_at is None:
                xls_seen_at = time.time()

            sizes = _sizes(files)
            stable = sizes == previous_sizes
            previous_sizes = sizes

            if ".xps" in files and stable:
                return files[".xls"], files[".xps"]

            if ".xps" not in files and stable \
                    and time.time() - xls_seen_at > xps_grace_s:
                print("Warning: the XPS file did not appear.")
                return files[".xls"], None

        time.sleep(pause_s)

    raise TimeoutError(
        f"The GIMP*_{quantity} files did not appear in {dump_folder} "
        f"after {timeout_s} seconds."
    )


def move_files(paths, destination, attempts=10, pause_s=2):
    """Move the files to `destination`, retrying if they are still in use."""
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    new_paths = []

    for file_path in paths:
        if file_path is None:
            continue
        for attempt in range(1, attempts + 1):
            try:
                new_paths.append(
                    Path(shutil.move(str(file_path), str(destination / file_path.name)))
                )
                break
            except OSError:
                if attempt == attempts:
                    raise
                time.sleep(pause_s)

    return new_paths


# ============================================================
# 2. READING THE XLS
# ============================================================

def _to_text(value):
    """Cell value as clean text, or None if it is empty."""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    return text or None


def _is_sedol(value):
    return value is not None and not value.upper().startswith(NO_DATA)


def extract_securities(matrix):
    """From the cell matrix (row 1 = matrix[0]) build one dict per security.

    Each column from B onward is a security. "SEDOL" is a list with the
    distinct SEDOLs of rows 82 to 86, in the order they appear. Empty
    columns at the end are dropped; empty ones in the middle are kept so
    rows do not get misaligned.
    """
    columns = len(matrix[0]) if matrix else 0
    securities = []

    for col in range(1, columns):
        security = {
            key: _to_text(matrix[row - 1][col])
            for key, row in XLS_ROWS.items()
        }
        sedols = [_to_text(matrix[row - 1][col]) for row in SEDOL_ROWS]
        security["SEDOL"] = list(dict.fromkeys(v for v in sedols if _is_sedol(v)))
        securities.append(security)

    while securities and not any(securities[-1].values()):
        securities.pop()

    return securities


def build_rows(securities):
    """Turn the securities into PreStaging File rows.

    First row of each security: all the identifiers and the first SEDOL.
    For each additional SEDOL, a row with the same ISIN, that SEDOL and the
    rest empty.
    """
    rows = []

    for security in securities:
        sedols = security["SEDOL"]
        rows.append({**security, "SEDOL": sedols[0] if sedols else None})

        for sedol in sedols[1:]:
            rows.append({
                "ISIN": security["ISIN"],
                "SEDOL": sedol,
                "CUSIP": None,
                "COMMON": None,
                "FIGI_ID": None,
            })

    return rows


def read_xls_matrix(file_path):
    """Open the XLS in a hidden Excel instance and return its cells
    (rows 1 to LAST_XLS_ROW; row 1 is matrix[0])."""
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        workbook = app.books.open(str(file_path), read_only=True, update_links=False)
        sheet = workbook.sheets[0]
        last_col = max(sheet.used_range.last_cell.column, 2)
        matrix = sheet.range((1, 1), (LAST_XLS_ROW, last_col)) \
            .options(ndim=2).value
        workbook.close()
    finally:
        app.quit()

    print("Mnemonics read from the XLS:")
    for key, row in XLS_ROWS.items():
        print(f"  {key} (row {row}): {matrix[row - 1][0]}")
    for row in SEDOL_ROWS:
        print(f"  SEDOL (row {row}): {matrix[row - 1][0]}")

    return matrix


# ============================================================
# 3. PRESTAGING FILE
# ============================================================

def row_owners(securities):
    """For each row of build_rows, the index of the security it belongs to."""
    owners = []
    for i, security in enumerate(securities):
        owners.extend([i] * max(1, len(security["SEDOL"])))
    return owners


def _for_excel(value):
    """None clears the cell; codes made only of digits are stored as text
    (apostrophe) so leading zeros are not lost."""
    if value is None:
        return None
    return "'" + value if value.isdigit() else value


def paste_into_prestaging(securities, sheet=None):
    """Paste the identifiers starting at row 5 and save the workbook.

    A security with several SEDOLs takes several rows (see build_rows).

    It overwrites the destination cells (if the value is empty, the cell is
    left empty) and clears any content further down in those same columns
    left from previous runs. `sheet` is the sheet name; by default, the
    active sheet.
    """
    import xlwings as xw

    workbook = None
    for app in xw.apps:
        for candidate in app.books:
            if PRESTAGING_NAME.lower() in candidate.name.lower():
                workbook = candidate
                break
        if workbook:
            break
    if workbook is None:
        raise RuntimeError(f"The file '{PRESTAGING_NAME}' is not open.")

    destination = workbook.sheets[sheet] if sheet else workbook.sheets.active

    rows = build_rows(securities)

    for i, values in enumerate(rows):
        row = PRESTAGING_FIRST_ROW + i
        for key, column in PRESTAGING_COLUMNS.items():
            destination.range(f"{column}{row}").value = _for_excel(values[key])

    # Clear leftovers from previous runs (only the 5 identifier columns, and
    # only the content, not the formatting).
    first_leftover_row = PRESTAGING_FIRST_ROW + len(rows)
    last_row = destination.used_range.last_cell.row
    if last_row >= first_leftover_row:
        for column in PRESTAGING_COLUMNS.values():
            destination.range(
                f"{column}{first_leftover_row}:{column}{last_row}"
            ).clear_contents()

    workbook.save()
    print(
        f"{len(securities)} securities ({len(rows)} rows) pasted into "
        f"'{workbook.name}' (sheet '{destination.name}') and file saved."
    )


# ============================================================
# FULL FLOW
# ============================================================

def process_gimp(quantity, since, evidence_folder, dump_folder=DUMP_FOLDER,
                 sheet=None):
    """Move the GIMP files from DUMP to the evidence folder and fill the
    PreStaging File.

    Returns the XLS cells (to classify the securities), or None if the XLS
    had no securities.
    """
    print()
    print(f"Waiting for GIMP*_{quantity} files in {dump_folder}...", flush=True)
    xls, xps = wait_for_files(quantity, since, dump_folder)

    moved = move_files([xls, xps], evidence_folder)
    for file_path in moved:
        print(f"File moved: {file_path}")

    moved_xls = next(r for r in moved if r.suffix.lower() == ".xls")
    matrix = read_xls_matrix(moved_xls)
    securities = extract_securities(matrix)

    if len(securities) != quantity:
        print(
            f"Warning: {quantity} securities were sent to BLOOMBOT but the "
            f"XLS has {len(securities)} columns with data."
        )
    if not securities:
        print("The XLS has no securities; nothing is pasted.")
        return None

    paste_into_prestaging(securities, sheet=sheet)
    return matrix

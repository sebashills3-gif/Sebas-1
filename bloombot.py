"""Label the codes (CUSIP/ISIN) of the output Excel and send them to the
BLR00441 - GDS BLOOMBOT workbook ("Request File" sheet), then launch the
macro from the "Request Builder" sheet.

Requires Windows with Excel installed and: pip install pandas openpyxl xlwings
"""

import pandas as pd

WORKBOOK_NAME = "BLR00441 - GDS BLOOMBOT"
REQUEST_FILE_SHEET = "Request File"
REQUEST_BUILDER_SHEET = "Request Builder"
MACRO_SUBMIT = "creates_req"


def _text(value):
    """Return the value as clean text, or '' if it is empty."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def _column(df, name):
    """Find a column ignoring case and surrounding spaces."""
    for col in df.columns:
        if str(col).strip().lower() == name.lower():
            return col
    raise KeyError(f"Column '{name}' was not found in the Excel file.")


def label_code(code):
    """8 or 9 characters -> 'X |CUSIP|'; 12 characters -> 'X |ISIN|'.

    Any other length returns None (the code cannot be identified).
    """
    if len(code) in (8, 9):
        return f"{code} |CUSIP|"
    if len(code) == 12:
        return f"{code} |ISIN|"
    return None


def get_codes(excel_path):
    """Read the output Excel and return (labeled_codes, skipped, entities).

    It takes the Name column and, only when that cell is empty, the Cusip
    column. `entities` is aligned with `labeled_codes` (same position = same
    column of the BLOOMBOT XLS): each element is the set of entities (Entity
    column) that code comes from.
    """
    df = pd.read_excel(excel_path, dtype=str)
    col_name = _column(df, "Name")
    col_cusip = _column(df, "Cusip")
    col_entity = _column(df, "Entity")

    # First get the codes and remove duplicates (keeping the order); the
    # |CUSIP| or |ISIN| label is added afterwards.
    codes = []
    entities_by_code = {}
    for name, cusip, entity in zip(
        df[col_name], df[col_cusip], df[col_entity]
    ):
        code = _text(name) or _text(cusip)
        if code:
            codes.append(code)
            entities_by_code.setdefault(code.upper(), set()).add(
                _text(entity).upper()
            )

    unique = list(dict.fromkeys(c.upper() for c in codes))
    if len(unique) < len(codes):
        print(f"Duplicates removed: {len(codes) - len(unique)}")

    labeled_codes, skipped, entities = [], [], []
    for code in unique:
        labeled = label_code(code)
        if labeled is None:
            skipped.append(code)
        else:
            labeled_codes.append(labeled)
            entities.append(entities_by_code[code])
    return labeled_codes, skipped, entities


def _find_sheet(workbook, name, idx):
    """Sheet by name; if it does not exist, by position (1 = first)."""
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == name.lower():
            return sheet
    return workbook.sheets[idx - 1]


def _click_submit(sheet, macro=MACRO_SUBMIT):
    """Run the Submit button macro (creates_req by default).

    It is called through Application.Run qualifying the workbook, so it
    works even with other workbooks open. If that fails, the sheet's Submit
    button is looked up and the macro assigned to it is run.
    """
    workbook = sheet.book
    app = workbook.app.api
    errors = []

    if macro:
        for name in (f"'{workbook.name}'!{macro}", macro):
            try:
                app.Run(name)
                return
            except Exception as error:
                errors.append(f"{name}: {error}")

    for shape in sheet.api.Shapes:
        try:
            text = shape.TextFrame2.TextRange.Text
        except Exception:
            continue
        if "submit" in text.lower() and shape.OnAction:
            app.Run(shape.OnAction.strip("'"))
            return

    raise RuntimeError(
        "Could not run the Submit macro:\n" + "\n".join(errors)
    )


def paste_and_submit(codes, macro=MACRO_SUBMIT, clear_first=True):
    """Paste the codes into Request File!A and press Submit in Request Builder.

    The BLR00441 - GDS BLOOMBOT workbook must already be open in Excel.
    """
    import xlwings as xw

    workbook = None
    for app in xw.apps:
        for candidate in app.books:
            if candidate.name.lower().startswith(WORKBOOK_NAME.lower()):
                workbook = candidate
                break
        if workbook:
            break
    if workbook is None:
        raise RuntimeError(f"The file '{WORKBOOK_NAME}' is not open.")

    request_file = _find_sheet(workbook, REQUEST_FILE_SHEET, 3)
    request_builder = _find_sheet(workbook, REQUEST_BUILDER_SHEET, 1)

    if clear_first:
        request_file.range("A:A").clear_contents()

    cell_range = request_file.range(f"A1:A{len(codes)}")
    cell_range.number_format = "@"  # text, so Excel does not alter the codes
    cell_range.value = [[code] for code in codes]

    request_builder.activate()
    _click_submit(request_builder, macro)


def process_and_send(excel_path, macro=MACRO_SUBMIT):
    """Full flow starting from the Excel produced by the extraction.

    Returns (how many codes were sent, entities of each one in the XLS
    column order), or None if nothing was sent.
    """
    codes, skipped, entities = get_codes(excel_path)

    if skipped:
        print(f"Codes without 8, 9 or 12 characters (not sent): {skipped}")
    if not codes:
        print("There are no valid codes to send.")
        return

    paste_and_submit(codes, macro=macro)
    print(f"{len(codes)} codes sent to {WORKBOOK_NAME}.")
    return len(codes), entities

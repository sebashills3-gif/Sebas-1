from pathlib import Path
from datetime import datetime
import shutil
import time

import pandas as pd
from playwright.sync_api import sync_playwright

from bloombot import process_and_send
from prestaging import process_gimp
from genie import click_sanction_check
from sanctions import process_sanctions
from sanctions_pdf import verify_securities
from skeletons import process_matrix
from impact import process_securities_in_impact
from setup_securities import (
    FILTERED_NAME,
    choose_file,
    new_bloombot_folder,
    prepare_setup,
)


# ============================================================
# CONFIGURATION
# ============================================================

SYSTEM_URL = (
    "https://prd.impact.spsfi.broadridge.com/"
    "impact_hsbpro/controller"
)

# The folder name is kept in Spanish on purpose: it holds the saved Chrome
# session, and renaming it would force a new login (and MFA).
PROFILE = (Path.cwd() / "perfil_chrome_sistema").resolve()

COLUMNS = [
    "Entity",
    "External Trade ID",
    "Trader",
    "Cusip",
    "Name",
    "Date",
    "Status",
    "Security Id",
    "Amount_External",
    "Depository",
    "Amount_Normalized",
    "Cur",
]

TARGET_ERRORS = {
    "SECURITY NOT FOUND",
    "SECURITY DOES NOT EXIST IN MSD, BROADCAST SENT",
}

END_MESSAGE = "LAST QUALIFYING TRANSACTION ALREADY SHOWN"

MAX_PAGES = 100

# Detection and extraction retries
DETECTION_ATTEMPTS = 5
DETECTION_PAUSE_S = 2
EXTRACTION_ATTEMPTS = 3


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(text):
    return " ".join(text.split()).strip()


def visible_text(page):
    """
    Text of the page body and of all its iframes.
    """

    parts = []

    for frame in page.frames:
        try:
            parts.append(
                frame.locator("body").inner_text(timeout=30000)
            )
        except Exception:
            continue

    return "\n".join(parts)


def wait_page_ready(page, timeout_ms=15000):
    """
    Wait for the page to finish loading.
    It does not fail if some state is not reached.
    """

    for status in ("domcontentloaded", "load", "networkidle"):
        try:
            page.wait_for_load_state(status, timeout=timeout_ms)
        except Exception:
            pass


def get_target_table(page):
    """
    Look for the table that contains 'External Trade ID'
    in the page and in all its iframes.
    """

    for frame_number, frame in enumerate(page.frames):
        try:
            tables = frame.locator("table")
            total = tables.count()
        except Exception:
            continue

        for i in range(total):
            try:
                table = tables.nth(i)
                content = clean_text(table.inner_text())

                if "External Trade ID" in content:
                    return table, f"{frame_number}:{i}"

            except Exception:
                continue

    return None, None


def table_visible(
    page,
    attempts=DETECTION_ATTEMPTS,
    pause_s=DETECTION_PAUSE_S,
):
    """
    Check whether the main table is visible.

    It retries several times waiting for the page to load, because the
    table may take a while to be drawn.
    """

    for attempt in range(1, attempts + 1):
        wait_page_ready(page)

        table, _ = get_target_table(page)

        if table is not None:
            return True

        if attempt < attempts:
            print(
                f"Table not detected "
                f"(attempt {attempt}/{attempts}); retrying...",
                flush=True,
            )
            time.sleep(pause_s)

    return False


def extract_page(page, page_number):
    """
    Extract the rows of the table on the current page.

    Returns:
        records: list of dictionaries.
        signature: text that identifies the content of the page.
    """

    table, table_index = get_target_table(page)

    # While a page is loading the table may disappear for a moment;
    # retry before giving it up.
    if table is None and page_number > 0:
        for _ in range(3):
            time.sleep(2)
            table, table_index = get_target_table(page)

            if table is not None:
                break

    if table is None:
        if page_number > 0:
            print(
                f"[Page {page_number}] "
                "The target table was not found.",
                flush=True,
            )
        return [], ""

    if page_number > 0:
        print(
            f"[Page {page_number}] "
            f"Target table found at index {table_index}.",
            flush=True,
        )

    rows = table.locator("tr")
    records = []

    for i in range(rows.count()):
        row = rows.nth(i)
        cells = row.locator("th, td")
        values = []

        for j in range(cells.count()):
            try:
                value = clean_text(cells.nth(j).inner_text())
            except Exception:
                value = ""

            values.append(value)

        # Ignore headers and auxiliary rows
        if len(values) != len(COLUMNS):
            continue

        if values[0] == "Entity":
            continue

        record = dict(zip(COLUMNS, values))
        record["Page"] = page_number
        records.append(record)

    signature = "|".join(
        f"{record['External Trade ID']}::"
        f"{record['Status']}::"
        f"{record['Cusip']}"
        for record in records
    )

    if page_number > 0:
        print(
            f"[Page {page_number}] "
            f"Rows extracted: {len(records)}",
            flush=True,
        )

    return records, signature


def get_next_link(page):
    """
    Look for the link whose visible text is Next
    in the page and in all its iframes.
    """

    for frame in page.frames:
        try:
            links = frame.locator("a")
            total = links.count()
        except Exception:
            continue

        for i in range(total):
            link = links.nth(i)

            try:
                text = clean_text(link.inner_text())

                if text.upper() == "NEXT":
                    return link

            except Exception:
                continue

    return None


def save_screenshot(page, folder, file_name):
    """
    Save a full screenshot of the page.
    """

    file_path = folder / file_name

    page.screenshot(
        path=str(file_path),
        full_page=True,
    )

    print(
        f"Screenshot saved: {file_path.name}",
        flush=True,
    )

    return file_path


def wait_for_page_change(
    page,
    previous_signature,
    max_seconds=30,
):
    """
    Wait for the table to change after pressing Next.

    Returns:
        CHANGED    if the table changed.
        END        if the last-page message appeared.
        NO_CHANGE  if the time ran out.
    """

    start_time = time.time()

    while time.time() - start_time < max_seconds:
        time.sleep(0.5)

        try:
            current_text = visible_text(page).upper()

            if END_MESSAGE in current_text:
                return "END"

            _, current_signature = extract_page(page, -1)

            if current_signature and current_signature != previous_signature:
                return "CHANGED"

        except Exception:
            pass

    return "NO_CHANGE"


def deduplicate_records(df):
    """
    Remove duplicates when an External Trade ID exists.

    Rows without an External Trade ID are kept.
    """

    id_column = "External Trade ID"

    with_id = df[
        df[id_column]
        .fillna("")
        .astype(str)
        .str.strip()
        != ""
    ].copy()

    without_id = df[
        df[id_column]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    ].copy()

    with_id = with_id.drop_duplicates(
        subset=[id_column],
        keep="first",
    )

    return pd.concat(
        [with_id, without_id],
        ignore_index=True,
    )


# ============================================================
# FULL EXTRACTION OF ONE RUN
# ============================================================

def run_extraction(page):
    """
    Process all the pages from the current position
    and generate an independent evidence folder.

    Returns the path of the filtered Excel, or None if nothing was
    extracted (in that case no empty folder is left behind).
    """

    run_timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    evidence = (
        Path.cwd() / f"TUR Evidence_{run_timestamp}"
    )

    print()
    print("=" * 60)
    print("STARTING NEW EXTRACTION")
    print("=" * 60)
    print(f"Evidence folder: {evidence}")
    print()

    if not table_visible(page):
        print(
            "The table with 'External Trade ID' was not detected."
        )
        print(
            "Go manually to the first page of the table "
            "and try again."
        )
        return None

    evidence.mkdir(
        parents=True,
        exist_ok=True,
    )

    all_records = []
    seen_signatures = set()
    pages_processed = 0

    for page_number in range(1, MAX_PAGES + 1):
        print()
        print("-" * 60)
        print(f"Processing page {page_number}")
        print("-" * 60)

        # Time for the table to finish rendering
        page.wait_for_timeout(2000)

        records, current_signature = extract_page(
            page,
            page_number,
        )

        if not records:
            print(
                "No records found. "
                "This extraction stops."
            )
            break

        if current_signature in seen_signatures:
            print(
                "This page had already been processed. "
                "Stopping to avoid duplicates."
            )
            break

        seen_signatures.add(current_signature)
        all_records.extend(records)
        pages_processed = page_number

        # Screenshot of the current page
        save_screenshot(
            page,
            evidence,
            f"TUR_page_{page_number:03d}.png",
        )

        current_text = visible_text(page).upper()

        # If the notice is already showing, capture it and finish
        if END_MESSAGE in current_text:
            print(
                "The last-page notice was detected."
            )

            save_screenshot(
                page,
                evidence,
                f"TUR_last_page_notice_{page_number:03d}.png",
            )

            break

        next_link = get_next_link(page)

        if next_link is None:
            print(
                "The Next link was not found. "
                "This extraction stops."
            )
            break

        previous_signature = current_signature

        print("Pressing Next...", flush=True)

        try:
            next_link.click(timeout=30000)

        except Exception as error:
            print(
                f"Could not press Next: {error}"
            )
            break

        result = wait_for_page_change(
            page,
            previous_signature,
            max_seconds=30,
        )

        if result == "END":
            print(
                "Detected: "
                "Last Qualifying Transaction Already Shown"
            )

            save_screenshot(
                page,
                evidence,
                f"TUR_last_page_notice_{page_number:03d}.png",
            )

            break

        if result == "NO_CHANGE":
            print(
                "The table did not change after pressing Next. "
                "Stopping for safety."
            )
            break

        print(
            "The new page loaded correctly.",
            flush=True,
        )

    if not all_records:
        print(
            "No records were extracted in this run."
        )
        # Do not leave an empty evidence folder behind
        shutil.rmtree(evidence, ignore_errors=True)
        return None

    # ========================================================
    # GENERATE EXCEL
    # ========================================================

    df = pd.DataFrame(
        all_records,
        columns=COLUMNS + ["Page"],
    )

    df = deduplicate_records(df)

    df["Status_normalized"] = (
        df["Status"]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    df_filtered = df[
        df["Status_normalized"].isin(TARGET_ERRORS)
    ].copy()

    df.drop(
        columns=["Status_normalized"],
        inplace=True,
    )

    df_filtered.drop(
        columns=["Status_normalized"],
        inplace=True,
    )

    all_pages_file = (
        evidence / "TUR_requests_all_pages.xlsx"
    )

    filtered_file = (
        evidence / FILTERED_NAME
    )

    df.to_excel(
        all_pages_file,
        index=False,
    )

    df_filtered.to_excel(
        filtered_file,
        index=False,
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    print()
    print("=" * 60)
    print("EXTRACTION FINISHED")
    print("=" * 60)
    print(f"Pages processed: {pages_processed}")
    print(f"Total rows: {len(df)}")
    print(f"Filtered rows: {len(df_filtered)}")
    print()
    print("Evidence folder:")
    print(evidence)
    print()
    print("Full Excel:")
    print(all_pages_file)
    print()
    print("Filtered Excel:")
    print(filtered_file)
    print()

    print("Statuses found in the filtered Excel:")

    if len(df_filtered) > 0:
        print(
            df_filtered["Status"]
            .value_counts()
            .to_string()
        )
    else:
        print("None")

    return filtered_file


def run_extraction_with_retries(
    page,
    attempts=EXTRACTION_ATTEMPTS,
):
    """
    Run the extraction and, if it gets nothing, detect the table
    again and repeat it.

    It only retries when no row was extracted, that is, when the
    browser is still on the first page.
    Returns the path of the filtered Excel, or None.
    """

    for attempt in range(1, attempts + 1):
        filtered_file = run_extraction(page)

        if filtered_file is not None:
            return filtered_file

        if attempt < attempts:
            print()
            print(
                f"Extraction without results "
                f"(attempt {attempt}/{attempts}). "
                "Detecting the table again...",
                flush=True,
            )
            time.sleep(3)

    print()
    print(
        "No results were obtained after "
        f"{attempts} attempts."
    )
    return None


def bloombot_to_impact(
    filtered_file,
    page,
    evidence_folder=None,
    auto_save=False,
):
    """
    Send the codes of the filtered Excel to BLOOMBOT and, when the macro
    finishes, move the GIMP files to the evidence folder, fill the
    PreStaging File, classify the securities, press "Perform Sanctions
    Check" in Genie, move the resulting PDF to the evidence folder and
    prepare the creation of the securities in IMPACT.

    `evidence_folder` is where the GIMP files, Sanctions and reviews are
    saved (by default, the folder of the filtered Excel). With `auto_save`
    IMPACT saves each security without waiting for the manual review.
    An error here does not interrupt the program.
    """

    start_time = datetime.now()

    try:
        result = process_and_send(filtered_file)

    except Exception as error:
        print()
        print(f"Could not send to BLOOMBOT: {error}")
        return

    if not result:
        return

    quantity, entities = result

    if evidence_folder is None:
        evidence_folder = Path(filtered_file).parent

    try:
        matrix = process_gimp(
            quantity,
            start_time,
            evidence_folder,
        )

    except Exception as error:
        print()
        print(f"Could not complete the PreStaging File: {error}")
        return

    # Classification by skeleton and data cleaning (does not stop the flow)
    securities = None

    if matrix:
        try:
            securities = process_matrix(
                matrix,
                evidence_folder,
                entities,
            )

        except Exception as error:
            print()
            print(f"Could not classify the securities: {error}")

    sanctions_start = datetime.now()
    pdfs = []

    try:
        click_sanction_check()

        pdfs = process_sanctions(
            sanctions_start,
            evidence_folder,
        )

    except Exception as error:
        print()
        print(f"The Sanctions Check did not complete: {error}")

    if not securities:
        return

    # PDF review: only securities with PASS can be created. If there is no
    # PDF or it cannot be read, none is created.
    try:
        verify_securities(pdfs, matrix, securities, evidence_folder)

    except Exception as error:
        print()
        print(f"Could not review the Sanctions Check: {error}")
        for sec in securities:
            sec.block_reason = f"the Sanctions Check could not be reviewed: {error}"

    try:
        process_securities_in_impact(
            page,
            securities,
            auto_save=auto_save,
        )

    except Exception as error:
        print()
        print(f"Could not complete the work in IMPACT: {error}")


# ============================================================
# MENU OPTIONS
# ============================================================

def prepare_table(page):
    """Ask to leave Chrome on the first page of the table and check it.
    Returns True if it is possible to start."""

    print()
    print("Leave Chrome on the first page of the table.")
    print("When it is ready, come back here and press Enter.")
    input()

    while not table_visible(page):
        print()
        print(
            "The table was not detected in the tab "
            "controlled by Playwright."
        )

        check_again = input(
            "Do you want to check it again? (y/n): "
        ).strip().lower()

        if check_again != "y":
            print("The process was not started.")
            return False

    return True


def ask_auto_save():
    """True if IMPACT should save each security without a manual review."""

    answer = input(
        "Save each security in IMPACT automatically? (y/N): "
    ).strip().lower()

    return answer == "y"


def full_process(page, auto_save=False):
    """Extraction, BLOOMBOT, Sanctions and IMPACT."""

    filtered_file = run_extraction_with_retries(page)

    if filtered_file is not None:
        bloombot_to_impact(
            filtered_file,
            page,
            auto_save=auto_save,
        )


def extraction_only(page):
    """Only screenshots and filtered Excel files (before BLOOMBOT)."""

    run_extraction_with_retries(page)


def from_bloombot(page):
    """BLOOMBOT, Sanctions and IMPACT starting from an existing filtered
    Excel; everything is saved in a new folder."""

    excel_file = choose_file(
        f"Choose the filtered Excel ({FILTERED_NAME})"
    )

    if excel_file is None:
        print("No file was chosen.")
        return

    auto_save = ask_auto_save()
    folder, copy_path = new_bloombot_folder(excel_file)

    print()
    print(f"New folder: {folder}")

    bloombot_to_impact(
        copy_path,
        page,
        evidence_folder=folder,
        auto_save=auto_save,
    )


def setup_only(page):
    """Only the securities set up in IMPACT, starting from an existing
    evidence folder."""

    excel_file = choose_file(
        "Choose the base Excel (BLOOMBOT XLS GIMP*.xls or the filtered Excel)"
    )

    if excel_file is None:
        print("No file was chosen.")
        return

    auto_save = ask_auto_save()

    try:
        securities, _ = prepare_setup(excel_file)

    except Exception as error:
        print()
        print(f"Could not prepare the set up: {error}")
        return

    try:
        process_securities_in_impact(
            page,
            securities,
            auto_save=auto_save,
        )

    except Exception as error:
        print()
        print(f"Could not complete the work in IMPACT: {error}")


def show_menu():
    print()
    print("=" * 60)
    print("WHAT DO YOU WANT TO DO?")
    print("=" * 60)
    print("1 - Full process")
    print("2 - Full process without manual check (IMPACT saves by itself)")
    print("3 - Extraction only: screenshots and filtered Excel files")
    print("4 - From BLOOMBOT: Sanctions and creation (choose the filtered Excel)")
    print("5 - Securities set up in IMPACT only (choose the base Excel)")
    print("6 - Close the process")


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():
    with sync_playwright() as p:
        print("Opening Chrome and the system...", flush=True)

        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE),
            channel="chrome",
            headless=False,
            timeout=30000,
        )

        page = (
            context.pages[0]
            if context.pages
            else context.new_page()
        )

        page.goto(
            SYSTEM_URL,
            wait_until="commit",
            timeout=60000,
        )

        print()
        print("1. Enter your credentials manually in Chrome.")
        print("2. Complete MFA if the system asks for it.")
        print("3. Come back to this console and press Enter to see the menu.")
        input()

        while True:
            show_menu()

            option = input("Choose an option (1-6): ").strip()

            if option == "1":
                if prepare_table(page):
                    full_process(page)

            elif option == "2":
                if prepare_table(page):
                    full_process(page, auto_save=True)

            elif option == "3":
                if prepare_table(page):
                    extraction_only(page)

            elif option == "4":
                from_bloombot(page)

            elif option == "5":
                setup_only(page)

            elif option == "6":
                close_chrome = input(
                    "Close Chrome as well? (y/n): "
                ).strip().lower()

                if close_chrome == "y":
                    print()
                    print("Closing Chrome...")
                    context.close()

                else:
                    print()
                    print(
                        "The script will end and Chrome will be "
                        "left open."
                    )
                    # context.close() is not called, so Chrome is not
                    # closed explicitly.

                break

            else:
                print()
                print("Invalid option. Choose a number from 1 to 6.")


if __name__ == "__main__":
    main()

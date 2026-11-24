"""Check the Sanctions Check PDF before creating securities in IMPACT.

The PDF has the structure of the PreStaging File: one block per identifier
(By ISIN, By SEDOL, By CUSIP, By COMMON, By FIGI ID), each with the columns
No / identifier / Sanctions Check, plus a "Combine Status" block.
Row "No" k of the PDF corresponds to the k-th row pasted into the
PreStaging File (Excel row 5 is No 1).

A security only passes if ALL of its rows say PASS. At any doubt (unreadable
PDF, missing row, result other than PASS) the security is blocked: the rule
is not to touch anything that could not be verified.

Requires: pip install pdfplumber
"""

import re
import statistics
from pathlib import Path

from prestaging import build_rows, row_owners, extract_securities

IDENTIFIERS = ("ISIN", "SEDOL", "CUSIP", "COMMON", "FIGI_ID")

_TITLE = re.compile(
    r"^(ISIN|SEDOL|CUSIP|COMMON|FIGI(?:_ID)?)S?$", re.IGNORECASE
)
_GLUED_TITLE = re.compile(
    r"^BY(ISIN|SEDOL|CUSIP|COMMON|FIGI(?:_ID)?)S?$", re.IGNORECASE
)
_KEY = {
    "ISIN": "ISIN", "SEDOL": "SEDOL", "CUSIP": "CUSIP",
    "COMMON": "COMMON", "FIGI": "FIGI_ID", "FIGI_ID": "FIGI_ID",
}
_HEADERS = {"NO", "SANCTIONS", "CHECK", "STATUS", "COMMENTS", "ID"}
_NO_DATA = ("#N/A", "N.A.", "N/A")


class PdfSanctionsError(Exception):
    pass


# ============================================================
# READING THE PDF
# ============================================================

def _words_by_page(file_path):
    import pdfplumber

    pdf_pages = []
    with pdfplumber.open(str(file_path)) as pdf:
        for pdf_page in pdf.pages:
            words = pdf_page.extract_words(use_text_flow=False)
            pdf_pages.append(
                sorted(words, key=lambda w: (round(w["top"]), w["x0"]))
            )
    return pdf_pages


def _cx(w):
    return (w["x0"] + w["x1"]) / 2


def _cy(w):
    return (w["top"] + w["bottom"]) / 2


def _titles(words):
    """Block titles of the page: {key: {center, top, bottom}}."""
    titles = {}

    for i, w in enumerate(words):
        text = w["text"].strip()

        glued = _GLUED_TITLE.match(text.upper())
        if glued:
            key = _KEY[glued.group(1).upper()]
            titles[key] = {"center": _cx(w), "top": w["top"], "bottom": w["bottom"]}
            continue

        if text.upper() == "BY":
            for nxt in words[i + 1:]:
                if abs(nxt["top"] - w["top"]) <= 3 and nxt["x0"] >= w["x1"] - 1:
                    matched = _TITLE.match(nxt["text"].strip())
                    if matched:
                        key = _KEY[matched.group(1).upper()]
                        titles[key] = {
                            "center": (w["x0"] + nxt["x1"]) / 2,
                            "top": w["top"],
                            "bottom": w["bottom"],
                        }
                    break

        if text.upper() == "COMBINE":
            for nxt in words[i + 1:]:
                if abs(nxt["top"] - w["top"]) <= 3 and nxt["x0"] >= w["x1"] - 1:
                    if nxt["text"].strip().upper() == "STATUS":
                        titles["COMBINE"] = {
                            "center": (w["x0"] + nxt["x1"]) / 2,
                            "top": w["top"],
                            "bottom": w["bottom"],
                        }
                    break

    return titles


def _block_of(w, titles):
    return min(titles, key=lambda k: abs(_cx(w) - titles[k]["center"]))


def _geometry(words, titles):
    """Status columns (with their limits), the No column and the bottom edge
    of the header.

    Results sit against the left edge of their cell (they are not centered
    like the header), so the status column of each block is delimited from
    just before its "Sanctions" header up to where the next block starts.
    """
    if not any(k in titles for k in IDENTIFIERS):
        raise PdfSanctionsError("the 'By ...' titles were not found.")

    titles_bottom = max(t["bottom"] for t in titles.values())
    height = titles_bottom - min(t["top"] for t in titles.values())
    limit = titles_bottom + 3 * max(height, 8) + 6

    order = sorted(titles, key=lambda k: titles[k]["center"])
    sub = {k: [] for k in titles}
    header_bottom = titles_bottom

    for w in words:
        if titles_bottom - 0.5 <= w["top"] <= limit:
            sub[_block_of(w, titles)].append(w)
            text = w["text"].strip().upper()
            if text in _HEADERS or _TITLE.match(text):
                header_bottom = max(header_bottom, w["bottom"])

    left, no_col, no_left = {}, None, {}
    combine = None

    for block in order:
        block_words = sorted(sub[block], key=lambda w: w["x0"])

        if block == "COMBINE":
            status_words = [
                w for w in block_words
                if w["text"].strip().upper() == "STATUS"
                and w["top"] > titles_bottom - 0.5
            ]
            comment_words = [
                w for w in block_words
                if w["text"].strip().upper() == "COMMENTS"
            ]
            if status_words:
                sc = _cx(status_words[0])
                if comment_words:
                    cc = _cx(comment_words[0])
                    combine = (sc - (cc - sc) / 2, (sc + cc) / 2)
                else:
                    combine = (status_words[0]["x0"] - 15, status_words[0]["x1"] + 40)
            continue

        no = next((w for w in block_words if w["text"].strip().upper() == "NO"), None)
        sanction_words = [
            w for w in block_words
            if w["text"].strip().upper() in ("SANCTIONS", "CHECK")
        ]
        if no is not None:
            no_left[block] = no["x0"] - 12
            if no_col is None:
                no_col = (no["x0"] - 12, no["x1"] + 12)
        if sanction_words:
            left[block] = min(w["x0"] for w in sanction_words) - 20

    if no_col is None:
        raise PdfSanctionsError("the 'No' column was not found.")
    if not left:
        raise PdfSanctionsError("the 'Sanctions Check' columns were not found.")

    status_cols = {}
    for pos, block in enumerate(order):
        if block == "COMBINE" or block not in left:
            continue
        following = order[pos + 1] if pos + 1 < len(order) else None
        if following is None:
            right = left[block] + 120
        elif following == "COMBINE":
            right = (combine[0] if combine else left[block] + 120) - 1
        else:
            right = no_left.get(following, titles[following]["center"] - 60) - 1
        status_cols[block] = (left[block], right)

    if combine:
        status_cols["COMBINE"] = combine

    return {"status": status_cols, "no": no_col, "bottom": header_bottom}


def read_statuses(file_path):
    """Read the PDF and return (statuses, has_combine).

    statuses: {row_number: {"ISIN": "PASS", ..., "COMBINE": "PASS"}}
    """
    pdf_pages = _words_by_page(file_path)
    if not any(pdf_pages):
        raise PdfSanctionsError("the PDF has no text (is it an image?).")

    statuses = {}
    has_combine = False
    geometry = None

    for words in pdf_pages:
        titles = _titles(words)
        if titles:
            try:
                geometry = _geometry(words, titles)
                bottom = geometry["bottom"]
            except PdfSanctionsError:
                if geometry is None:
                    continue
                bottom = 0
        elif geometry is not None:
            bottom = 0  # continuation page, without a header
        else:
            continue

        x0, x1 = geometry["no"]
        anchors = [
            (int(w["text"]), _cy(w))
            for w in words
            if w["text"].strip().isdigit() and w["top"] >= bottom - 1
            and x0 <= _cx(w) <= x1
        ]
        if not anchors:
            continue

        pitches = [b[1] - a[1] for a, b in zip(anchors, anchors[1:]) if b[1] > a[1]]
        pitch = statistics.median(pitches) if pitches else 12
        tolerance = pitch * 0.45

        has_combine = has_combine or "COMBINE" in geometry["status"]

        for number, cy in anchors:
            if number in statuses:
                continue
            row = {}
            for block, (bx0, bx1) in geometry["status"].items():
                texts = [
                    w["text"] for w in words
                    if w["top"] >= bottom - 1 and bx0 <= _cx(w) <= bx1
                    and abs(_cy(w) - cy) <= tolerance
                ]
                row[block] = " ".join(texts).strip()
            statuses[number] = row

    if not statuses:
        raise PdfSanctionsError("no result rows could be read.")

    return statuses, has_combine


def _dump_diagnostic(file_path, folder):
    """Save the PDF words with their coordinates so they can be reviewed."""
    destination = Path(folder) / "sanctions_diagnostic.txt"
    lines = []
    try:
        for n, words in enumerate(_words_by_page(file_path), start=1):
            lines.append(f"--- page {n}: {len(words)} words")
            for w in words[:400]:
                lines.append(
                    f"{w['text']!r} x0={w['x0']:.0f} x1={w['x1']:.0f} "
                    f"top={w['top']:.0f}"
                )
    except Exception as error:
        lines.append(f"(could not read the PDF: {error})")
    destination.write_text("\n".join(lines), encoding="utf-8")
    return destination


# ============================================================
# VERIFICATION
# ============================================================

def _has_data(value):
    return value is not None and not str(value).strip().upper().startswith(_NO_DATA)


def verify_rows(statuses, has_combine, rows, owners, quantity):
    """Blocking reasons per security: {index: [reasons]} (empty list = passes)."""
    reasons = {i: [] for i in range(quantity)}

    for number, (row, owner) in enumerate(zip(rows, owners), start=1):
        status = statuses.get(number)
        if status is None:
            reasons[owner].append(f"row {number}: not in the PDF")
            continue

        for key in IDENTIFIERS:
            result = status.get(key, "").strip().upper()
            value = row.get(key)

            if _has_data(value):
                if result != "PASS":
                    reasons[owner].append(
                        f"row {number} {key} {value}: "
                        f"{result or 'no result'}"
                    )
            elif result and result != "PASS":
                reasons[owner].append(
                    f"row {number} {key} (empty in the PreStaging): {result}"
                )

        if has_combine:
            combined = status.get("COMBINE", "").strip().upper()
            if combined != "PASS":
                reasons[owner].append(
                    f"row {number} Combine Status: {combined or 'no result'}"
                )

    return reasons


def _alert(lines):
    bar = "!" * 70
    print()
    print(bar)
    for line in lines:
        print(f"!!  {line}")
    print(bar)


def verify_securities(pdfs, matrix, securities, folder):
    """Check the PDF and set `Security.block_reason` on those that do not pass.

    If the PDF cannot be read or interpreted, ALL securities are blocked.
    Returns how many securities ended up blocked.
    """
    def block_all(reason):
        for sec in securities:
            sec.block_reason = reason
        _alert([
            "THE SANCTIONS CHECK COULD NOT BE VERIFIED",
            reason,
            "NO SECURITY WILL BE CREATED IN IMPACT.",
        ])
        return len(securities)

    if not pdfs:
        return block_all("there is no Sanctions Check PDF to review")

    base = extract_securities(matrix)
    if len(base) != len(securities):
        return block_all(
            f"the PreStaging has {len(base)} securities and the "
            f"classification has {len(securities)}; the PDF cannot be matched"
        )

    rows = build_rows(base)
    owners = row_owners(base)
    reasons = {i: [] for i in range(len(securities))}

    for pdf in pdfs:
        try:
            statuses, has_combine = read_statuses(pdf)
        except Exception as error:
            diag_path = _dump_diagnostic(pdf, folder)
            return block_all(
                f"could not interpret '{Path(pdf).name}': {error} "
                f"(diagnostic in {diag_path})"
            )

        if len(pdfs) > 1:
            print(f"Reviewing {Path(pdf).name}...")
        for i, pdf_reasons in verify_rows(
            statuses, has_combine, rows, owners, len(securities)
        ).items():
            reasons[i].extend(pdf_reasons)

    blocked = 0
    for i, sec in enumerate(securities):
        if reasons[i]:
            sec.block_reason = "Sanctions Check NO PASS: " + "; ".join(reasons[i])
            blocked += 1

    if blocked:
        lines = [f"SANCTIONS CHECK: {blocked} SECURITY(IES) WITHOUT PASS - NOT TOUCHED"]
        for i, sec in enumerate(securities):
            if sec.block_reason and reasons[i]:
                lines.append(f"- {sec.rows.get(21) or 'column ' + str(sec.column)}")
                lines.extend(f"    {m}" for m in reasons[i])
        _alert(lines)
    else:
        print()
        print(f"Sanctions Check: all {len(securities)} securities say PASS.")

    return blocked

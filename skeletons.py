"""Classify each security of the BLOOMBOT XLS into a skeleton (template) and
clean its data before creating it in IMPACT.

Structure (to add new skeletons only the SKELETONS section is touched):

* GENERAL TRANSFORMATIONS: apply to all skeletons.
* Each skeleton defines its CRITERIA (XLS row -> expected value) and its
  SPECIFIC TRANSFORMATIONS.
* The final cell formatting (dates without slashes and clean decimals) is
  always applied last, so the earlier transformations work with the
  original values.

Rows are the ones of the BLOOMBOT XLS (1 = first); column A holds the
mnemonic and from column B onward there is one security per column.
"""

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

# ============================================================
# XLS ROWS
# ============================================================

IS_UNIT_TRADED = 4
CPN_TYP = 6
CRNCY = 7
MARKET_SECTOR_DES = 12
SECURITY_DES = 15
SECURITY_TYP = 8
COLLAT_TYP = 9
TYPE_OF_BOND = 10
CALC_TYP_DES = 13
MATURITY_ROW = 19
DAY_COUNT_ROW = 20
ID_ISIN = 21
ID_CUSIP = 22
FIRST_CPN_DT = 27
IS_PERPETUAL = 52
SINKABLE = 56

# Expected mnemonic in column A for each row used (to warn if the XLS layout
# changes).
EXPECTED_MNEMONICS = {
    IS_UNIT_TRADED: "IS_UNIT_TRADED",
    CPN_TYP: "CPN_TYP",
    CRNCY: "CRNCY",
    MARKET_SECTOR_DES: "MARKET_SECTOR_DES",
    SECURITY_TYP: "SECURITY_TYP",
    COLLAT_TYP: "COLLAT_TYP",
    TYPE_OF_BOND: "TYPE_OF_BOND",
    CALC_TYP_DES: "CALC_TYP_DES",
    SECURITY_DES: "SECURITY_DES",
    MATURITY_ROW: "MATURITY",
    DAY_COUNT_ROW: "DAY_CNT_DES",
    ID_ISIN: "ID_ISIN",
    ID_CUSIP: "ID_CUSIP",
    FIRST_CPN_DT: "FIRST_CPN_DT",
    IS_PERPETUAL: "IS_PERPETUAL",
    SINKABLE: "SINKABLE",
}

# Identifier rows: decimals are never "cleaned" there (a numeric CUSIP or
# SEDOL must not turn into 12345.0).
IDENTIFIER_ROWS = {21, 22, 23, 82, 83, 84, 85, 86, 91}

PERPETUAL_YEARS = 30


# ============================================================
# MODEL
# ============================================================

@dataclass
class Security:
    column: int                     # 2 = column B
    rows: dict                      # XLS row -> value
    mnemonics: dict                 # XLS row -> mnemonic (column A)
    skeleton: str = None
    entities: set = field(default_factory=set)   # Entity column of the TUR
    block_reason: str = None        # reason why it must NOT be created
    derived: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)

    def text(self, row):
        return _text(self.rows.get(row))


# ============================================================
# HELPERS
# ============================================================

def _text(value):
    """Value as clean text ('' if empty or a date/other object)."""
    if value is None or isinstance(value, (date, datetime)):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def _column_letter(column):
    letters = ""
    while column:
        column, rest = divmod(column - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


_STANDALONE_DATE = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$")


def to_date(value):
    """date if the value IS a standalone date (date object or 'MM/DD/YYYY');
    None otherwise (for example 'BXT 12/05/2030' does not count)."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        matched = _STANDALONE_DATE.match(value)
        if matched:
            month_value, day_value, year_value = (int(g) for g in matched.groups())
            try:
                return date(year_value, month_value, day_value)
            except ValueError:
                return None
    return None


def format_date_no_slashes(dt):
    """12/15/2030 -> 12152030."""
    return dt.strftime("%m%d%Y")


def format_decimal(value):
    """Remove extra zeros: 5.87000 -> 5.87, 5 -> 5.0, .78 -> 0.78."""
    decimal = Decimal(str(value).strip())
    text = format(decimal.normalize(), "f")
    if "." not in text:
        text += ".0"
    return text


_NUMBER_WITH_DOT = re.compile(r"^[+-]?(\d+\.\d*|\.\d+)$")


def _is_number_to_clean(value):
    """True if the value is a number (numeric cell or text with a decimal
    point). Texts made only of digits are left alone: they may be codes."""
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return isinstance(value, str) and bool(_NUMBER_WITH_DOT.match(value.strip()))


def add_years(dt, years):
    try:
        return dt.replace(year=dt.year + years)
    except ValueError:  # February 29
        return dt.replace(year=dt.year + years, day=28)


# ============================================================
# GENERAL TRANSFORMATIONS (all skeletons)
# ============================================================

def derive_cusip_from_isin(sec):
    """CRNCY = USD and an ISIN that starts with US: the CUSIP is the ISIN
    without its first 2 characters and without the last one
    (US1234567890 -> 123456789)."""
    isin = sec.text(ID_ISIN)

    if sec.text(CRNCY).upper() == "USD" and isin.upper().startswith("US"):
        if len(isin) != 12:
            sec.warnings.append(
                f"ISIN '{isin}' does not have 12 characters; the CUSIP was "
                "not derived."
            )
            return
        sec.derived["CUSIP_ORIGINAL"] = sec.text(ID_CUSIP)
        sec.rows[ID_CUSIP] = isin[2:-1]


_DAY_COUNT_CODES = {
    "ACT/ACT": ("A", "A"),
    "ACT/360": ("A", "360"),
    "ACT/365": ("A", "365"),
    "30/360": ("30", "360"),
    "30/365": ("30", "365"),
    "ISMA-30/360": ("3E", "360"),
    "ISMA-30/365": ("3E", "365"),
}

# ISMA-30/36x goes first and the lookarounds prevent 30/360 from being taken
# from inside ISMA-30/360.
_DAY_COUNT_RE = re.compile(
    r"(?<![\w/-])(ISMA-30/36[05]|30/36[05]|ACT/ACT|ACT/36[05])(?![\w/])",
    re.IGNORECASE,
)


def split_day_count(sec):
    """Row 20 (alone or inside other text, e.g. 'BXT 4V ACT/ACT 57') ->
    DAY COUNT and YEAR BASIS."""
    text = sec.text(DAY_COUNT_ROW)
    if not text:
        return

    codes = {m.upper() for m in _DAY_COUNT_RE.findall(text)}

    if not codes:
        sec.warnings.append(
            f"Row {DAY_COUNT_ROW}: the day count was not recognized in '{text}'."
        )
        return
    if len(codes) > 1:
        sec.warnings.append(
            f"Row {DAY_COUNT_ROW}: several day counts in '{text}' "
            f"({sorted(codes)}); the first one is used."
        )

    first_one = _DAY_COUNT_RE.search(text).group(1).upper()
    sec.derived["DAY COUNT"], sec.derived["YEAR BASIS"] = \
        _DAY_COUNT_CODES[first_one]


GENERAL_TRANSFORMATIONS = [
    derive_cusip_from_isin,
    split_day_count,
]


def format_cells(sec):
    """Always last: standalone dates -> MMDDYYYY and decimals without extra
    zeros."""
    for row, value in list(sec.rows.items()):
        dt = to_date(value)

        if dt:
            sec.rows[row] = format_date_no_slashes(dt)
        elif row not in IDENTIFIER_ROWS and _is_number_to_clean(value):
            try:
                sec.rows[row] = format_decimal(value)
            except InvalidOperation:
                pass
        elif value is not None:
            sec.rows[row] = _text(value)


# ============================================================
# SPECIFIC TRANSFORMATIONS OF 5/80 (reused by EUR 3/01 and 8/04)
# ============================================================

_MIXED_FRACTION = re.compile(r"(?<![\d/.])(\d+)\s+(\d+)/(\d+)(?![\d/])")
_PLAIN_FRACTION = re.compile(r"(?<![\d/.])(\d+)/(\d+)(?![\d/])")


def _fraction_to_decimal(whole, numerator, denominator):
    return Decimal(whole) + Decimal(numerator) / Decimal(denominator)


def coupon_fraction_to_decimal(sec):
    """SECURITY_DES 'ISSUER 5 1/2 12/15/2030' -> 'ISSUER 5.5 12/15/2030'.
    If the coupon is already a decimal it is left alone."""
    description = sec.text(SECURITY_DES)

    mixed = _MIXED_FRACTION.search(description)
    plain = None if mixed else _PLAIN_FRACTION.search(description)
    matched = mixed or plain
    if not matched:
        return

    parts = matched.groups()
    whole, numerator, denominator = (
        parts if mixed else ("0",) + parts
    )
    if int(denominator) == 0:
        return

    coupon = format_decimal(
        _fraction_to_decimal(whole, numerator, denominator)
    )
    sec.rows[SECURITY_DES] = (
        description[:matched.start()] + coupon + description[matched.end():]
    )


def perpetual_to_maturity(sec):
    """IS_PERPETUAL = Y: MATURITY (row 19) = FIRST_CPN_DT + 30 years, and if
    SECURITY_DES ends in PERP it is replaced by the date as MM/DD/YYYY."""
    if sec.text(IS_PERPETUAL).upper() != "Y":
        return

    first_coupon = to_date(sec.rows.get(FIRST_CPN_DT))
    if first_coupon is None:
        sec.warnings.append(
            "IS_PERPETUAL = Y but FIRST_CPN_DT is not a valid date; "
            "MATURITY was not calculated."
        )
        return

    maturity_date = add_years(first_coupon, PERPETUAL_YEARS)
    sec.rows[MATURITY_ROW] = maturity_date  # formatted at the end

    description = sec.text(SECURITY_DES)
    if re.search(r"(?i)(?<!\S)PERP$", description):
        sec.rows[SECURITY_DES] = re.sub(
            r"(?i)(?<!\S)PERP$", maturity_date.strftime("%m/%d/%Y"), description
        )


@dataclass(frozen=True)
class Contains:
    """Criteria value: the cell text must contain `text` somewhere."""
    text: str


def _meets(sec, criteria):
    """True if the security meets every (row -> expected) of `criteria`.
    An expected value is an exact text or a Contains(...)."""
    for row, expected in criteria.items():
        actual = sec.text(row).upper()
        if isinstance(expected, Contains):
            if expected.text.upper() not in actual:
                return False
        elif actual != expected.upper():
            return False
    return True


@dataclass
class Skeleton:
    name: str
    criteria: dict                  # XLS row -> expected value
    transformations: list
    # row -> value (or list of values) it must NOT have
    excludes: dict = field(default_factory=dict)
    # Forbidden combinations: each dict (row -> value) excludes the security
    # only if it meets ALL of its conditions at once.
    excluded_combinations: list = field(default_factory=list)
    # Alternative ways to meet the criteria: each dict is a full set of
    # criteria; meeting `criteria` OR any variant identifies the skeleton.
    # The exclusions apply to all of them.
    variants: list = field(default_factory=list)

    def matches(self, sec):
        meets = _meets(sec, self.criteria) or any(
            _meets(sec, variant) for variant in self.variants
        )
        not_excluded = all(
            sec.text(row).upper() not in {
                v.upper() for v in (
                    [forbidden_values] if isinstance(forbidden_values, str)
                    else forbidden_values
                )
            }
            for row, forbidden_values in self.excludes.items()
        )
        no_combination = not any(
            all(
                sec.text(row).upper() == value.upper()
                for row, value in combination.items()
            )
            for combination in self.excluded_combinations
        )
        return meets and not_excluded and no_combination


# Identification shared by the fixed-rate bond skeletons; each one only
# changes the market sector (MARKET_SECTOR_DES) and the currency (CRNCY).
_FIXED_RATE_CRITERIA = {
    IS_UNIT_TRADED: "N",
    CPN_TYP: "FIXED",
    SINKABLE: "N",
}

# A floating-rate bond (CPN_TYP = VARIABLE) whose CALC_TYP_DES contains FIX
# is treated as fixed. Both conditions are required.
_VARIABLE_BUT_FIX_CRITERIA = {
    **_FIXED_RATE_CRITERIA,
    CPN_TYP: "VARIABLE",
    CALC_TYP_DES: Contains("FIX"),
}

# With any of these values (even if it comes alone) the security is not one
# of these skeletons. An empty value does pass.
_FIXED_RATE_EXCLUSIONS = {
    SECURITY_TYP: "US DOMESTIC",
    COLLAT_TYP: ["CERT OF DEPOSIT", "COVERED"],
    TYPE_OF_BOND: ["Convertible", "Covered"],
}

# Government bonds: same exclusions, except US DOMESTIC (it is normal for a
# US government bond).
_GOVT_EXCLUSIONS = {
    key: value for key, value in _FIXED_RATE_EXCLUSIONS.items()
    if key != SECURITY_TYP
}

_FIXED_RATE_TRANSFORMATIONS = [
    coupon_fraction_to_decimal,
    perpetual_to_maturity,
]


def _fixed_rate_skeleton(name, sector, currency, excludes, variants=()):
    """A fixed-rate skeleton for a market sector and a currency."""
    sector_and_currency = {MARKET_SECTOR_DES: sector, CRNCY: currency}
    return Skeleton(
        name=name,
        criteria={**_FIXED_RATE_CRITERIA, **sector_and_currency},
        excludes=excludes,
        transformations=_FIXED_RATE_TRANSFORMATIONS,
        variants=[{**v, **sector_and_currency} for v in variants],
    )


SKELETONS = [
    _fixed_rate_skeleton(
        "5/80", "Corp", "USD", _FIXED_RATE_EXCLUSIONS,
        variants=[_VARIABLE_BUT_FIX_CRITERIA],
    ),
    # EUR 3/01: currency EUR, market 3, instrument 01.
    _fixed_rate_skeleton(
        "EUR 3/01", "Corp", "EUR", _FIXED_RATE_EXCLUSIONS,
        variants=[_VARIABLE_BUT_FIX_CRITERIA],
    ),
    # 8/04: same as 5/80 but for government bonds (Govt instead of Corp) and
    # without the US DOMESTIC exclusion.
    _fixed_rate_skeleton("8/04", "Govt", "USD", _GOVT_EXCLUSIONS),
]


# ============================================================
# FLOW
# ============================================================

def classify(sec):
    """First skeleton whose criteria are met, or None."""
    for skeleton in SKELETONS:
        if skeleton.matches(sec):
            return skeleton
    return None


def clean_security(sec):
    """Classify the security and apply its transformations."""
    skeleton = classify(sec)
    sec.skeleton = skeleton.name if skeleton else None

    for transformation in GENERAL_TRANSFORMATIONS:
        transformation(sec)

    if skeleton:
        for transformation in skeleton.transformations:
            transformation(sec)

    format_cells(sec)
    return sec


def verify_mnemonics(matrix):
    """Warnings if column A does not hold the expected mnemonics."""
    notices = []
    for row, expected in EXPECTED_MNEMONICS.items():
        real = _text(matrix[row - 1][0]) if len(matrix) >= row else ""
        if real.upper() != expected:
            notices.append(
                f"Row {row}: expected '{expected}' but the XLS has '{real}'."
            )
    return notices


def build_securities(matrix):
    """One Security per column (from B); empty columns at the end are
    dropped."""
    columns = len(matrix[0]) if matrix else 0
    mnemonics = {i + 1: _text(row[0]) for i, row in enumerate(matrix)}
    securities = []

    for col in range(1, columns):
        values = {i + 1: row[col] for i, row in enumerate(matrix)}
        securities.append(
            Security(column=col + 1, rows=values, mnemonics=mnemonics)
        )

    while securities and all(
        v is None for v in securities[-1].rows.values()
    ):
        securities.pop()

    return securities


def export_excel(securities, file_path):
    """One row per security with the cleaning result."""
    import pandas as pd

    records = []
    for sec in securities:
        record = {
            "XLS Column": _column_letter(sec.column),
            "Skeleton": sec.skeleton or "NO SKELETON",
            "Entity": ", ".join(sorted(sec.entities)),
            "ISIN": sec.rows.get(ID_ISIN),
            "CUSIP": sec.rows.get(ID_CUSIP),
            "DAY COUNT": sec.derived.get("DAY COUNT"),
            "YEAR BASIS": sec.derived.get("YEAR BASIS"),
            "MATURITY": sec.rows.get(MATURITY_ROW),
            "Warnings": " | ".join(sec.warnings),
        }
        for row, value in sec.rows.items():
            record[f"{row:03d} {sec.mnemonics.get(row, '')}"] = value
        records.append(record)

    pd.DataFrame(records).to_excel(file_path, index=False)


def process_matrix(matrix, destination_folder, entities=None):
    """Classify and clean all the securities of the XLS and save a review
    Excel in `destination_folder`. Returns the list of Security.

    `entities` (optional) is the list of entities of each security in the
    XLS column order; it is stored in `Security.entities`.
    """
    print()
    print("Classifying securities by skeleton...")

    for notice in verify_mnemonics(matrix):
        print(f"Warning: {notice}")

    securities = [clean_security(sec) for sec in build_securities(matrix)]

    if entities is not None:
        if len(entities) != len(securities):
            print(
                f"Warning: {len(entities)} codes were sent to BLOOMBOT but "
                f"the XLS has {len(securities)} securities; entities are "
                "not assigned."
            )
        else:
            for sec, entity in zip(securities, entities):
                sec.entities = set(entity)

    for sec in securities:
        print(
            f"  Column {_column_letter(sec.column)}: "
            f"{sec.rows.get(ID_ISIN) or '(no ISIN)'} -> "
            f"{sec.skeleton or 'NO SKELETON'}"
            + (f" [{', '.join(sorted(sec.entities))}]" if sec.entities else "")
        )
        for warning in sec.warnings:
            print(f"      Warning: {warning}")

    file_path = Path(destination_folder) / "Securities_classified.xlsx"
    export_excel(securities, file_path)
    print(f"Review Excel: {file_path}")

    return securities

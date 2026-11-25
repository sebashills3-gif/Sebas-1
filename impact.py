"""Creation of securities in IMPACT (screens NSEX02 to NSEX05) for the
skeletons listed in IMPACT_CONFIG (5/80, EUR 3/01 and 8/04).

For each security: navigate with changeProgram (the first time), choose the
market and instrument of its skeleton, fill the screens with the already
cleaned data (see skeletons.py), resolve the known warnings and errors and
leave the NSEX05 screen ready. Saving is done by the person by hand, unless
auto-save is requested.
"""

import calendar
import re
import time
from dataclasses import dataclass
from datetime import date, datetime

PROGRAM = ("NSEX", "HBSI", "")
# Only the securities of these entities are CREATED; the others go through
# the whole previous flow but are not created.
ENTITIES_TO_CREATE = {"HBSI"}


@dataclass(frozen=True)
class ImpactConfig:
    """How each skeleton is created in IMPACT."""
    market: str
    instrument: str
    currency: str = None        # ici_selectCurrency box; None = not touched
    fills_nsex05: bool = True   # False: nothing is entered in NSEX05


# Skeletons IMPACT knows how to create (the key is the name used in
# skeletons.SKELETONS).
IMPACT_CONFIG = {
    "5/80": ImpactConfig(market="5", instrument="80"),
    "8/04": ImpactConfig(market="8", instrument="04"),
    "EUR 3/01": ImpactConfig(
        market="3", instrument="01", currency="EUR", fills_nsex05=False
    ),
}

PAUSE_AFTER_ENTER_S = 1.5
MAX_ENTERS_PER_SCREEN = 12
MAX_RETRIES_PER_ERROR = 2

# Field that identifies each screen.
SCREEN_FINGERPRINTS = (
    ("NSEX05", "ici_trailerCode"),
    ("NSEX03", "ici_informationCommonStrikePrice"),
    ("NSEX02", "ici_securityName"),
    ("SELECTION", "ici_selectMarket"),
)

FREQUENCIES = {1: "A", 2: "S", 4: "Q", 12: "M"}
MONTHS_BY_FREQUENCY = {"A": 12, "S": 6, "Q": 3, "M": 1}

# Mnemonic aliases in column A of the XLS.
ALIAS = {
    "NAME": ("NAME",),
    "ID_BB": ("ID_BB", "ID_BB_UNIQUE"),
    "ID_COMMON": ("ID_COMMON", "COMMON"),
}

# Result labels shown in the final summary.
AUTO_SAVED = "Created (auto-saved)"
CONFIRMED_SAVED = "Saved (confirmed by you)"
PREPARED = "Prepared, not saved"
FAILED = "Error"
NOT_TAKEN = "Not taken"


class ImpactError(Exception):
    pass


# ============================================================
# SECURITY DATA
# ============================================================

def _is_na(value):
    """Value Bloomberg returns when there is no data."""
    if value is None:
        return True
    return str(value).strip().upper().startswith(("#N/A", "N.A.", "N/A"))


@dataclass
class ImpactData:
    """Already cleaned values of a security, as they are entered in IMPACT."""
    security_des: str = None
    name: str = None
    isin: str = None
    cusip: str = None
    cpn: str = None
    maturity: str = None
    day_count: str = None
    year_basis: str = None
    commodity: str = None
    nickname: str = None
    series: str = None
    country: str = None
    int_acc: str = None
    issue_dt: str = None
    first_cpn: str = None
    call_dt: str = None
    call_px: str = None
    frequency: str = None
    crncy: str = None
    notices: list = None


def build_data(sec):
    """IMPACT data from an already cleaned Security (skeletons.py)."""
    row_of = {}
    for row, mnemonic in sec.mnemonics.items():
        row_of.setdefault(mnemonic.strip().upper(), row)

    notices = []

    def value(*names, keep_na=False):
        for name in names:
            row = row_of.get(name)
            if row is not None:
                v = sec.rows.get(row)
                if v is None:
                    return None
                v = str(v).strip()
                if not v or (_is_na(v) and not keep_na):
                    return None
                return v
        notices.append(f"The mnemonic {names[0]} was not found in the XLS.")
        return None

    security_des = value("SECURITY_DES")
    cusip = value("ID_CUSIP")  # the final one: the ISIN-derived one if applicable

    bb = value(*ALIAS["ID_BB"])
    common = value(*ALIAS["ID_COMMON"])
    # ID_COMMON only when ID_BB equals the CUSIP that will be entered (the
    # ISIN-derived one when applicable; otherwise the XLS ID_CUSIP).
    nickname = common if (bb is not None and bb == cusip) else bb

    series = (value("SERIES") or "").upper()

    frequency = None
    cpn_freq = value("CPN_FREQ")
    if cpn_freq is not None:
        try:
            frequency = FREQUENCIES.get(int(float(cpn_freq)))
        except ValueError:
            pass
        if frequency is None:
            notices.append(f"CPN_FREQ '{cpn_freq}' is not 1, 2, 4 or 12.")

    return ImpactData(
        security_des=security_des,
        name=value(*ALIAS["NAME"]),
        isin=value("ID_ISIN"),
        cusip=cusip,
        cpn=value("CPN", keep_na=True),
        maturity=value("MATURITY"),
        day_count=sec.derived.get("DAY COUNT"),
        year_basis=sec.derived.get("YEAR BASIS"),
        commodity=(security_des or "")[:4] or None,
        nickname=nickname,
        series=series,
        country=value("CNTRY_OF_INCORPORATION"),
        int_acc=value("INT_ACC_DT"),
        issue_dt=value("ISSUE_DT"),
        first_cpn=value("FIRST_CPN_DT"),
        call_dt=value("NXT_CALL_DT"),
        call_px=value("NXT_CALL_PX"),
        frequency=frequency,
        crncy=(value("CRNCY") or "").upper(),
        notices=notices,
    )


def nsex02_fields(d):
    return {
        "ici_securityName": d.security_des,
        "ici_description1": d.name,
        "ici_description4": d.isin,
        "ici_iisnCode": d.isin,
        "ici_customerIPNumber": d.cusip,
        "ici_sortCouponStrikePrice": d.cpn,
        "ici_sortMaturity": d.maturity,
        "ici_dayCount": d.day_count,
        "ici_yearBasis": d.year_basis,
        "ici_sortIssueCommodityCodePart1": d.commodity,
        "ici_nickname": d.nickname,
        "ici_description2": d.series if d.series in ("144A", "REGS") else None,
    }


def nsex03_fields(d):
    return {
        "ici_informationCommonStrikePrice": d.cpn,
        "ici_informationMaturity": d.maturity,
        "ici_countryofOrigin": d.country,
        "ici_datedDate": d.int_acc,
        "ici_issueDate": d.issue_dt,
        "ici_firstCouponDate": d.first_cpn,
        "ici_callDate": d.call_dt,
        "ici_callPrice": d.call_px,
        "ici_informationIssueCommodityCode": d.commodity,
        "ici_elig144a": "Y" if d.series == "144A" else None,
        "ici_RegS": "Y" if d.series == "REGS" else None,
        "ici_paymentFrequencyCode": d.frequency,
    }


def nsex05_fields(d):
    """Only the fields that apply; the others are not touched."""
    fields = {}
    if "HSBC" in (d.security_des or "").upper():
        fields["ici_trailerCode"] = "REP"
    if d.crncy == "USD" and (d.isin or "").upper().startswith("XS"):
        fields["ici_depositoryAccountPointer"] = "FORG"
    return fields


# ============================================================
# PAGE ACCESS
# ============================================================

def _selector(name):
    return f'[name="{name}"], [id="{name}"]'


def _find_field(page, name):
    """Field by name or id, in the page and in all its iframes."""
    for frame in page.frames:
        try:
            matches = frame.locator(_selector(name))
            if matches.count():
                return matches.first
        except Exception:
            continue
    return None


def _wait_for_field(page, name, timeout_s=15):
    start_time = time.time()
    while True:
        input_field = _find_field(page, name)
        if input_field is not None:
            return input_field
        if time.time() - start_time >= timeout_s:
            return None
        time.sleep(0.5)


def set_value(page, name, value):
    """Write the value into the field (None leaves it empty)."""
    input_field = _wait_for_field(page, name)
    if input_field is None:
        raise ImpactError(f"The field {name} was not found on the page.")

    text = "" if value is None else str(value)
    if input_field.evaluate("e => e.tagName") == "SELECT":
        input_field.select_option(text)
    else:
        input_field.fill(text)


def fill_fields(page, fields):
    for name, value in fields.items():
        set_value(page, name, value)


def _wait_page_ready(page, timeout_ms=15000):
    for state in ("domcontentloaded", "load"):
        try:
            page.wait_for_load_state(state, timeout=timeout_ms)
        except Exception:
            pass


def enter(page, anchor_field=None):
    """Press Enter (from a field of the screen, if given) and wait."""
    input_field = _find_field(page, anchor_field) if anchor_field else None
    if input_field is not None:
        input_field.press("Enter")
    else:
        page.keyboard.press("Enter")
    time.sleep(PAUSE_AFTER_ENTER_S)
    _wait_page_ready(page)


def read_text(page):
    parts = []
    for frame in page.frames:
        try:
            parts.append(frame.locator("body").inner_text(timeout=5000))
        except Exception:
            continue
    return "\n".join(parts)


def current_screen(page):
    """NSEX02, NSEX03, NSEX05, SELECTION; NSEX04 and NSEX06 only if their
    name appears in the text; None if not recognized."""
    for name, input_field in SCREEN_FINGERPRINTS:
        if _find_field(page, input_field) is not None:
            return name

    matched = re.search(r"NSEX0[1-6]", read_text(page).upper())
    return matched.group(0) if matched else None


def wait_for_screen(page, name, timeout_s=30):
    start_time = time.time()
    while time.time() - start_time < timeout_s:
        if current_screen(page) == name:
            return
        time.sleep(0.5)
    raise ImpactError(
        f"The screen {name} was not reached "
        f"(current screen: {current_screen(page)})."
    )


# ============================================================
# NAVIGATION
# ============================================================

def go_to_selection(page):
    """Run changeProgram in the frame that defines it and wait for the
    market and instrument selection screen."""
    for frame in page.frames:
        try:
            defined = frame.evaluate("typeof changeProgram") == "function"
        except Exception:
            continue
        if defined:
            try:
                frame.evaluate(
                    "([a, b, c]) => changeProgram(a, b, c)", list(PROGRAM)
                )
            except Exception as error:  # the page may navigate when calling it
                if "context was destroyed" not in str(error).lower() \
                        and "navigation" not in str(error).lower():
                    raise
            break
    else:
        raise ImpactError("changeProgram was not found on the IMPACT page.")

    if _wait_for_field(page, "ici_selectMarket", timeout_s=30) is None:
        raise ImpactError("The market selection screen did not appear.")


def _four_digit_year(text):
    """ISSUER 5.0 12/03/26 -> ISSUER 5.0 12/03/2026."""
    return re.sub(r"(\d{1,2}/\d{1,2}/)(\d{2})(?!\d)", r"\g<1>20\2", text or "")


def _subtract_months(dt, months):
    idx = dt.year * 12 + (dt.month - 1) - months
    year_value, month_value = divmod(idx, 12)
    month_value += 1
    day_value = min(dt.day, calendar.monthrange(year_value, month_value)[1])
    return date(year_value, month_value, day_value)


# --- Known errors: error text -> fix ------------------------------------

def _fix_duplicate_sort(page, d):
    """Sort suffix = CUSIP without its first 2 and last characters, and a
    4-digit year in the name. The date with a 2-digit year may be in the
    security name (ici_securityName) or in description1, so both are
    fixed."""
    set_value(page, "ici_sortSuffix", (d.cusip or "")[2:-1])

    d.security_des = _four_digit_year(d.security_des)
    set_value(page, "ici_securityName", d.security_des)

    d.name = _four_digit_year(d.name)
    set_value(page, "ici_description1", d.name)


def _fix_commodity(page, d):
    d.commodity = (d.name or "")[:4]
    set_value(page, "ici_sortIssueCommodityCodePart1", d.commodity)


def _fix_nickname(page, d):
    d.nickname = None
    set_value(page, "ici_nickname", None)


def _fix_coupon(page, d):
    d.cpn = "0"
    set_value(page, "ici_sortCouponStrikePrice", "0")


def _fix_next_to_last_coupon(page, d):
    months = MONTHS_BY_FREQUENCY.get(d.frequency)
    if months is None or not d.maturity:
        raise ImpactError(
            "NEXT TO LAST COUPON DATE cannot be calculated "
            "(MATURITY or a valid CPN_FREQ is missing)."
        )
    maturity_date = datetime.strptime(d.maturity, "%m%d%Y").date()
    set_value(
        page,
        "ici_nextToLastCouponDate",
        _subtract_months(maturity_date, months).strftime("%m%d%Y"),
    )


NSEX02_ERRORS = {
    "DUPLICATE SORT SEQUENCE NOT ALLOWED": _fix_duplicate_sort,
    "INVALID ISSUE/COMMODITY CODE ENTRY": _fix_commodity,
    "DUPLICATE NICKNAME NOT ALLOWED": _fix_nickname,
    "INVALID COUPON/STRIKE PRICE": _fix_coupon,
}

NSEX03_ERRORS = {
    "NEXT TO LAST COUPON DATE ENTRY REQUIRED": _fix_next_to_last_coupon,
}


def advance(page, origin, destination, errors, data, anchor_field=None):
    """Press Enter until `destination` is reached: repeat on warnings and fix
    the known errors. An unknown error (or one that does not get fixed)
    stops the security with ImpactError."""
    repetitions = {}

    for _ in range(MAX_ENTERS_PER_SCREEN):
        enter(page, anchor_field)
        screen = current_screen(page)

        if screen == destination:
            return
        if destination == "NSEX04" and screen not in (origin, "NSEX05"):
            return  # NSEX04 has no fields of its own to recognize it by
        if destination == "NSEX06" and screen != origin:
            return  # neither does NSEX06; it is enough that the screen changed

        messages = re.findall(r"ERROR:?\s*([^\n\r]+)", read_text(page).upper())
        if not messages:
            continue  # warning: press Enter again

        for message in messages:
            for pattern, fix in errors.items():
                if pattern in message:
                    repetitions[pattern] = repetitions.get(pattern, 0) + 1
                    if repetitions[pattern] > MAX_RETRIES_PER_ERROR:
                        raise ImpactError(
                            f"The error '{pattern}' keeps appearing after "
                            "fixing it."
                        )
                    fix(page, data)
                    break
            else:
                raise ImpactError(f"Unhandled error in {origin}: {message}")

    raise ImpactError(f"Could not advance from {origin} to {destination}.")


# ============================================================
# CREATING ONE SECURITY
# ============================================================

def _already_on_selection(page, timeout_s=20):
    """True if the market and instrument selection screen is already open
    (after saving a security IMPACT goes back to it)."""
    return _wait_for_field(page, "ici_selectMarket", timeout_s=timeout_s) is not None


def save_automatically(page, d):
    """Press Enter on NSEX05 (saves), then on NSEX06, and wait for IMPACT to
    go back to the selection screen."""
    advance(page, "NSEX05", "NSEX06", {}, d, anchor_field="ici_trailerCode")

    if current_screen(page) != "SELECTION":
        advance(page, "NSEX06", "SELECTION", {}, d)


def create_security(page, sec, already_on_selection=False, auto_save=False):
    """Fill NSEX02 to NSEX05 for a security of a skeleton in IMPACT_CONFIG
    and leave the last screen unsaved (or, with `auto_save`, save it by
    pressing Enter on NSEX05 and NSEX06). Returns the notices about missing
    data (list).

    With `already_on_selection` the page is assumed to already be on the
    selection screen and it goes straight to entering the currency, market
    and instrument; otherwise it navigates with changeProgram.
    """
    config = IMPACT_CONFIG[sec.skeleton]
    d = build_data(sec)

    if not (already_on_selection and _already_on_selection(page)):
        go_to_selection(page)
    if config.currency:
        set_value(page, "ici_selectCurrency", config.currency)
    set_value(page, "ici_selectMarket", config.market)
    set_value(page, "ici_selectInstrument", config.instrument)
    enter(page, "ici_selectInstrument")
    wait_for_screen(page, "NSEX02")

    fill_fields(page, nsex02_fields(d))
    advance(page, "NSEX02", "NSEX03", NSEX02_ERRORS, d,
            anchor_field="ici_securityName")

    fill_fields(page, nsex03_fields(d))
    advance(page, "NSEX03", "NSEX04", NSEX03_ERRORS, d,
            anchor_field="ici_informationCommonStrikePrice")

    advance(page, "NSEX04", "NSEX05", {}, d)

    wait_for_screen(page, "NSEX05")
    if config.fills_nsex05:
        fill_fields(page, nsex05_fields(d))

    if auto_save:
        save_automatically(page, d)

    return d.notices


# ============================================================
# FLOW
# ============================================================

def _label(sec):
    isin = sec.rows.get(21)
    return str(isin) if isin else f"column {sec.column}"


def process_securities_in_impact(page, securities, auto_save=False):
    """Create in IMPACT the securities of the skeletons in IMPACT_CONFIG
    (5/80, EUR 3/01 and 8/04), one by one.

    By default it leaves each one ready on NSEX05 and waits for the person
    to save it by hand. With `auto_save` it presses Enter to save it (NSEX05
    and NSEX06) and moves on to the next one without asking. At the end it
    prints which securities were taken and which were not.
    """
    results = []   # (status, label, detail)
    pending = []

    for sec in securities:
        if sec.skeleton not in IMPACT_CONFIG:
            results.append(
                (NOT_TAKEN, _label(sec),
                 f"does not meet any skeleton ({', '.join(IMPACT_CONFIG)})")
            )
        elif sec.block_reason:
            results.append((NOT_TAKEN, _label(sec), sec.block_reason))
        elif not sec.entities & ENTITIES_TO_CREATE:
            entities = ", ".join(sorted(sec.entities)) or "no entity"
            results.append(
                (NOT_TAKEN, _label(sec),
                 f"entity {entities}: only {sorted(ENTITIES_TO_CREATE)} are created")
            )
        else:
            pending.append(sec)

    stop = False
    # After saving a security IMPACT goes back to the selection screen, and
    # from there changeProgram is no longer needed (nor available).
    on_selection = False
    for i, sec in enumerate(pending, start=1):
        label = _label(sec)

        if stop:
            results.append((NOT_TAKEN, label, "process interrupted"))
            continue

        print()
        print("=" * 60)
        print(f"IMPACT: security {i} of {len(pending)} - {label} "
              f"[{sec.skeleton}]")
        print("=" * 60)

        try:
            notices = create_security(
                page,
                sec,
                already_on_selection=on_selection,
                auto_save=auto_save,
            )

        except Exception as error:
            print(f"The security could not be prepared: {error}")
            results.append((FAILED, label, str(error)))
            on_selection = False
            continue

        for notice in notices:
            print(f"Warning: {notice}")

        if auto_save:
            print("Security saved automatically.")
            results.append((AUTO_SAVED, label, ""))
            on_selection = True
            continue

        print()
        print("Security ready on NSEX05. Review it and save it by hand in IMPACT.")
        answer = input(
            "Enter = I saved it | n = I did not save it (skip) | "
            "q = quit: "
        ).strip().lower()

        on_selection = False
        if answer == "q":
            results.append((NOT_TAKEN, label, "interrupted by you"))
            stop = True
        elif answer == "n":
            results.append((PREPARED, label, ""))
        else:
            results.append((CONFIRMED_SAVED, label, ""))
            on_selection = True

    print()
    print("=" * 60)
    print("IMPACT SUMMARY")
    print("=" * 60)
    for status in (AUTO_SAVED, CONFIRMED_SAVED, PREPARED, FAILED, NOT_TAKEN):
        group = [r for r in results if r[0] == status]
        if not group:
            continue
        print(f"{status}: {len(group)}")
        for _, label, detail in group:
            print(f"  - {label}" + (f"  ({detail})" if detail else ""))

    return results

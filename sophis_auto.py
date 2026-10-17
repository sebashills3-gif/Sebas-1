#!/usr/bin/env python3
"""
Sophis Risque automation (runs over Citrix) -- STAGE 1:
  you select the code's row -> read its Term -> set the dropdown to match -> open the update window
  (Maturity / Minimum / Maximum / Ric).

It does NOT read the Excel file or fill in the dates yet (that is stage 2).

Because Sophis runs over Citrix, this script cannot see the application's controls (it only receives
the picture of the session). So it works "like a person": it takes screenshots, reads text with OCR
and moves the mouse / presses keys. It runs on YOUR computer, with the Citrix window visible and in front.

Install (Windows):
    pip install -r requirements_sophis.txt
    (if rapidocr does not install, use Python 3.10-3.12)

Usage (in this order the first time):
    1) Open the Sophis "Listed market ..." window. It may open anywhere: the script finds it again every
       time by looking for the table header. It must be visible (not minimized) and on the primary monitor.
    2) python sophis_auto.py calibrate          # once: point the mouse at 12 spots
    3) python sophis_auto.py diagnose           # checks how your screen is read (send me what it prints)
    4) python sophis_auto.py open               # YOU click the right row; it reads its Term, sets the dropdown,
                                                # clicks the button next to 'Available contracts' and waits
                                                # for the update window (--open-with double = double click)
       python sophis_auto.py open --term "Long term share"   # same, but you tell it the Term (skips reading it)
       python sophis_auto.py open-code ZYS ZTS     # (older flow) types the code to find its row; not reliable yet
    6) python sophis_auto.py fill --entry "20 Nov 2026#1" 100 120 --entry "27 Nov 2026" 95 130 [--dry-run]
                                                # (update window open) writes Minimum/Maximum per date: double
                                                # select the cell, double click, type, Enter, read it back; never presses OK
    5) python sophis_auto.py window --date "20 Nov 2026" --show
                                                # (update window open) reads its dates and shows where that date's
                                                # Minimum and Maximum cells are; nothing is clicked or typed

Old Spanish command names (calibrar / diagnostico / abrir) still work as aliases.

Safety: move the mouse to a screen corner to ABORT at any time (pyautogui failsafe).
If anything does not match, the script stops and saves screenshots in ./sophis_debug (it never guesses).
"""

from __future__ import annotations

import argparse
import datetime as dt
import difflib
import json
import re
import statistics
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image

# ════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ════════════════════════════════════════════════════════════════════════════

CONFIG_FILE = Path("sophis_config.json")
DEBUG_DIR = Path("sophis_debug")

# The 5 Term types (as they appear in the column and in the dropdown).
TERMS = [
    "Short term index",
    "Long term index",
    "Short term share",
    "Long term share",
    "Future",
]

# Dropdown order (only used as a keyboard fallback when OCR cannot find the option in the list).
DROPDOWN_ORDER = TERMS

# The window does not always open in the same position, so NOTHING is stored as a fixed screen position:
# two "anchors" are stored (image crops of fixed parts of the header) and every point is computed relative
# to them. On every run the anchors are searched on screen and all points are shifted accordingly.
ANCHOR_DIR = Path("sophis_anchors")
ANCHOR_THRESHOLD = 0.80   # minimum match score (0-1) to consider an anchor found

# Points to calibrate (name, instruction). You point at them with the mouse on YOUR screen, once.
CALIBRATION_POINTS = [
    ("anchor_grid_a", "ANCHOR 1 (TOP-LEFT corner): the grey table header, only 'Mnemo | Extended Mnemo | "
                      "Description' (do NOT include RIC or Term). Mouse on the top-left corner of 'Mnemo'"),
    ("anchor_grid_b", "ANCHOR 1 (BOTTOM-RIGHT corner): bottom-right corner of that header, at the end of 'Description'"),
    ("row1", "Mouse on the TEXT of the Mnemo cell of the FIRST visible row (left part of the text)"),
    ("row2", "Mouse on the text of the Mnemo cell of the SECOND row (same horizontal position)"),
    ("rowN", "Mouse on the text of the Mnemo cell of the LAST fully visible row"),
    ("ext", "Mouse on the 'Extended Mnemo' column (empty), left part, on any row"),
    ("ric_left", "Mouse on the HEADER, exactly on the line that separates Code | RIC (left border of RIC)"),
    ("term_left", "Mouse on the HEADER, exactly on the line that separates RIC | Term (left border of Term)"),
    ("term_right", "Mouse on the HEADER, on the right border of the Term column (before 'Cyc')"),
    ("anchor_right_a", "ANCHOR 2 (TOP-LEFT corner): header of the right-hand table 'Date | Cycle | Name'"),
    ("anchor_right_b", "ANCHOR 2 (BOTTOM-RIGHT corner): bottom-right corner of that header (after 'Name')"),
    ("dd_left", "Mouse INSIDE the dropdown ('Short term index'), at the start of the text, vertically centered"),
    ("dd_arrow", "Mouse on the dropdown button / arrow"),
    ("avail_btn", "Mouse on the FIRST small button next to the 'Available contracts' label (the leftmost of the "
                  "little icons above the table). It opens the update window of the selected row"),
]
# The corner that closes each anchor (its image crop is saved when it is captured).
ANCHOR_CLOSING_POINT = {"anchor_grid_b": "grid", "anchor_right_b": "right"}
# Which anchor each point belongs to (points move together with their anchor).
POINT_ANCHOR = {"row1": "grid", "row2": "grid", "rowN": "grid", "ext": "grid", "ric_left": "grid",
                "term_left": "grid", "term_right": "grid", "dd_left": "right", "dd_arrow": "right",
                "avail_btn": "grid"}

# Names used by the earlier (Spanish) version, so an existing calibration keeps working.
LEGACY_CONFIG_KEYS = {"puntos": "points", "anclas": "anchors", "ensanchar_px": "widen_px",
                      "intervalo_tecla": "key_interval", "pausa": "slowdown", "lector": "reader",
                      "verificar_mnemo": "verify_mnemonic"}
LEGACY_POINT_NAMES = {"ancla_grid_a": "anchor_grid_a", "ancla_grid_b": "anchor_grid_b", "fila1": "row1",
                      "fila2": "row2", "filaN": "rowN", "term_borde": "term_left", "term_der": "term_right",
                      "ancla_der_a": "anchor_right_a", "ancla_der_b": "anchor_right_b", "dd_izq": "dd_left",
                      "dd_flecha": "dd_arrow"}
LEGACY_ANCHOR_NAMES = {"der": "right"}
LEGACY_ANCHOR_FIELDS = {"archivo": "file", "origen": "origin"}
LEGACY_READER_VALUES = {"portapapeles": "clipboard"}


class SophisError(RuntimeError):
    pass


@dataclass
class Config:
    points: dict                                  # screen positions at calibration time
    anchors: dict = field(default_factory=dict)   # {"grid": {"file": ..., "origin": [x, y]}, "right": {...}}
    widen_px: int = 140           # how far the Term column is widened to the left (0 = leave it alone)
    key_interval: float = 0.12    # seconds between key presses (Citrix may drop keys if too fast)
    slowdown: float = 1.0         # multiplier for all waits (raise it if Citrix is slow)
    reader: str = "auto"          # auto | clipboard | ocr
    verify_mnemonic: bool = True  # False = accept the row even if its Mnemo cannot be read (NOT recommended)
    weak_verify: bool = True      # accept a Mnemo OCR could not read exactly if its shape fits (same glyph count)
    auto_focus: bool = False      # False = YOU click a row once and the script only types; True = it clicks too
    open_method: str = "button"   # how the update window is opened: "button" (first button next to 'Available
                                  # contracts') or "double" (double click on the row)
    type_gap: float = 2.5         # seconds between two typed codes (the grid's type-ahead buffer must expire)

    @staticmethod
    def _migrate(d: dict) -> dict:
        """Accepts a config saved by the earlier Spanish version."""
        d = {LEGACY_CONFIG_KEYS.get(k, k): v for k, v in d.items()}
        d["points"] = {LEGACY_POINT_NAMES.get(k, k): v for k, v in d.get("points", {}).items()}
        d["anchors"] = {LEGACY_ANCHOR_NAMES.get(k, k): {LEGACY_ANCHOR_FIELDS.get(f, f): x for f, x in v.items()}
                        for k, v in d.get("anchors", {}).items()}
        if "reader" in d:
            d["reader"] = LEGACY_READER_VALUES.get(d["reader"], d["reader"])
        return d

    @staticmethod
    def load(path: Path = CONFIG_FILE, strict: bool = True) -> "Config":
        """strict=False allows missing points; `calibrate` uses it to fill in the points that are missing
        (otherwise a calibration that lacks one point could never be repaired)."""
        if not path.exists():
            raise SophisError(f"{path} does not exist. Run first:  python sophis_auto.py calibrate")
        d = Config._migrate(json.loads(path.read_text(encoding="utf-8")))
        missing = [n for n, _ in CALIBRATION_POINTS if n not in d.get("points", {})]
        if missing and strict:
            raise SophisError(f"Incomplete calibration (missing {missing}). "
                              f"Run:  python sophis_auto.py calibrate --only {' '.join(missing)}")
        if set(d.get("anchors", {})) != {"grid", "right"}:
            raise SophisError("The calibration comes from an older version (no image anchors). "
                              "Run again:  python sophis_auto.py calibrate")
        return Config(**d)

    def save(self, path: Path = CONFIG_FILE) -> None:
        path.write_text(json.dumps(self.__dict__, indent=2), encoding="utf-8")


# ════════════════════════════════════════════════════════════════════════════
# PURE HELPERS (text / color) -- no screen needed, easy to test
# ════════════════════════════════════════════════════════════════════════════

def _norm(s: str) -> str:
    """Lowercase letters/digits only: OCR sometimes joins or splits words ('ShortTermIndex')."""
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _norm_code(s: str) -> str:
    return re.sub(r"\s+", "", s).upper()


def _best_window(text_norm: str, target_norm: str) -> float:
    """Best similarity (0-1) between `target_norm` and any stretch of `text_norm` of similar length."""
    length = len(target_norm)
    if not text_norm:
        return 0.0
    best = 0.0
    for size in (length - 1, length, length + 1):
        if size <= 0:
            continue
        if len(text_norm) <= size:
            best = max(best, difflib.SequenceMatcher(None, text_norm, target_norm).ratio())
            continue
        for i in range(len(text_norm) - size + 1):
            best = max(best, difflib.SequenceMatcher(None, text_norm[i:i + size], target_norm).ratio())
    return best


def classify_term(text: str, threshold: float = 0.82, margin: float = 0.06) -> str | None:
    """Returns one of TERMS (tolerating OCR errors) or None if it is not clear."""
    n = _norm(text).replace("shares", "share")
    if not n:
        return None
    scores = sorted(((_best_window(n, _norm(t)), t) for t in TERMS), reverse=True)
    (s1, t1), (s2, _) = scores[0], scores[1]
    if s1 >= threshold and (s1 - s2) >= margin:
        return t1
    return None


def term_truncated(text: str) -> bool:
    """True if it looks like 'Short ter...' (the Term column is too narrow)."""
    for piece in text.split(" | "):      # several readings of the same crop are joined with ' | '
        n = _norm(piece)
        if classify_term(piece) is None and (n.startswith("short") or n.startswith("long")) and len(n) < 12:
            return True
    return False


def contains_word(text: str, word: str, threshold: float = 0.8) -> bool:
    return _best_window(_norm(text), _norm(word)) >= threshold


_LOOKALIKE = str.maketrans({"0": "O", "1": "I", "L": "I", "5": "S", "2": "Z", "8": "B"})


def _canon(s: str) -> str:
    """Upper-case letters/digits only, with look-alikes merged (0/O, 1/I/L, 5/S, 2/Z, 8/B)."""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper()).translate(_LOOKALIKE)


def match_code(readings: list[str], glyphs: int, code: str, allow_weak: bool = True) -> str | None:
    """Does what OCR read on a Mnemo cell correspond to `code`?

    "exact": a reading equals the code. The glyph count must not exceed the code's length (an extra glyph means
             a longer code, e.g. 'ZMH' when searching 'ZM'); one fewer is tolerated because neighbouring
             letters sometimes touch ('YY').
    "weak" : OCR could not read it exactly (very short codes like 'Z' often fail) but the cell holds as many
             glyphs as the code has characters and the readings do not contradict it.
    None   : no match."""
    want = _canon(code)
    cleaned = [_canon(r) for r in readings]
    if any(c == want for c in cleaned) and (glyphs == 0 or len(want) - 1 <= glyphs <= len(want)):
        return "exact"
    doubled = any(a == b for a, b in zip(want, want[1:]))     # 'YY': identical neighbours often touch
    if allow_weak and (glyphs == len(want) or (doubled and glyphs == len(want) - 1)):
        readable = [c for c in cleaned if c]
        if not readable or any(c[:1] == want[:1] or difflib.SequenceMatcher(None, c, want).ratio() > 0.5
                               for c in readable):
            return "weak"
    return None


def looks_like_extension(readings: list[str], glyphs: int, code: str) -> bool:
    """True if the row holds the code plus more characters (searching 'ZM' and landing on 'ZMH')."""
    want = _canon(code)
    cleaned = [_canon(r) for r in readings]
    if glyphs:
        return glyphs > len(want) and all(c[:1] == want[:1] for c in cleaned if c)
    return any(c.startswith(want) and len(c) > len(want) for c in cleaned)


def clean_cell(img: Image.Image):
    """Prepares a one-line table cell for OCR. Returns (gray image, glyph count).

    Removes what is not text: the cell background (whatever its color), vertical/horizontal grid lines, the
    selection's focus rectangle and anything touching the crop edge. That clutter is what made OCR read
    '[ZJOL]' instead of 'ZJO'. The glyph count is the number of separate characters left."""
    import cv2
    import numpy as np

    g = np.asarray(img.convert("L"))
    background = float(np.median(g))
    mask = (np.abs(g.astype(np.int16) - background) > 45).astype(np.uint8)
    h, w = mask.shape
    n, labels, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    keep = np.zeros_like(mask)
    glyphs = 0
    tiny = []                                               # 1-pixel dots: decimal points, commas (or noise)
    for i in range(1, n):
        x, y, cw, ch, area = stats[i]
        touches_edge = x <= 0 or y <= 0 or x + cw >= w or y + ch >= h
        is_line = (cw <= 1 and ch >= 0.7 * h) or (ch <= 1 and cw >= 0.7 * w)
        is_frame = cw > 0.55 * w and ch > 0.6 * h
        if touches_edge or is_line or is_frame:
            continue
        if area < 2:
            tiny.append(i)
            continue
        keep[labels == i] = 1
        glyphs += 1
    if glyphs:
        # A '.' or ',' inside a number ("100.00") is a component of a pixel or two. Keep such dots only where they
        # can be part of the text: inside the box of the real characters (the focus rectangle's dotted outline and
        # other specks lie outside it). They are not counted as glyphs.
        ys, xs = np.nonzero(keep)
        x0, x1, y0, y1 = xs.min() - 1, xs.max() + 1, ys.min() - 1, ys.max() + 1
        for i in tiny:
            x, y, cw, ch, area = stats[i]
            if x0 <= x <= x1 and y0 <= y <= y1:
                keep[labels == i] = 1
        soft = (np.abs(g.astype(np.int16) - background) > 28).astype(np.uint8)       # faint, anti-aliased dots
        sn, slabels, sstats, _ = cv2.connectedComponentsWithStats(soft, connectivity=8)
        for i in range(1, sn):
            x, y, cw, ch, area = sstats[i]
            if area <= 6 and x0 <= x <= x1 and y0 <= y <= y1 and not keep[slabels == i].any():
                keep[slabels == i] = 1
    if glyphs == 0:
        return np.full_like(g, 255), 0
    core = g[keep > 0].astype(np.float32)               # the pixels of the letters themselves
    keep = cv2.dilate(keep, np.ones((3, 3), np.uint8))  # plus a 1px halo to keep the anti-aliased edges
    gf = g.astype(np.float32)
    # Output: dark text on a pure white background, whatever the cell colors were. A selected row is often
    # WHITE text on a DARK background, so the polarity must be detected (a halo of the cell color, or text
    # painted white on white, makes the recognizer read nothing).
    if float(np.median(core)) <= background:            # dark text on a lighter cell
        lo = float(np.percentile(core, 5))
        stretched = np.clip((gf - lo) * 255.0 / max(1.0, background - lo), 0, 255)
    else:                                               # light text on a darker cell: invert
        hi = float(np.percentile(core, 95))
        stretched = 255.0 - np.clip((gf - background) * 255.0 / max(1.0, hi - background), 0, 255)
    return np.where(keep > 0, stretched, 255).astype(np.uint8), glyphs


def text_from_ocr_result(result) -> str:
    """The recognized text from a RapidOCR result, whatever the version's layout is:
    [[box, 'text', 'score']] (older) or [['text', score]] / [('text', score)] (newer). Never raises."""
    for item in result or []:
        parts = item if isinstance(item, (list, tuple)) else [item]
        for part in parts:
            if isinstance(part, str):
                return part
    return ""


# ── the update window (Maturity | Minimum | Maximum | Ric) ───────────────────

_MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
_DIGIT_FIX = str.maketrans({"O": "0", "o": "0", "l": "1", "I": "1", "|": "1", "S": "5", "B": "8"})


def _to_int(text: str) -> int:
    return int(text.translate(_DIGIT_FIX))


def parse_date(text) -> tuple[dt.date | None, str]:
    """Reads a date written as '19 Oct 2026', '19-Oct-2026', '2026-10-19' or '19/10/2026' (day first).
    Tolerates OCR slips (O for 0, l for 1). A trailing '-1' ('19 Feb 2027-1') is returned separately as the
    suffix. Returns (date or None, suffix)."""
    t = str(text).strip()
    try:
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", t)                       # ISO
        if m:
            return dt.date(int(m[1]), int(m[2]), int(m[3])), ""
        m = re.search(r"([0-9OolI|S]{1,2})\s*[-/. ]?\s*([A-Za-z0]{3,9})\.?\s*[-/. ]?\s*([0-9OolI|SB]{4})"
                      r"(?:\s*(-\s*\d+))?", t)
        if m:                                                                      # 19 Oct 2026 (-1)
            word = m[2].lower().replace("0", "o")[:3]
            best = max(_MONTHS, key=lambda mo: difflib.SequenceMatcher(None, word, mo).ratio())
            if difflib.SequenceMatcher(None, word, best).ratio() >= 0.66:
                day, year = _to_int(m[1]), _to_int(m[3])
                if 1990 <= year <= 2100:
                    return dt.date(year, _MONTHS.index(best) + 1, day), re.sub(r"\s+", "", m[4] or "")
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$", t)               # 19/10/2026 (day first)
        if m:
            a, b, y = int(m[1]), int(m[2]), int(m[3])
            try:
                return dt.date(y, b, a), ""
            except ValueError:
                return dt.date(y, a, b), ""
    except ValueError:
        pass
    return None, ""


def parse_number(text) -> float | None:
    """A number as written in a cell: '100', '100.00', '1,25' (decimal comma) or '-0.5'. None if it is not one."""
    t = re.sub(r"\s+", "", str(text))
    t = re.sub(r"(?<=\d)[Oo](?=\d|$)", "0", t)               # OCR: O for 0 between digits
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", t):         # 1,234.50 -> thousands separators
        t = t.replace(",", "")
    elif re.fullmatch(r"-?\d+,\d+", t):                          # 1,25 -> decimal comma
        t = t.replace(",", ".")
    try:
        return float(t) if re.fullmatch(r"-?\d+(\.\d+)?", t) else None
    except ValueError:
        return None


@dataclass
class FillEntry:
    """The values for one maturity date. `occurrence` picks the row when the date appears more than once."""
    date: dt.date
    minimum: str | None
    maximum: str | None
    occurrence: int | None = None


def parse_entry(date_text: str, minimum: str | None, maximum: str | None) -> FillEntry:
    """'20 Nov 2026' or '20 Nov 2026#1' (the second row with that date) plus the two values."""
    text, _, occ = str(date_text).partition("#")
    date, _ = parse_date(text)
    if date is None:
        raise SophisError(f"Could not understand the date {date_text!r} (try 19 Oct 2026, 2026-10-19 or 19/10/2026).")
    for name, value in (("Minimum", minimum), ("Maximum", maximum)):
        if value is not None and parse_number(value) is None:
            raise SophisError(f"The {name} value {value!r} for {text} is not a number.")
    return FillEntry(date, minimum, maximum, int(occ) if occ.strip().isdigit() else None)


def find_header(lines: list) -> dict | None:
    """Finds the header row 'Maturity | Minimum | Maximum | Ric' among OCR lines (each a Line). Returns the
    center x of every column and the header's y (the texts are centered in their cells), or None."""
    def candidates(word):
        return [l for l in lines if contains_word(l.text, word) and len(_norm(l.text)) <= len(word) + 3]

    for mat in candidates("Maturity"):
        for mn in (l for l in candidates("Minimum") if abs(l.cy - mat.cy) <= 10 and l.cx > mat.cx):
            for mx in (l for l in candidates("Maximum") if abs(l.cy - mat.cy) <= 10 and l.cx > mn.cx):
                ric = next((l for l in lines if abs(l.cy - mat.cy) <= 10 and l.cx > mx.cx
                            and len(_norm(l.text)) <= 4 and _best_window(_norm(l.text), "ric") >= 0.8), None)
                return {"y": mat.cy, "maturity": mat.cx, "minimum": mn.cx, "maximum": mx.cx,
                        "ric": ric.cx if ric else None}
    return None


@dataclass
class WindowRow:
    index: int                # 0 = first visible row
    cy: int                   # y of the row's center on screen
    text: str                 # what was read in the Maturity cell
    date: dt.date | None
    suffix: str = ""          # e.g. "-1" in '19 Feb 2027-1'


@dataclass
class UpdateWindow:
    header_y: int
    x_maturity: int
    x_minimum: int
    x_maximum: int
    x_ric: int | None
    pitch: float
    rows: list
    image: Image.Image | None = None
    title: tuple | None = None      # center of the window's title text ("Strikes ...")
    ok: tuple | None = None         # center of the OK button
    cancel: tuple | None = None

    def find(self, date: dt.date) -> list:
        return [r for r in self.rows if r.date == date]

    def dated(self) -> list:
        return [r for r in self.rows if r.date]


@dataclass
class CellTarget:
    """Where to put the values for one date: the centers of its Minimum and Maximum cells."""
    row: WindowRow
    minimum: tuple
    maximum: tuple
    matches: list             # every row that has this date
    ambiguous: bool           # more than one row has this date and no `occurrence` was chosen
    win: "UpdateWindow | None" = None   # the view of the window the cells were found in


def median_color(img: Image.Image) -> tuple:
    """Per-channel median color of an image (robust to a few text pixels)."""
    import numpy as np
    a = np.asarray(img.convert("RGB")).reshape(-1, 3)
    return tuple(int(v) for v in np.median(a, axis=0))


def color_distance(a, b) -> int:
    return sum(abs(int(x) - int(y)) for x, y in zip(a[:3], b[:3]))


def row_colors(img: Image.Image, row_height: float, n: int) -> list[tuple]:
    """Average color at the center of each row in a vertical strip (img starts at the top edge of row 0)."""
    px = img.convert("RGB")
    cx = px.width // 2
    colors = []
    for i in range(n):
        yc = int(round(row_height / 2 + i * row_height))
        box = (max(0, cx - 2), max(0, yc - 2), min(px.width, cx + 3), min(px.height, yc + 3))
        colors.append(px.crop(box).resize((1, 1), Image.BOX).getpixel((0, 0)))
    return colors


def find_selected_row(colors: list[tuple], tol: int = 30) -> int | None:
    """Index of the one row whose color stands out from the grid's own row colors, or None.

    Needs no reference color: grids alternate two stripe colors, so each row is compared with the median
    color of the rows of the same parity (the median ignores the single selected row). The selected row is
    the only one that differs by more than `tol`."""
    def median(values):
        values = sorted(values)
        mid = len(values) // 2
        return values[mid] if len(values) % 2 else (values[mid - 1] + values[mid]) / 2

    refs = {}
    for parity in (0, 1):
        group = [c[:3] for j, c in enumerate(colors) if j % 2 == parity]
        if not group:
            return None
        refs[parity] = tuple(median([c[k] for c in group]) for k in range(3))
    outliers = [i for i, c in enumerate(colors) if color_distance(c, refs[i % 2]) > tol]
    return outliers[0] if len(outliers) == 1 else None


def locate_template(screenshot: Image.Image, template: Image.Image) -> tuple[int, int, float]:
    """Looks for `template` inside `screenshot`. Returns (x, y, score 0-1) of its top-left corner."""
    import cv2
    import numpy as np

    a = cv2.cvtColor(np.asarray(screenshot.convert("RGB")), cv2.COLOR_RGB2GRAY)
    t = cv2.cvtColor(np.asarray(template.convert("RGB")), cv2.COLOR_RGB2GRAY)
    if t.shape[0] > a.shape[0] or t.shape[1] > a.shape[1] or float(t.std()) < 1.0:
        return 0, 0, 0.0                                  # too big, or a flat image that matches "everywhere"
    res = cv2.matchTemplate(a, t, cv2.TM_CCOEFF_NORMED)
    _, best, _, pos = cv2.minMaxLoc(res)
    if not np.isfinite(best):
        return 0, 0, 0.0
    return int(pos[0]), int(pos[1]), float(best)


def compute_points(screenshot: Image.Image, cfg: "Config") -> tuple[dict, dict]:
    """Finds the anchors in the screenshot and returns (points in CURRENT coordinates, info per anchor).
    Raises SophisError if the window is not in view."""
    points, info = {}, {}
    for name, anchor in cfg.anchors.items():
        path = Path(anchor["file"])
        if not path.exists():
            raise SophisError(f"Anchor image {path} is missing. Run again:  python sophis_auto.py calibrate")
        x, y, score = locate_template(screenshot, Image.open(path))
        info[name] = (x - anchor["origin"][0], y - anchor["origin"][1], score)
        if score < ANCHOR_THRESHOLD:
            raise SophisError(
                f"Cannot find the Sophis window on screen (anchor '{name}', match {score:.2f} < {ANCHOR_THRESHOLD}). "
                "Is it visible (not minimized or covered) and on the primary monitor? If the Citrix zoom/resolution "
                "changed, calibrate again.")
    for point, name in POINT_ANCHOR.items():
        if point not in cfg.points:        # partial calibration: the missing point is about to be recorded
            continue
        dx, dy, _ = info[name]
        px, py = cfg.points[point]
        points[point] = [px + dx, py + dy]
    return points, info


# ════════════════════════════════════════════════════════════════════════════
# OCR
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Line:
    text: str
    conf: float
    x0: int
    y0: int
    x1: int
    y1: int

    @property
    def cx(self) -> int:
        return (self.x0 + self.x1) // 2

    @property
    def cy(self) -> int:
        return (self.y0 + self.y1) // 2


class OCR:
    """RapidOCR (pip only, no separate program to install)."""

    def __init__(self) -> None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as e:
            raise SophisError("OCR is missing. Install:  pip install rapidocr-onnxruntime") from e
        self._engine = RapidOCR()

    def read(self, img: Image.Image, scale: int = 3, margin: int = 24) -> list[Line]:
        import numpy as np

        img = img.convert("RGB")
        if scale != 1:
            img = img.resize((img.width * scale, img.height * scale), Image.LANCZOS)
        # The border is padded with the background color: small crops without a margin are almost never detected.
        a = np.asarray(img)
        border = np.concatenate([a[0], a[-1], a[:, 0], a[:, -1]])
        background = tuple(int(v) for v in np.median(border, axis=0))
        canvas = Image.new("RGB", (img.width + 2 * margin, img.height + 2 * margin), background)
        canvas.paste(img, (margin, margin))
        result, _ = self._engine(np.asarray(canvas)[:, :, ::-1])  # RapidOCR expects BGR
        lines = []
        for box, text, conf in result or []:
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            lines.append(Line(
                text=text, conf=float(conf),
                x0=int((min(xs) - margin) / scale), y0=int((min(ys) - margin) / scale),
                x1=int((max(xs) - margin) / scale), y1=int((max(ys) - margin) / scale),
            ))
        return lines

    def text(self, img: Image.Image, scale: int = 3) -> str:
        lines = sorted(self.read(img, scale), key=lambda l: (l.x0, l.y0))
        return " ".join(l.text for l in lines)

    def read_cell(self, img: Image.Image) -> tuple[list[str], int]:
        """Reads ONE short line of text (a table cell). Returns (several readings, glyph count).

        The crop is cleaned first (see clean_cell) and sent straight to the text recognizer, skipping the
        text-detection step, which is what fails on tiny crops such as 'ZS' or 'YY'."""
        import cv2
        import numpy as np

        cleaned, glyphs = clean_cell(img)
        if glyphs == 0:
            return [], 0
        readings = []
        engine = self._engine
        # Old rapidocr versions ignore use_det/use_cls and always run the angle classifier, which turns small
        # crops upside down ('ZJO' -> 'orz'). Force "recognize only" through their attributes as well.
        saved = {name: getattr(engine, name) for name in ("min_height", "use_angle_cls") if hasattr(engine, name)}
        try:
            if "min_height" in saved:
                engine.min_height = 10 ** 6      # "image is too short for detection" -> skip detection
            if "use_angle_cls" in saved:
                engine.use_angle_cls = False
            for scale in (3, 4):
                big = cv2.resize(cleaned, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
                big = cv2.copyMakeBorder(big, 16, 16, 16, 16, cv2.BORDER_CONSTANT, value=255)
                result, _ = engine(cv2.cvtColor(big, cv2.COLOR_GRAY2BGR), use_det=False, use_cls=False,
                                   use_rec=True, text_score=0.0)
                readings.append(text_from_ocr_result(result))
        finally:
            for name, value in saved.items():
                setattr(engine, name, value)
        return readings, glyphs


# ════════════════════════════════════════════════════════════════════════════
# SCREEN (mouse / keyboard / screenshots / clipboard)
# ════════════════════════════════════════════════════════════════════════════

def _enable_dpi_awareness() -> None:
    """Without this, on Windows with scaling >100% mouse and screenshot coordinates do not match."""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class Screen:
    def __init__(self, key_interval: float = 0.12) -> None:
        _enable_dpi_awareness()
        try:
            import pyautogui
        except ImportError as e:
            raise SophisError("pyautogui is missing. Install:  pip install pyautogui pillow") from e
        pyautogui.FAILSAFE = True      # mouse in a corner = abort
        pyautogui.PAUSE = 0.05
        self.pg = pyautogui
        self.key_interval = key_interval
        self._tk = None
        self._last_click = -1e9
        self._double_interval = self._double_click_time()

    def position(self) -> tuple[int, int]:
        p = self.pg.position()
        return int(p.x), int(p.y)

    def capture(self, region: tuple[int, int, int, int] | None = None) -> Image.Image:
        """region = (x0, y0, x1, y1) in screen coordinates."""
        if region is None:
            return self.pg.screenshot().convert("RGB")
        x0, y0, x1, y1 = (int(v) for v in region)
        return self.pg.screenshot(region=(x0, y0, max(1, x1 - x0), max(1, y1 - y0))).convert("RGB")

    @staticmethod
    def _double_click_time() -> float:
        """Time (s) within which Windows merges two clicks into a double click (normally 0.5 s)."""
        if sys.platform == "win32":
            try:
                import ctypes
                return ctypes.windll.user32.GetDoubleClickTime() / 1000.0
            except Exception:
                pass
        return 0.5

    def _wait_out_double_click(self) -> None:
        """Two single clicks in quick succession are read as ONE double click: wait for that window to pass."""
        remaining = self._last_click + self._double_interval + 0.15 - time.monotonic()
        if remaining > 0:
            time.sleep(remaining)

    def click(self, x: int, y: int) -> None:
        """A single click (never merged with the previous click)."""
        self._wait_out_double_click()
        self.pg.moveTo(x, y, duration=0.15)
        self.pg.click()
        self._last_click = time.monotonic()

    def double_click(self, x: int, y: int) -> None:
        self._wait_out_double_click()
        self.pg.moveTo(x, y, duration=0.15)
        self.pg.doubleClick(interval=0.12)
        self._last_click = time.monotonic()

    def type_text(self, text: str) -> None:
        for ch in text:
            self.pg.write(ch)
            time.sleep(self.key_interval)

    def press(self, key: str) -> None:
        self.pg.press(key)

    def hotkey(self, *keys: str) -> None:
        self.pg.hotkey(*keys)

    def scroll(self, clicks: int, x: int, y: int) -> None:
        """Mouse wheel over (x, y): positive scrolls up, negative scrolls down."""
        self.pg.moveTo(x, y, duration=0.15)
        self.pg.scroll(clicks)

    def drag(self, x0: int, y0: int, x1: int, y1: int) -> None:
        self.pg.moveTo(x0, y0, duration=0.2)
        self.pg.mouseDown()
        self.pg.moveTo(x1, y1, duration=0.5)
        self.pg.mouseUp()

    def _root(self):
        if self._tk is None:
            import tkinter
            self._tk = tkinter.Tk()
            self._tk.withdraw()
        return self._tk

    def clipboard_set(self, text: str) -> None:
        r = self._root()
        r.clipboard_clear()
        r.clipboard_append(text)
        r.update()

    def clipboard_get(self) -> str | None:
        r = self._root()
        r.update()
        try:
            return r.clipboard_get()
        except Exception:
            return None


# ════════════════════════════════════════════════════════════════════════════
# AUTOMATOR
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Geometry:
    x_click: int      # x where a Mnemo cell is clicked
    x_mnemo0: int
    x_mnemo1: int
    x_ext: int
    y1: int           # y of the center of row 0
    h: float          # row height
    n: int            # visible rows
    term_x0: int
    term_x1: int

    def row_y(self, i: int) -> int:
        return int(round(self.y1 + i * self.h))


@dataclass
class Reading:
    mnemo_ok: bool
    term: str | None
    source: str
    text: str
    verified: bool = True
    seen: tuple = ()     # mnemonics read on the row (to detect a prefix-only match)
    glyphs: int = 0      # how many separate characters the Mnemo cell holds (0 = unknown)
    match: str | None = None  # "exact" | "weak" | None (how the Mnemo was confirmed)
    term_text: str = ""  # raw text read from the Term cell (to detect a truncated column)


@dataclass
class Result:
    code: str
    term: str
    row: int
    window: dict


class Automator:
    def __init__(self, screen, ocr, cfg: Config, log=print) -> None:
        self.s = screen
        self.ocr = ocr
        self.cfg = cfg
        self.log = log
        self.pts: dict = {}
        self.g: Geometry
        self._widened = False          # the widening was already attempted this session
        self._widened_by = 0           # how many px Term's right border was dragged
        self._screen_width = 10 ** 6
        self._win_cache: dict | None = None    # what is known about the update window (see read_update_window)
        self.trace = False                      # save a screenshot at every step of fill_values
        self._trace_n = 0
        self._first_click_done = False
        self._last_typed_at = -1e9
        self._clipboard_works = True

    # ── helpers ─────────────────────────────────────────────────────────────
    def _wait(self, seconds: float) -> None:
        time.sleep(seconds * self.cfg.slowdown)

    def _pt(self, name: str) -> tuple[int, int]:
        x, y = self.pts[name]
        return int(x), int(y)

    def locate(self) -> None:
        """Finds the window on screen (it may be somewhere else than at calibration) and recomputes the points."""
        shot = self.s.capture()
        self._screen_width = shot.width
        self.pts, info = compute_points(shot, self.cfg)
        self.g = self._geometry()
        moved = ", ".join(f"{n}: shifted ({dx:+d},{dy:+d}) match {sc:.2f}" for n, (dx, dy, sc) in info.items())
        self.log(f"  window located -> {moved}")

    def _geometry(self) -> Geometry:
        r1, r2, rn = self._pt("row1"), self._pt("row2"), self._pt("rowN")
        ext, ric, tr = self._pt("ext"), self._pt("ric_left"), self._pt("term_right")
        h = r2[1] - r1[1]
        data = f"(row1 y={r1[1]}, row2 y={r2[1]}, rowN y={rn[1]})"
        if h < 8:
            raise SophisError(
                f"Invalid calibration {data}: 'row2' must be about 15-25px BELOW 'row1'. "
                "Redo only those points:  python sophis_auto.py calibrate --only row1 row2 rowN")
        n = int(round((rn[1] - r1[1]) / h)) + 1
        if n < 3:
            raise SophisError(
                f"Invalid calibration {data}: row height={h}px, so 'rowN' falls on row {n}. "
                "'rowN' must be the LAST visible row of the table (several rows below 'row1'). "
                "Redo only those points:  python sophis_auto.py calibrate --only row1 row2 rowN")
        return Geometry(
            x_click=r1[0], x_mnemo0=r1[0] - 30, x_mnemo1=max(r1[0] + 12, ext[0] - 2), x_ext=ext[0],
            y1=r1[1], h=float(h), n=n,
            term_x0=ric[0] + 3, term_x1=tr[0] + self._widened_by - 1,
        )

    def save_debug(self, name: str, img: Image.Image | None = None) -> Path:
        DEBUG_DIR.mkdir(exist_ok=True)
        path = DEBUG_DIR / f"{time.strftime('%H%M%S')}_{name}.png"
        (img or self.s.capture()).save(path)
        return path

    # ── preparation ─────────────────────────────────────────────────────────
    def widen_term(self) -> None:
        """The Term column shows 'Short ter...'. Widen it the way you do by hand: drag the RIGHT border of Term
        (the line between Term and Cyc) to the right. Only that border is touched."""
        if self._widened or self.cfg.widen_px <= 0:
            return
        tr = self._pt("term_right")
        room = self._screen_width - 5 - tr[0]                  # do not drag off the screen
        d = max(0, min(self.cfg.widen_px, room))
        self._widened = True                                   # one attempt per session
        if d < 30:
            self.log("  not enough room on screen to widen Term; widen it by hand instead")
            return
        self.log(f"Widening the Term column by {d}px (dragging its right border) ...")
        self.s.drag(tr[0], tr[1], tr[0] + d, tr[1])
        self._widened_by = d
        self.g = self._geometry()                              # the Term strip now reaches the new right edge
        self._wait(0.6)

    def prepare(self) -> None:
        self.locate()

    def _type_code(self, code: str) -> int | None:
        """Types `code` once the type-ahead buffer of the previous code has expired.

        Manual mode (default) TRUSTS that you already selected a row: there is no check that can stop the
        typing (the checks used to give false alarms and block it). The row that was selected before typing is
        only looked up for information (None if it cannot be told); the result is verified AFTER typing."""
        pause = self._last_typed_at + self.cfg.type_gap * self.cfg.slowdown - time.time()
        if pause > 0:
            time.sleep(pause)
        start = self.focus_grid() if self.cfg.auto_focus else self.selected_row()
        self.s.type_text(code)
        self._last_typed_at = time.time()
        return start

    def _locate_row_by_reading(self, code: str) -> int | None:
        """Fallback when the highlighted row cannot be told by color: read the Mnemo of every visible row and
        return the one that exactly equals `code` (None if none or several do)."""
        hits = []
        for i in range(self.g.n):
            readings, glyphs = self._mnemonic_by_ocr(i)
            if match_code(readings, glyphs, code, allow_weak=False) == "exact":
                hits.append(i)
        return hits[0] if len(hits) == 1 else None

    def focus_grid(self) -> int:
        """--auto-focus only: gives the grid keyboard focus with ONE click and returns the clicked row.

        In this grid a click on a row that is ALREADY SELECTED opens its cell for editing, so the click always
        goes to a row that is not selected (a click on an unselected row only selects it). It never blocks the
        typing: if the selection cannot be confirmed it only logs a warning."""
        selected = self.selected_row()
        candidates = [i for i in (1, 2, 3) if i != selected and i < self.g.n]
        target = candidates[0]
        self.click_row(target)
        self._wait(0.4)
        if not self._first_click_done:
            # The very first click may only bring the Citrix window to the front; give it a moment so the
            # first characters are not typed into a window that is not ready yet.
            self._first_click_done = True
            self._wait(1.0)
        if self.selected_row() == target:
            return target
        alt = next(i for i in (1, 2, 3, 4) if i not in (target, selected) and i < self.g.n)
        self.log(f"  row {target} does not look selected yet; clicking row {alt} instead")
        self.click_row(alt)
        self._wait(0.5)
        if self.selected_row() != alt:
            self.log("  WARNING: could not confirm the selected row; typing anyway")
        return alt

    def _row_in_edit_mode(self, i: int) -> bool:
        """Informational only (shown by `diagnose`, never used to block anything): True if the Mnemo cell of
        row i looks different from its neighbour, which is what an open edit box looks like. It can give false
        alarms when the columns simply have different background colors."""
        g = self.g
        yc = g.row_y(i)
        cell = self.s.capture(self._cell_region(i, g.x_mnemo0, g.x_mnemo1))
        beside = self.s.capture((g.x_ext - 3, yc - 3, g.x_ext + 4, yc + 4))
        return color_distance(median_color(cell), median_color(beside)) > 60

    # ── grid ────────────────────────────────────────────────────────────────
    def click_row(self, i: int, double: bool = False) -> None:
        y = self.g.row_y(i)
        (self.s.double_click if double else self.s.click)(self.g.x_click, y)

    def _colors(self) -> list[tuple]:
        g = self.g
        top = int(round(g.y1 - g.h / 2))
        img = self.s.capture((g.x_ext - 3, top, g.x_ext + 4, int(round(top + g.h * g.n))))
        return row_colors(img, g.h, g.n)

    def selected_row(self) -> int | None:
        return find_selected_row(self._colors())

    def _cell_region(self, i: int, x0: int, x1: int) -> tuple[int, int, int, int]:
        yc = self.g.row_y(i)
        half = int(self.g.h / 2) - 1
        return (x0, yc - half, x1, yc + half)

    def _read_clipboard(self) -> str | None:
        """Copies the selected row (Ctrl+C, then Ctrl+Insert) and waits for Citrix to hand the text over."""
        sentinel = f"__sophis_{time.time()}__"
        for keys in (("ctrl", "c"), ("ctrl", "insert")):
            self.s.clipboard_set(sentinel)
            self.s.hotkey(*keys)
            deadline = time.time() + 1.5 * self.cfg.slowdown
            while True:
                self._wait(0.2)
                text = self.s.clipboard_get()
                if text and text != sentinel:
                    return text
                if time.time() >= deadline:
                    break
        return None

    def _column_image(self) -> Image.Image:
        """The visible part of the Mnemo column (used to tell whether typing made the grid move)."""
        g = self.g
        top = int(round(g.y1 - g.h / 2))
        return self.s.capture((g.x_mnemo0, top, g.x_mnemo1, int(round(top + g.h * g.n)))).convert("L")

    @staticmethod
    def _differs(a: Image.Image, b: Image.Image) -> bool:
        import numpy as np
        return float(np.mean(np.abs(np.asarray(a, dtype=np.int16) - np.asarray(b, dtype=np.int16)))) > 1.0

    def _mnemonic_by_ocr(self, i: int) -> tuple[list[str], int]:
        """(readings, glyph count) of the Mnemo cell of row i."""
        region = self._cell_region(i, self.g.x_mnemo0, self.g.x_mnemo1)
        return self.ocr.read_cell(self.s.capture(region))

    def _term_by_ocr(self, i: int) -> tuple[str | None, str]:
        img = self.s.capture(self._cell_region(i, self.g.term_x0, self.g.term_x1))
        texts = []
        for scale in (3, 4):
            text = self.ocr.text(img, scale)
            texts.append(text)
            term = classify_term(text)
            if term:
                return term, text
        return None, " | ".join(texts)

    def read_row(self, i: int, code: str, allow_weak: bool = True) -> Reading:
        mode = self.cfg.reader
        fallback = None
        if mode == "clipboard" or (mode == "auto" and self._clipboard_works):
            text = self._read_clipboard()
            if text:
                cells = [c.strip() for c in re.split(r"[\t\r\n]+", text) if c.strip()]
                if any(_norm_code(c) == _norm_code(code) for c in cells[:3]):
                    return Reading(True, classify_term(text), "clipboard", text, term_text=text, match="exact")
                if cells:
                    fallback = Reading(False, classify_term(text), "clipboard", text, seen=(cells[0],),
                                       term_text=text)
            elif mode == "auto":
                self._clipboard_works = False
                self.log("  clipboard gave nothing (Citrix may block it); reading with OCR from now on")
        if mode in ("auto", "ocr"):
            readings, glyphs = self._mnemonic_by_ocr(i)
            term, term_text = self._term_by_ocr(i)
            match = match_code(readings, glyphs, code, self.cfg.weak_verify and allow_weak)
            text = f"mnemonic={readings} glyphs={glyphs} term='{term_text}'"
            if match:
                return Reading(True, term, "ocr", text, verified=(match == "exact"), seen=tuple(readings),
                               glyphs=glyphs, term_text=term_text, match=match)
            if not self.cfg.verify_mnemonic and term:
                self.log("  WARNING: Mnemo NOT verified (verify_mnemonic=False)")
                return Reading(True, term, "ocr", text, verified=False, term_text=term_text)
            return Reading(False, term, "ocr", text, seen=tuple(readings), glyphs=glyphs, term_text=term_text)
        return fallback or Reading(False, None, mode, "")

    # ── search ──────────────────────────────────────────────────────────────
    @staticmethod
    def _is_extension(reading: Reading, code: str) -> bool:
        """True if the row holds the code plus more characters (e.g. you search 'ZM' and land on 'ZMH')."""
        return looks_like_extension(list(reading.seen), reading.glyphs, code)

    def _walk_down_to(self, code: str, max_rows: int = 15) -> tuple[int, Reading] | None:
        """Incremental search lands on the first prefix match; in a descending list the exact code (ZM)
        comes AFTER its extensions (ZMH). Walk down row by row, checking each one."""
        for _ in range(max_rows):
            self.s.press("down")
            self._wait(0.35)
            j = self.selected_row()
            if j is None:
                return None
            reading = self.read_row(j, code)
            if reading.mnemo_ok:
                return j, reading
            if not self._is_extension(reading, code):
                return None
        return None

    def search(self, code: str, attempts: int = 3) -> tuple[int, Reading]:
        """Type the code (the grid already has focus because you clicked a row, or --auto-focus clicked an
        unselected one). If the row cannot be confirmed, the next attempt
        first just re-reads (the grid may still have been scrolling); only the last attempt, or one that
        found no selected row at all, clicks and types again."""
        last = ""
        retype = True
        moved = False
        clicked = None
        for k in range(1, attempts + 1):
            if retype or k == attempts:
                before = self._column_image()
                clicked = self._type_code(code)
                self._wait(0.8)
                moved = self._differs(before, self._column_image())
            else:
                self._wait(0.8)
            i = self.selected_row()
            if i is None:
                i = self._locate_row_by_reading(code)
                if i is not None:
                    self.log("  highlighted row not detected by color; found it by reading the Mnemo column")
            if i is None:
                last, retype = "no selected row detected", True
                self.log(f"  attempt {k}: {last}")
                continue
            retype = False
            # A match by shape alone is only allowed if typing really moved the grid (otherwise the code was
            # probably not found and the highlighted row is just where it already was).
            reading = self.read_row(i, code, allow_weak=moved or (clicked is not None and i != clicked))
            if not reading.mnemo_ok and self._is_extension(reading, code):
                self.log(f"  landed on {list(reading.seen)[:1]}; walking down to '{code}' ...")
                found = self._walk_down_to(code)
                if found:
                    i, reading = found
            if reading.mnemo_ok:
                note = "" if reading.match != "weak" else "  (Mnemo confirmed by shape only)"
                self.log(f"  row {i} ok ({reading.source}): term={reading.term}{note}")
                return i, reading
            last = f"the selected row is not '{code}' ({reading.text})"
            self.log(f"  attempt {k}: {last}")
        path = self.save_debug(f"not_found_{code}")
        raise SophisError(f"Could not locate '{code}': {last}. Screenshot: {path}")

    # ── dropdown ────────────────────────────────────────────────────────────
    def read_dropdown(self) -> tuple[str | None, str]:
        left, arrow = self._pt("dd_left"), self._pt("dd_arrow")
        img = self.s.capture((left[0] - 6, left[1] - 11, arrow[0] - 8, left[1] + 11))
        for scale in (3, 4):
            text = self.ocr.text(img, scale)
            term = classify_term(text)
            if term:
                return term, text
        return None, text

    def set_dropdown(self, term: str) -> None:
        current, _ = self.read_dropdown()
        if current == term:
            self.log(f"  dropdown is already on '{term}'")
            return
        left, arrow = self._pt("dd_left"), self._pt("dd_arrow")
        region = (left[0] - 8, arrow[1] + 6, arrow[0] + 10, arrow[1] + 6 + 260)
        for k in range(1, 4):
            self.s.click(*arrow)
            self._wait(0.8)
            lines = self.ocr.read(self.s.capture(region), scale=2)
            target = next((l for l in lines if classify_term(l.text) == term), None)
            if target:
                self.s.click(region[0] + target.cx, region[1] + target.cy)
            else:
                self.log(f"  attempt {k}: did not see '{term}' in the list; using the keyboard")
                self.s.press("home")
                for _ in range(DROPDOWN_ORDER.index(term)):
                    self.s.press("down")
                self.s.press("enter")
            self._wait(0.8)
            current, text = self.read_dropdown()
            if current == term:
                self.log(f"  dropdown -> '{term}'")
                return
            self.log(f"  attempt {k}: dropdown ended on {current!r} ({text!r})")
        path = self.save_debug("dropdown_mismatch")
        raise SophisError(f"Could not set the dropdown to '{term}'. Screenshot: {path}")

    # ── update window ───────────────────────────────────────────────────────
    def wait_for_window(self, timeout: float = 20.0) -> dict:
        """Waits for the window with columns Maturity / Minimum / Maximum / Ric to appear."""
        t0 = time.time()
        while time.time() - t0 < timeout * max(1.0, self.cfg.slowdown):
            self._wait(1.0)
            img = self.s.capture()
            for scale in (1, 2):      # quick pass first, then a more precise one (the text is small)
                found = {}
                for line in self.ocr.read(img, scale=scale, margin=8):
                    for word in ("Maturity", "Minimum", "Maximum"):
                        if word not in found and contains_word(line.text, word):
                            found[word] = (line.cx, line.cy)
                if len(found) >= 2:
                    return found
        path = self.save_debug("window_did_not_open")
        raise SophisError(f"The Maturity/Minimum/Maximum window did not appear. Screenshot: {path}")

    # ── full stage-1 flow ───────────────────────────────────────────────────
    def open_contract(self, code: str) -> Result:
        code = code.strip().upper()
        self.log(f"[{code}] searching ...")
        self.locate()      # the window may have moved since the previous code
        i, reading = self.search(code)
        if not reading.term and self.cfg.widen_px > 0 and not self._widened:
            # Unreadable Term (usually because it is truncated: "Short ter..."): widen the column once per
            # session. With the clipboard reader this is almost never needed.
            self.widen_term()
            reading = self.read_row(i, code)
        if not reading.term:
            if term_truncated(reading.term_text):
                raise SophisError(f"[{code}] the Term column is truncated ('{reading.term_text}'). "
                                  f"Increase widen_px in {CONFIG_FILE} or widen it by hand.")
            path = self.save_debug(f"term_unreadable_{code}")
            raise SophisError(f"[{code}] could not read the Term ('{reading.text}'). Screenshot: {path}")
        self.set_dropdown(reading.term)

        # Changing the dropdown may reload/move the list: re-check the row before opening.
        j = self.selected_row()
        if j is None or not self.read_row(j, code).mnemo_ok:
            j = self._locate_row_by_reading(code)
        if j is None:
            if not self.cfg.auto_focus:
                # The dropdown click took the keyboard focus away from the grid, so typing the code again
                # would go nowhere. Nothing is typed or clicked blindly.
                path = self.save_debug(f"row_lost_{code}")
                raise SophisError(f"[{code}] the row could not be found again after setting the dropdown. "
                                  f"Click a row, then run this code again. Screenshot: {path}")
            self.log("  the list changed after the dropdown; searching again")
            j, _ = self.search(code)
        self.open_update_window(j)
        window = self.wait_for_window()
        self.log(f"[{code}] done: term='{reading.term}', update window is open")
        return Result(code, reading.term, j, window)

    # ── the flow you drive: YOU select the row, the script continues from there ──
    def row_under_mouse(self) -> int | None:
        """The visible row the mouse pointer is over (your pointer stays on the row you just clicked)."""
        x, y = self.s.position()
        g = self.g
        if not (g.x_mnemo0 - 40 <= x <= g.term_x1 + 40):
            return None
        i = int(round((y - g.y1) / g.h))
        return i if 0 <= i < g.n and abs(y - g.row_y(i)) <= g.h / 2 else None

    def find_selected_row_or_mouse(self) -> tuple[int, str]:
        """Which row did you select? The highlighted row if it can be seen, otherwise the row under the mouse."""
        mouse = self.row_under_mouse()          # read first: nothing has moved the pointer yet
        color = self.selected_row()
        if color is not None:
            if mouse is not None and mouse != color:
                self.log(f"  note: the highlight is on row {color} but the mouse is on row {mouse}; using the highlight")
            return color, "highlight"
        if mouse is not None:
            return mouse, "mouse position"
        path = self.save_debug("row_not_found")
        raise SophisError("Could not tell which row you selected (no highlighted row found and the mouse is not "
                          "over the table). Click the row once and leave the mouse on it. Screenshot: " + str(path))

    def _read_term_of_row(self, i: int) -> str | None:
        """The Term of row i: from the clipboard if it works, else by OCR (widening the column if it is cut)."""
        mode = self.cfg.reader
        if mode == "clipboard" or (mode == "auto" and self._clipboard_works):
            text = self._read_clipboard()
            if text:
                term = classify_term(text)
                if term:
                    return term
            elif mode == "auto":
                self._clipboard_works = False
        if mode in ("auto", "ocr"):
            term, _ = self._term_by_ocr(i)
            if not term and self.cfg.widen_px > 0 and not self._widened:
                self.widen_term()
                term, _ = self._term_by_ocr(i)
            return term
        return None

    def open_update_window(self, row: int) -> None:
        """Opens the update window (Maturity / Minimum / Maximum / Ric) of the selected row."""
        if self.cfg.open_method == "double":
            self.log(f"  opening the window (double click on row {row}) ...")
            self.click_row(row, double=True)
        else:
            self.log("  opening the window (click on the first button next to 'Available contracts') ...")
            self.s.click(*self._pt("avail_btn"))

    def open_selected(self, term: str | None = None) -> Result:
        """Assumes YOU already selected the right row. Reads its Term (or uses `term`), matches the dropdown,
        double-clicks the row and waits for the update window. Nothing is typed and no row is searched."""
        self.locate()
        i, how = self.find_selected_row_or_mouse()
        readings, glyphs = self._mnemonic_by_ocr(i)
        shown = next((r for r in readings if str(r).strip()), "?")
        before = clean_cell(self.s.capture(self._cell_region(i, self.g.x_mnemo0, self.g.x_mnemo1)))[0]
        if term is None:
            term = self._read_term_of_row(i)
            if not term:
                path = self.save_debug("term_unreadable")
                raise SophisError(f"Could not read the Term of row {i}. Tell it yourself with  --term \"Short term "
                                  f"index\"  (or widen the Term column by hand). Screenshot: {path}")
        self.log(f"  row {i} ({how}); Mnemo reads {shown!r}; Term = {term!r}")
        self.set_dropdown(term)

        # The dropdown click moved the keyboard focus; the list should be unchanged. Warn (never block) if the
        # row's text is not the same any more, since the double click would then open another contract.
        import numpy as np
        after = clean_cell(self.s.capture(self._cell_region(i, self.g.x_mnemo0, self.g.x_mnemo1)))[0]
        if float(np.mean(np.abs(before.astype(np.int16) - after.astype(np.int16)))) > 12:
            self.log("  WARNING: the text of the row changed after setting the dropdown; check the window that opens")
        self.open_update_window(i)
        window = self.wait_for_window()
        self.log(f"done: Mnemo '{shown}', Term '{term}', update window is open")
        self.report_update_window()
        return Result(str(shown), term, i, window)

    # ── stage 2: find where the values go in the update window (nothing is clicked or typed) ──
    def _ocr_tiled(self, shot: Image.Image, cols: int = 3, rows: int = 2, scale: int = 2) -> list:
        """OCR a big screenshot in overlapping tiles. One pass over a whole high-resolution screen gets shrunk by
        the text detector until small text is unreadable; tiles keep the text legible."""
        w, h = shot.size
        tw, th = w / cols, h / rows
        ox, oy = int(tw * 0.15), int(th * 0.15)
        lines = []
        for r in range(rows):
            for c in range(cols):
                x0, y0 = max(0, int(c * tw) - ox), max(0, int(r * th) - oy)
                x1, y1 = min(w, int((c + 1) * tw) + ox), min(h, int((r + 1) * th) + oy)
                for l in self.ocr.read(shot.crop((x0, y0, x1, y1)), scale=scale, margin=8):
                    lines.append(Line(l.text, l.conf, l.x0 + x0, l.y0 + y0, l.x1 + x0, l.y1 + y0))
        return lines

    def _dump_lines(self, name: str, lines: list) -> Path:
        DEBUG_DIR.mkdir(exist_ok=True)
        path = DEBUG_DIR / f"{time.strftime('%H%M%S')}_{name}.txt"
        path.write_text("\n".join(f"{l.text!r:40} cx={l.cx} cy={l.cy} conf={l.conf:.2f}"
                                  for l in sorted(lines, key=lambda l: (l.cy, l.cx))), encoding="utf-8")
        return path

    def _window_cache_files(self) -> tuple[Path, Path]:
        return ANCHOR_DIR / "update_window.png", ANCHOR_DIR / "update_window.json"

    def _load_window_cache(self) -> None:
        """Remembers from an earlier run how the update window's header looks (see _save_window_cache)."""
        png, meta = self._window_cache_files()
        try:
            if png.exists() and meta.exists():
                m = json.loads(meta.read_text(encoding="utf-8"))
                self._win_cache = {"template": Image.open(png).convert("RGB"), "origin": tuple(m["origin"]),
                                   "header": m["header"], "extras": {k: tuple(v) if v else None
                                                                     for k, v in m["extras"].items()},
                                   "pitch": m.get("pitch"), "n_rows": m.get("n_rows")}
        except (OSError, ValueError, KeyError):
            self._win_cache = None

    def _save_window_cache(self) -> None:
        """Keeps the header template so that, once it has been read ONCE, later runs follow it as an image and
        never depend on the OCR recognizing the header words again."""
        c = self._win_cache
        if not c:
            return
        try:
            ANCHOR_DIR.mkdir(exist_ok=True)
            png, meta = self._window_cache_files()
            c["template"].save(png)
            meta.write_text(json.dumps({"origin": list(c["origin"]), "header": c["header"],
                                        "extras": {k: list(v) if v else None for k, v in c["extras"].items()},
                                        "pitch": c["pitch"], "n_rows": c["n_rows"]}, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _find_window_header(self, shot: Image.Image) -> tuple[dict, dict, list | None]:
        """Where the Maturity | Minimum | Maximum | Ric header is. The first time it is recognized with OCR (whole
        screen, then in tiles). Afterwards it is followed as an IMAGE (template matching), which is fast and does
        not depend on the OCR getting every word right each time. Returns (header, extras, ocr lines or None)."""
        if self._win_cache is None:
            self._load_window_cache()
        cache = self._win_cache
        if cache:
            x, y, score = locate_template(shot, cache["template"])
            if score >= 0.85:
                dx, dy = x - cache["origin"][0], y - cache["origin"][1]
                h = cache["header"]
                header = {"y": h["y"] + dy, "maturity": h["maturity"] + dx, "minimum": h["minimum"] + dx,
                          "maximum": h["maximum"] + dx, "ric": h["ric"] + dx if h["ric"] else None}
                extras = {k: (v[0] + dx, v[1] + dy) if v else None for k, v in cache["extras"].items()}
                return header, extras, None
        lines, header = [], None
        for attempt in range(4):                   # whole screen x2, whole screen x1, then tiles x2, tiles x3
            lines = (self.ocr.read(shot, scale=2, margin=8) if attempt == 0
                     else self.ocr.read(shot, scale=1, margin=8) if attempt == 1
                     else self._ocr_tiled(shot, scale=2 if attempt == 2 else 3))
            header = find_header(lines)
            if header:
                break
        if header is None:
            shot_path = self.save_debug("update_window_header_not_found", shot)
            lines_path = self._dump_lines("update_window_ocr_lines", lines)
            raise SophisError("Could not find the Maturity | Minimum | Maximum | Ric header on screen. Is the "
                              f"update window open and fully visible? Screenshot: {shot_path}; what the OCR "
                              f"read: {lines_path}")
        hy = header["y"]
        title = next((l for l in lines if l.cy < hy and len(_norm(l.text)) > 8 and contains_word(l.text, "Strikes")),
                     None)
        cancel = next((l for l in lines if l.cy > hy and len(_norm(l.text)) <= 8
                       and _best_window(_norm(l.text), "cancel") >= 0.8), None)
        ok = next((l for l in lines if l.cy > hy and _norm(l.text) == "ok"
                   and (cancel is None or (abs(l.cy - cancel.cy) <= 12 and l.cx < cancel.cx))), None)
        extras = {"title": (title.cx, title.cy) if title else None, "ok": (ok.cx, ok.cy) if ok else None,
                  "cancel": (cancel.cx, cancel.cy) if cancel else None}
        d = header["minimum"] - header["maturity"]
        box = (max(0, int(header["maturity"] - 0.85 * d)), max(0, hy - 10),
               min(shot.width, int(header["maximum"] + 0.8 * d)), min(shot.height, hy + 10))
        self._win_cache = {"template": shot.crop(box), "origin": (box[0], box[1]), "header": header,
                           "extras": extras, "pitch": None, "n_rows": None}
        self._save_window_cache()
        return header, extras, lines

    def _read_update_window_once(self) -> UpdateWindow:
        shot = self.s.capture()
        header, extras, header_lines = self._find_window_header(shot)
        cache = self._win_cache
        hy, xm = header["y"], header["maturity"]
        d = header["minimum"] - xm
        half = 0.6 * d                                            # a bit less than half the Maturity column

        # The dates are read from a thin strip over the Maturity column (fast, and the text is big in it).
        pitch_guess = cache["pitch"] or 0.19 * d
        n_hint = (cache["n_rows"] or 14) + 2
        sx0, sy0 = int(xm - half), int(hy + pitch_guess * 0.5)
        sy1 = min(shot.height, int(hy + pitch_guess * (n_hint + 0.5)))
        strip_lines = [Line(l.text, l.conf, l.x0 + sx0, l.y0 + sy0, l.x1 + sx0, l.y1 + sy0)
                       for l in self.ocr.read(shot.crop((sx0, sy0, int(xm + half), sy1)), scale=3, margin=8)]
        found = self._dates_in(strip_lines, hy, xm, half)
        if len(found) < 2 and header_lines:                       # first look: the whole-screen OCR can help
            found = self._dates_in(header_lines, hy, xm, half)
        gaps = [b[0] - a[0] for a, b in zip(found, found[1:]) if 8 <= b[0] - a[0] <= 80]
        if not gaps:
            path = self.save_debug("update_window_no_dates", shot)
            lines_path = self._dump_lines("update_window_strip_lines", strip_lines)
            raise SophisError("Found the header but could not read at least two dates below it. "
                              f"Screenshot: {path}; what the OCR read: {lines_path}")
        small = min(gaps)
        pitch = float(statistics.median([g for g in gaps if g <= 1.4 * small]))

        by_index = {}
        for cy, text, date, suffix in found:
            idx = int(round((cy - hy) / pitch)) - 1
            by_index.setdefault(idx, WindowRow(idx, cy, text, date, suffix))
        # The detector often misses the highlighted row (white text on dark blue) and any row it merged:
        # read those row slots one by one, which handles light-on-dark text.
        for idx in range(0, max(by_index) + 2):
            if idx in by_index:
                continue
            cy = int(round(hy + (idx + 1) * pitch))
            crop = shot.crop((int(xm - half), int(cy - pitch / 2 + 1), int(xm + half), int(cy + pitch / 2 - 1)))
            readings, _ = self.ocr.read_cell(crop)
            for text in readings:
                date, suffix = parse_date(text)
                if date:
                    by_index[idx] = WindowRow(idx, cy, text, date, suffix)
                    break
        rows = [by_index[i] for i in sorted(by_index)]
        learned = cache["pitch"] is None
        cache["pitch"], cache["n_rows"] = pitch, len(rows)
        if learned:
            self._save_window_cache()
        return UpdateWindow(hy, xm, header["minimum"], header["maximum"], header["ric"], pitch, rows, shot,
                            title=extras["title"], ok=extras["ok"], cancel=extras["cancel"])

    @staticmethod
    def _dates_in(lines: list, hy: int, xm: int, half: float) -> list:
        found = []
        for l in lines:
            if l.cy > hy + 6 and abs(l.cx - xm) <= half:
                date, suffix = parse_date(l.text)
                if date:
                    found.append((l.cy, l.text, date, suffix))
        return sorted(found)

    def read_update_window(self) -> UpdateWindow:
        """Reads the open update window: the Maturity / Minimum / Maximum / Ric columns and the date of every
        visible row, with the on-screen position of each. Read-only. Once the window has been seen, a failed
        read is retried (the screen may simply have been mid-redraw) before giving up."""
        tries = 3
        for k in range(1, tries + 1):
            try:
                return self._read_update_window_once()
            except SophisError as e:
                if k == tries:
                    raise
                self.log(f"  could not read the update window (attempt {k}/{tries}); trying again")
                self._wait(1.0)
        raise AssertionError("unreachable")

    def _trace(self, step: str) -> None:
        """With trace on, saves a picture of the window at each step so a failed run can be inspected."""
        if not self.trace:
            return
        shot = self.s.capture()
        c = self._win_cache
        if c:
            h, d = c["header"], c["header"]["minimum"] - c["header"]["maturity"]
            p = c["pitch"] or 0.19 * d
            box = (max(0, int(h["maturity"] - 2.2 * d)), max(0, int(h["y"] - 6 * p)),
                   min(shot.width, int((h["ric"] or h["maximum"] + 2 * d) + 1.2 * d)),
                   min(shot.height, int(h["y"] + ((c["n_rows"] or 14) + 4) * p)))
            shot = shot.crop(box)
        self._trace_n += 1
        path = self.save_debug(f"fill{self._trace_n:02d}_{step}", shot)
        self.log(f"    [trace] {path}")

    def locate_date(self, target: dt.date, occurrence: int | None = None, max_scrolls: int = 15) -> CellTarget:
        """Finds the row whose Maturity is `target` (scrolling the list if needed) and returns where its Minimum
        and Maximum cells are. If several rows have that date, `occurrence` (0, 1, ...) picks one; without it the
        first is returned and `ambiguous` is True. Nothing is clicked or typed."""
        previous = None
        for _ in range(max_scrolls + 1):
            win = self.read_update_window()
            matches = win.find(target)
            if matches:
                win, matches = self._reveal_duplicates(win, target)
                if occurrence is not None and occurrence >= len(matches):
                    raise SophisError(f"{target:%d %b %Y} appears {len(matches)} time(s); there is no occurrence "
                                      f"{occurrence} (0 is the first).")
                row = matches[occurrence or 0]
                return CellTarget(row, (win.x_minimum, row.cy), (win.x_maximum, row.cy), matches,
                                  ambiguous=len(matches) > 1 and occurrence is None, win=win)
            dated = win.dated()
            if not dated:
                raise SophisError("No dates could be read in the update window.")
            if dated[0].date <= target <= dated[-1].date:
                raise SophisError(f"{target:%d %b %Y} is not in the list (it falls between "
                                  f"{dated[0].date:%d %b %Y} and {dated[-1].date:%d %b %Y} but is not there).")
            signature = tuple(r.date for r in dated)
            if signature == previous:
                raise SophisError(f"{target:%d %b %Y} was not found: the list stopped scrolling at "
                                  f"{dated[0].date:%d %b %Y} ... {dated[-1].date:%d %b %Y}.")
            previous = signature
            down = target > dated[-1].date
            self.log(f"  {target:%d %b %Y} not visible ({dated[0].date:%d %b %Y} ... {dated[-1].date:%d %b %Y}); "
                     f"scrolling {'down' if down else 'up'}")
            mid = win.rows[len(win.rows) // 2].cy
            self.s.scroll(-3 if down else 3, win.x_maturity, mid)
            self._wait(0.7)
        raise SophisError(f"{target:%d %b %Y} was not found after {max_scrolls} scrolls.")

    def _reveal_duplicates(self, win: UpdateWindow, target: dt.date) -> tuple[UpdateWindow, list]:
        """The same date can appear on several adjacent rows. If the date sits on the first or last visible row,
        more rows with it may be hidden just beyond, so nudge the list one wheel click at a time until the edge
        row has a different date (or the list stops moving). Returns the final view and the rows with the date."""
        for direction in (-1, +1):                                      # -1 = down, +1 = up
            for _ in range(4):
                dated = win.dated()
                if not dated:
                    break
                edge = dated[-1] if direction == -1 else dated[0]
                if edge.date != target:
                    break
                before = tuple(r.date for r in dated)
                self.s.scroll(direction, win.x_maturity, win.rows[len(win.rows) // 2].cy)
                self._wait(0.6)
                win = self.read_update_window()
                if tuple(r.date for r in win.dated()) == before:         # the list did not move: end of the list
                    break
        return win, win.find(target)

    # ── stage 2b: write the values (double click a cell, type, click outside) ─────────────────
    def _commit_point(self, win: UpdateWindow, row: WindowRow, how: str = "title") -> tuple[int, int]:
        """Where to click to leave a cell after typing ("click outside the cell"): the window's title text, which
        is not part of the grid, or failing that the Ric cell of the same row."""
        if how == "title" and win.title:
            return win.title
        return (win.x_ric or win.x_maturity, row.cy)

    def _read_cell_number(self, win: UpdateWindow, x: int, y: int, half: int) -> tuple[float | None, list[str]]:
        """What the cell at (x, y) shows now, as a number (None if unreadable), plus the raw readings."""
        h = int(win.pitch / 2) - 1
        crop = self.s.capture((x - half, y - h, x + half, y + h))
        readings, _ = self.ocr.read_cell(crop)
        for r in readings:
            value = parse_number(r)
            if value is not None:
                return value, readings
        return None, readings

    def _cell_box(self, win: UpdateWindow, x: int, y: int, half: int) -> tuple[int, int, int, int]:
        h = max(3, int(win.pitch / 2) - 2)
        return (x - half, y - h, x + half, y + h)

    def _open_cell_for_editing(self, win: UpdateWindow, row: WindowRow, column: str, x: int, half: int) -> bool:
        """Gets the Minimum / Maximum cell into edit mode and CHECKS that it is in edit mode before anything is
        typed. A double click on a row that is not selected yet only selects it (the editor never opens), and typing
        then just jumps through the list instead of filling the cell, so the order is: ONE click to select the
        cell, then a double click, then (if still not open) a slower double click, then F2. "Open" means the
        cell visibly changed (it turns into an edit box). Returns False if it never opened."""
        y = row.cy
        box = self._cell_box(win, x, y, half)
        self.s.click(x, y)
        self._wait(0.6)
        self._trace(f"{column}_1b_selected")
        base = median_color(self.s.capture(box))
        self._open_debug = f"cell color after selecting it {base}"

        def editing() -> bool:
            now = median_color(self.s.capture(box))
            self._open_debug += f"; later {now}"
            return color_distance(now, base) > 45

        if editing():                                   # (a click on an already selected cell can open it)
            return True
        for how in ("double click", "double click again, slower", "F2"):
            if how == "double click":
                self.s.double_click(x, y)
            elif how.startswith("double click again"):
                self._wait(0.8)
                self.s.double_click(x, y)
            else:
                self.s.press("f2")
            self._wait(0.6)
            self._trace(f"{column}_2_after_{how.split(',')[0].replace(' ', '_')}")
            if editing():
                self.log(f"    the cell opened for editing ({how})")
                return True
            self.log(f"    the cell did not open after: {how}")
        return False

    def write_cell(self, win: UpdateWindow, row: WindowRow, column: str, value: str, commit: str = "enter",
                   verify: bool = True, check_edit: bool = True) -> str:
        """Open the cell for editing (checked), replace its content with `value`, confirm it (Enter by default, or a
        click outside the cell), and read the cell back. Returns "ok", "unverified" (could not read the cell) or raises if it shows another value, or
        if the cell would not open (then NOTHING is typed, unless check_edit is False)."""
        x = win.x_minimum if column == "minimum" else win.x_maximum
        other = win.x_maximum if column == "minimum" else (win.x_ric or win.x_maximum + (win.x_maximum - win.x_minimum))
        half = max(8, int(0.4 * abs(other - x)))
        self.log(f"  {column} of {row.date:%d %b %Y}: opening the cell at ({x}, {row.cy}) to type {value!r}")
        self._trace(f"{column}_1_before")
        existing, _ = self._read_cell_number(win, x, row.cy, half) if verify else (None, [])
        opened = self._open_cell_for_editing(win, row, column, x, half)
        if not opened and win.title:
            # An editor may have been left open somewhere (some grids keep editing after Enter), which also hides
            # the "cell changed" signal. Clicking the window's title closes it without changing anything.
            self.log("    the cell did not open; clicking the window title (closes any editor left open) and retrying")
            self.s.click(*win.title)
            self._wait(0.7)
            self._trace(f"{column}_1c_title_clicked")
            opened = self._open_cell_for_editing(win, row, column, x, half)
        if not opened:
            if check_edit:
                path = self.save_debug(f"cell_would_not_open_{column}", self.s.capture())
                raise SophisError(
                    f"The {column} cell of {row.date:%d %b %Y} did not open for editing (tried: select + double click, "
                    "slower double click, F2; also after clicking the window title). NOTHING was typed. If your "
                    f"editor does not look different when it is open, run again with --no-edit-check. "
                    f"({getattr(self, '_open_debug', '')}) Screenshot: {path}")
            self.log("    WARNING: could not tell whether the cell opened; typing anyway (--no-edit-check)")
        if existing is not None:
            self.s.hotkey("ctrl", "a")              # the cell already held a value: replace it
            self._wait(0.15)
        self.s.type_text(str(value))
        self._wait(0.25)
        self._trace(f"{column}_3_typed")
        if commit == "enter":
            self.s.press("enter")                   # leaves the edit box, keeps the typed value
        else:
            self.s.click(*self._commit_point(win, row, commit))
        self._wait(0.7)
        self._trace(f"{column}_4_committed")
        if not verify:
            return "unverified"
        seen, readings = self._read_cell_number(win, x, row.cy, half)
        wanted = parse_number(value)
        if seen is None:
            self.log(f"    WARNING: could not read the {column} cell back (OCR gave {readings}); check it by eye")
            return "unverified"
        if abs(seen - wanted) > 1e-9:
            path = self.save_debug(f"cell_mismatch_{column}")
            raise SophisError(f"The {column} cell of {row.date:%d %b %Y} shows {seen} but {wanted} was typed "
                              f"(OCR readings: {readings}). Stopped before touching anything else. Screenshot: {path}")
        self.log(f"    {column} reads back as {seen:g}")
        return "ok"

    def _reuse_row(self, previous: CellTarget, view) -> CellTarget | None:
        """Same row, next cell: one quick read is enough IF the list looks exactly as before. Enter may have moved
        the selection or scrolled the list by a row (then the rows are shifted); in that case, or if the quick read
        fails for any reason, it returns None and the caller searches the date again."""
        try:
            win = self.read_update_window()
        except SophisError as e:
            self.log(f"  (quick re-read failed: {e}; searching the date again)")
            return None
        if view(win) != view(previous.win):
            return None
        return CellTarget(previous.row, (win.x_minimum, previous.row.cy), (win.x_maximum, previous.row.cy),
                          previous.matches, ambiguous=False, win=win)

    def fill_values(self, entries: list, verify: bool = True, dry_run: bool = False, commit: str = "enter",
                    check_edit: bool = True) -> list:
        """Writes Minimum and Maximum for each entry in the open update window. The row is found again before
        every cell (the list may have scrolled). A date that appears more than once needs its `occurrence`;
        nothing is guessed. dry_run only locates the cells. It NEVER presses OK: the window is left open so you
        can check the values and confirm it yourself."""
        done = []
        # 1) Check everything that can be checked BEFORE changing anything.
        plan = []
        for e in entries:
            t = self.locate_date(e.date, e.occurrence)
            if t.ambiguous:
                raise SophisError(
                    f"{e.date:%d %b %Y} appears {len(t.matches)} times (rows {[m.index for m in t.matches]}, "
                    f"read as {[m.text for m in t.matches]}). Choose one by writing the date as "
                    f"'{e.date:%d %b %Y}#0' (the first) or '#1' (the second). Nothing was changed.")
            plan.append((e, t))
            self.log(f"plan: {e.date:%d %b %Y} -> row {t.row.index}; Minimum {e.minimum} at {t.minimum}, "
                     f"Maximum {e.maximum} at {t.maximum}")
        if dry_run:
            return [{"date": e.date, "row": t.row.index, "minimum": t.minimum, "maximum": t.maximum,
                     "written": False} for e, t in plan]
        # 2) Write. The row is found again before every cell (the list may have scrolled), with a fresh read.
        def view(win):
            return (tuple((r.index, r.date, r.suffix) for r in win.rows), round(win.header_y), round(win.x_maturity))

        for e, _ in plan:
            result = {"date": e.date, "minimum": None, "maximum": None, "written": True}
            previous = None                                              # the target of the cell written just before
            for column, value in (("minimum", e.minimum), ("maximum", e.maximum)):
                if value is None:
                    continue
                t = self._reuse_row(previous, view) if previous is not None else None
                if t is None:
                    t = self.locate_date(e.date, e.occurrence)
                    if t.ambiguous:
                        raise SophisError(f"{e.date:%d %b %Y} became ambiguous while writing; stopped.")
                result[column] = self.write_cell(t.win, t.row, column, value, commit, verify, check_edit)
                previous = t
            done.append(result)
        self.log("done. The window was left open: check the values and press OK yourself.")
        return done

    def annotate_window(self, win: UpdateWindow, target: CellTarget | None = None) -> Path:
        """Saves a screenshot with what was understood drawn on it, so you can check it by eye."""
        from PIL import ImageDraw
        img = (win.image or self.s.capture()).copy()
        d = ImageDraw.Draw(img)
        half = int(0.55 * (win.x_minimum - win.x_maturity))
        for x, name in ((win.x_maturity, "Maturity"), (win.x_minimum, "Minimum"), (win.x_maximum, "Maximum"),
                        (win.x_ric, "Ric")):
            if x:
                d.line((x, win.header_y - 12, x, win.header_y + 12), fill=(255, 0, 0), width=2)
                d.text((x - 20, win.header_y - 28), name, fill=(255, 0, 0))
        for r in win.rows:
            d.rectangle((win.x_maturity - half, int(r.cy - win.pitch / 2), (win.x_ric or win.x_maximum + 40),
                         int(r.cy + win.pitch / 2)), outline=(0, 160, 0))
            d.text(((win.x_ric or win.x_maximum + 40) + 6, r.cy - 6), f"{r.index}: {r.date:%d %b %Y}{r.suffix}" if r.date
                   else f"{r.index}: ?", fill=(0, 120, 0))
        if target:
            for (x, y), name in ((target.minimum, "MIN"), (target.maximum, "MAX")):
                d.ellipse((x - 7, y - 7, x + 7, y + 7), outline=(255, 0, 0), width=3)
                d.text((x + 10, y - 6), name, fill=(255, 0, 0))
        return self.save_debug("update_window_map", img)

    def report_update_window(self) -> None:
        """After opening the window: log what was understood. Never raises (it is only informative)."""
        try:
            win = self.read_update_window()
            dated = win.dated()
            if dated:
                self.log(f"  update window read: {len(dated)} visible dates, {dated[0].date:%d %b %Y} ... "
                         f"{dated[-1].date:%d %b %Y}; Minimum x={win.x_minimum}, Maximum x={win.x_maximum}")
            self.log(f"  map saved in {self.annotate_window(win)}")
        except SophisError as e:
            self.log(f"  (could not read the update window yet: {e})")


# ════════════════════════════════════════════════════════════════════════════
# COMMANDS
# ════════════════════════════════════════════════════════════════════════════

def cmd_calibrate(args) -> None:
    screen = Screen()
    names = [n for n, _ in CALIBRATION_POINTS]
    only = [LEGACY_POINT_NAMES.get(n, n) for n in (args.only or [])]
    if only:
        unknown = [n for n in only if n not in names]
        if unknown:
            raise SophisError(f"Unknown points: {unknown}. Valid: {names}")
        for anchor in ("grid", "right"):
            if f"anchor_{anchor}_b" in only and f"anchor_{anchor}_a" not in only:
                raise SophisError(f"To redo anchor_{anchor}_b also include anchor_{anchor}_a (both corners).")
        # Start from the saved calibration, moved to where the window is NOW (it may have moved).
        base = Config.load(strict=False)
        current, info = compute_points(screen.capture(), base)
        points = {**base.points, **current}
        anchors = {n: {"file": a["file"], "origin": [a["origin"][0] + info[n][0], a["origin"][1] + info[n][1]]}
                   for n, a in base.anchors.items()}
        todo = [(n, t) for n, t in CALIBRATION_POINTS if n in only]
        previous = base
    else:
        points, anchors, todo = {}, {}, CALIBRATION_POINTS
        try:
            previous = Config.load(strict=False)
        except SophisError:
            previous = None
    print("CALIBRATION: for each point, put the mouse where indicated and DO NOT move it.")
    print(f"The position is captured automatically when the {args.wait}s countdown ends.")
    print("The window may be anywhere, but it must be fully visible.\n")
    for k, (name, text) in enumerate(todo, 1):
        print(f"[{k}/{len(todo)}] {name}: {text}")
        for sec in range(args.wait, 0, -1):
            print(f"   {sec}...", end="\r", flush=True)
            time.sleep(1)
        x, y = screen.position()
        points[name] = [x, y]
        print(f"   {name} = ({x}, {y})        ")
        if name in ANCHOR_CLOSING_POINT:
            anchor = ANCHOR_CLOSING_POINT[name]
            ax, ay = points[f"anchor_{anchor}_a"]
            x0, y0, x1, y1 = min(ax, x), min(ay, y), max(ax, x), max(ay, y)
            if x1 - x0 < 40 or y1 - y0 < 8:
                raise SophisError(f"Anchor '{anchor}' is too small ({x1 - x0}x{y1 - y0}px). "
                                  "Mark both corners of the header properly and run 'calibrate' again.")
            ANCHOR_DIR.mkdir(exist_ok=True)
            path = ANCHOR_DIR / f"{anchor}.png"
            screen.capture((x0, y0, x1, y1)).save(path)
            anchors[anchor] = {"file": str(path), "origin": [x0, y0]}
    cfg = Config(points=points, anchors=anchors)
    if previous:
        cfg.widen_px, cfg.key_interval = previous.widen_px, previous.key_interval
        cfg.slowdown, cfg.reader, cfg.verify_mnemonic = previous.slowdown, previous.reader, previous.verify_mnemonic
    # Validate before saving: if something is wrong, the previous calibration is not overwritten.
    Automator(screen, None, cfg, log=lambda *_: None).locate()
    cfg.save()
    _, info = compute_points(screen.capture(), cfg)
    for name, (dx, dy, score) in info.items():
        print(f"Anchor '{name}': match {score:.2f}, shift ({dx:+d},{dy:+d})  <- should be ~1.00 and (0,0)")
    print(f"\nSaved to {CONFIG_FILE} and {ANCHOR_DIR}/. You can now move the window around.")
    print("Next:  python sophis_auto.py diagnose <CODE>")


def _create(args) -> Automator:
    cfg = Config.load()
    if getattr(args, "reader", None):
        cfg.reader = LEGACY_READER_VALUES.get(args.reader, args.reader)
    if getattr(args, "no_widen", False):
        cfg.widen_px = 0
    if getattr(args, "slowdown", None):
        cfg.slowdown = args.slowdown
    if getattr(args, "auto_focus", False):
        cfg.auto_focus = True
    if getattr(args, "open_with", None):
        cfg.open_method = args.open_with
    print("Loading OCR ...")
    return Automator(Screen(cfg.key_interval), OCR(), cfg)


def ask_for_click(a: Automator, args, first: bool = True, typing: bool = False) -> None:
    """You give the grid the selection with ONE click; the script continues from there."""
    if typing and a.cfg.auto_focus:
        print("You have 3 seconds to bring Sophis (Citrix) to the front ...")
        time.sleep(3)
        return
    wait = max(1, int(getattr(args, "click_wait", 6)))
    what = "Typing starts" if typing else "The script continues"
    print(f"\nNow, in Sophis: SELECT THE ROW of the code you want (one click on it) and then DO NOT touch "
          f"anything, not even the mouse.\n{what} in {wait} seconds ...")
    for sec in range(wait, 0, -1):
        print(f"   {sec}...", end="\r", flush=True)
        time.sleep(1)
    print("   going         ")


def cmd_diagnose_selected(args) -> None:
    """Shows what the script sees for the row YOU selected (does not type or open anything)."""
    a = _create(args)
    ask_for_click(a, args)
    DEBUG_DIR.mkdir(exist_ok=True)
    a.save_debug("diag_screen")
    a.prepare()
    mouse = a.row_under_mouse()
    color = a.selected_row()
    print(f"\nRow under the mouse: {mouse} | highlighted row: {color}  (None = not detected)")
    row = color if color is not None else mouse
    if row is None:
        print("Could not tell the row. Click it once and leave the mouse on it.")
        return
    g = a.g
    raw = a.s.capture(a._cell_region(row, g.x_mnemo0, g.x_mnemo1))
    a.save_debug("diag_mnemo", raw)
    a.save_debug("diag_mnemo_x6", raw.resize((raw.width * 6, raw.height * 6), Image.NEAREST))
    a.save_debug("diag_term", a.s.capture(a._cell_region(row, g.term_x0, g.term_x1)))
    a.save_debug("diag_mnemo_cleaned", Image.fromarray(clean_cell(raw)[0]))
    readings, glyphs = a._mnemonic_by_ocr(row)
    print("Mnemo crop size          :", raw.size, "| background:", median_color(raw))
    print("OCR Mnemo readings       :", readings, "| glyphs:", glyphs)
    print("OCR Term                 :", a._term_by_ocr(row))
    print("Clipboard (Ctrl+C/Insert):", repr(a._read_clipboard()))
    print("Dropdown (OCR)           :", a.read_dropdown())
    print(f"\nScreenshots saved in {DEBUG_DIR}/ (screen, mnemo, mnemo_x6, mnemo_cleaned, term).")


def cmd_diagnose(args) -> None:
    """Shows what the script sees; useful to tune the calibration (does not open any window)."""
    if not args.code:
        return cmd_diagnose_selected(args)
    a = _create(args)
    code = args.code.upper()
    ask_for_click(a, args, typing=True)
    DEBUG_DIR.mkdir(exist_ok=True)
    a.save_debug("diag_screen")
    a.prepare()
    start = a._type_code(code)
    print(f"Row selected before typing: {start} (None = not detected). The script typed without any check.")
    a._wait(0.8)
    i = a.selected_row()
    print(f"\nSelected row after typing '{code}': {i}")
    if i is not None:
        g = a.g
        raw = a.s.capture(a._cell_region(i, g.x_mnemo0, g.x_mnemo1))
        a.save_debug("diag_mnemo", raw)
        a.save_debug("diag_mnemo_x6", raw.resize((raw.width * 6, raw.height * 6), Image.NEAREST))
        a.save_debug("diag_term", a.s.capture(a._cell_region(i, g.term_x0, g.term_x1)))
        cleaned, _ = clean_cell(raw)
        a.save_debug("diag_mnemo_cleaned", Image.fromarray(cleaned))
        print("Mnemo crop size          :", raw.size, "| background color:", median_color(raw),
              "| ext color:", median_color(a.s.capture((g.x_ext - 3, g.row_y(i) - 3, g.x_ext + 4, g.row_y(i) + 4))))
        print("Edit mode on this row    :", a._row_in_edit_mode(i))
        readings, glyphs = a._mnemonic_by_ocr(i)
        try:
            from importlib.metadata import version
            print("rapidocr version         :", version("rapidocr-onnxruntime"))
        except Exception:
            pass
        print("Clipboard (Ctrl+C/Insert):", repr(a._read_clipboard()))
        print("OCR Mnemo readings       :", readings, "| glyphs:", glyphs)
        print("Mnemo match              :", match_code(readings, glyphs, code))
        print("OCR Term                 :", a._term_by_ocr(i))
    print("Dropdown (OCR)           :", a.read_dropdown())
    print(f"\nScreenshots saved in {DEBUG_DIR}/ (screen, mnemo, term).")


def cmd_open(args) -> None:
    """You select the row; the script reads its Term, sets the dropdown, opens the window. Repeats on request."""
    if args.codes:
        print("`open` no longer takes codes: you select each row yourself. To search by typing use "
              "`open-code " + " ".join(args.codes) + "`.")
        sys.exit(2)
    term = None
    if args.term:
        matches = [t for t in TERMS if _norm(t) == _norm(args.term)]
        if not matches:
            print(f"ERROR: --term must be one of: {', '.join(TERMS)}")
            sys.exit(2)
        term = matches[0]
    a = _create(args)
    while True:
        ask_for_click(a, args)
        try:
            a.open_selected(term)
        except SophisError as e:
            print(f"\nERROR: {e}")
            sys.exit(1)
        answer = input("\nWhen you are done with the update window (close it / OK), press ENTER to do another "
                       "code (q = quit): ")
        if answer.strip().lower() == "q":
            break


def cmd_open_code(args) -> None:
    """Older flow: types each code to find its row (not reliable on this grid yet)."""
    a = _create(args)
    for k, code in enumerate(args.codes):
        ask_for_click(a, args, first=(k == 0), typing=True)
        try:
            a.open_contract(code)
        except SophisError as e:
            print(f"\nERROR: {e}")
            sys.exit(1)
        if k < len(args.codes) - 1:
            answer = input("\nWhen you are done with this window (close it / OK), press ENTER for the next one (q = quit): ")
            if answer.strip().lower() == "q":
                break


def cmd_window(args) -> None:
    """Reads the open update window and shows where each date's Minimum / Maximum cells are."""
    a = _create(args)
    _countdown_to_front(args)
    try:
        win = a.read_update_window()
        print(f"\nColumns (x): Maturity {win.x_maturity} | Minimum {win.x_minimum} | Maximum {win.x_maximum} | "
              f"Ric {win.x_ric}; header y {win.header_y}; row pitch {win.pitch:.1f}px")
        for r in win.rows:
            print(f"  row {r.index:2d}  y={r.cy}  {r.text!r:22} -> "
                  f"{r.date.strftime('%d %b %Y') if r.date else '?'}{r.suffix}")
        target = None
        if args.date:
            date, _ = parse_date(args.date)
            if not date:
                print(f"ERROR: could not understand --date {args.date!r} (try 19 Oct 2026, 2026-10-19 or 19/10/2026)")
                sys.exit(2)
            target = a.locate_date(date, args.occurrence)
            if len(target.matches) > 1:
                print(f"\n{date:%d %b %Y} appears {len(target.matches)} times: rows "
                      f"{[(m.index, m.text) for m in target.matches]}"
                      f"{'  -> pick one with --occurrence 0/1/...' if target.ambiguous else ''}")
            print(f"\n{date:%d %b %Y}: row {target.row.index}; Minimum cell at {target.minimum}, "
                  f"Maximum cell at {target.maximum}")
            win = a.read_update_window()                 # (the list may have scrolled) for the picture
        print("Map saved in", a.annotate_window(win, target if args.date else None))
        if target and args.show:
            for name, xy in (("Minimum", target.minimum), ("Maximum", target.maximum)):
                print(f"Moving the pointer to the {name} cell (no click) ...")
                a.s.pg.moveTo(*xy, duration=0.4)
                time.sleep(2)
    except SophisError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)


def _countdown_to_front(args) -> None:
    """The screenshot is of what is on top: if the terminal covers the Sophis window, nothing can be read."""
    wait = max(1, int(getattr(args, "click_wait", 5)))
    print(f"\nBring the Sophis update window (Maturity / Minimum / Maximum) to the FRONT, fully visible and NOT "
          f"covered by this terminal. Starting in {wait} seconds ...")
    for sec in range(wait, 0, -1):
        print(f"   {sec}...", end="\r", flush=True)
        time.sleep(1)
    print("   going         ")


def cmd_fill(args) -> None:
    """Writes Minimum / Maximum for the given dates in the open update window."""
    try:
        entries = [parse_entry(d, mn, mx) for d, mn, mx in (args.entry or [])]
        if args.occurrence is not None:
            for e in entries:
                e.occurrence = e.occurrence if e.occurrence is not None else args.occurrence
    except SophisError as e:
        print(f"ERROR: {e}")
        sys.exit(2)
    if not entries:
        print("ERROR: give at least one  --entry DATE MIN MAX   (e.g. --entry \"20 Nov 2026#1\" 100 120)")
        sys.exit(2)
    a = _create(args)
    a.trace = not args.no_trace
    _countdown_to_front(args)
    try:
        results = a.fill_values(entries, verify=not args.no_verify, dry_run=args.dry_run, commit=args.commit_click,
                                check_edit=not args.no_edit_check)
    except SophisError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)
    print("\nDry run: nothing was clicked or typed." if args.dry_run else "\nFinished.")
    for r in results:
        print(" ", {k: (v.strftime('%d %b %Y') if hasattr(v, 'strftime') else v) for k, v in r.items()})


def main() -> None:
    ap = argparse.ArgumentParser(description="Sophis automation (stage 1: search a code and open its window)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("fill", help="write Minimum / Maximum values in the open update window")
    s.add_argument("--entry", nargs=3, action="append", metavar=("DATE", "MIN", "MAX"),
                   help="a date and its two values; repeat for more dates. A repeated date: '20 Nov 2026#1'")
    s.add_argument("--occurrence", type=int, default=None,
                   help="row to use for ALL entries whose date appears more than once (or write DATE#n)")
    s.add_argument("--dry-run", dest="dry_run", action="store_true", help="only locate the cells; click/type nothing")
    s.add_argument("--no-verify", dest="no_verify", action="store_true", help="do not read each cell back")
    s.add_argument("--no-edit-check", dest="no_edit_check", action="store_true",
                   help="type even if the cell did not visibly open for editing (default: stop and type nothing)")
    s.add_argument("--no-trace", dest="no_trace", action="store_true",
                   help="do not save a screenshot at every step (they go to sophis_debug/ by default)")
    s.add_argument("--commit-click", dest="commit_click", choices=["enter", "title", "ric"], default="enter",
                   help="how to confirm a typed value: press Enter (default), or click outside the cell: on the "
                        "window's title or on the Ric cell of the row")
    s.add_argument("--reader", default=None, help=argparse.SUPPRESS)
    s.add_argument("--no-widen", dest="no_widen", action="store_true", help=argparse.SUPPRESS)
    s.add_argument("--slowdown", dest="slowdown", type=float, help="multiplier for waits (2 = twice as slow)")
    s.add_argument("--click-wait", dest="click_wait", type=int, default=5,
                   help="seconds you get to bring the update window to the front (default 5)")
    s.set_defaults(fn=cmd_fill)

    s = sub.add_parser("window", help="read the open update window: dates and where Minimum / Maximum go")
    s.add_argument("--date", help="a date to locate, e.g. '19 Oct 2026', 2026-10-19 or 19/10/2026")
    s.add_argument("--occurrence", type=int, default=None,
                   help="which row to use if the date appears more than once (0 = first)")
    s.add_argument("--show", action="store_true", help="move the pointer over the Minimum and Maximum cells (no click)")
    s.add_argument("--reader", default=None, help=argparse.SUPPRESS)
    s.add_argument("--no-widen", dest="no_widen", action="store_true", help=argparse.SUPPRESS)
    s.add_argument("--slowdown", dest="slowdown", type=float, help="multiplier for waits (2 = twice as slow)")
    s.add_argument("--click-wait", dest="click_wait", type=int, default=5,
                   help="seconds you get to bring the update window to the front (default 5)")
    s.set_defaults(fn=cmd_window)

    s = sub.add_parser("calibrate", aliases=["calibrar"], help="record the on-screen positions (once)")
    s.add_argument("--wait", "--espera", dest="wait", type=int, default=5, help="seconds per point (default 5)")
    s.add_argument("--only", "--solo", dest="only", nargs="+", metavar="POINT",
                   help="redo only these points (e.g. --only row1 row2 rowN), keeping the rest")
    s.set_defaults(fn=cmd_calibrate)

    for name, aliases, fn, help_text in (
            ("diagnose", ["diagnostico"], cmd_diagnose,
             "check how the screen is read for the row you select (give a CODE to test typing instead)"),
            ("open", ["abrir"], cmd_open,
             "you select the row; it reads the Term, matches the dropdown and opens the update window"),
            ("open-code", [], cmd_open_code, "older flow: type the code(s) to find the row (not reliable yet)")):
        s = sub.add_parser(name, aliases=aliases, help=help_text)
        if name == "diagnose":
            s.add_argument("code", nargs="?", default=None)
        elif name == "open":
            s.add_argument("codes", nargs="*")
            s.add_argument("--term", help=f"the row's Term, so it is not read from the screen ({', '.join(TERMS)})")
        else:
            s.add_argument("codes", nargs="+")
        s.add_argument("--reader", "--lector", dest="reader", choices=["auto", "clipboard", "ocr", "portapapeles"],
                       help="how to read the row")
        s.add_argument("--no-widen", "--no-ensanchar", dest="no_widen", action="store_true",
                       help="do not touch the Term column width")
        s.add_argument("--slowdown", "--pausa", dest="slowdown", type=float,
                       help="multiplier for waits (2 = twice as slow)")
        s.add_argument("--click-wait", dest="click_wait", type=int, default=6,
                       help="seconds you get to select the row in Sophis before the script continues (default 6)")
        s.add_argument("--open-with", dest="open_with", choices=["button", "double"],
                       help="how to open the update window: the button next to 'Available contracts' (default) "
                            "or a double click on the row")
        s.add_argument("--auto-focus", dest="auto_focus", action="store_true",
                       help="(open-code only) let the script click an unselected row itself")
        s.set_defaults(fn=fn)

    args = ap.parse_args()
    try:
        args.fn(args)
    except SophisError as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

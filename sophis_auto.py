#!/usr/bin/env python3
"""
Sophis Risque automation (runs over Citrix) -- STAGE 1:
  search a code (Mnemo) -> read its Term -> set the dropdown to match -> open the update window
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
    3) python sophis_auto.py diagnose ZYS       # checks how your screen is read (send me what it prints)
    4) python sophis_auto.py open ZYS ZTS YY    # searches, matches the dropdown, opens each code's window

Old Spanish command names (calibrar / diagnostico / abrir) still work as aliases.

Safety: move the mouse to a screen corner to ABORT at any time (pyautogui failsafe).
If anything does not match, the script stops and saves screenshots in ./sophis_debug (it never guesses).
"""

from __future__ import annotations

import argparse
import difflib
import json
import re
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
    ("term_left", "Mouse on the HEADER, exactly on the line that separates RIC | Term (left border of Term)"),
    ("term_right", "Mouse on the HEADER, on the right border of the Term column (before 'Cyc')"),
    ("anchor_right_a", "ANCHOR 2 (TOP-LEFT corner): header of the right-hand table 'Date | Cycle | Name'"),
    ("anchor_right_b", "ANCHOR 2 (BOTTOM-RIGHT corner): bottom-right corner of that header (after 'Name')"),
    ("dd_left", "Mouse INSIDE the dropdown ('Short term index'), at the start of the text, vertically centered"),
    ("dd_arrow", "Mouse on the dropdown button / arrow"),
]
# The corner that closes each anchor (its image crop is saved when it is captured).
ANCHOR_CLOSING_POINT = {"anchor_grid_b": "grid", "anchor_right_b": "right"}
# Which anchor each point belongs to (points move together with their anchor).
POINT_ANCHOR = {"row1": "grid", "row2": "grid", "rowN": "grid", "ext": "grid", "term_left": "grid",
                "term_right": "grid", "dd_left": "right", "dd_arrow": "right"}

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
    def load(path: Path = CONFIG_FILE) -> "Config":
        if not path.exists():
            raise SophisError(f"{path} does not exist. Run first:  python sophis_auto.py calibrate")
        d = Config._migrate(json.loads(path.read_text(encoding="utf-8")))
        missing = [n for n, _ in CALIBRATION_POINTS if n not in d.get("points", {})]
        if missing:
            raise SophisError(f"Incomplete calibration (missing {missing}). Run:  python sophis_auto.py calibrate")
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
    if t.shape[0] > a.shape[0] or t.shape[1] > a.shape[1]:
        return 0, 0, 0.0
    res = cv2.matchTemplate(a, t, cv2.TM_CCOEFF_NORMED)
    _, best, _, pos = cv2.minMaxLoc(res)
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
        self._widened = False

    # ── helpers ─────────────────────────────────────────────────────────────
    def _wait(self, seconds: float) -> None:
        time.sleep(seconds * self.cfg.slowdown)

    def _pt(self, name: str) -> tuple[int, int]:
        x, y = self.pts[name]
        return int(x), int(y)

    def locate(self) -> None:
        """Finds the window on screen (it may be somewhere else than at calibration) and recomputes the points."""
        self.pts, info = compute_points(self.s.capture(), self.cfg)
        self.g = self._geometry()
        moved = ", ".join(f"{n}: shifted ({dx:+d},{dy:+d}) match {sc:.2f}" for n, (dx, dy, sc) in info.items())
        self.log(f"  window located -> {moved}")

    def _geometry(self) -> Geometry:
        r1, r2, rn = self._pt("row1"), self._pt("row2"), self._pt("rowN")
        ext, tl, tr = self._pt("ext"), self._pt("term_left"), self._pt("term_right")
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
        widen = max(0, self.cfg.widen_px)
        return Geometry(
            x_click=r1[0], x_mnemo0=r1[0] - 30, x_mnemo1=max(r1[0] + 12, ext[0] - 2), x_ext=ext[0],
            y1=r1[1], h=float(h), n=n,
            term_x0=tl[0] - widen - 6, term_x1=tr[0],
        )

    def save_debug(self, name: str, img: Image.Image | None = None) -> Path:
        DEBUG_DIR.mkdir(exist_ok=True)
        path = DEBUG_DIR / f"{time.strftime('%H%M%S')}_{name}.png"
        (img or self.s.capture()).save(path)
        return path

    # ── preparation ─────────────────────────────────────────────────────────
    def widen_term(self) -> None:
        """The Term column shows 'Short ter...': drag its left border to reveal the whole text."""
        if self._widened or self.cfg.widen_px <= 0:
            return
        bx, by = self._pt("term_left")
        self.log(f"Widening the Term column by {self.cfg.widen_px}px ...")
        self.s.drag(bx, by, bx - self.cfg.widen_px, by)
        self._widened = True
        self._wait(0.6)

    def prepare(self) -> None:
        self.locate()

    def focus_first_row(self) -> None:
        """The ONE click made before typing: focuses the grid and resets the incremental-search buffer.
        Nothing else happens between this click and typing, so typing can never be skipped."""
        self.click_row(0)
        self._wait(0.4)

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
        sentinel = f"__sophis_{time.time()}__"
        self.s.clipboard_set(sentinel)
        self.s.hotkey("ctrl", "c")
        self._wait(0.4)
        text = self.s.clipboard_get()
        if not text or text == sentinel:
            return None
        return text

    def _mnemonic_by_ocr(self, i: int, code: str) -> tuple[bool, list[str]]:
        region = self._cell_region(i, self.g.x_mnemo0, self.g.x_mnemo1)
        img = self.s.capture(region)
        seen = []
        for scale in (3, 4, 2):
            text = self.ocr.text(img, scale)
            seen.append(text)
            if _norm_code(text) == _norm_code(code):
                return True, seen
        return False, seen

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

    def read_row(self, i: int, code: str) -> Reading:
        mode = self.cfg.reader
        fallback = None
        if mode in ("auto", "clipboard"):
            text = self._read_clipboard()
            if text:
                cells = [c.strip() for c in re.split(r"[\t\r\n]+", text) if c.strip()]
                if any(_norm_code(c) == _norm_code(code) for c in cells[:3]):
                    return Reading(True, classify_term(text), "clipboard", text, term_text=text)
                if cells:
                    fallback = Reading(False, classify_term(text), "clipboard", text, seen=(cells[0],),
                                       term_text=text)
        if mode in ("auto", "ocr"):
            ok, seen = self._mnemonic_by_ocr(i, code)
            term, term_text = self._term_by_ocr(i)
            text = f"mnemonic={seen} term='{term_text}'"
            if ok:
                return Reading(True, term, "ocr", text, seen=tuple(seen), term_text=term_text)
            if not self.cfg.verify_mnemonic and term:
                self.log("  WARNING: Mnemo NOT verified (verify_mnemonic=False)")
                return Reading(True, term, "ocr", text, verified=False, term_text=term_text)
            return Reading(False, term, "ocr", text, seen=tuple(seen), term_text=term_text)
        return fallback or Reading(False, None, mode, "")

    # ── search ──────────────────────────────────────────────────────────────
    @staticmethod
    def _is_extension(reading: Reading, code: str) -> bool:
        """True if the row read starts with the code but is longer (e.g. you search 'ZM' and land on 'ZMH')."""
        c = _norm_code(code)
        return any(_norm_code(t).startswith(c) and len(_norm_code(t)) > len(c) for t in reading.seen)

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
        last = ""
        for k in range(1, attempts + 1):
            self.focus_first_row()
            self.s.type_text(code)
            self._wait(0.8)
            i = self.selected_row()
            if i is None:
                last = "no selected row detected"
                self.log(f"  attempt {k}: {last}")
                continue
            reading = self.read_row(i, code)
            if not reading.mnemo_ok and self._is_extension(reading, code):
                self.log(f"  landed on {reading.seen[0]!r}; walking down to '{code}' ...")
                found = self._walk_down_to(code)
                if found:
                    i, reading = found
            if reading.mnemo_ok:
                self.log(f"  row {i} ok ({reading.source}): term={reading.term}")
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
            self.log("  the list changed after the dropdown; searching again")
            j, _ = self.search(code)
        self.log(f"  opening window (double click on row {j}) ...")
        self.click_row(j, double=True)
        window = self.wait_for_window()
        self.log(f"[{code}] done: term='{reading.term}', update window is open")
        return Result(code, reading.term, j, window)


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
        base = Config.load()
        current, info = compute_points(screen.capture(), base)
        points = {**base.points, **current}
        anchors = {n: {"file": a["file"], "origin": [a["origin"][0] + info[n][0], a["origin"][1] + info[n][1]]}
                   for n, a in base.anchors.items()}
        todo = [(n, t) for n, t in CALIBRATION_POINTS if n in only]
        previous = base
    else:
        points, anchors, todo = {}, {}, CALIBRATION_POINTS
        try:
            previous = Config.load()
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
    print("Loading OCR ...")
    return Automator(Screen(cfg.key_interval), OCR(), cfg)


def cmd_diagnose(args) -> None:
    """Shows what the script sees; useful to tune the calibration (does not open any window)."""
    a = _create(args)
    code = args.code.upper()
    print("You have 3 seconds to bring Sophis (Citrix) to the front ...")
    time.sleep(3)
    DEBUG_DIR.mkdir(exist_ok=True)
    a.save_debug("diag_screen")
    a.prepare()
    a.focus_first_row()
    a.s.type_text(code)
    a._wait(0.8)
    i = a.selected_row()
    print(f"\nSelected row after typing '{code}': {i}")
    if i is not None:
        g = a.g
        a.save_debug("diag_mnemo", a.s.capture(a._cell_region(i, g.x_mnemo0, g.x_mnemo1)))
        a.save_debug("diag_term", a.s.capture(a._cell_region(i, g.term_x0, g.term_x1)))
        print("Clipboard after Ctrl+C   :", repr(a._read_clipboard()))
        print("OCR Mnemo (ok, readings) :", a._mnemonic_by_ocr(i, code))
        print("OCR Term                 :", a._term_by_ocr(i))
    print("Dropdown (OCR)           :", a.read_dropdown())
    print(f"\nScreenshots saved in {DEBUG_DIR}/ (screen, mnemo, term).")


def cmd_open(args) -> None:
    a = _create(args)
    print("You have 3 seconds to bring Sophis (Citrix) to the front ...")
    time.sleep(3)
    a.prepare()
    for k, code in enumerate(args.codes):
        try:
            a.open_contract(code)
        except SophisError as e:
            print(f"\nERROR: {e}")
            sys.exit(1)
        if k < len(args.codes) - 1:
            answer = input("\nWhen you are done with this window (close it / OK), press ENTER for the next one (q = quit): ")
            if answer.strip().lower() == "q":
                break


def main() -> None:
    ap = argparse.ArgumentParser(description="Sophis automation (stage 1: search a code and open its window)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("calibrate", aliases=["calibrar"], help="record the on-screen positions (once)")
    s.add_argument("--wait", "--espera", dest="wait", type=int, default=5, help="seconds per point (default 5)")
    s.add_argument("--only", "--solo", dest="only", nargs="+", metavar="POINT",
                   help="redo only these points (e.g. --only row1 row2 rowN), keeping the rest")
    s.set_defaults(fn=cmd_calibrate)

    for name, aliases, fn, help_text in (
            ("diagnose", ["diagnostico"], cmd_diagnose, "check how the screen is read"),
            ("open", ["abrir"], cmd_open, "search the code, match the dropdown and open the window")):
        s = sub.add_parser(name, aliases=aliases, help=help_text)
        if name == "diagnose":
            s.add_argument("code")
        else:
            s.add_argument("codes", nargs="+")
        s.add_argument("--reader", "--lector", dest="reader", choices=["auto", "clipboard", "ocr", "portapapeles"],
                       help="how to read the row")
        s.add_argument("--no-widen", "--no-ensanchar", dest="no_widen", action="store_true",
                       help="do not touch the Term column width")
        s.add_argument("--slowdown", "--pausa", dest="slowdown", type=float,
                       help="multiplier for waits (2 = twice as slow)")
        s.set_defaults(fn=fn)

    args = ap.parse_args()
    try:
        args.fn(args)
    except SophisError as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

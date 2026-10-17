"""
Tests for sophis_auto.py WITHOUT Sophis or a real screen: a "fake screen" draws an imitation of the
window (grid, dropdown, update window) with PIL and the real OCR reads it.

    pip install pytest rapidocr-onnxruntime pillow opencv-python
    pytest -q test_sophis_auto.py
"""

import json
import sys

import pytest
from PIL import Image, ImageDraw, ImageFont

import sophis_auto as sa
from sophis_auto import Automator, Config, OCR, SophisError, TERMS

# ── pure-function tests ─────────────────────────────────────────────────────

@pytest.mark.parametrize("text,expected", [
    ("Short term index", "Short term index"),
    ("ShortTermIndex", "Short term index"),            # OCR joins words
    ("short term lndex", "Short term index"),          # I -> l
    ("Long Term Shares", "Long term share"),           # plural
    ("  LONG TERM INDEX ", "Long term index"),
    ("Future", "Future"),
    ("Short term share", "Short term share"),
    ("ZYS.N Short term share", "Short term share"),    # extra text in the same crop
])
def test_classify_term_ok(text, expected):
    assert sa.classify_term(text) == expected


@pytest.mark.parametrize("text", ["", "Short ter...", "Short ter", "Cycle", "US92189H6071"])
def test_classify_term_ambiguous(text):
    assert sa.classify_term(text) is None


def test_term_truncated():
    assert sa.term_truncated("Short ter...")
    assert not sa.term_truncated("Short term index")


def test_the_five_terms_are_told_apart():
    for t in TERMS:
        assert sa.classify_term(t) == t


@pytest.mark.parametrize("selected", [0, 1, 4, 7])
def test_find_selected_row_needs_no_reference_color(selected):
    stripes = [(255, 255, 255), (235, 235, 235)]
    colors = [stripes[i % 2] for i in range(8)]
    colors[selected] = (170, 195, 235)
    assert sa.find_selected_row(colors) == selected


def test_find_selected_row_none_or_ambiguous():
    stripes = [(255, 255, 255), (235, 235, 235)]
    plain = [stripes[i % 2] for i in range(8)]
    assert sa.find_selected_row(plain) is None                       # nothing selected
    two = list(plain)
    two[1] = two[4] = (170, 195, 235)
    assert sa.find_selected_row(two) is None                         # ambiguous


def test_find_selected_row_with_compression_noise():
    colors = [(255, 255, 255) if i % 2 == 0 else (236, 235, 234) for i in range(10)]
    colors[3] = (168, 196, 233)
    colors[6] = (254, 255, 254)
    assert sa.find_selected_row(colors) == 3


@pytest.mark.parametrize("readings,glyphs,code,expected", [
    (["ZJO", "ZJO"], 3, "ZJO", "exact"),
    (["zs", "zS"], 2, "ZS", "exact"),                 # case does not matter
    (["ZJ0", ""], 3, "ZJO", "exact"),                 # 0 / O look-alike
    (["", ""], 1, "Z", "weak"),                       # OCR cannot read a lone 'Z' but one glyph is there
    (["Y", "Y"], 2, "YY", "weak"),                    # OCR collapses 'YY' into 'Y', two glyphs are there
    (["", ""], 1, "YY", "weak"),                      # the two Y's touch and count as one glyph
    (["Z1", "z"], 2, "ZI", "exact"),                  # 1 / I look-alike
    (["z", "z"], 2, "ZI", "weak"),
    (["ZMH", "ZMH"], 3, "ZM", None),                  # a longer code is not the code
    (["QS", "QS"], 2, "ZS", None),                    # first character disagrees
    (["", ""], 0, "ZS", None),                        # nothing in the cell
])
def test_match_code(readings, glyphs, code, expected):
    assert sa.match_code(readings, glyphs, code) == expected


def test_match_code_strict_mode_refuses_weak():
    assert sa.match_code(["", ""], 1, "Z", allow_weak=False) is None


@pytest.mark.parametrize("readings,glyphs,code,expected", [
    (["ZMH", "ZMH"], 3, "ZM", True),
    (["ZM", "ZM"], 2, "ZM", False),
    (["[ZJOL]", "ZJOL"], 3, "ZJO", False),            # noise character, but only 3 glyphs: not an extension
    (["ZMH"], 0, "ZM", True),                         # glyph count unknown: fall back to the text
])
def test_looks_like_extension(readings, glyphs, code, expected):
    assert sa.looks_like_extension(readings, glyphs, code) is expected


# ── fake screen ─────────────────────────────────────────────────────────────

FONT = ImageFont.load_default(size=13)
ROWS = [("ZYS", "Short term share"), ("ZTS", "Short term index"), ("ZTO", "Long term share"),
        ("ZS", "Long term index"), ("ZQK", "Future"), ("ZMH", "Short term share"),
        ("ZM", "Long term share"), ("ZJO", "Short term index")]
Y1, H, BORDER, RIGHT, RIC_LEFT = 100, 24, 470, 540, 400
ARROW = (836, 62)
AVAIL_BTN = (40, 45)         # first small button next to "Available contracts"

# The update window ("Strikes ...") with its Maturity | Minimum | Maximum | Ric grid
DATES = ["19 Oct 2026", "23 Oct 2026", "30 Oct 2026", "06 Nov 2026", "13 Nov 2026", "20 Nov 2026", "20 Nov 2026",
         "27 Nov 2026", "30 Nov 2026", "18 Dec 2026", "15 Jan 2027", "21 Jan 2027", "19 Feb 2027-1",
         "19 Feb 2027", "19 Mar 2027", "16 Apr 2027"]
HX = {"Maturity": 290, "Minimum": 380, "Maximum": 460, "Ric": 600}       # header text centers (x)
HEADER_Y, PITCH, VISIBLE = 175, 22, 10                                    # header y, row pitch, rows in view

ANCHOR_GRID = (22, 62, 330, 84)      # header "Mnemo | Extended Mnemo | Description"
ANCHOR_RIGHT = (700, 188, 850, 208)  # header "Date | Cycle | Name"


class FakeSophis:
    def __init__(self, clipboard=False, popup_visible=True, shift=(0, 0), hidden=False, always_edit=False,
                 tinted_mnemo=False, no_highlight=False):
        self.tinted_mnemo, self.no_highlight = tinted_mnemo, no_highlight
        self.dx, self.dy = shift
        self.editing = None          # row whose Mnemo cell is open for editing (a click on a SELECTED row does it)
        self.edited = ""             # what was typed into that edit box (i.e. what would overwrite the data)
        self.always_edit = always_edit
        self.hidden = hidden
        self.sel, self.typed, self.dd = 0, "", "Short term index"
        self.popup, self.dialog = False, False
        self.term_extra = 0                           # how much the Term column was widened (right border dragged)
        self.drags = []                               # every header drag: (start x in window coordinates, dx)
        self.popup_idx = 0
        self.clipboard, self.popup_visible = clipboard, popup_visible
        self._clip = None
        self.key_interval = 0
        self.events = []       # ("click" | "double" | "type", ...) in the order they happened
        self._last_type_t = -1e9
        self.mouse = (0, 0)    # where the pointer is (it stays on the row you clicked)
        self.avail_clicks = 0
        self.dscroll = 0       # first visible row of the update window's grid
        self.values = {}       # committed cell values of the update window: {(row number, "minimum"|"maximum"): text}
        self.cell_edit = None  # the cell being edited, (row number, column)
        self.cell_buf = ""     # text in the open edit box
        self.sel_all = False   # Ctrl+A pressed inside the edit box
        self.corrupt = False   # simulate a grid that stores something else than what was typed
        self.ok_clicks = 0
        self.dsel = 0          # selected row of the update window's grid (absolute index)
        self.dcell = None      # focused cell (row number, column)
        self.edit_mode = "double"   # how a cell opens: "double" | "select_first" | "f2" | "never"
        self.stray = ""        # what was typed while NO cell was open (it only moves the selection)
        self.keys = []         # hotkeys / key presses seen by the update window
        self.enter_moves = True      # Enter commits the edit AND moves the selection to the next row (like real grids)
        self.enter_scrolls = False   # ...and, on the last visible row, scrolls the list by one row

    # drawing
    def _image(self):
        im = Image.new("RGB", (1000, 600), (200, 200, 200))
        d = ImageDraw.Draw(im)
        term_left = BORDER
        term_width = RIGHT - BORDER + self.term_extra
        for i, (m, t) in enumerate(ROWS):
            y = Y1 - H // 2 + i * H
            color = (170, 195, 235) if i == self.sel else ((255, 255, 255) if i % 2 == 0 else (235, 235, 235))
            if self.no_highlight:                                              # selection not visible by color
                color = (255, 255, 255) if i % 2 == 0 else (235, 235, 235)
            d.rectangle((20, y, 640, y + H - 1), fill=color)
            if self.tinted_mnemo:                                              # Mnemo column has its own background
                d.rectangle((21, y, 133, y + H - 1), fill=(255, 230, 150))
            d.text((40, y + 4), m, fill=(0, 0, 160), font=FONT)
            d.line((135, y, 135, y + H - 1), fill=(90, 90, 90))               # border between Mnemo and Extended
            if i == self.sel:                                                  # focus rectangle of the selection
                d.rectangle((36, y + 1, 84, y + H - 3), outline=(0, 0, 0))
            if i == self.editing:                                              # edit box in the Mnemo cell
                d.rectangle((21, y + 1, 133, y + H - 2), fill=(255, 255, 255), outline=(0, 0, 0))
                d.text((24, y + 4), self.edited, fill=(0, 0, 0), font=FONT)
            d.text((200, y + 4), "VANECK OIL SERVIC...", fill=(0, 0, 160), font=FONT)   # Description
            d.text((300, y + 4), "US92189H6071", fill=(0, 0, 160), font=FONT)           # Code
            d.text((term_left + 6, y + 4), t if term_width >= 110 else t[:9] + "...", fill=(0, 0, 160), font=FONT)
        d.rectangle(ANCHOR_GRID, fill=(215, 215, 215))
        d.text((30, 66), "Mnemo        Extended Mnemo        Description", fill=(0, 0, 0), font=FONT)
        d.rectangle(ANCHOR_RIGHT, fill=(215, 215, 215))
        d.text((706, 192), "Date     Cycle     Name", fill=(0, 0, 0), font=FONT)
        d.text((term_left + 6, 62), "Term", fill=(0, 0, 0), font=FONT)
        for k in range(3):                                                  # the three little buttons
            d.rectangle((32 + 25 * k, 37, 48 + 25 * k, 53), fill=(225, 225, 225), outline=(0, 0, 0))
        d.rectangle((700, 50, 850, 74), fill=(255, 255, 255), outline=(0, 0, 0))
        d.text((706, 55), self.dd, fill=(0, 0, 0), font=FONT)
        d.polygon([(830, 58), (842, 58), (836, 67)], fill=(0, 0, 0))
        if self.popup and self.popup_visible:
            d.rectangle((700, 76, 850, 76 + 22 * 5), fill=(255, 255, 255), outline=(0, 0, 0))
            for i, t in enumerate(TERMS):
                d.text((706, 80 + 22 * i), t, fill=(0, 0, 0), font=FONT)
        if self.dialog:
            d.rectangle((200, 150, 800, 450), fill=(240, 240, 240), outline=(0, 0, 0))
            d.text((220, 158), "Strikes ZIMMER BIOMET HOLDINGS INC", fill=(0, 0, 0), font=FONT, anchor="lm")
            d.rectangle((232, HEADER_Y - 11, 660, HEADER_Y + 11 + PITCH * VISIBLE), fill=(255, 255, 255),
                        outline=(0, 0, 0))
            for name, x in HX.items():
                d.text((x, HEADER_Y), name, fill=(0, 0, 0), font=FONT, anchor="mm")
            for x in (335, 420, 500):                                       # column borders
                d.line((x, HEADER_Y - 11, x, HEADER_Y + 11 + PITCH * VISIBLE), fill=(120, 120, 120))
            for k in range(VISIBLE):
                n = self.dscroll + k
                if n >= len(DATES):
                    break
                y = HEADER_Y + PITCH * (k + 1)
                selected = (n == self.dsel)                                 # the first date starts highlighted
                if selected:
                    d.rectangle((233, y - PITCH // 2 + 1, 659, y + PITCH // 2 - 1), fill=(10, 36, 106))
                d.text((HX["Maturity"], y), DATES[n], fill=(255, 255, 255) if selected else (0, 0, 0), font=FONT,
                       anchor="mm")
                for col in ("minimum", "maximum"):
                    x = HX[col.capitalize()]
                    if self.dcell == (n, col) and self.cell_edit != (n, col):       # focused cell: light blue
                        d.rectangle((x - 40, y - PITCH // 2 + 1, x + 40, y + PITCH // 2 - 1), fill=(170, 200, 235))
                    if self.cell_edit == (n, col):                                    # the open edit box
                        d.rectangle((x - 38, y - 9, x + 38, y + 9), fill=(255, 255, 255), outline=(0, 0, 0))
                        d.text((x, y), self.cell_buf, fill=(0, 0, 0), font=FONT, anchor="mm")
                    elif (n, col) in self.values:                                     # grid shows two decimals
                        txt = self.values[(n, col)]
                        try:
                            txt = f"{float(txt):.2f}"
                        except ValueError:
                            pass
                        d.text((x, y), txt, fill=(255, 255, 255) if selected else (0, 0, 0), font=FONT, anchor="mm")
            d.text((560, 425), "OK", fill=(0, 0, 0), font=FONT, anchor="mm")
            d.text((640, 425), "Cancel", fill=(0, 0, 0), font=FONT, anchor="mm")
        canvas = Image.new("RGB", (1200, 800), (30, 60, 90))
        if self.hidden:                        # minimized window: only the desktop is visible
            return canvas
        canvas.paste(im, (self.dx, self.dy))   # the window appears shifted
        return canvas

    # Screen API (receives ABSOLUTE screen coordinates)
    def position(self):
        return self.mouse

    def scroll(self, clicks, x, y):
        x, y = x - self.dx, y - self.dy
        if self.dialog and 232 <= x <= 660 and HEADER_Y <= y <= HEADER_Y + PITCH * VISIBLE:
            self.dscroll = max(0, min(len(DATES) - VISIBLE, self.dscroll - clicks))     # one row per wheel click
            self.events.append(("scroll", clicks))

    def mouse_on_row(self, i):
        self.mouse = (45 + self.dx, Y1 + i * H + self.dy)

    def capture(self, region=None):
        im = self._image()
        return im if region is None else im.crop(tuple(int(v) for v in region))

    def _dialog_cell(self, x, y):
        """(row number, column) of the Minimum / Maximum cell at window coordinates (x, y), else None."""
        k = int(round((y - HEADER_Y) / PITCH)) - 1
        if not (0 <= k < VISIBLE and self.dscroll + k < len(DATES) and abs(y - (HEADER_Y + PITCH * (k + 1))) <= 10):
            return None
        col = "minimum" if 335 <= x < 420 else "maximum" if 420 <= x < 500 else None
        return (self.dscroll + k, col) if col else None

    def _commit_cell(self):
        if self.cell_edit is not None:
            v = self.cell_buf
            if self.corrupt:
                try:
                    v = str(float(v) + 1)
                except ValueError:
                    pass
            self.values[self.cell_edit] = v
            self.cell_edit, self.cell_buf, self.sel_all = None, "", False

    def _row_at(self, x, y):
        for i in range(len(ROWS)):
            if abs(y - (Y1 + i * H)) <= H // 2 and 20 <= x <= 640:
                return i
        return None

    def click(self, x, y):
        self.events.append(("click", x, y))
        self._click(x, y)

    def _click(self, x, y):
        x, y = x - self.dx, y - self.dy
        if self.dialog:                                         # the update window is modal: only it reacts
            if 520 <= x <= 600 and 415 <= y <= 435:
                self.ok_clicks += 1                             # (the script must never do this for now)
            self._commit_cell()                                 # any click outside an edit box commits it
            cell = self._dialog_cell(x, y)
            if cell:                                            # a click on a cell selects its row and focuses it
                self.dsel, self.dcell = cell[0], cell
            return
        if self.popup:
            if 700 <= x <= 850 and 76 <= y <= 76 + 110:
                self.dd = TERMS[min(4, (y - 76) // 22)]
            self.popup = False
            return
        if abs(x - AVAIL_BTN[0]) <= 8 and abs(y - AVAIL_BTN[1]) <= 8:     # first button: opens the selected row
            self.avail_clicks += 1
            if self.sel is not None:
                self.dialog = True
            return
        if abs(x - ARROW[0]) <= 14 and abs(y - ARROW[1]) <= 14:
            self.popup, self.popup_idx = True, 0
            return
        i = self._row_at(x, y)
        if i is not None:
            if self.always_edit or (i == self.sel and self.editing is None):
                self.editing = i                       # clicking an already selected cell opens it for editing
            else:
                self.sel, self.typed, self.editing = i, "", None

    def double_click(self, x, y):
        self.events.append(("double", x, y))
        if self.dialog:                                         # double click on a Minimum / Maximum cell
            self._commit_cell()
            cell = self._dialog_cell(x - self.dx, y - self.dy)
            if cell:
                was_selected = (self.dsel == cell[0])
                self.dsel, self.dcell = cell[0], cell
                if self.edit_mode == "double" or (self.edit_mode == "select_first" and was_selected):
                    self.cell_edit, self.cell_buf, self.sel_all = cell, self.values.get(cell, ""), False
            return
        i = self._row_at(x - self.dx, y - self.dy)
        if i is not None:                              # a double click opens the window; it does not edit
            self.sel, self.editing, self.typed = i, None, ""
            self.dialog = True

    def type_text(self, text):
        self.events.append(("type", text))
        if self.cell_edit is not None:                           # typing into a cell of the update window
            if self.sel_all:
                self.cell_buf, self.sel_all = "", False          # Ctrl+A first: the typed text replaces the old
            self.cell_buf += text
            return
        if self.dialog:                                          # no cell open: typing only moves the selection
            self.stray += text
            for n, d_ in enumerate(DATES):
                if d_.startswith(text[0]):
                    self.dsel = n
                    break
            return
        if self.editing is not None:                   # typing into the open edit box overwrites the cell
            self.edited += text
            return
        if sa.time.time() - self._last_type_t > 0.02:  # the grid's type-ahead buffer expires after a pause
            self.typed = ""
        self._last_type_t = sa.time.time()
        for ch in text:
            self.typed += ch.upper()
            for i, (m, _) in enumerate(ROWS):
                if m.startswith(self.typed):
                    self.sel = i
                    break

    def press(self, key):
        if self.dialog:
            self.keys.append(key)
            if key == "enter":
                row = self.cell_edit[0] if self.cell_edit else None
                self._commit_cell()
                if self.enter_moves:
                    self.dsel = min(len(DATES) - 1, self.dsel + 1)
                if self.enter_scrolls and row is not None and row == self.dscroll + VISIBLE - 1:
                    self.dscroll = min(len(DATES) - VISIBLE, self.dscroll + 1)
                return
            if key == "f2" and self.dcell and self.edit_mode == "f2" and self.cell_edit is None:
                self.cell_edit, self.cell_buf, self.sel_all = self.dcell, self.values.get(self.dcell, ""), False
            return
        if not self.popup and key == "down":
            self.sel = min(self.sel + 1, len(ROWS) - 1)
        elif self.popup and key == "home":
            self.popup_idx = 0
        elif self.popup and key == "down":
            self.popup_idx += 1
        elif self.popup and key == "enter":
            self.dd, self.popup = TERMS[self.popup_idx], False

    def hotkey(self, *keys):
        if tuple(keys) == ("ctrl", "a"):
            self.keys.append("ctrl+a")
            self.sel_all = self.cell_edit is not None
            return
        if self.clipboard:
            m, t = ROWS[self.sel]
            self._clip = f"{m}\t\tDESCRIPTION\tUS0000\t\t{t}"

    def drag(self, x0, y0, x1, y1):
        self.drags.append((x0 - self.dx, x1 - x0))
        if abs(x0 - self.dx - RIGHT) <= 4:                         # right border of Term -> Term gets wider
            self.term_extra += max(0, x1 - x0)
        # any other drag (e.g. the RIC | Term border) would change nothing useful here

    def clipboard_set(self, text):
        self._clip = text

    def clipboard_get(self):
        return self._clip


@pytest.fixture(scope="module")
def ocr():
    return OCR()


@pytest.fixture(autouse=True)
def _isolated_anchor_dir(tmp_path_factory, monkeypatch):
    """The update window's header template is remembered on disk; every test gets its own empty folder."""
    monkeypatch.setattr(sa, "ANCHOR_DIR", tmp_path_factory.mktemp("anchor_dir") / "sophis_anchors")


@pytest.fixture(scope="module")
def anchors(tmp_path_factory):
    """Calibration: crops of the two anchors taken with the window at its original position."""
    folder = tmp_path_factory.mktemp("anchors")
    base = FakeSophis().capture()
    result = {}
    for name, box in (("grid", ANCHOR_GRID), ("right", ANCHOR_RIGHT)):
        path = folder / f"{name}.png"
        base.crop(box).save(path)
        result[name] = {"file": str(path), "origin": [box[0], box[1]]}
    return result


def make(ocr, anchors, **kw):
    cfg = Config(anchors=anchors, points={
        "row1": [45, Y1], "row2": [45, Y1 + H], "rowN": [45, Y1 + 7 * H], "ext": [140, Y1],
        "ric_left": [RIC_LEFT, 70], "term_left": [BORDER, 70], "term_right": [RIGHT, 70], "dd_left": [706, 62], "dd_arrow": list(ARROW),
        "avail_btn": list(AVAIL_BTN),
    }, slowdown=0.01, **{k: v for k, v in kw.items() if k in ("reader", "verify_mnemonic", "auto_focus",
                                                              "open_method")})
    screen = FakeSophis(**{k: v for k, v in kw.items() if k in ("clipboard", "popup_visible", "shift", "hidden", "always_edit", "tinted_mnemo", "no_highlight")})
    a = Automator(screen, ocr, cfg, log=lambda *_: None)
    a.prepare()
    return a, screen


# ── full flow against the fake screen ───────────────────────────────────────

def test_geometry(ocr, anchors):
    a, _ = make(ocr, anchors)
    assert a.g.n == 8 and a.g.h == 24


@pytest.mark.parametrize("code,term", [
    ("ZYS", "Short term share"), ("ZS", "Long term index"), ("ZQK", "Future"), ("ZM", "Long term share"),
])
def test_open_by_ocr(ocr, anchors, code, term):
    a, fake = make(ocr, anchors, reader="ocr")
    r = a.open_contract(code)
    assert r.term == term and fake.dd == term and fake.dialog
    assert ROWS[fake.sel][0] == code


def test_open_by_clipboard(ocr, anchors):
    a, fake = make(ocr, anchors, clipboard=True)
    r = a.open_contract("ZTO")
    assert r.term == "Long term share" and fake.dd == "Long term share" and fake.dialog


def test_dropdown_by_keyboard_when_the_list_is_not_visible(ocr, anchors):
    a, fake = make(ocr, anchors, popup_visible=False, reader="ocr")
    a.open_contract("ZTO")
    assert fake.dd == "Long term share"


def test_unknown_code_aborts_without_opening(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake = make(ocr, anchors, reader="ocr")
    with pytest.raises(SophisError):
        a.open_contract("QQQ")
    assert not fake.dialog


# ── exactly one click before typing ─────────────────────────────────────────

@pytest.mark.parametrize("reader,clipboard", [("ocr", False), ("auto", True)])
def test_auto_focus_makes_a_single_click_before_typing(ocr, anchors, reader, clipboard):
    a, fake = make(ocr, anchors, reader=reader, clipboard=clipboard, auto_focus=True)
    assert fake.events == []                      # preparing must not click anything or type
    a.search("ZS")
    kinds = [e[0] for e in fake.events]
    assert kinds[:2] == ["click", "type"], kinds  # one click, then straight to typing


def test_each_code_gets_one_click_then_typing_until_the_window_opens_in_auto_focus(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr", auto_focus=True)
    a.open_contract("ZYS")
    first_type = [e[0] for e in fake.events].index("type")
    assert [e[0] for e in fake.events[:first_type]] == ["click"]


# ── the window is NOT always in the same position ───────────────────────────

@pytest.mark.parametrize("shift", [(57, 33), (-15, 120), (200, 0)])
@pytest.mark.parametrize("reader", ["ocr", "auto"])
def test_window_at_another_position(ocr, anchors, shift, reader):
    a, fake = make(ocr, anchors, shift=shift, reader=reader, clipboard=(reader == "auto"))
    r = a.open_contract("ZS")
    assert r.term == "Long term index" and fake.dd == "Long term index" and fake.dialog
    assert ROWS[fake.sel][0] == "ZS"


def test_window_that_moves_between_codes(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    a.open_contract("ZYS")
    fake.dialog, fake.dx, fake.dy = False, 80, 45        # the user closes the window and reopens it elsewhere
    r = a.open_contract("ZTO")
    assert r.term == "Long term share" and fake.dd == "Long term share" and fake.dialog


def test_minimized_window_gives_a_clear_error(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    with pytest.raises(SophisError, match="Cannot find the Sophis window"):
        make(ocr, anchors, hidden=True)


# ── calibrate (full and partial) with the fake screen ───────────────────────

class _Args:
    def __init__(self, only=None):
        self.wait, self.only = 0, only


def _calibrate_with(monkeypatch, tmp_path, fake, positions, only=None):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sa.time, "sleep", lambda *_: None)
    it = iter(positions)
    fake.position = lambda: next(it)
    monkeypatch.setattr(sa, "Screen", lambda *a, **k: fake)
    sa.cmd_calibrate(_Args(only))
    return Config.load()


GOOD_POINTS = [(22, 62), (330, 84), (45, Y1), (45, Y1 + H), (45, Y1 + 7 * H), (140, Y1), (RIC_LEFT, 70),
               (BORDER, 70), (RIGHT, 70),
               (700, 188), (850, 208), (706, 62), ARROW, AVAIL_BTN]


def test_full_then_partial_calibration_with_moved_window(ocr, monkeypatch, tmp_path):
    cfg = _calibrate_with(monkeypatch, tmp_path, FakeSophis(), GOOD_POINTS)
    assert set(cfg.anchors) == {"grid", "right"} and cfg.points["rowN"] == [45, Y1 + 7 * H]
    # the window is now shifted by (40, 25); only rowN is redone, pointing at its new position
    moved = FakeSophis(shift=(40, 25))
    cfg = _calibrate_with(monkeypatch, tmp_path, moved, [(45 + 40, Y1 + 5 * H + 25)], only=["rowN"])
    a = Automator(moved, ocr, cfg, log=lambda *_: None)
    a.prepare()
    assert a.g.n == 6 and a.g.y1 == Y1 + 25            # new rowN; the rest was translated to the current position


def test_invalid_calibration_does_not_overwrite_the_previous_one(ocr, monkeypatch, tmp_path):
    _calibrate_with(monkeypatch, tmp_path, FakeSophis(), GOOD_POINTS)
    before = (tmp_path / "sophis_config.json").read_text()
    with pytest.raises(SophisError, match="rowN"):
        _calibrate_with(monkeypatch, tmp_path, FakeSophis(), [(45, Y1 + H)], only=["rowN"])
    assert (tmp_path / "sophis_config.json").read_text() == before


def test_legacy_spanish_config_still_loads_and_works(ocr, anchors, tmp_path):
    legacy = {
        "puntos": {"ancla_grid_a": [0, 0], "ancla_grid_b": [0, 0], "fila1": [45, Y1], "fila2": [45, Y1 + H],
                   "filaN": [45, Y1 + 7 * H], "ext": [140, Y1], "term_borde": [BORDER, 70],
                   "term_der": [RIGHT, 70], "ancla_der_a": [0, 0], "ancla_der_b": [0, 0],
                   "dd_izq": [706, 62], "dd_flecha": list(ARROW)},
        "anclas": {"grid": {"archivo": anchors["grid"]["file"], "origen": anchors["grid"]["origin"]},
                   "der": {"archivo": anchors["right"]["file"], "origen": anchors["right"]["origin"]}},
        "ensanchar_px": 140, "intervalo_tecla": 0.12, "pausa": 0.01, "lector": "portapapeles",
        "verificar_mnemo": True,
    }
    path = tmp_path / "legacy.json"
    path.write_text(json.dumps(legacy), encoding="utf-8")
    with pytest.raises(SophisError, match="calibrate --only ric_left"):     # the one point the old version lacked
        Config.load(path)
    legacy["puntos"]["ric_left"] = [RIC_LEFT, 70]
    legacy["puntos"]["avail_btn"] = list(AVAIL_BTN)
    path.write_text(json.dumps(legacy), encoding="utf-8")
    cfg = Config.load(path)
    assert cfg.reader == "clipboard" and cfg.widen_px == 140 and set(cfg.anchors) == {"grid", "right"}
    assert cfg.points["row1"] == [45, Y1] and cfg.anchors["grid"]["origin"] == anchors["grid"]["origin"]
    fake = FakeSophis(shift=(30, 10), clipboard=True)
    a = Automator(fake, ocr, cfg, log=lambda *_: None)
    a.prepare()
    assert a.open_contract("ZQK").term == "Future" and fake.dialog


# ── two single clicks in a row must not become a double click ───────────────

class _FakePyAutoGui:
    FAILSAFE = False
    PAUSE = 0

    def __init__(self):
        self.events = []

    def moveTo(self, *a, **k):
        pass

    def click(self):
        self.events.append(("click", sa.time.monotonic()))

    def doubleClick(self, **k):
        self.events.append(("double", sa.time.monotonic()))


def test_two_single_clicks_are_kept_apart(monkeypatch):
    import sys
    pg = _FakePyAutoGui()
    monkeypatch.setitem(sys.modules, "pyautogui", pg)
    clock = {"t": 1000.0}
    waits = []
    monkeypatch.setattr(sa.time, "monotonic", lambda: clock["t"])
    monkeypatch.setattr(sa.time, "sleep", lambda s: (waits.append(s), clock.__setitem__("t", clock["t"] + s)))
    screen = sa.Screen()
    screen.click(10, 10)
    clock["t"] += 0.4                   # the script only waits 0.4 s between clicks (less than Windows' double-click time)
    screen.click(10, 10)
    t1, t2 = (e[1] for e in pg.events)
    assert t2 - t1 >= screen._double_interval + 0.1
    assert waits, "it should have waited before the second click"


def test_clicks_that_are_already_apart_do_not_wait(monkeypatch):
    import sys
    pg = _FakePyAutoGui()
    monkeypatch.setitem(sys.modules, "pyautogui", pg)
    clock = {"t": 1000.0}
    waits = []
    monkeypatch.setattr(sa.time, "monotonic", lambda: clock["t"])
    monkeypatch.setattr(sa.time, "sleep", lambda s: waits.append(s))
    screen = sa.Screen()
    screen.click(10, 10)
    clock["t"] += 5
    screen.click(10, 10)
    assert not waits


def test_typing_is_never_skipped_after_the_click(ocr, anchors):
    """Whatever the screen looks like, the click is followed by typing (no pre-check can abort before it)."""
    a, fake = make(ocr, anchors, reader="ocr")
    a.focus_grid()
    a.s.type_text("ZS")
    assert [e[0] for e in fake.events] == ["click", "type"]


# ── reading a Mnemo cell with grid lines and a focus rectangle around the text ──

def _cell(text, selected=True, width=96, height=18, dark=False):
    """A Mnemo cell. dark=True is Windows' classic selection: WHITE text on a DARK blue background."""
    if dark:
        bg, ink, frame = (10, 36, 106), (255, 255, 255), (255, 255, 255)
    else:
        bg, ink, frame = ((170, 195, 235) if selected else (255, 255, 255)), (0, 0, 160), (0, 0, 0)
    im = Image.new("RGB", (width, height), bg)
    d = ImageDraw.Draw(im)
    d.text((14, 2), text, fill=ink, font=ImageFont.load_default(size=12))
    d.line((0, 0, 0, height), fill=(90, 90, 90))                # column border at the crop edge
    d.line((width - 8, 2, width - 8, height - 2), fill=(60, 60, 60))   # column border inside the crop
    d.rectangle((3, 1, width - 12, height - 2), outline=frame)         # focus rectangle
    return im


@pytest.mark.parametrize("text", ["ZYS", "ZTS", "ZJO", "ZMH", "ZQK", "ZM", "ZS", "ZION", "ZBRA"])
@pytest.mark.parametrize("selected", [True, False])
def test_read_cell_confirms_the_code_despite_borders(ocr, text, selected):
    readings, glyphs = ocr.read_cell(_cell(text, selected))
    assert len(text) - 1 <= glyphs <= len(text)       # touching letters can merge into one glyph
    assert sa.match_code(readings, glyphs, text) in ("exact", "weak"), readings


@pytest.mark.parametrize("text", ["Z", "YY", "ZI"])
def test_read_cell_short_codes_are_confirmed_by_shape(ocr, text):
    readings, glyphs = ocr.read_cell(_cell(text))
    assert len(text) - 1 <= glyphs <= len(text)       # touching letters can merge into one glyph
    assert sa.match_code(readings, glyphs, text) in ("exact", "weak"), readings


def test_read_cell_does_not_confirm_a_different_code(ocr):
    readings, glyphs = ocr.read_cell(_cell("ZMH"))
    assert sa.match_code(readings, glyphs, "ZM") is None
    assert sa.looks_like_extension(readings, glyphs, "ZM")


def test_clean_cell_drops_borders_and_frame():
    _, glyphs = sa.clean_cell(_cell("ZJO"))
    assert glyphs == 3


# ── widening the Term column and reading only the Term cell ─────────────────

def test_term_strip_does_not_reach_description_or_code(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    assert a.g.term_x0 > RIC_LEFT                         # starts at the RIC column, never in Code/Description
    a.search("ZYS")
    term, text = a._term_by_ocr(0 if fake.sel == 0 else fake.sel)
    assert "VANECK" not in text and "US92189" not in text


def test_widen_term_drags_only_the_right_border_of_term(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    a.search("ZYS")
    i = a.selected_row()
    assert a._term_by_ocr(i)[0] is None                   # truncated: "Short ter..."
    a.widen_term()
    assert len(fake.drags) == 1                           # one drag...
    start, moved = fake.drags[0]
    assert abs(start - RIGHT) <= 1 and moved > 0          # ...of the Term | Cyc border, to the right
    assert fake.term_extra == moved
    assert a._term_by_ocr(i)[0] == "Short term share"     # and the strip follows the widened column


def test_widen_term_is_attempted_once_per_session(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    a.widen_term()
    a.widen_term()
    assert len(fake.drags) == 1


def test_widen_term_never_drags_off_the_screen(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    a._screen_width = RIGHT + 20                          # almost no room on the right
    a.widen_term()
    assert fake.drags == []                               # not enough room: it does nothing instead of dragging blindly


def test_an_unconfirmed_row_is_reread_not_clicked_again(ocr, anchors, tmp_path, monkeypatch):
    """The repeated-click problem: when the Mnemo cannot be confirmed, attempt 2 only re-reads the screen."""
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake = make(ocr, anchors, reader="ocr", auto_focus=True)
    with pytest.raises(SophisError):
        a.search("QQQ")                                   # not in the list: never confirmed
    clicks = [e for e in fake.events if e[0] == "click"]
    types = [e for e in fake.events if e[0] == "type"]
    assert len(clicks) == len(types) == 2                 # attempt 1 and the last attempt, nothing in between


def test_a_calibration_missing_one_point_can_be_repaired_with_only(ocr, anchors, monkeypatch, tmp_path):
    """Regression: `calibrate --only ric_left` used to fail with 'missing ric_left' because it could not load
    the very calibration it was meant to complete."""
    legacy = {
        "puntos": {"ancla_grid_a": [0, 0], "ancla_grid_b": [0, 0], "fila1": [45, Y1], "fila2": [45, Y1 + H],
                   "filaN": [45, Y1 + 7 * H], "ext": [140, Y1], "term_borde": [BORDER, 70],
                   "term_der": [RIGHT, 70], "ancla_der_a": [0, 0], "ancla_der_b": [0, 0],
                   "dd_izq": [706, 62], "dd_flecha": list(ARROW)},
        "anclas": {"grid": {"archivo": anchors["grid"]["file"], "origen": anchors["grid"]["origin"]},
                   "der": {"archivo": anchors["right"]["file"], "origen": anchors["right"]["origin"]}},
        "ensanchar_px": 140, "intervalo_tecla": 0.12, "pausa": 0.01, "lector": "ocr", "verificar_mnemo": True,
    }
    (tmp_path / "sophis_config.json").write_text(json.dumps(legacy), encoding="utf-8")
    monkeypatch.chdir(tmp_path)
    with pytest.raises(SophisError, match="calibrate --only ric_left"):
        Config.load()                                         # normal use is still blocked until repaired
    fake = FakeSophis(shift=(25, 15))                         # and the window is not where it was
    monkeypatch.setattr(sa.time, "sleep", lambda *_: None)
    spots = iter([(RIC_LEFT + 25, 70 + 15), (AVAIL_BTN[0] + 25, AVAIL_BTN[1] + 15)])
    fake.position = lambda: next(spots)
    monkeypatch.setattr(sa, "Screen", lambda *a, **k: fake)
    sa.cmd_calibrate(_Args(["ric_left", "avail_btn"]))
    cfg = Config.load()                                       # now complete and in the new format
    assert cfg.points["ric_left"] == [RIC_LEFT + 25, 70 + 15]
    assert cfg.points["avail_btn"] == [AVAIL_BTN[0] + 25, AVAIL_BTN[1] + 15]
    assert cfg.reader == "ocr" and cfg.widen_px == 140        # old settings kept
    a = Automator(fake, ocr, cfg, log=lambda *_: None)
    a.prepare()
    assert a.open_contract("ZJO").term == "Short term index" and fake.dialog


# ── light text on a dark selection (the cleaned image used to come out blank) ──

@pytest.mark.parametrize("text", ["ZYS", "ZJO", "ZMH", "ZS", "ZION"])
def test_white_text_on_dark_selection_is_read(ocr, text):
    readings, glyphs = ocr.read_cell(_cell(text, dark=True))
    assert glyphs >= len(text) - 1
    assert sa.match_code(readings, glyphs, text) in ("exact", "weak"), readings


def test_cleaned_cell_is_never_blank_for_either_polarity():
    import numpy as np
    for dark in (False, True):
        cleaned, glyphs = sa.clean_cell(_cell("ZJO", dark=dark))
        assert glyphs == 3
        assert (cleaned < 128).sum() > 20, f"dark={dark}: the letters vanished"     # real black ink exists
        assert cleaned.max() == 255


@pytest.mark.parametrize("result,expected", [
    ([[[[0, 0], [1, 0], [1, 1], [0, 1]], "ZJO", "0.62"]], "ZJO"),     # older layout: [box, text, score-as-str]
    ([["ZJO", 0.71]], "ZJO"),                                          # newer layout: [text, score]
    ([("ZJO", 0.71, 0.0)], "ZJO"),
    ([["", 0.0]], ""),
    (None, ""),
    ([], ""),
    ([[0.0, 0.0]], ""),                                                # numbers only: no text, and no crash
])
def test_text_from_ocr_result(result, expected):
    assert sa.text_from_ocr_result(result) == expected


def test_match_code_survives_non_string_readings():
    assert sa.match_code([0.0, None, "ZJO"], 3, "ZJO") == "exact"


# ── a click on an already selected cell opens it for editing: never do that ──

def test_the_click_never_lands_on_the_row_that_is_already_selected(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr", auto_focus=True)   # the fake starts with row 0 selected
    clicked = a.focus_grid()
    assert clicked == 1 and fake.editing is None
    assert fake.events[0] == ("click", 45, Y1 + H)        # row 1, not the selected row 0


def test_when_row_one_is_the_selected_one_it_clicks_another(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr", auto_focus=True)
    fake.sel = 1
    assert a.focus_grid() == 2 and fake.editing is None


@pytest.mark.parametrize("auto_focus", [False, True])
@pytest.mark.parametrize("code", ["ZS", "ZQK", "ZYS", "ZM"])
def test_repeated_searches_never_overwrite_a_cell(ocr, anchors, code, auto_focus):
    """The reported bug: after a search the found row sits selected at the top; clicking it again opened it for
    editing and the next search was typed INTO the cell."""
    a, fake = make(ocr, anchors, reader="ocr", auto_focus=auto_focus)
    for c in ("ZJO", code, "ZTS", code):
        a.search(c)
    assert fake.edited == "" and fake.editing is None


def test_row_in_edit_mode_detection(ocr, anchors):
    a, fake = make(ocr, anchors)
    assert not a._row_in_edit_mode(0)                     # selected normally: every cell has the highlight color
    assert not a._row_in_edit_mode(3)                     # an ordinary row
    fake.editing = 0                                      # the Mnemo cell became a white edit box
    assert a._row_in_edit_mode(0)


# ── manual mode (default): YOU click a row once, the script only types ───────

def test_manual_mode_never_clicks_before_typing(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")            # the fake starts with a row selected, as after your click
    a.search("ZS")
    assert fake.events[0][0] == "type"
    assert not [e for e in fake.events if e[0] in ("click", "double")]
    assert fake.edited == ""


@pytest.mark.parametrize("start_row", [0, 1, 4])
def test_manual_mode_works_whichever_row_you_clicked(ocr, anchors, start_row):
    a, fake = make(ocr, anchors, reader="ocr")
    fake.sel = start_row
    r = a.search("ZQK")
    assert ROWS[fake.sel][0] == "ZQK" and r[1].match in ("exact", "weak")
    assert fake.edited == ""


def test_manual_mode_open_contract_clicks_only_the_dropdown_and_the_button(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    a.open_contract("ZTO")
    kinds = [e[0] for e in fake.events]
    assert kinds[0] == "type" and kinds.count("double") == 0 and fake.avail_clicks == 1
    assert fake.dialog and fake.edited == ""


def test_manual_mode_waits_for_the_typeahead_buffer_between_codes(ocr, anchors, monkeypatch):
    a, fake = make(ocr, anchors, reader="ocr")
    a.cfg.type_gap, a.cfg.slowdown = 2.5, 1.0
    slept = []
    real_sleep = sa.time.sleep
    monkeypatch.setattr(sa.time, "sleep", lambda x: (slept.append(x), real_sleep(0.01)))
    a._last_typed_at = sa.time.time()                     # a code was typed just now
    a._type_code("ZS")
    assert any(x > 1.5 for x in slept), slept             # it let the grid's search buffer expire first


# ── manual mode trusts your click: nothing can stop the typing ───────────────

def test_typing_is_never_blocked_by_a_false_edit_mode_alarm(ocr, anchors):
    """The Mnemo column has its own background color, so the edit-mode heuristic would raise a false alarm.
    The script must type anyway and still find the row."""
    a, fake = make(ocr, anchors, reader="ocr", tinted_mnemo=True)
    assert a._row_in_edit_mode(0)                          # the (informational) heuristic is fooled...
    r = a.search("ZQK")
    assert fake.events[0][0] == "type" and ROWS[fake.sel][0] == "ZQK"     # ...and the script typed regardless
    assert r[1].match in ("exact", "weak")


@pytest.mark.parametrize("sel", [None, 5])
def test_typing_does_not_depend_on_a_row_being_detected_as_selected(ocr, anchors, sel):
    a, fake = make(ocr, anchors, reader="ocr")
    fake.sel = sel
    a.search("ZS")
    assert fake.events[0][0] == "type" and ROWS[fake.sel][0] == "ZS"


def test_row_found_by_reading_when_the_highlight_cannot_be_seen(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr", no_highlight=True)
    assert a.selected_row() is None
    i, reading = a.search("ZQK")
    assert ROWS[i][0] == "ZQK" and reading.match == "exact"
    assert [e[0] for e in fake.events].count("type") == 1          # typed once, not retried blindly


def test_open_contract_without_row_highlight_never_types_after_the_dropdown(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr", no_highlight=True)
    a.open_contract("ZTO")
    assert [e[0] for e in fake.events].count("type") == 1 and fake.dialog and fake.dd == "Long term share"


# ── the flow YOU drive: you select the row, the script continues from there ──

def _logged(ocr, anchors, **kw):
    a, fake = make(ocr, anchors, **kw)
    lines = []
    a.log = lines.append
    return a, fake, lines


@pytest.mark.parametrize("row,term", [(0, "Short term share"), (3, "Long term index"), (4, "Future"),
                                      (6, "Long term share"), (7, "Short term index")])
def test_open_selected_uses_the_row_you_selected(ocr, anchors, row, term):
    a, fake, _ = _logged(ocr, anchors, reader="ocr")
    fake.sel = row
    r = a.open_selected()
    assert r.row == row and r.term == term and fake.dd == term and fake.dialog
    assert not [e for e in fake.events if e[0] == "click" and e[1] < 700 and e[2] >= Y1 - H // 2]   # no row clicked
    assert [e[0] for e in fake.events].count("type") == 0 and fake.edited == ""
    assert fake.avail_clicks == 1 and not [e for e in fake.events if e[0] == "double"]    # opened with the button


def test_open_selected_uses_the_mouse_when_the_highlight_cannot_be_seen(ocr, anchors):
    a, fake, lines = _logged(ocr, anchors, reader="ocr", no_highlight=True)
    fake.sel = 5
    fake.mouse_on_row(5)
    r = a.open_selected()
    assert r.row == 5 and r.term == "Short term share" and fake.dialog
    assert any("mouse position" in l for l in lines)


def test_open_selected_prefers_the_highlight_if_the_mouse_disagrees(ocr, anchors):
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.sel = 2
    fake.mouse_on_row(6)
    assert a.open_selected().row == 2
    assert any("using the highlight" in l for l in lines)


def test_open_selected_says_so_when_it_cannot_tell_the_row(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake, _ = _logged(ocr, anchors, reader="ocr", no_highlight=True)
    fake.mouse = (5, 5)                                    # pointer nowhere near the table
    with pytest.raises(SophisError, match="Could not tell which row"):
        a.open_selected()
    assert not fake.dialog and not [e for e in fake.events if e[0] == "double"]


def test_open_selected_widens_the_term_column_when_it_is_cut(ocr, anchors):
    a, fake, _ = _logged(ocr, anchors, reader="ocr")
    fake.sel = 3
    assert fake.term_extra == 0                            # truncated: "Long term..."
    r = a.open_selected()
    assert r.term == "Long term index" and fake.term_extra > 0


def test_open_selected_with_term_given_never_reads_it(ocr, anchors):
    a, fake, _ = _logged(ocr, anchors, reader="ocr")
    a.cfg.widen_px = 0                                     # and no widening: the Term stays unreadable
    fake.sel = 3
    with pytest.raises(SophisError, match="--term"):
        a.open_selected()
    r = a.open_selected(term="Long term share")            # ...but you can tell it yourself
    assert r.term == "Long term share" and fake.dd == "Long term share" and fake.dialog and r.row == 3


def test_open_selected_works_with_the_window_somewhere_else(ocr, anchors):
    a, fake, _ = _logged(ocr, anchors, reader="ocr", shift=(60, 35))
    fake.sel = 4
    r = a.open_selected()
    assert r.row == 4 and r.term == "Future" and fake.dialog


def test_open_selected_does_not_depend_on_the_clipboard_or_reading_the_mnemo(ocr, anchors):
    """Even if the Mnemo OCR returns nothing, the flow continues: it never needs the code."""
    a, fake, _ = _logged(ocr, anchors, reader="ocr")
    fake.sel = 1
    a._mnemonic_by_ocr = lambda i: ([], 0)
    r = a.open_selected()
    assert r.row == 1 and r.term == "Short term index" and fake.dialog and r.code == "?"


def test_row_under_mouse(ocr, anchors):
    a, fake = make(ocr, anchors)
    fake.mouse_on_row(3)
    assert a.row_under_mouse() == 3
    fake.mouse = (45, Y1 - 60)
    assert a.row_under_mouse() is None
    fake.mouse = (900, Y1 + H)
    assert a.row_under_mouse() is None


def test_cli_open_refuses_codes_and_points_to_open_code(monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["sophis_auto.py", "open", "ZYS"])
    with pytest.raises(SystemExit) as e:
        sa.main()
    assert e.value.code == 2 and "open-code ZYS" in capsys.readouterr().out


# ── how the update window is opened ──────────────────────────────────────────

def test_the_window_is_opened_with_the_button_next_to_available_contracts(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    fake.sel = 3
    a.open_selected()
    assert fake.avail_clicks == 1 and fake.dialog
    assert not [e for e in fake.events if e[0] == "double"]
    assert fake.events[-1] == ("click", AVAIL_BTN[0], AVAIL_BTN[1])        # the last thing the script did


def test_the_button_click_follows_the_window_when_it_moves(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr", shift=(70, 40))
    fake.sel = 2
    a.open_selected()
    assert fake.events[-1] == ("click", AVAIL_BTN[0] + 70, AVAIL_BTN[1] + 40) and fake.dialog


def test_the_button_is_clicked_only_after_the_dropdown_matches(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
    fake.sel = 3                                           # Long term index; the fake starts on Short term index
    a.open_selected()
    assert fake.dd == "Long term index" and fake.avail_clicks == 1


def test_double_click_can_still_be_chosen(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr", open_method="double")
    fake.sel = 4
    a.open_selected()
    assert fake.avail_clicks == 0 and [e for e in fake.events if e[0] == "double"] == [("double", 45, Y1 + 4 * H)]
    assert fake.dialog


def test_cli_open_with_option_is_parsed(monkeypatch):
    seen = {}
    monkeypatch.setattr(sa, "cmd_open", lambda args: seen.update(open_with=args.open_with, term=args.term))
    monkeypatch.setattr(sys, "argv", ["sophis_auto.py", "open", "--open-with", "double", "--term", "Future"])
    sa.main()
    assert seen == {"open_with": "double", "term": "Future"}


# ── stage 2: where do the values go in the update window? ────────────────────

import datetime as dt


def D(text):
    return sa.parse_date(text)[0]


@pytest.mark.parametrize("text,date,suffix", [
    ("19 Oct 2026", dt.date(2026, 10, 19), ""),
    ("19-Oct-2026", dt.date(2026, 10, 19), ""),
    ("2026-10-19", dt.date(2026, 10, 19), ""),
    ("19/10/2026", dt.date(2026, 10, 19), ""),
    ("19 Feb 2027-1", dt.date(2027, 2, 19), "-1"),
    ("l9 0ct 2O26", dt.date(2026, 10, 19), ""),        # OCR slips: l for 1, 0 for O, O for 0
    ("19 Ocl 2026", dt.date(2026, 10, 19), ""),
    ("12/31/2026", dt.date(2026, 12, 31), ""),         # month/day order is guessed when the day cannot be second
])
def test_parse_date(text, date, suffix):
    assert sa.parse_date(text) == (date, suffix)


@pytest.mark.parametrize("text", ["", "Maturity", "Short term index", "19 Xyz 2026", "99/99/2026", "19 Oct 1850"])
def test_parse_date_rejects_what_is_not_a_date(text):
    assert sa.parse_date(text)[0] is None


def _open_window(ocr, anchors, **kw):
    a, fake = make(ocr, anchors, reader="ocr", **kw)
    fake.dialog = True
    return a, fake


def test_read_update_window_finds_columns_and_every_visible_date(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    win = a.read_update_window()
    for name, x in (("x_maturity", HX["Maturity"]), ("x_minimum", HX["Minimum"]), ("x_maximum", HX["Maximum"]),
                    ("x_ric", HX["Ric"])):
        assert abs(getattr(win, name) - x) <= 4, name
    assert abs(win.header_y - HEADER_Y) <= 4 and abs(win.pitch - PITCH) <= 2
    assert [r.date for r in win.rows] == [D(t) for t in DATES[:VISIBLE]]      # all ten, in order
    for r in win.rows:
        assert abs(r.cy - (HEADER_Y + PITCH * (r.index + 1))) <= 3


def test_the_highlighted_first_row_is_read_too(ocr, anchors):
    """White text on dark blue (the selected row) is what the plain detector tends to miss."""
    a, fake = _open_window(ocr, anchors)
    win = a.read_update_window()
    assert win.rows[0].index == 0 and win.rows[0].date == D("19 Oct 2026")


def test_locate_date_gives_the_minimum_and_maximum_cells(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    t = a.locate_date(D("30 Nov 2026"))
    y = HEADER_Y + PITCH * (8 + 1)
    assert t.row.index == 8 and not t.ambiguous
    assert abs(t.minimum[0] - HX["Minimum"]) <= 4 and abs(t.maximum[0] - HX["Maximum"]) <= 4
    assert abs(t.minimum[1] - y) <= 3 and t.minimum[1] == t.maximum[1]


def test_locate_date_on_the_highlighted_row(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    t = a.locate_date(D("19 Oct 2026"))
    assert t.row.index == 0 and abs(t.minimum[1] - (HEADER_Y + PITCH)) <= 3


def test_a_repeated_date_is_reported_as_ambiguous_until_you_choose(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    t = a.locate_date(D("20 Nov 2026"))
    assert t.ambiguous and [m.index for m in t.matches] == [5, 6] and t.row.index == 5
    second = a.locate_date(D("20 Nov 2026"), occurrence=1)
    assert not second.ambiguous and second.row.index == 6 and second.minimum[1] > t.minimum[1]
    with pytest.raises(SophisError, match="no occurrence 2"):
        a.locate_date(D("20 Nov 2026"), occurrence=2)


def test_the_list_is_scrolled_to_reach_a_date_that_is_not_visible(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    t = a.locate_date(D("16 Apr 2027"))
    assert fake.dscroll > 0 and t.row.date == D("16 Apr 2027")
    assert t.minimum[1] <= HEADER_Y + PITCH * VISIBLE + 3                          # the cell is on screen now


def test_scrolling_back_up_to_an_earlier_date(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    fake.dscroll = len(DATES) - VISIBLE
    t = a.locate_date(D("23 Oct 2026"))
    assert fake.dscroll == 0 and t.row.date == D("23 Oct 2026")


def test_a_suffixed_duplicate_is_found_and_flagged(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    t = a.locate_date(D("19 Feb 2027"))
    assert t.ambiguous and sorted(m.suffix for m in t.matches) == ["", "-1"]
    plain = a.locate_date(D("19 Feb 2027"), occurrence=[m.suffix for m in t.matches].index(""))
    assert plain.row.suffix == ""


def test_a_date_that_is_not_in_the_list_says_so(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    with pytest.raises(SophisError, match="is not in the list"):
        a.locate_date(D("24 Oct 2026"))                                            # between two existing dates
    with pytest.raises(SophisError, match="was not found"):
        a.locate_date(D("01 Jan 2030"))                                            # beyond the end: stops scrolling
    assert not [e for e in fake.events if e[0] in ("click", "double", "type")]     # read-only: nothing clicked


def test_update_window_cells_follow_the_window_when_it_moves(ocr, anchors):
    a, fake = _open_window(ocr, anchors, shift=(55, 30))
    t = a.locate_date(D("30 Nov 2026"))
    assert abs(t.minimum[0] - (HX["Minimum"] + 55)) <= 4
    assert abs(t.minimum[1] - (HEADER_Y + PITCH * 9 + 30)) <= 3


def test_the_annotated_map_is_saved(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake = _open_window(ocr, anchors)
    win = a.read_update_window()
    path = a.annotate_window(win, a.locate_date(D("27 Nov 2026")))
    assert path.exists() and Image.open(path).size == win.image.size


def test_a_missing_window_is_a_clear_error(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake = make(ocr, anchors, reader="ocr")                                     # the update window is NOT open
    with pytest.raises(SophisError, match="Could not find the Maturity"):
        a.read_update_window()


def test_opening_the_window_reports_what_it_sees(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.sel = 3
    a.open_selected()
    assert any("update window read: 10 visible dates" in l for l in lines)
    assert any("map saved" in l for l in lines) and list(tmp_path.glob("*update_window_map.png"))


def test_cli_window_options_are_parsed(monkeypatch):
    seen = {}
    monkeypatch.setattr(sa, "cmd_window", lambda args: seen.update(date=args.date, occ=args.occurrence, show=args.show))
    monkeypatch.setattr(sys, "argv", ["sophis_auto.py", "window", "--date", "20 Nov 2026", "--occurrence", "1", "--show"])
    sa.main()
    assert seen == {"date": "20 Nov 2026", "occ": 1, "show": True}


# ── stage 2b: writing the values ─────────────────────────────────────────────

def _fill(ocr, anchors, entries, **kw):
    a, fake, lines = _logged(ocr, anchors, reader="ocr", **{k: v for k, v in kw.items() if k in ("shift",)})
    fake.dialog = True
    for attr in ("corrupt",):
        if attr in kw:
            setattr(fake, attr, kw[attr])
    result = a.fill_values([sa.parse_entry(*e) for e in entries], **{k: v for k, v in kw.items()
                                                                       if k in ("verify", "dry_run", "commit")})
    return a, fake, lines, result


def test_fill_writes_minimum_and_maximum_of_the_chosen_row(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("30 Nov 2026", "100", "120")])
    assert fake.values == {(8, "minimum"): "100", (8, "maximum"): "120"}
    assert res[0]["minimum"] == "ok" and res[0]["maximum"] == "ok"        # read back: grid shows 100.00 == 100


def test_each_cell_is_selected_double_clicked_typed_and_confirmed_with_enter(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("30 Nov 2026", "100", "120")])
    ev = [e for e in fake.events if e[0] in ("double", "type", "click")]
    # per cell: ONE click to select it, a double click to open it, then type; Enter confirms (no click after typing)
    assert [e[0] for e in ev] == ["click", "double", "type", "click", "double", "type"]
    assert fake.keys.count("enter") == 2
    y = HEADER_Y + PITCH * 9
    for sel, dbl in ((ev[0], ev[1]), (ev[3], ev[4])):
        assert sel[1:] == dbl[1:]                                                # same cell for both
    assert abs(ev[1][1] - HX["Minimum"]) <= 4 and abs(ev[1][2] - y) <= 3        # Minimum first
    assert abs(ev[4][1] - HX["Maximum"]) <= 4 and abs(ev[4][2] - y) <= 3        # then Maximum
    assert ev[2][1] == "100" and ev[5][1] == "120"
    assert fake.values == {(8, "minimum"): "100", (8, "maximum"): "120"}


def test_clicking_outside_the_cell_is_still_available(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("30 Nov 2026", "100", "120")], commit="title")
    ev = [e for e in fake.events if e[0] in ("double", "type", "click")]
    assert [e[0] for e in ev] == ["click", "double", "type", "click", "click", "double", "type", "click"]
    assert "enter" not in fake.keys
    for click in (ev[3], ev[7]):                                                 # the click outside: the window title
        assert abs(click[2] - 158) <= 8
    assert fake.values == {(8, "minimum"): "100", (8, "maximum"): "120"}


def test_the_ric_cell_can_be_used_to_leave_a_cell_instead(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("30 Nov 2026", "100", "120")], commit="ric")
    kinds = [e[0] for e in fake.events]
    outside = [fake.events[i] for i in range(1, len(kinds)) if kinds[i] == "click" and kinds[i - 1] == "type"]
    assert len(outside) == 2 and all(abs(e[1] - HX["Ric"]) <= 6 for e in outside) and "enter" not in fake.keys
    assert fake.values[(8, "maximum")] == "120"


def test_existing_values_are_replaced_not_appended(ocr, anchors):
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.dialog = True
    fake.values[(7, "minimum")] = "55"
    a.fill_values([sa.parse_entry("27 Nov 2026", "95", "130")])
    assert fake.values[(7, "minimum")] == "95" and fake.values[(7, "maximum")] == "130"


def test_a_date_that_needs_scrolling_is_filled_too(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("16 Apr 2027", "10", "20")])
    assert fake.values == {(15, "minimum"): "10", (15, "maximum"): "20"} and fake.dscroll > 0


def test_several_dates_in_one_go(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("23 Oct 2026", "1", "2"), ("27 Nov 2026", "3.5", "4.5"),
                                              ("16 Apr 2027", "7", "8")])
    assert fake.values == {(1, "minimum"): "1", (1, "maximum"): "2", (7, "minimum"): "3.5", (7, "maximum"): "4.5",
                           (15, "minimum"): "7", (15, "maximum"): "8"}


def test_a_repeated_date_is_never_guessed(ocr, anchors):
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.dialog = True
    with pytest.raises(SophisError, match="appears 2 times.*#0.*#1"):
        a.fill_values([sa.parse_entry("23 Oct 2026", "1", "2"), sa.parse_entry("20 Nov 2026", "5", "6")])
    assert fake.values == {} and not [e for e in fake.events if e[0] in ("double", "type")]   # not even the first


@pytest.mark.parametrize("which,row", [("#0", 5), ("#1", 6)])
def test_a_repeated_date_with_its_occurrence(ocr, anchors, which, row):
    a, fake, lines, res = _fill(ocr, anchors, [("20 Nov 2026" + which, "5", "6")])
    assert fake.values == {(row, "minimum"): "5", (row, "maximum"): "6"}


def test_a_value_that_did_not_take_stops_everything(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.dialog, fake.corrupt = True, True                   # the grid stores 101 when 100 was typed
    with pytest.raises(SophisError, match="shows 101.0 but 100.0 was typed"):
        a.fill_values([sa.parse_entry("23 Oct 2026", "100", "120"), sa.parse_entry("27 Nov 2026", "1", "2")])
    assert list(fake.values) == [(1, "minimum")]            # stopped right after the first cell


def test_an_unreadable_cell_only_warns(ocr, anchors):
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.dialog = True
    a._read_cell_number = lambda *args: (None, [])
    res = a.fill_values([sa.parse_entry("23 Oct 2026", "100", "120")])
    assert res[0]["minimum"] == "unverified" and fake.values[(1, "minimum")] == "100"
    assert any("check it by eye" in l for l in lines)


def test_dry_run_only_locates(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("23 Oct 2026", "1", "2")], dry_run=True)
    assert res[0]["written"] is False and res[0]["row"] == 1 and fake.values == {}
    assert not [e for e in fake.events if e[0] in ("double", "type", "click")]


def test_fill_works_with_the_window_somewhere_else(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("30 Nov 2026", "100", "120")], shift=(45, 25))
    assert fake.values == {(8, "minimum"): "100", (8, "maximum"): "120"}


def test_ok_is_never_pressed(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("30 Nov 2026", "100", "120")])
    assert fake.ok_clicks == 0 and fake.dialog                 # the window stays open for you to check
    assert any("press OK yourself" in l for l in lines)
    assert not hasattr(a, "press_ok")


def test_the_values_must_be_numbers(ocr, anchors):
    with pytest.raises(SophisError, match="not a number"):
        sa.parse_entry("23 Oct 2026", "abc", "2")
    with pytest.raises(SophisError, match="Could not understand the date"):
        sa.parse_entry("someday", "1", "2")


@pytest.mark.parametrize("text,value", [("100", 100.0), ("100.00", 100.0), ("1,25", 1.25), ("1,234.50", 1234.5),
                                        ("-0.5", -0.5), ("1O0", 100.0), ("abc", None), ("", None)])
def test_parse_number(text, value):
    assert sa.parse_number(text) == value


def test_cli_fill_has_no_ok_option_and_parses_entries(monkeypatch):
    seen = {}
    monkeypatch.setattr(sa, "cmd_fill", lambda args: seen.update(entry=args.entry, dry=args.dry_run,
                                                                 commit=args.commit_click))
    monkeypatch.setattr(sys, "argv", ["sophis_auto.py", "fill", "--entry", "20 Nov 2026#1", "100", "120",
                                      "--entry", "27 Nov 2026", "95", "130", "--dry-run"])
    sa.main()
    assert seen == {"entry": [["20 Nov 2026#1", "100", "120"], ["27 Nov 2026", "95", "130"]], "dry": True,
                    "commit": "enter"}
    monkeypatch.setattr(sys, "argv", ["sophis_auto.py", "fill", "--entry", "20 Nov 2026", "1", "2", "--ok"])
    with pytest.raises(SystemExit):
        sa.main()                                              # --ok no longer exists


def test_cli_fill_without_entries_explains_itself(monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["sophis_auto.py", "fill"])
    with pytest.raises(SystemExit) as e:
        sa.main()
    assert e.value.code == 2 and "--entry" in capsys.readouterr().out


# ── robustness of reading the update window (it used to depend on one full-screen OCR every time) ──

class _CountingOCR:
    """Wraps the real OCR and counts how big each `read` was, so the tests can see what is OCR'd and how often."""

    def __init__(self, real):
        self.real, self.calls, self.fail_next = real, [], 0

    def read(self, img, *args, **kwargs):
        self.calls.append(img.size)
        if self.fail_next > 0:
            self.fail_next -= 1
            return []
        return self.real.read(img, *args, **kwargs)

    def __getattr__(self, name):
        return getattr(self.real, name)


def _window_with_counting_ocr(ocr, anchors, **kw):
    a, fake = _open_window(ocr, anchors, **kw)
    counter = _CountingOCR(ocr)
    a.ocr = counter
    return a, fake, counter


def test_after_the_first_look_the_header_is_followed_as_an_image_not_with_full_screen_ocr(ocr, anchors):
    a, fake, counter = _window_with_counting_ocr(ocr, anchors)
    a.read_update_window()
    first = len(counter.calls)
    counter.calls.clear()
    win = a.read_update_window()
    assert first >= 1 and len(win.dated()) == VISIBLE
    screen = fake.capture().size
    assert all(w * h < 0.25 * screen[0] * screen[1] for w, h in counter.calls), counter.calls   # strips only
    assert len(counter.calls) <= 3                                                          # strip + a few row slots


def test_the_followed_header_tracks_the_window_when_it_moves_between_reads(ocr, anchors):
    a, fake, counter = _window_with_counting_ocr(ocr, anchors)
    first = a.read_update_window()
    fake.dx, fake.dy = 35, 22                                    # the window is dragged somewhere else
    moved = a.read_update_window()
    assert moved.x_minimum - first.x_minimum == 35 and moved.header_y - first.header_y == 22
    assert [r.date for r in moved.rows] == [r.date for r in first.rows]


def test_a_failed_read_is_retried_instead_of_aborting_everything(ocr, anchors):
    a, fake, counter = _window_with_counting_ocr(ocr, anchors)
    a.read_update_window()                                       # the window is known now
    counter.fail_next = 1                                        # the next OCR pass sees nothing (screen redraw...)
    win = a.read_update_window()
    assert len(win.dated()) == VISIBLE


def test_a_header_missed_by_the_whole_screen_pass_is_found_in_tiles(ocr, anchors):
    a, fake, counter = _window_with_counting_ocr(ocr, anchors)
    counter.fail_next = 1                                        # the whole-screen pass fails; the tiles must work
    win = a.read_update_window()
    assert abs(win.x_minimum - HX["Minimum"]) <= 4 and len(win.dated()) == VISIBLE


def test_when_the_header_cannot_be_found_you_get_the_screenshot_and_what_the_ocr_read(ocr, anchors, tmp_path,
                                                                                     monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake = make(ocr, anchors, reader="ocr")                   # update window not open
    with pytest.raises(SophisError, match="what the OCR read"):
        a.read_update_window()
    assert list(tmp_path.glob("*update_window_header_not_found.png")) and list(tmp_path.glob("*ocr_lines.txt"))


def test_fill_saves_a_picture_of_every_step_when_tracing(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.dialog, a.trace = True, True
    a.fill_values([sa.parse_entry("30 Nov 2026", "100", "120")])
    names = sorted(p.name for p in tmp_path.glob("*fill*.png"))
    assert len(names) == 10                                      # 5 steps x 2 cells
    assert any("minimum_1b_selected" in n for n in names) and any("minimum_2_after_double_click" in n for n in names)
    assert any("maximum_3_typed" in n for n in names) and any("maximum_4_committed" in n for n in names)


def test_without_tracing_no_step_pictures_are_saved(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.dialog = True
    a.fill_values([sa.parse_entry("30 Nov 2026", "100", "120")])
    assert not list(tmp_path.glob("*fill*.png"))


def test_fill_reads_the_window_with_strips_not_with_the_whole_screen_every_time(ocr, anchors):
    a, fake, counter = _window_with_counting_ocr(ocr, anchors)
    a.fill_values([sa.parse_entry("30 Nov 2026", "100", "120")])
    screen = fake.capture().size
    big = [c for c in counter.calls if c[0] * c[1] >= 0.25 * screen[0] * screen[1]]
    assert len(big) <= 1 + 6 + 6, big                            # only the first look may use the full screen / tiles


# ── the header is learned once and remembered between runs ──────────────────

def test_the_header_is_remembered_on_disk_after_the_first_good_read(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    a.read_update_window()
    png, meta = a._window_cache_files()
    assert png.exists() and meta.exists()
    saved = json.loads(meta.read_text())
    assert saved["pitch"] and saved["n_rows"] and saved["header"]["maturity"]


def test_a_later_run_reads_the_window_even_if_the_ocr_cannot_find_the_header_words(ocr, anchors):
    """The reported failure: header words not recognized this time. After one good read, that no longer matters."""
    a, fake = _open_window(ocr, anchors)
    a.read_update_window()                                          # run 1: learns the header
    b, fake_b = make(ocr, anchors, reader="ocr")                    # run 2: a brand new Automator, same folder
    fake_b.dialog = True
    counter = _CountingOCR(ocr)
    b.ocr = counter
    screen = fake_b.capture().size
    real_read = counter.real.read

    def read_but_blind_on_big_images(img, *args, **kwargs):          # whole-screen / tile OCR sees nothing
        counter.calls.append(img.size)
        return [] if img.size[0] * img.size[1] >= 0.2 * screen[0] * screen[1] else real_read(img, *args, **kwargs)
    counter.read = read_but_blind_on_big_images
    win = b.read_update_window()
    assert [r.date for r in win.rows] == [D(t) for t in DATES[:VISIBLE]]


def test_a_remembered_header_that_no_longer_matches_is_relearned(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    a.read_update_window()
    png, _ = a._window_cache_files()
    Image.new("RGB", (120, 20), (0, 0, 0)).save(png)                # a stale / wrong template
    b, fake_b = make(ocr, anchors, reader="ocr")
    fake_b.dialog = True
    win = b.read_update_window()                                     # falls back to OCR and relearns
    assert len(win.dated()) == VISIBLE and Image.open(png).size != (120, 20)


def test_fill_works_from_a_remembered_header_without_any_big_ocr(ocr, anchors):
    a, fake = _open_window(ocr, anchors)
    a.read_update_window()
    b, fake_b = make(ocr, anchors, reader="ocr")
    fake_b.dialog = True
    counter = _CountingOCR(ocr)
    b.ocr = counter
    b.fill_values([sa.parse_entry("27 Nov 2026", "95", "130")])
    assert fake_b.values == {(7, "minimum"): "95", (7, "maximum"): "130"}
    screen = fake_b.capture().size
    assert not [c for c in counter.calls if c[0] * c[1] >= 0.2 * screen[0] * screen[1]]


# ── the cell must really be open before anything is typed ───────────────────

def _grid(ocr, anchors, mode, **kw):
    a, fake, lines = _logged(ocr, anchors, reader="ocr")
    fake.dialog, fake.edit_mode = True, mode
    return a, fake, lines


def test_a_double_click_on_an_unselected_row_only_selects_so_the_cell_is_selected_first(ocr, anchors):
    """The reported bug: double click on a row that is not selected yet never opens the editor."""
    a, fake, lines = _grid(ocr, anchors, "select_first")
    assert fake.dsel == 0
    a.fill_values([sa.parse_entry("30 Nov 2026", "100", "120")])               # row 8 is NOT the selected row
    assert fake.values == {(8, "minimum"): "100", (8, "maximum"): "120"} and fake.stray == ""


def test_the_old_way_double_click_only_would_have_failed_here(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "select_first")
    fake.double_click(HX["Minimum"], HEADER_Y + PITCH * 9)                      # what the old code did first
    assert fake.cell_edit is None and fake.dsel == 8                            # selected, but not open


def test_f2_is_tried_when_double_click_does_not_open_the_cell(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "f2")
    a.fill_values([sa.parse_entry("27 Nov 2026", "95", "130")])
    assert fake.values == {(7, "minimum"): "95", (7, "maximum"): "130"} and "f2" in fake.keys and fake.stray == ""


def test_if_the_cell_never_opens_nothing_is_typed(ocr, anchors, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "DEBUG_DIR", tmp_path)
    a, fake, lines = _grid(ocr, anchors, "never")
    with pytest.raises(SophisError, match="did not open for editing.*NOTHING was typed.*--no-edit-check"):
        a.fill_values([sa.parse_entry("27 Nov 2026", "95", "130")])
    assert not [e for e in fake.events if e[0] == "type"] and fake.stray == "" and fake.values == {}
    assert fake.dsel == 7                                                       # it only selected the row


def test_no_edit_check_types_anyway(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "never")
    a.fill_values([sa.parse_entry("27 Nov 2026", "95", "130")], check_edit=False, verify=False)
    assert fake.stray == "95130" and [e for e in fake.events if e[0] == "type"]   # typed (and went nowhere here)
    assert any("typing anyway" in l for l in lines)


def test_ctrl_a_is_only_used_when_the_cell_already_holds_a_value(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "double")
    a.fill_values([sa.parse_entry("30 Nov 2026", "100", "120")])               # empty cells
    assert "ctrl+a" not in fake.keys
    fake.keys.clear()
    fake.values[(8, "minimum")] = "55"                                         # only Minimum holds a value now
    del fake.values[(8, "maximum")]
    a.fill_values([sa.parse_entry("30 Nov 2026", "77", "120")])
    assert fake.keys.count("ctrl+a") == 1 and fake.values[(8, "minimum")] == "77"


def test_an_already_selected_row_works_too(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "select_first")
    a.fill_values([sa.parse_entry("19 Oct 2026", "1", "2")])                    # row 0 is the selected one
    assert fake.values == {(0, "minimum"): "1", (0, "maximum"): "2"} and fake.stray == ""


def test_cli_no_edit_check_option(monkeypatch):
    seen = {}
    monkeypatch.setattr(sa, "cmd_fill", lambda args: seen.update(flag=args.no_edit_check))
    monkeypatch.setattr(sys, "argv", ["sophis_auto.py", "fill", "--entry", "20 Nov 2026", "1", "2", "--no-edit-check"])
    sa.main()
    assert seen == {"flag": True}


# ── decimal points must survive the cell cleaning ("100.00" used to be read as "10000") ──

@pytest.mark.parametrize("text,value", [("100.00", 100.0), ("95.50", 95.5), ("1234.50", 1234.5), ("0.25", 0.25),
                                        ("120.00", 120.0), ("7.00", 7.0), ("15.75", 15.75)])
@pytest.mark.parametrize("dark", [False, True])
def test_numbers_with_decimals_are_read_back_whole(ocr, text, value, dark):
    readings, glyphs = ocr.read_cell(_cell(text, dark=dark))
    assert any(sa.parse_number(r) == value for r in readings), readings


def test_the_dotted_focus_rectangle_is_not_mistaken_for_decimal_points(ocr):
    cell = _cell("100", dark=False)
    d = ImageDraw.Draw(cell)
    for x in range(3, cell.width - 12, 3):                              # a dotted focus rectangle around the text
        d.point((x, 1), fill=(0, 0, 0))
        d.point((x, cell.height - 2), fill=(0, 0, 0))
    readings, glyphs = ocr.read_cell(cell)
    assert any(sa.parse_number(r) == 100.0 for r in readings), readings
    assert glyphs == 3


# ── Enter moves the selection / may scroll the list: the next cell must still land on the right row ──

def test_enter_moves_the_selection_but_both_cells_of_the_row_get_their_value(ocr, anchors):
    a, fake, lines, res = _fill(ocr, anchors, [("27 Nov 2026", "95", "130"), ("30 Nov 2026", "1", "2")])
    assert fake.values == {(7, "minimum"): "95", (7, "maximum"): "130", (8, "minimum"): "1", (8, "maximum"): "2"}
    assert fake.stray == ""


def test_the_list_scrolling_on_enter_does_not_send_the_maximum_to_another_row(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "double")
    fake.enter_scrolls = True                                    # Enter on the last visible row shifts the list
    a.fill_values([sa.parse_entry("18 Dec 2026", "5", "6")])     # row 9 is the last visible row
    assert fake.dscroll == 1
    assert fake.values == {(9, "minimum"): "5", (9, "maximum"): "6"}      # the date's own row, not the row below it


def test_when_nothing_moved_the_second_cell_does_not_search_the_date_again(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "double")
    calls = []
    real = a.locate_date
    a.locate_date = lambda *args, **kw: (calls.append(1), real(*args, **kw))[1]
    a.fill_values([sa.parse_entry("27 Nov 2026", "95", "130")])
    assert len(calls) == 2                                       # the plan, and the first cell; Maximum reuses the row
    assert fake.values == {(7, "minimum"): "95", (7, "maximum"): "130"}


def test_a_repeated_date_is_safe_with_the_fast_second_cell(ocr, anchors):
    a, fake, lines = _grid(ocr, anchors, "double")
    fake.enter_scrolls = False
    a.fill_values([sa.parse_entry("20 Nov 2026#1", "5", "6")])   # the second of two identical dates
    assert fake.values == {(6, "minimum"): "5", (6, "maximum"): "6"}

"""
Tests for sophis_auto.py WITHOUT Sophis or a real screen: a "fake screen" draws an imitation of the
window (grid, dropdown, update window) with PIL and the real OCR reads it.

    pip install pytest rapidocr-onnxruntime pillow opencv-python
    pytest -q test_sophis_auto.py
"""

import json

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


def test_rows_with_color():
    img = Image.new("RGB", (7, 72), (255, 255, 255))
    ImageDraw.Draw(img).rectangle((0, 24, 6, 47), fill=(170, 195, 235))
    colors = sa.row_colors(img, 24, 3)
    assert sa.rows_with_color(colors, (170, 195, 235), 20) == [1]


# ── fake screen ─────────────────────────────────────────────────────────────

FONT = ImageFont.load_default(size=13)
ROWS = [("ZYS", "Short term share"), ("ZTS", "Short term index"), ("ZTO", "Long term share"),
        ("ZS", "Long term index"), ("ZQK", "Future"), ("ZMH", "Short term share"),
        ("ZM", "Long term share"), ("ZJO", "Short term index")]
Y1, H, BORDER, RIGHT = 100, 24, 470, 620
ARROW = (836, 62)

ANCHOR_GRID = (22, 62, 330, 84)      # header "Mnemo | Extended Mnemo | Description"
ANCHOR_RIGHT = (700, 188, 850, 208)  # header "Date | Cycle | Name"


class FakeSophis:
    def __init__(self, clipboard=False, popup_visible=True, shift=(0, 0), hidden=False):
        self.dx, self.dy = shift
        self.hidden = hidden
        self.sel, self.typed, self.dd = 0, "", "Short term index"
        self.popup, self.widened, self.dialog = False, False, False
        self.popup_idx = 0
        self.clipboard, self.popup_visible = clipboard, popup_visible
        self._clip = None
        self.key_interval = 0
        self.events = []       # ("click" | "double" | "type", ...) in the order they happened

    # drawing
    def _image(self):
        im = Image.new("RGB", (1000, 600), (200, 200, 200))
        d = ImageDraw.Draw(im)
        border = BORDER - (140 if self.widened else 0)
        for i, (m, t) in enumerate(ROWS):
            y = Y1 - H // 2 + i * H
            color = (170, 195, 235) if i == self.sel else ((255, 255, 255) if i % 2 == 0 else (235, 235, 235))
            d.rectangle((20, y, 640, y + H - 1), fill=color)
            d.text((40, y + 4), m, fill=(0, 0, 160), font=FONT)
            d.text((border + 6, y + 4), t if self.widened else t[:9] + "...", fill=(0, 0, 160), font=FONT)
        d.rectangle(ANCHOR_GRID, fill=(215, 215, 215))
        d.text((30, 66), "Mnemo        Extended Mnemo        Description", fill=(0, 0, 0), font=FONT)
        d.rectangle(ANCHOR_RIGHT, fill=(215, 215, 215))
        d.text((706, 192), "Date     Cycle     Name", fill=(0, 0, 0), font=FONT)
        d.text((BORDER + 6, 62), "Term", fill=(0, 0, 0), font=FONT)
        d.rectangle((700, 50, 850, 74), fill=(255, 255, 255), outline=(0, 0, 0))
        d.text((706, 55), self.dd, fill=(0, 0, 0), font=FONT)
        d.polygon([(830, 58), (842, 58), (836, 67)], fill=(0, 0, 0))
        if self.popup and self.popup_visible:
            d.rectangle((700, 76, 850, 76 + 22 * 5), fill=(255, 255, 255), outline=(0, 0, 0))
            for i, t in enumerate(TERMS):
                d.text((706, 80 + 22 * i), t, fill=(0, 0, 0), font=FONT)
        if self.dialog:
            d.rectangle((200, 150, 800, 450), fill=(240, 240, 240), outline=(0, 0, 0))
            d.text((220, 170), "Maturity        Minimum        Maximum        Ric", fill=(0, 0, 0), font=FONT)
        canvas = Image.new("RGB", (1200, 800), (30, 60, 90))
        if self.hidden:                        # minimized window: only the desktop is visible
            return canvas
        canvas.paste(im, (self.dx, self.dy))   # the window appears shifted
        return canvas

    # Screen API (receives ABSOLUTE screen coordinates)
    def capture(self, region=None):
        im = self._image()
        return im if region is None else im.crop(tuple(int(v) for v in region))

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
        if self.popup:
            if 700 <= x <= 850 and 76 <= y <= 76 + 110:
                self.dd = TERMS[min(4, (y - 76) // 22)]
            self.popup = False
            return
        if abs(x - ARROW[0]) <= 14 and abs(y - ARROW[1]) <= 14:
            self.popup, self.popup_idx = True, 0
            return
        i = self._row_at(x, y)
        if i is not None:
            self.sel, self.typed = i, ""

    def double_click(self, x, y):
        self.events.append(("double", x, y))
        self._click(x, y)
        if self._row_at(x - self.dx, y - self.dy) is not None:
            self.dialog = True

    def type_text(self, text):
        self.events.append(("type", text))
        for ch in text:
            self.typed += ch.upper()
            for i, (m, _) in enumerate(ROWS):
                if m.startswith(self.typed):
                    self.sel = i
                    break

    def press(self, key):
        if not self.popup and key == "down":
            self.sel = min(self.sel + 1, len(ROWS) - 1)
        elif self.popup and key == "home":
            self.popup_idx = 0
        elif self.popup and key == "down":
            self.popup_idx += 1
        elif self.popup and key == "enter":
            self.dd, self.popup = TERMS[self.popup_idx], False

    def hotkey(self, *keys):
        if self.clipboard:
            m, t = ROWS[self.sel]
            self._clip = f"{m}\t\tDESCRIPTION\tUS0000\t\t{t}"

    def drag(self, x0, y0, x1, y1):
        if abs(x0 - self.dx - BORDER) <= 4:
            self.widened = True

    def clipboard_set(self, text):
        self._clip = text

    def clipboard_get(self):
        return self._clip


@pytest.fixture(scope="module")
def ocr():
    return OCR()


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
        "term_left": [BORDER, 70], "term_right": [RIGHT, 70], "dd_left": [706, 62], "dd_arrow": list(ARROW),
    }, slowdown=0.01, **{k: v for k, v in kw.items() if k in ("reader", "verify_mnemonic")})
    screen = FakeSophis(**{k: v for k, v in kw.items() if k in ("clipboard", "popup_visible", "shift", "hidden")})
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
def test_a_single_click_before_typing(ocr, anchors, reader, clipboard):
    a, fake = make(ocr, anchors, reader=reader, clipboard=clipboard)
    assert fake.events == []                      # preparing must not click anything
    a.search("ZS")
    kinds = [e[0] for e in fake.events]
    assert kinds[:2] == ["click", "type"], kinds  # one click, then straight to typing


def test_each_code_gets_one_click_then_typing_until_the_window_opens(ocr, anchors):
    a, fake = make(ocr, anchors, reader="ocr")
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


GOOD_POINTS = [(22, 62), (330, 84), (45, Y1), (45, Y1 + H), (45, Y1 + 7 * H), (140, Y1), (BORDER, 70), (RIGHT, 70),
               (700, 188), (850, 208), (706, 62), ARROW]


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

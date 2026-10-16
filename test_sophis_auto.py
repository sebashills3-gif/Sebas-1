"""
Pruebas de sophis_auto.py SIN Sophis ni pantalla real: una "pantalla falsa" dibuja con PIL una
imitación de la ventana (grid, dropdown, ventana de actualización) y el OCR real la lee.

    pip install pytest rapidocr-onnxruntime pillow
    pytest -q test_sophis_auto.py
"""

import pytest
from PIL import Image, ImageDraw, ImageFont

import sophis_auto as sa
from sophis_auto import Automatizador, Config, ErrorSophis, OCR, TERMS

# ── pruebas de funciones puras ──────────────────────────────────────────────

@pytest.mark.parametrize("texto,esperado", [
    ("Short term index", "Short term index"),
    ("ShortTermIndex", "Short term index"),            # OCR junta palabras
    ("short term lndex", "Short term index"),          # I -> l
    ("Long Term Shares", "Long term share"),           # plural
    ("  LONG TERM INDEX ", "Long term index"),
    ("Future", "Future"),
    ("Short term share", "Short term share"),
    ("ZYS.N Short term share", "Short term share"),    # texto extra en el mismo recorte
])
def test_clasificar_term_ok(texto, esperado):
    assert sa.clasificar_term(texto) == esperado


@pytest.mark.parametrize("texto", ["", "Short ter...", "Short ter", "Cycle", "US92189H6071"])
def test_clasificar_term_ambiguo(texto):
    assert sa.clasificar_term(texto) is None


def test_term_truncado():
    assert sa.term_truncado("Short ter...")
    assert not sa.term_truncado("Short term index")


def test_los_cinco_terms_se_distinguen_entre_si():
    for t in TERMS:
        assert sa.clasificar_term(t) == t


def test_filas_con_color():
    img = Image.new("RGB", (7, 72), (255, 255, 255))
    ImageDraw.Draw(img).rectangle((0, 24, 6, 47), fill=(170, 195, 235))
    cols = sa.colores_filas(img, 24, 3)
    assert sa.filas_con_color(cols, (170, 195, 235), 20) == [1]


# ── pantalla falsa ──────────────────────────────────────────────────────────

FUENTE = ImageFont.load_default(size=13)
FILAS = [("ZYS", "Short term share"), ("ZTS", "Short term index"), ("ZTO", "Long term share"),
         ("ZS", "Long term index"), ("ZQK", "Future"), ("ZMH", "Short term share"),
         ("ZM", "Long term share"), ("ZJO", "Short term index")]
Y1, H, BORDE, DER = 100, 24, 470, 620
FLECHA = (836, 62)


class FakeSophis:
    def __init__(self, portapapeles=False, popup_visible=True):
        self.sel, self.typed, self.dd = 0, "", "Short term index"
        self.popup, self.ancho, self.dialogo = False, False, False
        self.popup_idx = 0
        self.portapapeles, self.popup_visible = portapapeles, popup_visible
        self._clip = None
        self.intervalo_tecla = 0

    # dibujo
    def _imagen(self):
        im = Image.new("RGB", (1000, 600), (200, 200, 200))
        d = ImageDraw.Draw(im)
        borde = BORDE - (140 if self.ancho else 0)
        for i, (m, t) in enumerate(FILAS):
            y = Y1 - H // 2 + i * H
            color = (170, 195, 235) if i == self.sel else ((255, 255, 255) if i % 2 == 0 else (235, 235, 235))
            d.rectangle((20, y, 640, y + H - 1), fill=color)
            d.text((40, y + 4), m, fill=(0, 0, 160), font=FUENTE)
            d.text((borde + 6, y + 4), t if self.ancho else t[:9] + "...", fill=(0, 0, 160), font=FUENTE)
        d.text((BORDE + 6, 62), "Term", fill=(0, 0, 0), font=FUENTE)
        d.rectangle((700, 50, 850, 74), fill=(255, 255, 255), outline=(0, 0, 0))
        d.text((706, 55), self.dd, fill=(0, 0, 0), font=FUENTE)
        d.polygon([(830, 58), (842, 58), (836, 67)], fill=(0, 0, 0))
        if self.popup and self.popup_visible:
            d.rectangle((700, 76, 850, 76 + 22 * 5), fill=(255, 255, 255), outline=(0, 0, 0))
            for i, t in enumerate(TERMS):
                d.text((706, 80 + 22 * i), t, fill=(0, 0, 0), font=FUENTE)
        if self.dialogo:
            d.rectangle((200, 150, 800, 450), fill=(240, 240, 240), outline=(0, 0, 0))
            d.text((220, 170), "Maturity        Minimum        Maximum        Ric", fill=(0, 0, 0), font=FUENTE)
        return im

    # API de Pantalla
    def captura(self, region=None):
        im = self._imagen()
        return im if region is None else im.crop(tuple(int(v) for v in region))

    def _fila_en(self, x, y):
        for i in range(len(FILAS)):
            if abs(y - (Y1 + i * H)) <= H // 2 and 20 <= x <= 640:
                return i
        return None

    def click(self, x, y):
        if self.popup:
            if 700 <= x <= 850 and 76 <= y <= 76 + 110:
                self.dd = TERMS[min(4, (y - 76) // 22)]
            self.popup = False
            return
        if abs(x - FLECHA[0]) <= 14 and abs(y - FLECHA[1]) <= 14:
            self.popup, self.popup_idx = True, 0
            return
        i = self._fila_en(x, y)
        if i is not None:
            self.sel, self.typed = i, ""

    def doble_click(self, x, y):
        self.click(x, y)
        if self._fila_en(x, y) is not None:
            self.dialogo = True

    def escribir(self, texto):
        for ch in texto:
            self.typed += ch.upper()
            for i, (m, _) in enumerate(FILAS):
                if m.startswith(self.typed):
                    self.sel = i
                    break

    def tecla(self, nombre):
        if not self.popup and nombre == "down":
            self.sel = min(self.sel + 1, len(FILAS) - 1)
        elif self.popup and nombre == "home":
            self.popup_idx = 0
        elif self.popup and nombre == "down":
            self.popup_idx += 1
        elif self.popup and nombre == "enter":
            self.dd, self.popup = TERMS[self.popup_idx], False

    def hotkey(self, *teclas):
        if self.portapapeles:
            m, t = FILAS[self.sel]
            self._clip = f"{m}\t\tDESCRIPCION\tUS0000\t\t{t}"

    def arrastrar(self, x0, y0, x1, y1):
        if abs(x0 - BORDE) <= 4:
            self.ancho = True

    def portapapeles_set(self, txt):
        self._clip = txt

    def portapapeles_get(self):
        return self._clip


@pytest.fixture(scope="module")
def ocr():
    return OCR()


def crear(ocr, **kw):
    cfg = Config(puntos={
        "fila1": [45, Y1], "fila2": [45, Y1 + H], "filaN": [45, Y1 + 7 * H], "ext": [140, Y1],
        "term_borde": [BORDE, 70], "term_der": [DER, 70], "dd_izq": [706, 62], "dd_flecha": list(FLECHA),
    }, pausa=0.01, **{k: v for k, v in kw.items() if k in ("lector", "verificar_mnemo")})
    pantalla = FakeSophis(**{k: v for k, v in kw.items() if k in ("portapapeles", "popup_visible")})
    a = Automatizador(pantalla, ocr, cfg, log=lambda *_: None)
    a.preparar()
    return a, pantalla


# ── flujo completo contra la pantalla falsa ─────────────────────────────────

def test_geometria(ocr):
    a, _ = crear(ocr)
    assert a.g.n == 8 and a.g.h == 24


@pytest.mark.parametrize("codigo,term", [
    ("ZYS", "Short term share"), ("ZS", "Long term index"), ("ZQK", "Future"), ("ZM", "Long term share"),
])
def test_abrir_por_ocr(ocr, codigo, term):
    a, fake = crear(ocr, lector="ocr")
    r = a.abrir_contrato(codigo)
    assert r.term == term and fake.dd == term and fake.dialogo
    assert FILAS[fake.sel][0] == codigo


def test_abrir_por_portapapeles(ocr):
    a, fake = crear(ocr, portapapeles=True)
    r = a.abrir_contrato("ZTO")
    assert r.term == "Long term share" and fake.dd == "Long term share" and fake.dialogo


def test_dropdown_con_teclado_si_no_se_ve_la_lista(ocr):
    a, fake = crear(ocr, popup_visible=False, lector="ocr")
    a.abrir_contrato("ZTO")
    assert fake.dd == "Long term share"


def test_codigo_inexistente_aborta_sin_abrir(ocr, tmp_path, monkeypatch):
    monkeypatch.setattr(sa, "CARPETA_DEBUG", tmp_path)
    a, fake = crear(ocr, lector="ocr")
    with pytest.raises(ErrorSophis):
        a.abrir_contrato("QQQ")
    assert not fake.dialogo

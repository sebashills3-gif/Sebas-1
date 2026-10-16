#!/usr/bin/env python3
"""
Automatización de Sophis Risque (corre en Citrix) — ETAPA 1:
  buscar un código (Mnemo) -> leer su Term -> poner el dropdown igual -> abrir la ventana de
  actualización (Maturity / Minimum / Maximum / Ric).

Todavía NO lee el Excel ni llena las fechas (eso es la etapa 2).

Como Sophis corre en Citrix, este script NO puede ver los controles de la aplicación (solo recibe la
imagen). Por eso trabaja "como una persona": captura la pantalla, lee texto con OCR, y mueve
mouse/teclado. Se ejecuta en TU computadora, con la ventana de Citrix visible y al frente.

Instalación (Windows):
    pip install pyautogui pillow rapidocr-onnxruntime
    (si rapidocr no instala, usa Python 3.10–3.12)

Uso (en este orden la primera vez):
    1) Abre en Sophis la ventana "Listed market ..." (puede abrirse en cualquier posición: el script la
       vuelve a encontrar cada vez buscando el encabezado de la tabla; debe estar visible, no minimizada).
    2) python sophis_auto.py calibrar          # una sola vez: apuntas con el mouse a 12 puntos
    3) python sophis_auto.py diagnostico ZYS   # prueba cómo se lee tu pantalla (mándame lo que imprime)
    4) python sophis_auto.py abrir ZYS ZTS YY  # busca, iguala el dropdown y abre la ventana de cada código

Seguridad: mueve el mouse a una esquina de la pantalla para ABORTAR en cualquier momento (pyautogui failsafe).
Si algo no coincide, el script se detiene y guarda capturas en ./sophis_debug (nunca "adivina").
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
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════

ARCHIVO_CONFIG = Path("sophis_config.json")
CARPETA_DEBUG = Path("sophis_debug")

# Los 5 tipos de Term (así aparecen en la columna y en el dropdown).
TERMS = [
    "Short term index",
    "Long term index",
    "Short term share",
    "Long term share",
    "Future",
]

# Orden del dropdown (solo se usa como respaldo por teclado si el OCR no encuentra la opción).
ORDEN_DROPDOWN = TERMS

# La ventana no siempre se abre en la misma posición, así que NADA se guarda como posición fija de pantalla:
# se guardan dos "anclas" (recortes de imagen de zonas fijas del encabezado) y cada punto se calcula
# relativo a ellas. En cada ejecución se busca el ancla en la pantalla y se mueven todos los puntos.
ARCHIVO_ANCLAS = Path("sophis_anclas")
UMBRAL_ANCLA = 0.80   # parecido mínimo (0-1) para dar por encontrada una ancla

# Puntos que se calibran (nombre, instrucción). Se apunta con el mouse en TU pantalla, una sola vez.
PUNTOS_CALIBRACION = [
    ("ancla_grid_a", "ANCLA 1 (esquina SUPERIOR IZQUIERDA): el encabezado gris de la tabla, solo 'Mnemo | Extended Mnemo | "
                     "Description' (NO incluyas RIC ni Term). Mouse en la esquina superior izquierda de 'Mnemo'"),
    ("ancla_grid_b", "ANCLA 1 (esquina INFERIOR DERECHA): esquina inferior derecha de ese encabezado, al final de 'Description'"),
    ("fila1", "Mouse sobre el TEXTO de la celda Mnemo de la PRIMERA fila visible (parte izquierda del texto)"),
    ("fila2", "Mouse sobre el texto de la celda Mnemo de la SEGUNDA fila (misma posición horizontal)"),
    ("filaN", "Mouse sobre el texto de la celda Mnemo de la ÚLTIMA fila completamente visible"),
    ("ext", "Mouse sobre la columna 'Extended Mnemo' (vacía), parte izquierda, en cualquier fila"),
    ("term_borde", "Mouse en el ENCABEZADO, justo en la línea que separa RIC | Term (borde izquierdo de Term)"),
    ("term_der", "Mouse en el ENCABEZADO, en el borde derecho de la columna Term (antes de 'Cyc')"),
    ("ancla_der_a", "ANCLA 2 (esquina SUPERIOR IZQUIERDA): encabezado de la tabla de la derecha 'Date | Cycle | Name'"),
    ("ancla_der_b", "ANCLA 2 (esquina INFERIOR DERECHA): esquina inferior derecha de ese encabezado (después de 'Name')"),
    ("dd_izq", "Mouse DENTRO del dropdown ('Short term index'), al inicio del texto y centrado verticalmente"),
    ("dd_flecha", "Mouse sobre el botón/flecha del dropdown"),
]
# Esquina que cierra cada ancla (al capturarla se guarda el recorte de imagen).
ANCLA_CIERRE = {"ancla_grid_b": "grid", "ancla_der_b": "der"}
# A qué ancla pertenece cada punto (los puntos se mueven junto con su ancla).
PUNTO_ANCLA = {"fila1": "grid", "fila2": "grid", "filaN": "grid", "ext": "grid", "term_borde": "grid",
               "term_der": "grid", "dd_izq": "der", "dd_flecha": "der"}


@dataclass
class Config:
    puntos: dict                   # posiciones (de pantalla) al momento de calibrar
    anclas: dict = field(default_factory=dict)   # {"grid": {"archivo": ..., "origen": [x, y]}, "der": {...}}
    ensanchar_px: int = 140        # cuánto se ensancha la columna Term hacia la izquierda (0 = no tocar)
    intervalo_tecla: float = 0.12  # segundos entre teclas (Citrix puede perder teclas si es muy rápido)
    pausa: float = 1.0             # multiplicador de todas las esperas (súbelo si Citrix va lento)
    lector: str = "auto"           # auto | portapapeles | ocr
    verificar_mnemo: bool = True   # False = acepta la fila aunque no se pueda leer el Mnemo (NO recomendado)

    @staticmethod
    def cargar(ruta: Path = ARCHIVO_CONFIG) -> "Config":
        if not ruta.exists():
            raise ErrorSophis(f"No existe {ruta}. Corre primero:  python sophis_auto.py calibrar")
        d = json.loads(ruta.read_text(encoding="utf-8"))
        faltan = [n for n, _ in PUNTOS_CALIBRACION if n not in d.get("puntos", {})]
        if faltan:
            raise ErrorSophis(f"Calibración incompleta (faltan {faltan}). Corre:  python sophis_auto.py calibrar")
        if set(d.get("anclas", {})) != {"grid", "der"}:
            raise ErrorSophis("La calibración es de una versión anterior (sin anclas de imagen). "
                              "Corre de nuevo:  python sophis_auto.py calibrar")
        return Config(**d)

    def guardar(self, ruta: Path = ARCHIVO_CONFIG) -> None:
        ruta.write_text(json.dumps(self.__dict__, indent=2), encoding="utf-8")


class ErrorSophis(RuntimeError):
    pass


# ════════════════════════════════════════════════════════════════════════════
# UTILIDADES PURAS (texto / color) — sin pantalla, fáciles de probar
# ════════════════════════════════════════════════════════════════════════════

def _norm(s: str) -> str:
    """Minúsculas y solo letras/números: el OCR a veces junta o separa palabras ('ShortTermIndex')."""
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _norm_codigo(s: str) -> str:
    return re.sub(r"\s+", "", s).upper()


def _mejor_ventana(texto_norm: str, objetivo_norm: str) -> float:
    """Mejor parecido (0-1) entre `objetivo_norm` y cualquier tramo de `texto_norm` de largo similar."""
    L = len(objetivo_norm)
    if not texto_norm:
        return 0.0
    mejor = 0.0
    for largo in (L - 1, L, L + 1):
        if largo <= 0:
            continue
        if len(texto_norm) <= largo:
            mejor = max(mejor, difflib.SequenceMatcher(None, texto_norm, objetivo_norm).ratio())
            continue
        for i in range(len(texto_norm) - largo + 1):
            r = difflib.SequenceMatcher(None, texto_norm[i:i + largo], objetivo_norm).ratio()
            mejor = max(mejor, r)
    return mejor


def clasificar_term(texto: str, umbral: float = 0.82, margen: float = 0.06) -> str | None:
    """Devuelve uno de TERMS (tolerando errores de OCR) o None si no es claro."""
    n = _norm(texto).replace("shares", "share")
    if not n:
        return None
    puntajes = sorted(((_mejor_ventana(n, _norm(t)), t) for t in TERMS), reverse=True)
    (p1, t1), (p2, _) = puntajes[0], puntajes[1]
    if p1 >= umbral and (p1 - p2) >= margen:
        return t1
    return None


def term_truncado(texto: str) -> bool:
    """True si parece 'Short ter...' (la columna Term está angosta)."""
    for trozo in texto.split(" | "):      # varias lecturas del mismo recorte van unidas con ' | '
        n = _norm(trozo)
        if clasificar_term(trozo) is None and (n.startswith("short") or n.startswith("long")) and len(n) < 12:
            return True
    return False


def contiene_palabra(texto: str, palabra: str, umbral: float = 0.8) -> bool:
    return _mejor_ventana(_norm(texto), _norm(palabra)) >= umbral


def distancia_color(a, b) -> int:
    return sum(abs(int(x) - int(y)) for x, y in zip(a[:3], b[:3]))


def colores_filas(img: Image.Image, alto_fila: float, n: int) -> list[tuple]:
    """Color promedio del centro de cada fila en una tira vertical (img empieza en el borde superior de la fila 0)."""
    px = img.convert("RGB")
    cx = px.width // 2
    colores = []
    for i in range(n):
        yc = int(round(alto_fila / 2 + i * alto_fila))
        caja = (max(0, cx - 2), max(0, yc - 2), min(px.width, cx + 3), min(px.height, yc + 3))
        colores.append(px.crop(caja).resize((1, 1), Image.BOX).getpixel((0, 0)))
    return colores


def filas_con_color(colores: list[tuple], ref: tuple, tol: int) -> list[int]:
    return [i for i, c in enumerate(colores) if distancia_color(c, ref) <= tol]


def localizar_plantilla(captura: Image.Image, plantilla: Image.Image) -> tuple[int, int, float]:
    """Busca `plantilla` dentro de `captura`. Devuelve (x, y, parecido 0-1) de su esquina superior izquierda."""
    import cv2
    import numpy as np

    a = cv2.cvtColor(np.asarray(captura.convert("RGB")), cv2.COLOR_RGB2GRAY)
    t = cv2.cvtColor(np.asarray(plantilla.convert("RGB")), cv2.COLOR_RGB2GRAY)
    if t.shape[0] > a.shape[0] or t.shape[1] > a.shape[1]:
        return 0, 0, 0.0
    res = cv2.matchTemplate(a, t, cv2.TM_CCOEFF_NORMED)
    _, mejor, _, pos = cv2.minMaxLoc(res)
    return int(pos[0]), int(pos[1]), float(mejor)


def calcular_puntos(captura: Image.Image, cfg: "Config") -> tuple[dict, dict]:
    """Encuentra las anclas en la captura y devuelve (puntos en coordenadas ACTUALES, info por ancla).
    Lanza ErrorSophis si no está la ventana a la vista."""
    puntos, info = {}, {}
    for nombre, anc in cfg.anclas.items():
        ruta = Path(anc["archivo"])
        if not ruta.exists():
            raise ErrorSophis(f"Falta la imagen de ancla {ruta}. Corre de nuevo:  python sophis_auto.py calibrar")
        x, y, score = localizar_plantilla(captura, Image.open(ruta))
        info[nombre] = (x - anc["origen"][0], y - anc["origen"][1], score)
        if score < UMBRAL_ANCLA:
            raise ErrorSophis(
                f"No encuentro la ventana de Sophis en pantalla (ancla '{nombre}', parecido {score:.2f} < {UMBRAL_ANCLA}). "
                "¿Está visible (no minimizada ni tapada) y en el monitor principal? Si cambió el zoom/resolución "
                "de Citrix, vuelve a calibrar.")
    for pt, nombre in PUNTO_ANCLA.items():
        dx, dy, _ = info[nombre]
        px, py = cfg.puntos[pt]
        puntos[pt] = [px + dx, py + dy]
    return puntos, info


# ════════════════════════════════════════════════════════════════════════════
# OCR
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Linea:
    texto: str
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
    """RapidOCR (pip, sin instalar programas aparte)."""

    def __init__(self) -> None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as e:
            raise ErrorSophis("Falta el OCR. Instala:  pip install rapidocr-onnxruntime") from e
        self._motor = RapidOCR()

    def leer(self, img: Image.Image, escala: int = 3, margen: int = 24) -> list[Linea]:
        import numpy as np

        img = img.convert("RGB")
        if escala != 1:
            img = img.resize((img.width * escala, img.height * escala), Image.LANCZOS)
        # Se rellena el borde con el color de fondo: los recortes pequeños sin margen casi nunca se detectan.
        a = np.asarray(img)
        borde = np.concatenate([a[0], a[-1], a[:, 0], a[:, -1]])
        fondo = tuple(int(v) for v in np.median(borde, axis=0))
        lienzo = Image.new("RGB", (img.width + 2 * margen, img.height + 2 * margen), fondo)
        lienzo.paste(img, (margen, margen))
        resultado, _ = self._motor(np.asarray(lienzo)[:, :, ::-1])  # RapidOCR espera BGR
        lineas = []
        for caja, texto, conf in resultado or []:
            xs = [p[0] for p in caja]
            ys = [p[1] for p in caja]
            lineas.append(Linea(
                texto=texto, conf=float(conf),
                x0=int((min(xs) - margen) / escala), y0=int((min(ys) - margen) / escala),
                x1=int((max(xs) - margen) / escala), y1=int((max(ys) - margen) / escala),
            ))
        return lineas

    def texto(self, img: Image.Image, escala: int = 3) -> str:
        lineas = sorted(self.leer(img, escala), key=lambda l: (l.x0, l.y0))
        return " ".join(l.texto for l in lineas)


# ════════════════════════════════════════════════════════════════════════════
# PANTALLA (mouse / teclado / capturas / portapapeles)
# ════════════════════════════════════════════════════════════════════════════

def _activar_dpi() -> None:
    """Sin esto, en Windows con escala >100% las coordenadas del mouse y de las capturas no coinciden."""
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


class Pantalla:
    def __init__(self, intervalo_tecla: float = 0.12) -> None:
        _activar_dpi()
        try:
            import pyautogui
        except ImportError as e:
            raise ErrorSophis("Falta pyautogui. Instala:  pip install pyautogui pillow") from e
        pyautogui.FAILSAFE = True      # mouse a una esquina = abortar
        pyautogui.PAUSE = 0.05
        self.pg = pyautogui
        self.intervalo_tecla = intervalo_tecla
        self._tk = None

    def posicion(self) -> tuple[int, int]:
        p = self.pg.position()
        return int(p.x), int(p.y)

    def captura(self, region: tuple[int, int, int, int] | None = None) -> Image.Image:
        """region = (x0, y0, x1, y1) en coordenadas de pantalla."""
        if region is None:
            return self.pg.screenshot().convert("RGB")
        x0, y0, x1, y1 = (int(v) for v in region)
        return self.pg.screenshot(region=(x0, y0, max(1, x1 - x0), max(1, y1 - y0))).convert("RGB")

    def click(self, x: int, y: int) -> None:
        self.pg.moveTo(x, y, duration=0.15)
        self.pg.click()

    def doble_click(self, x: int, y: int) -> None:
        self.pg.moveTo(x, y, duration=0.15)
        self.pg.doubleClick(interval=0.12)

    def escribir(self, texto: str) -> None:
        for ch in texto:
            self.pg.write(ch)
            time.sleep(self.intervalo_tecla)

    def tecla(self, nombre: str) -> None:
        self.pg.press(nombre)

    def hotkey(self, *teclas: str) -> None:
        self.pg.hotkey(*teclas)

    def arrastrar(self, x0: int, y0: int, x1: int, y1: int) -> None:
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

    def portapapeles_set(self, texto: str) -> None:
        r = self._root()
        r.clipboard_clear()
        r.clipboard_append(texto)
        r.update()

    def portapapeles_get(self) -> str | None:
        r = self._root()
        r.update()
        try:
            return r.clipboard_get()
        except Exception:
            return None


# ════════════════════════════════════════════════════════════════════════════
# AUTOMATIZADOR
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class Geometria:
    x_click: int      # x donde se hace click en una celda Mnemo
    x_mnemo0: int
    x_mnemo1: int
    x_ext: int
    y1: int           # y del centro de la fila 0
    h: float          # alto de fila
    n: int            # filas visibles
    term_x0: int
    term_x1: int

    def yc(self, i: int) -> int:
        return int(round(self.y1 + i * self.h))


@dataclass
class Lectura:
    mnemo_ok: bool
    term: str | None
    fuente: str
    texto: str
    verificado: bool = True
    vistos: tuple = ()   # mnemos leídos en la fila (para detectar coincidencia solo por prefijo)
    term_txt: str = ""   # texto crudo leído de la celda Term (para detectar columna truncada)


@dataclass
class Resultado:
    codigo: str
    term: str
    fila: int
    ventana: dict


class Automatizador:
    def __init__(self, pantalla, ocr, cfg: Config, log=print) -> None:
        self.p = pantalla
        self.ocr = ocr
        self.cfg = cfg
        self.log = log
        self.abs: dict = {}
        self.g: Geometria
        self._ancho_aplicado = False
        self._color_sel: tuple | None = None
        self._tol = 30

    # ── utilidades ──────────────────────────────────────────────────────────
    def _esp(self, seg: float) -> None:
        time.sleep(seg * self.cfg.pausa)

    def _pt(self, nombre: str) -> tuple[int, int]:
        x, y = self.abs[nombre]
        return int(x), int(y)

    def localizar(self) -> None:
        """Encuentra la ventana en pantalla (puede estar en otra posición que al calibrar) y recalcula los puntos."""
        self.abs, info = calcular_puntos(self.p.captura(), self.cfg)
        self.g = self._geometria()
        mov = ", ".join(f"{n}: desplazada ({dx:+d},{dy:+d}) parecido {sc:.2f}" for n, (dx, dy, sc) in info.items())
        self.log(f"  ventana localizada -> {mov}")

    def _geometria(self) -> Geometria:
        f1, f2, fn = self._pt("fila1"), self._pt("fila2"), self._pt("filaN")
        ext, tb, td = self._pt("ext"), self._pt("term_borde"), self._pt("term_der")
        h = f2[1] - f1[1]
        if h < 8:
            raise ErrorSophis("Calibración inválida: 'fila2' debe estar debajo de 'fila1'. Repite 'calibrar'.")
        n = int(round((fn[1] - f1[1]) / h)) + 1
        if n < 3:
            raise ErrorSophis("Calibración inválida: 'filaN' debe ser la última fila visible. Repite 'calibrar'.")
        ensanche = max(0, self.cfg.ensanchar_px)
        return Geometria(
            x_click=f1[0], x_mnemo0=f1[0] - 30, x_mnemo1=max(f1[0] + 12, ext[0] - 2), x_ext=ext[0],
            y1=f1[1], h=float(h), n=n,
            term_x0=tb[0] - ensanche - 6, term_x1=td[0],
        )

    def guardar_debug(self, nombre: str, img: Image.Image | None = None) -> Path:
        CARPETA_DEBUG.mkdir(exist_ok=True)
        ruta = CARPETA_DEBUG / f"{time.strftime('%H%M%S')}_{nombre}.png"
        (img or self.p.captura()).save(ruta)
        return ruta

    # ── preparación ─────────────────────────────────────────────────────────
    def ensanchar_term(self) -> None:
        """La columna Term muestra 'Short ter...': se arrastra su borde izquierdo para ver el texto completo."""
        if self._ancho_aplicado or self.cfg.ensanchar_px <= 0:
            return
        bx, by = self._pt("term_borde")
        self.log(f"Ensanchando columna Term {self.cfg.ensanchar_px}px ...")
        self.p.arrastrar(bx, by, bx - self.cfg.ensanchar_px, by)
        self._ancho_aplicado = True
        self._esp(0.6)

    def calibrar_color_seleccion(self) -> None:
        g = self.g
        self.click_fila(0)
        self._esp(0.4)
        colores = self._colores()
        sel, otros = colores[0], colores[1:3]
        dist = min(distancia_color(sel, o) for o in otros)
        if dist < 30:
            ruta = self.guardar_debug("seleccion_no_distinguible")
            raise ErrorSophis(
                f"No logro distinguir la fila seleccionada de las demás por color (diferencia={dist}). "
                f"Revisa que 'ext' y 'fila1' estén bien calibrados. Captura: {ruta}")
        self._color_sel = sel
        self._tol = max(12, min(40, dist // 2))
        self.log(f"Color de fila seleccionada: {sel} (tolerancia {self._tol})")

    def preparar(self) -> None:
        self.localizar()
        self.calibrar_color_seleccion()

    # ── grid ────────────────────────────────────────────────────────────────
    def click_fila(self, i: int, doble: bool = False) -> None:
        y = self.g.yc(i)
        (self.p.doble_click if doble else self.p.click)(self.g.x_click, y)

    def _colores(self) -> list[tuple]:
        g = self.g
        top = int(round(g.y1 - g.h / 2))
        img = self.p.captura((g.x_ext - 3, top, g.x_ext + 4, int(round(top + g.h * g.n))))
        return colores_filas(img, g.h, g.n)

    def fila_seleccionada(self) -> int | None:
        if self._color_sel is None:
            raise ErrorSophis("Falta preparar() (color de selección).")
        idx = filas_con_color(self._colores(), self._color_sel, self._tol)
        return idx[0] if len(idx) == 1 else None

    def _region_celda(self, i: int, x0: int, x1: int) -> tuple[int, int, int, int]:
        yc = self.g.yc(i)
        mitad = int(self.g.h / 2) - 1
        return (x0, yc - mitad, x1, yc + mitad)

    def _leer_portapapeles(self) -> str | None:
        centinela = f"__sophis_{time.time()}__"
        self.p.portapapeles_set(centinela)
        self.p.hotkey("ctrl", "c")
        self._esp(0.4)
        txt = self.p.portapapeles_get()
        if not txt or txt == centinela:
            return None
        return txt

    def _mnemo_por_ocr(self, i: int, codigo: str) -> tuple[bool, list[str]]:
        region = self._region_celda(i, self.g.x_mnemo0, self.g.x_mnemo1)
        img = self.p.captura(region)
        vistos = []
        for escala in (3, 4, 2):
            txt = self.ocr.texto(img, escala)
            vistos.append(txt)
            if _norm_codigo(txt) == _norm_codigo(codigo):
                return True, vistos
        return False, vistos

    def _term_por_ocr(self, i: int) -> tuple[str | None, str]:
        img = self.p.captura(self._region_celda(i, self.g.term_x0, self.g.term_x1))
        textos = []
        for escala in (3, 4):
            txt = self.ocr.texto(img, escala)
            textos.append(txt)
            term = clasificar_term(txt)
            if term:
                return term, txt
        return None, " | ".join(textos)

    def leer_fila(self, i: int, codigo: str) -> Lectura:
        modo = self.cfg.lector
        ultima = None
        if modo in ("auto", "portapapeles"):
            txt = self._leer_portapapeles()
            if txt:
                celdas = [c.strip() for c in re.split(r"[\t\r\n]+", txt) if c.strip()]
                if any(_norm_codigo(c) == _norm_codigo(codigo) for c in celdas[:3]):
                    return Lectura(True, clasificar_term(txt), "portapapeles", txt, term_txt=txt)
                if celdas:
                    ultima = Lectura(False, clasificar_term(txt), "portapapeles", txt, vistos=(celdas[0],), term_txt=txt)
        if modo in ("auto", "ocr"):
            ok, vistos = self._mnemo_por_ocr(i, codigo)
            term, txt_t = self._term_por_ocr(i)
            texto = f"mnemo={vistos} term='{txt_t}'"
            if ok:
                return Lectura(True, term, "ocr", texto, vistos=tuple(vistos), term_txt=txt_t)
            if not self.cfg.verificar_mnemo and term:
                self.log("  ⚠ Mnemo NO verificado (verificar_mnemo=False)")
                return Lectura(True, term, "ocr", texto, verificado=False, term_txt=txt_t)
            return Lectura(False, term, "ocr", texto, vistos=tuple(vistos), term_txt=txt_t)
        return ultima or Lectura(False, None, modo, "")

    # ── búsqueda ────────────────────────────────────────────────────────────
    @staticmethod
    def _es_extension(lect: Lectura, codigo: str) -> bool:
        """True si la fila leída empieza con el código pero es más larga (ej. buscas 'ZM' y cae en 'ZMH')."""
        c = _norm_codigo(codigo)
        return any(_norm_codigo(t).startswith(c) and len(_norm_codigo(t)) > len(c) for t in lect.vistos)

    def _bajar_hasta(self, codigo: str, max_filas: int = 15) -> tuple[int, Lectura] | None:
        """La búsqueda incremental cae en el primer prefijo; en una lista descendente el código exacto
        (ZM) viene DESPUÉS de sus extensiones (ZMH). Se baja fila por fila verificando cada una."""
        for _ in range(max_filas):
            self.p.tecla("down")
            self._esp(0.35)
            j = self.fila_seleccionada()
            if j is None:
                return None
            lect = self.leer_fila(j, codigo)
            if lect.mnemo_ok:
                return j, lect
            if not self._es_extension(lect, codigo):
                return None
        return None

    def buscar(self, codigo: str, intentos: int = 3) -> tuple[int, Lectura]:
        ultimo = ""
        for k in range(1, intentos + 1):
            self.click_fila(0)       # enfoca el grid y reinicia el buffer de búsqueda incremental
            self._esp(0.4)
            self.p.escribir(codigo)
            self._esp(0.8)
            i = self.fila_seleccionada()
            if i is None:
                ultimo = "no se detectó fila seleccionada"
                self.log(f"  intento {k}: {ultimo}")
                continue
            lect = self.leer_fila(i, codigo)
            if not lect.mnemo_ok and self._es_extension(lect, codigo):
                self.log(f"  cayó en {lect.vistos[0]!r}; bajando hasta '{codigo}' ...")
                hallado = self._bajar_hasta(codigo)
                if hallado:
                    i, lect = hallado
            if lect.mnemo_ok:
                self.log(f"  fila {i} ok ({lect.fuente}): term={lect.term}")
                return i, lect
            ultimo = f"la fila seleccionada no es '{codigo}' ({lect.texto})"
            self.log(f"  intento {k}: {ultimo}")
        ruta = self.guardar_debug(f"no_encontrado_{codigo}")
        raise ErrorSophis(f"No pude ubicar '{codigo}': {ultimo}. Captura: {ruta}")

    # ── dropdown ────────────────────────────────────────────────────────────
    def leer_dropdown(self) -> tuple[str | None, str]:
        izq, fle = self._pt("dd_izq"), self._pt("dd_flecha")
        img = self.p.captura((izq[0] - 6, izq[1] - 11, fle[0] - 8, izq[1] + 11))
        for escala in (3, 4):
            txt = self.ocr.texto(img, escala)
            term = clasificar_term(txt)
            if term:
                return term, txt
        return None, txt

    def fijar_dropdown(self, term: str) -> None:
        actual, _ = self.leer_dropdown()
        if actual == term:
            self.log(f"  dropdown ya está en '{term}'")
            return
        izq, fle = self._pt("dd_izq"), self._pt("dd_flecha")
        region = (izq[0] - 8, fle[1] + 6, fle[0] + 10, fle[1] + 6 + 260)
        for k in range(1, 4):
            self.p.click(*fle)
            self._esp(0.8)
            lineas = self.ocr.leer(self.p.captura(region), escala=2)
            objetivo = next((l for l in lineas if clasificar_term(l.texto) == term), None)
            if objetivo:
                self.p.click(region[0] + objetivo.cx, region[1] + objetivo.cy)
            else:
                self.log(f"  intento {k}: no vi '{term}' en la lista; uso teclado")
                self.p.tecla("home")
                for _ in range(ORDEN_DROPDOWN.index(term)):
                    self.p.tecla("down")
                self.p.tecla("enter")
            self._esp(0.8)
            actual, txt = self.leer_dropdown()
            if actual == term:
                self.log(f"  dropdown -> '{term}'")
                return
            self.log(f"  intento {k}: dropdown quedó en {actual!r} ({txt!r})")
        ruta = self.guardar_debug("dropdown_no_coincide")
        raise ErrorSophis(f"No logré poner el dropdown en '{term}'. Captura: {ruta}")

    # ── ventana de actualización ────────────────────────────────────────────
    def esperar_ventana(self, timeout: float = 20.0) -> dict:
        """Espera a que aparezca la ventana con columnas Maturity / Minimum / Maximum / Ric."""
        t0 = time.time()
        while time.time() - t0 < timeout * max(1.0, self.cfg.pausa):
            self._esp(1.0)
            img = self.p.captura()
            for escala in (1, 2):      # pasada rápida y, si no alcanza, una más precisa (el texto es pequeño)
                hallado = {}
                for l in self.ocr.leer(img, escala=escala, margen=8):
                    for palabra in ("Maturity", "Minimum", "Maximum"):
                        if palabra not in hallado and contiene_palabra(l.texto, palabra):
                            hallado[palabra] = (l.cx, l.cy)
                if len(hallado) >= 2:
                    return hallado
        ruta = self.guardar_debug("ventana_no_abrio")
        raise ErrorSophis(f"No apareció la ventana Maturity/Minimum/Maximum. Captura: {ruta}")

    # ── flujo completo de la etapa 1 ────────────────────────────────────────
    def abrir_contrato(self, codigo: str) -> Resultado:
        codigo = codigo.strip().upper()
        self.log(f"[{codigo}] buscando ...")
        self.localizar()      # la ventana pudo moverse desde el código anterior
        i, lect = self.buscar(codigo)
        if not lect.term and self.cfg.ensanchar_px > 0 and not self._ancho_aplicado:
            # Term ilegible (normalmente porque sale truncado: "Short ter..."): se ensancha la columna una vez
            # por sesión. Con el lector de portapapeles esto casi nunca hace falta.
            self.ensanchar_term()
            lect = self.leer_fila(i, codigo)
        if not lect.term:
            if term_truncado(lect.term_txt):
                raise ErrorSophis(f"[{codigo}] la columna Term está truncada ('{lect.term_txt}'). "
                                  f"Aumenta ensanchar_px en {ARCHIVO_CONFIG} o ensánchala a mano.")
            ruta = self.guardar_debug(f"term_ilegible_{codigo}")
            raise ErrorSophis(f"[{codigo}] no pude leer el Term ('{lect.texto}'). Captura: {ruta}")
        self.fijar_dropdown(lect.term)

        # Cambiar el dropdown puede recargar/mover la lista: se re-verifica la fila antes de abrir.
        j = self.fila_seleccionada()
        if j is None or not self.leer_fila(j, codigo).mnemo_ok:
            self.log("  la lista cambió tras el dropdown; busco otra vez")
            j, _ = self.buscar(codigo)
        self.log(f"  abriendo ventana (doble click fila {j}) ...")
        self.click_fila(j, doble=True)
        ventana = self.esperar_ventana()
        self.log(f"[{codigo}] listo: term='{lect.term}', ventana de actualización abierta")
        return Resultado(codigo, lect.term, j, ventana)


# ════════════════════════════════════════════════════════════════════════════
# COMANDOS
# ════════════════════════════════════════════════════════════════════════════

def cmd_calibrar(args) -> None:
    p = Pantalla()
    print("CALIBRACIÓN: para cada punto pon el mouse donde se indica y NO lo muevas.")
    print(f"Se captura automáticamente al terminar la cuenta regresiva de {args.espera}s.")
    print("La ventana puede estar en cualquier posición, pero debe estar completa y visible.\n")
    puntos, anclas = {}, {}
    for k, (nombre, texto) in enumerate(PUNTOS_CALIBRACION, 1):
        print(f"[{k}/{len(PUNTOS_CALIBRACION)}] {texto}")
        for s in range(args.espera, 0, -1):
            print(f"   {s}...", end="\r", flush=True)
            time.sleep(1)
        x, y = p.posicion()
        puntos[nombre] = [x, y]
        print(f"   {nombre} = ({x}, {y})        ")
        if nombre in ANCLA_CIERRE:
            nom = ANCLA_CIERRE[nombre]
            ax, ay = puntos[f"ancla_{nom}_a"]
            x0, y0, x1, y1 = min(ax, x), min(ay, y), max(ax, x), max(ay, y)
            if x1 - x0 < 40 or y1 - y0 < 8:
                raise ErrorSophis(f"La ancla '{nom}' quedó muy pequeña ({x1 - x0}x{y1 - y0}px). "
                                  "Marca bien las dos esquinas del encabezado y repite 'calibrar'.")
            ARCHIVO_ANCLAS.mkdir(exist_ok=True)
            ruta = ARCHIVO_ANCLAS / f"{nom}.png"
            p.captura((x0, y0, x1, y1)).save(ruta)
            anclas[nom] = {"archivo": str(ruta), "origen": [x0, y0]}
    previa = None
    try:
        previa = Config.cargar()
    except ErrorSophis:
        pass
    cfg = Config(puntos=puntos, anclas=anclas)
    if previa:
        cfg.ensanchar_px, cfg.intervalo_tecla = previa.ensanchar_px, previa.intervalo_tecla
        cfg.pausa, cfg.lector, cfg.verificar_mnemo = previa.pausa, previa.lector, previa.verificar_mnemo
    cfg.guardar()
    # Comprobación inmediata: las anclas deben encontrarse a sí mismas con parecido ~1.0
    _, info = calcular_puntos(p.captura(), cfg)
    for nom, (dx, dy, sc) in info.items():
        print(f"Ancla '{nom}': parecido {sc:.2f}, desplazamiento ({dx:+d},{dy:+d})  <- debe ser ~1.00 y (0,0)")
    print(f"\nGuardado en {ARCHIVO_CONFIG} y {ARCHIVO_ANCLAS}/. Ahora puedes mover la ventana de lugar.")
    print("Siguiente:  python sophis_auto.py diagnostico <CODIGO>")


def _crear(args) -> Automatizador:
    cfg = Config.cargar()
    if getattr(args, "lector", None):
        cfg.lector = args.lector
    if getattr(args, "no_ensanchar", False):
        cfg.ensanchar_px = 0
    if getattr(args, "pausa", None):
        cfg.pausa = args.pausa
    print("Cargando OCR ...")
    return Automatizador(Pantalla(cfg.intervalo_tecla), OCR(), cfg)


def cmd_diagnostico(args) -> None:
    """Muestra qué ve el script; sirve para ajustar la calibración (no abre ninguna ventana)."""
    a = _crear(args)
    codigo = args.codigo.upper()
    print("Tienes 3 segundos para dejar Sophis (Citrix) al frente ...")
    time.sleep(3)
    CARPETA_DEBUG.mkdir(exist_ok=True)
    a.guardar_debug("diag_pantalla")
    a.preparar()
    a.click_fila(0)
    a._esp(0.4)
    a.p.escribir(codigo)
    a._esp(0.8)
    i = a.fila_seleccionada()
    print(f"\nFila seleccionada tras escribir '{codigo}': {i}")
    if i is not None:
        g = a.g
        a.guardar_debug("diag_mnemo", a.p.captura(a._region_celda(i, g.x_mnemo0, g.x_mnemo1)))
        a.guardar_debug("diag_term", a.p.captura(a._region_celda(i, g.term_x0, g.term_x1)))
        print("Portapapeles tras Ctrl+C :", repr(a._leer_portapapeles()))
        print("OCR Mnemo (ok, lecturas):", a._mnemo_por_ocr(i, codigo))
        print("OCR Term                :", a._term_por_ocr(i))
    print("Dropdown (OCR)          :", a.leer_dropdown())
    print(f"\nCapturas guardadas en {CARPETA_DEBUG}/ (pantalla, mnemo, term).")


def cmd_abrir(args) -> None:
    a = _crear(args)
    print("Tienes 3 segundos para dejar Sophis (Citrix) al frente ...")
    time.sleep(3)
    a.preparar()
    for k, codigo in enumerate(args.codigos):
        try:
            r = a.abrir_contrato(codigo)
        except ErrorSophis as e:
            print(f"\n✗ {e}")
            sys.exit(1)
        if k < len(args.codigos) - 1:
            resp = input("\nCuando termines con esta ventana (ciérrala/ok), ENTER para el siguiente (q = salir): ")
            if resp.strip().lower() == "q":
                break


def main() -> None:
    ap = argparse.ArgumentParser(description="Automatización Sophis (etapa 1: buscar código y abrir ventana)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("calibrar", help="registra las posiciones en pantalla (una vez)")
    s.add_argument("--espera", type=int, default=5, help="segundos por punto (default 5)")
    s.set_defaults(fn=cmd_calibrar)

    for nombre, fn, ayuda in (("diagnostico", cmd_diagnostico, "prueba cómo se lee la pantalla"),
                              ("abrir", cmd_abrir, "busca el código, iguala dropdown y abre la ventana")):
        s = sub.add_parser(nombre, help=ayuda)
        if nombre == "diagnostico":
            s.add_argument("codigo")
        else:
            s.add_argument("codigos", nargs="+")
        s.add_argument("--lector", choices=["auto", "portapapeles", "ocr"], help="cómo leer la fila")
        s.add_argument("--no-ensanchar", action="store_true", help="no tocar el ancho de la columna Term")
        s.add_argument("--pausa", type=float, help="multiplicador de esperas (2 = el doble de lento)")
        s.set_defaults(fn=fn)

    args = ap.parse_args()
    try:
        args.fn(args)
    except ErrorSophis as e:
        print(f"✗ {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

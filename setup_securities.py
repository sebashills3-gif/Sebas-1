"""Apoyo para las opciones 4 y 5 del menú:

* Elegir con un seleccionador de archivos el Excel base.
* Crear la carpeta de la corrida que arranca en BLOOMBOT.
* Reconstruir, a partir de una carpeta de evidencia ya existente, todo lo
  necesario para el set up de securities en IMPACT (XLS de BLOOMBOT, Excel
  filtrado con las entidades y PDF del Sanctions Check).
"""

import shutil
from datetime import datetime
from pathlib import Path

from bloombot import obtener_codigos
from prestaging import leer_matriz_xls
from sanctions_pdf import verificar_securities
from skeletons import procesar_matriz

NOMBRE_FILTRADO = "TUR_requests_filtrados.xlsx"


# ============================================================
# SELECCIONADOR DE ARCHIVOS
# ============================================================

def carpeta_evidencia_reciente(base=None):
    """La carpeta 'TUR Evidence_*' más reciente (por nombre), o `base`."""
    base = Path(base or Path.cwd())
    carpetas = sorted(
        (p for p in base.glob("TUR Evidence_*") if p.is_dir()),
        key=lambda p: p.name,
    )
    return carpetas[-1] if carpetas else base


def elegir_archivo(titulo, carpeta_inicial=None):
    """Abre el seleccionador de archivos de Windows. Si no está disponible,
    pide la ruta por consola. Devuelve un Path o None si se cancela."""
    inicial = Path(carpeta_inicial or carpeta_evidencia_reciente())
    tipos = [("Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]

    try:
        import tkinter as tk
        from tkinter import filedialog

        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes("-topmost", True)
        ruta = filedialog.askopenfilename(
            title=titulo,
            initialdir=str(inicial),
            filetypes=tipos,
            parent=raiz,
        )
        raiz.destroy()

    except Exception:
        ruta = input(f"{titulo}\nEscribe la ruta del archivo: ").strip().strip('"')

    return Path(ruta) if ruta else None


# ============================================================
# OPCIÓN 4: carpeta nueva para la corrida desde BLOOMBOT
# ============================================================

def nueva_carpeta_bloombot(excel_filtrado, base=None):
    """Crea 'TUR Evidence_<fecha>_BLOOMBOT' y copia ahí el Excel filtrado
    (así la opción 5 podrá usar esa carpeta después). Devuelve (carpeta,
    ruta de la copia)."""
    base = Path(base or Path.cwd())
    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    carpeta = base / f"TUR Evidence_{fecha}_BLOOMBOT"
    carpeta.mkdir(parents=True, exist_ok=True)

    copia = carpeta / NOMBRE_FILTRADO
    shutil.copy2(excel_filtrado, copia)
    return carpeta, copia


# ============================================================
# OPCIÓN 5: set up desde una carpeta existente
# ============================================================

def _mas_reciente(rutas):
    rutas = list(rutas)
    return max(rutas, key=lambda p: p.stat().st_mtime) if rutas else None


def localizar_archivos(ruta):
    """A partir del archivo elegido, busca en su carpeta lo que hace falta.

    Se puede elegir el XLS de BLOOMBOT (GIMP*.xls) o el Excel filtrado del
    TUR. Devuelve {"carpeta", "xls", "filtrado", "pdfs"}; los que no se
    encuentran quedan en None (o lista vacía).
    """
    ruta = Path(ruta)
    carpeta = ruta.parent

    if ruta.suffix.lower() == ".xls":
        xls = ruta
        filtrado = carpeta / NOMBRE_FILTRADO
        filtrado = filtrado if filtrado.exists() else None
    else:
        filtrado = ruta
        xls = _mas_reciente(
            p for p in carpeta.iterdir()
            if p.suffix.lower() == ".xls" and p.name.upper().startswith("GIMP")
        )

    pdf = _mas_reciente(
        p for p in carpeta.iterdir() if p.suffix.lower() == ".pdf"
    )

    return {
        "carpeta": carpeta,
        "xls": xls,
        "filtrado": filtrado,
        "pdfs": [pdf] if pdf else [],
    }


def preparar_setup(ruta):
    """Reconstruye los securities listos para IMPACT desde una carpeta de
    evidencia: lee el XLS de BLOOMBOT, clasifica, asigna entidades con el
    Excel filtrado y revisa el PDF del Sanctions Check.

    Devuelve (securities, carpeta). Lanza RuntimeError si falta algo
    imprescindible.
    """
    archivos = localizar_archivos(ruta)
    carpeta = archivos["carpeta"]

    if archivos["xls"] is None:
        raise RuntimeError(
            f"No encontré el XLS de BLOOMBOT (GIMP*.xls) en {carpeta}."
        )
    if archivos["filtrado"] is None:
        raise RuntimeError(
            f"No encontré {NOMBRE_FILTRADO} en {carpeta}; hace falta para "
            "saber la entidad de cada security."
        )

    print()
    print(f"Carpeta base: {carpeta}")
    print(f"XLS de BLOOMBOT: {archivos['xls'].name}")
    print(f"Excel filtrado: {archivos['filtrado'].name}")
    print(
        "PDF de Sanctions: "
        + (archivos["pdfs"][0].name if archivos["pdfs"] else "(ninguno)")
    )

    matriz = leer_matriz_xls(archivos["xls"])
    _, _, entidades = obtener_codigos(archivos["filtrado"])

    securities = procesar_matriz(matriz, carpeta, entidades)
    verificar_securities(archivos["pdfs"], matriz, securities, carpeta)

    return securities, carpeta

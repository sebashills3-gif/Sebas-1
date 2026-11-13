"""Después de correr el Sanctions Check en Genie:

1. Espera la carpeta "Sanctions Check_dd.mm.yyyy_hh.mm.ss" que Genie crea en
   C:\\IM Genie Bot\\Sanctions Check.
2. Toma el PDF que se generó dentro de ella.
3. Lo mueve a la carpeta de evidencia (donde están las capturas).
"""

import re
import time
from datetime import datetime, timedelta
from pathlib import Path

from prestaging import mover_archivos

CARPETA_SANCTIONS = Path(r"C:\IM Genie Bot\Sanctions Check")

PATRON_CARPETA = re.compile(
    r"^Sanctions Check_(\d{2}\.\d{2}\.\d{4}_\d{2}\.\d{2}\.\d{2})$",
    re.IGNORECASE,
)
FORMATO_FECHA = "%d.%m.%Y_%H.%M.%S"


def buscar_carpeta(desde, base=CARPETA_SANCTIONS):
    """Carpeta de Sanctions Check más reciente creada desde `desde`."""
    base = Path(base)
    if not base.exists():
        return None

    candidatas = []
    for ruta in base.iterdir():
        coincide = PATRON_CARPETA.match(ruta.name)
        if not ruta.is_dir() or not coincide:
            continue
        try:
            marca = datetime.strptime(coincide.group(1), FORMATO_FECHA)
        except ValueError:
            continue
        if marca >= desde - timedelta(seconds=2):
            candidatas.append((marca, ruta))

    return max(candidatas)[1] if candidatas else None


def esperar_pdfs(desde, base=CARPETA_SANCTIONS, timeout_s=900, pausa_s=3):
    """Espera la carpeta nueva y a que su PDF esté completo.

    Devuelve la lista de PDF (normalmente uno) de la carpeta.
    """
    inicio = time.time()
    tamanos_previos = None

    while time.time() - inicio < timeout_s:
        carpeta = buscar_carpeta(desde, base)

        if carpeta is not None:
            pdfs = sorted(
                p for p in carpeta.iterdir()
                if p.is_file() and p.suffix.lower() == ".pdf"
            )
            tamanos = {p: p.stat().st_size for p in pdfs}

            # El PDF debe existir y haber dejado de crecer.
            if pdfs and tamanos == tamanos_previos:
                return pdfs
            tamanos_previos = tamanos

        time.sleep(pausa_s)

    raise TimeoutError(
        f"No apareció la carpeta con el PDF de Sanctions Check en {base} "
        f"tras {timeout_s} segundos."
    )


def procesar_sanctions(desde, carpeta_evidencia, base=CARPETA_SANCTIONS):
    """Mueve el PDF del Sanctions Check a la carpeta de evidencia."""
    print()
    print(f"Esperando la carpeta de Sanctions Check en {base}...", flush=True)
    pdfs = esperar_pdfs(desde, base)

    if len(pdfs) > 1:
        print(f"Aviso: la carpeta trae {len(pdfs)} PDF; se mueven todos.")

    for ruta in mover_archivos(pdfs, carpeta_evidencia):
        print(f"PDF movido: {ruta}")

"""Pulsa el botón "Perform Sanctions Check" de la aplicación Genie, que ya
debe estar abierta.

Requiere Windows y: pip install pywinauto
"""

import re
import time

# Título de la ventana de Genie (en la captura aparece "PRODUCTION").
TITULO_VENTANA_RE = r"(?i).*(genie|production).*"
# En la app el botón se llama "Perfrom Sanctions Check" (con la errata), así
# que el patrón acepta también "Perform Sanction Check".
TEXTO_BOTON_RE = r"(?i)per\w+\s+sanction"


def _buscar_boton():
    """Devuelve (boton, ventana) o (None, titulos_visibles)."""
    from pywinauto import Desktop

    titulos = []

    for backend in ("uia", "win32"):
        escritorio = Desktop(backend=backend)

        for ventana in escritorio.windows(title_re=TITULO_VENTANA_RE):
            titulos.append(ventana.window_text())
            try:
                for control in ventana.descendants():
                    if re.search(TEXTO_BOTON_RE, control.window_text() or ""):
                        return control, ventana
            except Exception:
                continue

    return None, titulos


def hacer_click_sanction_check(timeout_s=30, pausa_s=2):
    """Busca el botón en la ventana de Genie y le da click.

    Reintenta durante `timeout_s` segundos por si la ventana tarda en
    aparecer. Lanza RuntimeError si no lo encuentra.
    """
    inicio = time.time()
    titulos = []

    while True:
        boton, resto = _buscar_boton()

        if boton is not None:
            ventana = resto
            try:
                ventana.set_focus()
            except Exception:
                pass
            boton.click_input()
            print("Click en 'Perform Sanctions Check' realizado.")
            return

        titulos = resto
        if time.time() - inicio >= timeout_s:
            break
        time.sleep(pausa_s)

    raise RuntimeError(
        "No se encontró el botón 'Perform Sanctions Check' en Genie. "
        f"Ventanas candidatas vistas: {sorted(set(titulos)) or 'ninguna'}."
    )

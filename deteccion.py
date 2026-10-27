"""Detección robusta de la tabla "External Trade ID" y reintento de la
extracción cuando no devuelve nada.

Uso en el script principal:

    from deteccion import tabla_visible_robusta, ejecutar_con_reintentos

    # sustituye a tabla_visible(page)
    if not tabla_visible_robusta(page): ...

    # sustituye a ejecutar_extraccion(page)
    ejecutar_con_reintentos(page, ejecutar_extraccion, carpeta_salida)
"""

import time
from pathlib import Path

TEXTO_TABLA = "External Trade ID"


def _estado_listo(page, timeout_ms=15000):
    """Espera a que la página termine de cargar; no falla si no lo logra."""
    for estado in ("domcontentloaded", "load", "networkidle"):
        try:
            page.wait_for_load_state(estado, timeout=timeout_ms)
        except Exception:
            pass


def _encabezado_en_frame(frame, espera_ms):
    """True si el frame muestra el encabezado, probando varios selectores."""
    selectores = (
        frame.get_by_role("columnheader", name=TEXTO_TABLA),
        frame.locator(f"th:has-text('{TEXTO_TABLA}')"),
        frame.get_by_text(TEXTO_TABLA, exact=False),
    )
    for localizador in selectores:
        try:
            localizador.first.wait_for(state="visible", timeout=espera_ms)
            return True
        except Exception:
            continue
    return False


def tabla_visible_robusta(page, intentos=5, espera_ms=3000, pausa_s=2):
    """Busca el encabezado en la página y en todos sus iframes.

    Reintenta varias veces esperando la carga y haciendo scroll, porque la
    tabla puede tardar en pintarse o estar dentro de un iframe.
    """
    for intento in range(1, intentos + 1):
        _estado_listo(page)

        for frame in page.frames:
            try:
                if _encabezado_en_frame(frame, espera_ms):
                    return True
            except Exception:
                continue

        print(
            f"Tabla no detectada (intento {intento}/{intentos})"
            + ("; reintentando..." if intento < intentos else ".")
        )
        try:
            page.mouse.wheel(0, 300)
            page.mouse.wheel(0, -300)
        except Exception:
            pass
        time.sleep(pausa_s)
    return False


def _archivos(carpeta):
    carpeta = Path(carpeta)
    if not carpeta.exists():
        return set()
    return {p for p in carpeta.rglob("*") if p.is_file()}


def ejecutar_con_reintentos(page, ejecutar_extraccion, carpeta_salida,
                            intentos=3, pausa_s=3):
    """Corre la extracción y, si no genera archivos nuevos, vuelve a
    detectar la tabla y la repite (hasta `intentos` veces).

    Devuelve True si la extracción produjo archivos, False si no.
    """
    for intento in range(1, intentos + 1):
        antes = _archivos(carpeta_salida)

        if not tabla_visible_robusta(page):
            print("No se detectó la tabla antes de extraer.")
        else:
            ejecutar_extraccion(page)

        if _archivos(carpeta_salida) - antes:
            return True

        print(
            f"La extracción no generó archivos (intento {intento}/{intentos})."
        )
        if intento < intentos:
            time.sleep(pausa_s)

    print("No se obtuvo ningún resultado tras los reintentos.")
    return False

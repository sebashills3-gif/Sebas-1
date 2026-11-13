"""Pulsa el botón "Perform Sanctions Check" de la aplicación Genie, que ya
debe estar abierta.

Genie (Genie_BOT_IM_x64.exe) tiene una ventana auxiliar "GDI+ Window" y su
ventana real puede no tener un título que sirva para buscarla, así que se
buscan todas las ventanas del proceso, sin importar su título.

Requiere Windows y: pip install pywinauto
Opcional (respaldo por imagen): pip install pyautogui opencv-python pillow
"""

import re
import time
from pathlib import Path

NOMBRE_EXE = "genie_bot_im_x64.exe"

# En la app el botón se llama "Perfrom Sanctions Check" (con la errata), así
# que el patrón acepta también "Perform Sanction Check".
TEXTO_BOTON_RE = r"(?i)per\w+\s+sanction"

# Respaldo: recorte de pantalla del botón, junto a este archivo.
IMAGEN_BOTON = Path(__file__).with_name("boton_sanction_check.png")
ARCHIVO_DIAGNOSTICO = Path(__file__).with_name("genie_diagnostico.txt")


def _exe_del_proceso(pid):
    """Ruta del ejecutable de un proceso, o '' si no se puede leer."""
    import ctypes
    from ctypes import wintypes

    kernel = ctypes.windll.kernel32
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]

    manejador = kernel.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFO
    if not manejador:
        return ""
    try:
        buffer = ctypes.create_unicode_buffer(1024)
        tamano = wintypes.DWORD(1024)
        if kernel.QueryFullProcessImageNameW(
            manejador, 0, buffer, ctypes.byref(tamano)
        ):
            return buffer.value
        return ""
    finally:
        kernel.CloseHandle(manejador)


def _ventanas_de_genie():
    """Ventanas de nivel superior del proceso de Genie, por backend."""
    from pywinauto import Desktop

    resultado = []

    for backend in ("uia", "win32"):
        for ventana in Desktop(backend=backend).windows():
            try:
                exe = _exe_del_proceso(ventana.process_id())
            except Exception:
                continue
            if exe.lower().endswith(NOMBRE_EXE):
                resultado.append((backend, ventana))

    return resultado


def _buscar_boton(ventanas):
    for _, ventana in ventanas:
        try:
            controles = [ventana] + ventana.descendants()
        except Exception:
            continue
        for control in controles:
            try:
                if re.search(TEXTO_BOTON_RE, control.window_text() or ""):
                    return control, ventana
            except Exception:
                continue
    return None, None


def _volcar_diagnostico(ventanas):
    """Guarda en un archivo las ventanas y controles que ve pywinauto."""
    lineas = [f"Ventanas de {NOMBRE_EXE}: {len(ventanas)}"]

    for backend, ventana in ventanas:
        lineas.append("")
        lineas.append(
            f"[{backend}] ventana: título={ventana.window_text()!r} "
            f"clase={ventana.friendly_class_name()!r} "
            f"visible={ventana.is_visible()} rect={ventana.rectangle()}"
        )
        try:
            for control in ventana.descendants():
                lineas.append(
                    f"    clase={control.friendly_class_name()!r} "
                    f"texto={control.window_text()!r} "
                    f"rect={control.rectangle()}"
                )
        except Exception as error:
            lineas.append(f"    (no se pudieron leer los controles: {error})")

    ARCHIVO_DIAGNOSTICO.write_text("\n".join(lineas), encoding="utf-8")
    return ARCHIVO_DIAGNOSTICO


def _click_por_imagen():
    """Respaldo: busca en pantalla el recorte del botón y le da click."""
    if not IMAGEN_BOTON.exists():
        return False

    import pyautogui

    try:
        posicion = pyautogui.locateCenterOnScreen(
            str(IMAGEN_BOTON), confidence=0.8
        )
    except TypeError:  # sin opencv no existe `confidence`
        posicion = pyautogui.locateCenterOnScreen(str(IMAGEN_BOTON))

    if posicion is None:
        return False

    pyautogui.click(posicion)
    return True


def hacer_click_sanction_check(timeout_s=30, pausa_s=2):
    """Busca el botón en las ventanas de Genie y le da click.

    Reintenta durante `timeout_s` segundos por si la ventana tarda en
    aparecer. Si no lo encuentra, prueba el respaldo por imagen (si existe
    boton_sanction_check.png) y, si tampoco, guarda un diagnóstico y lanza
    RuntimeError.
    """
    inicio = time.time()

    while True:
        ventanas = _ventanas_de_genie()
        boton, ventana = _buscar_boton(ventanas)

        if boton is not None:
            try:
                ventana.set_focus()
            except Exception:
                pass
            boton.click_input()
            print("Click en 'Perform Sanctions Check' realizado.")
            return

        if time.time() - inicio >= timeout_s:
            break
        time.sleep(pausa_s)

    if _click_por_imagen():
        print("Click en 'Perform Sanctions Check' realizado (por imagen).")
        return

    mensaje = "No se encontró el botón 'Perform Sanctions Check' en Genie."
    if ventanas:
        archivo = _volcar_diagnostico(ventanas)
        mensaje += f" Diagnóstico guardado en: {archivo}"
    else:
        mensaje += f" No hay ventanas visibles del proceso {NOMBRE_EXE}."
    raise RuntimeError(mensaje)

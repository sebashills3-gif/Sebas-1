"""Press the "Perform Sanctions Check" button of the Genie application, which
must already be open.

Genie (Genie_BOT_IM_x64.exe) has an auxiliary "GDI+ Window" and its real
window may not have a title that is useful for finding it, so all the windows
of the process are searched, whatever their title.

Requires Windows and: pip install pywinauto
Optional (image fallback): pip install pyautogui opencv-python pillow
"""

import re
import time
from pathlib import Path

EXE_NAME = "genie_bot_im_x64.exe"

# In the app the button is called "Perfrom Sanctions Check" (with the typo),
# so the pattern also accepts "Perform Sanction Check".
BUTTON_TEXT_RE = r"(?i)per\w+\s+sanction"

# Fallback: a screenshot crop of the button, next to this file. The old
# Spanish file name is still accepted.
BUTTON_IMAGES = (
    Path(__file__).with_name("sanction_check_button.png"),
    Path(__file__).with_name("boton_sanction_check.png"),
)
DIAGNOSTIC_FILE = Path(__file__).with_name("genie_diagnostic.txt")


def _process_exe(pid):
    """Path of a process executable, or '' if it cannot be read."""
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

    handle = kernel.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFO
    if not handle:
        return ""
    try:
        buffer = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(1024)
        if kernel.QueryFullProcessImageNameW(
            handle, 0, buffer, ctypes.byref(size)
        ):
            return buffer.value
        return ""
    finally:
        kernel.CloseHandle(handle)


def _genie_windows():
    """Top-level windows of the Genie process, per backend."""
    from pywinauto import Desktop

    result = []

    for backend in ("uia", "win32"):
        for window in Desktop(backend=backend).windows():
            try:
                exe = _process_exe(window.process_id())
            except Exception:
                continue
            if exe.lower().endswith(EXE_NAME):
                result.append((backend, window))

    return result


def _find_button(windows_found):
    for _, window in windows_found:
        try:
            controls = [window] + window.descendants()
        except Exception:
            continue
        for control in controls:
            try:
                if re.search(BUTTON_TEXT_RE, control.window_text() or ""):
                    return control, window
            except Exception:
                continue
    return None, None


def _dump_diagnostic(windows_found):
    """Save to a file the windows and controls that pywinauto sees."""
    lines = [f"Windows of {EXE_NAME}: {len(windows_found)}"]

    for backend, window in windows_found:
        lines.append("")
        lines.append(
            f"[{backend}] window: title={window.window_text()!r} "
            f"class={window.friendly_class_name()!r} "
            f"visible={window.is_visible()} rect={window.rectangle()}"
        )
        try:
            for control in window.descendants():
                lines.append(
                    f"    class={control.friendly_class_name()!r} "
                    f"text={control.window_text()!r} "
                    f"rect={control.rectangle()}"
                )
        except Exception as error:
            lines.append(f"    (could not read the controls: {error})")

    DIAGNOSTIC_FILE.write_text("\n".join(lines), encoding="utf-8")
    return DIAGNOSTIC_FILE


def _click_by_image():
    """Fallback: look for the button crop on screen and click it."""
    image = next((p for p in BUTTON_IMAGES if p.exists()), None)
    if image is None:
        return False

    import pyautogui

    try:
        position = pyautogui.locateCenterOnScreen(str(image), confidence=0.8)
    except TypeError:  # without opencv there is no `confidence`
        position = pyautogui.locateCenterOnScreen(str(image))

    if position is None:
        return False

    pyautogui.click(position)
    return True


def click_sanction_check(timeout_s=30, pause_s=2):
    """Find the button in the Genie windows and click it.

    It retries for `timeout_s` seconds in case the window takes a while to
    appear. If it cannot find it, it tries the image fallback (if
    sanction_check_button.png exists) and, failing that, saves a diagnostic
    and raises RuntimeError.
    """
    start_time = time.time()

    while True:
        windows_found = _genie_windows()
        button, window = _find_button(windows_found)

        if button is not None:
            try:
                window.set_focus()
            except Exception:
                pass
            button.click_input()
            print("Clicked 'Perform Sanctions Check'.")
            return

        if time.time() - start_time >= timeout_s:
            break
        time.sleep(pause_s)

    if _click_by_image():
        print("Clicked 'Perform Sanctions Check' (by image).")
        return

    message = "Could not find the 'Perform Sanctions Check' button in Genie."
    if windows_found:
        diagnostic = _dump_diagnostic(windows_found)
        message += f" Diagnostic saved to: {diagnostic}"
    else:
        message += f" There are no visible windows of the {EXE_NAME} process."
    raise RuntimeError(message)

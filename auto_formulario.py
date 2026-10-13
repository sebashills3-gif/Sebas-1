#!/usr/bin/env python3
"""
Auto-respondedor de formularios web con Playwright (sync API) + Claude.

Flujo:
  1. Abre un navegador visible; tú navegas/inicias sesión a mano.
  2. El script espera en la terminal hasta que presiones ENTER.
  3. Escanea la página, clasifica cada pregunta (abierta / única / múltiple),
     consulta a Claude y rellena las respuestas.

Instalación:
    pip install playwright anthropic python-pptx pypdf   # (python-pptx/pypdf solo si usas --material)
    playwright install chromium
    export ANTHROPIC_API_KEY="sk-ant-..."      # (Windows PowerShell: $env:ANTHROPIC_API_KEY="...")

Uso:
    python auto_formulario.py --url "https://tu-sitio.com/formulario"
    python auto_formulario.py --backend cli     # usa el CLI `claude -p` en vez de la API
    python auto_formulario.py --backend cli --material ./presentaciones   # responde usando tus .pptx/.pdf/.txt/.md
    python auto_formulario.py --demo            # prueba con un formulario local de ejemplo

Usa esta herramienta solo en formularios/evaluaciones donde esté permitido
(prácticas, encuestas, formularios propios, etc.).
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from playwright.sync_api import Locator, Page, sync_playwright

# ════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN — PERSONALIZA ESTOS SELECTORES SEGÚN EL SITIO REAL
# ════════════════════════════════════════════════════════════════════════════

# Contenedor de CADA pregunta (el bloque que agrupa enunciado + campos/opciones).
# Ej.: ".pregunta", ".question", "div.que", "fieldset", "[data-qa='question']"
SELECTOR_PREGUNTA = ".pregunta"

# Elemento (dentro del contenedor) con el texto del enunciado.
# Ej.: ".enunciado", ".qtext", "legend", "h3". Si no existe, se usa el texto
# completo del contenedor menos el de las opciones.
SELECTOR_ENUNCIADO = ".enunciado"

# Contenedor de CADA opción (radio/checkbox + su texto). Es opcional: si no
# coincide nada, se usan directamente los <input> y se deduce su etiqueta.
# Ej.: ".opcion", ".answer label", "li.choice"
SELECTOR_OPCION = ".opcion"

# Campos de las posibles respuestas (normalmente no hace falta cambiarlos).
SELECTOR_ABIERTA = "textarea, input[type='text'], input:not([type])"
SELECTOR_RADIO = "input[type='radio'], [role='radio']"
SELECTOR_CHECKBOX = "input[type='checkbox'], [role='checkbox']"

# Botón de enviar. Déjalo en None para NO enviar (recomendado: revisa tú).
SELECTOR_BOTON_ENVIAR = None  # Ej.: "button[type='submit']"

MODELO = os.getenv("CLAUDE_MODEL", "claude-sonnet-5-5")
TIMEOUT_ACCION_MS = 5000

SYSTEM_PROMPT = (
    "Eres un asistente experto que responde preguntas de formularios. "
    "Sigue estrictamente el formato de salida pedido, sin explicaciones, "
    "sin comillas ni texto adicional."
)


# ════════════════════════════════════════════════════════════════════════════
# a) ESPERAR EL ACTIVADOR DEL USUARIO
# ════════════════════════════════════════════════════════════════════════════
def esperar_activador() -> None:
    """Bloquea hasta que el usuario presione ENTER en la terminal."""
    print("\n" + "=" * 60)
    print(" Navega en el navegador hasta la página con las preguntas.")
    print(" Cuando estés en la pantalla correcta, vuelve aquí y presiona ENTER.")
    print("=" * 60)
    try:
        input(">>> Presiona ENTER para iniciar el escaneo... ")
    except EOFError:
        print("(No hay terminal interactiva; continuando.)")


# ════════════════════════════════════════════════════════════════════════════
# MATERIAL DE CONSULTA (presentaciones, PDFs, notas) — opcional con --material
# ════════════════════════════════════════════════════════════════════════════
CONTEXTO_MATERIAL = ""  # se llena en main() si usas --material
MAX_CHARS_MATERIAL = 400_000  # ~100k tokens; evita desbordar el contexto de Claude


def _texto_pptx(ruta: Path) -> str:
    from pptx import Presentation  # pip install python-pptx

    def textos(shapes):
        for sh in shapes:
            if sh.shape_type == 6:  # grupo de formas
                yield from textos(sh.shapes)
            if sh.has_text_frame and sh.text_frame.text.strip():
                yield sh.text_frame.text.strip()
            if getattr(sh, "has_table", False) and sh.has_table:
                for fila in sh.table.rows:
                    yield " | ".join(c.text.strip() for c in fila.cells)

    partes = []
    for n, sl in enumerate(Presentation(str(ruta)).slides, 1):
        lineas = list(textos(sl.shapes))
        if sl.has_notes_slide and sl.notes_slide.notes_text_frame.text.strip():
            lineas.append("Notas: " + sl.notes_slide.notes_text_frame.text.strip())
        partes.append(f"[Diapositiva {n}]\n" + "\n".join(lineas))
    return "\n\n".join(partes)


def _texto_pdf(ruta: Path) -> str:
    from pypdf import PdfReader  # pip install pypdf

    return "\n\n".join(f"[Página {n}]\n{(pg.extract_text() or '').strip()}"
                        for n, pg in enumerate(PdfReader(str(ruta)).pages, 1))


def cargar_material(ruta: str) -> str:
    """Lee un archivo o carpeta (.pptx, .pdf, .txt, .md) y devuelve todo su texto."""
    base = Path(ruta).expanduser()
    if not base.exists():
        sys.exit(f"--material: no existe la ruta '{ruta}'")
    archivos = [base] if base.is_file() else sorted(p for p in base.rglob("*") if p.is_file())
    bloques = []
    for a in archivos:
        ext = a.suffix.lower()
        try:
            if ext == ".pptx":
                txt = _texto_pptx(a)
            elif ext == ".pdf":
                txt = _texto_pdf(a)
            elif ext in (".txt", ".md"):
                txt = a.read_text(encoding="utf-8", errors="ignore")
            else:
                continue  # .ppt antiguo, imágenes, etc.: se ignoran
        except Exception as e:
            print(f"  (no se pudo leer {a.name}: {type(e).__name__}: {e})")
            continue
        if txt.strip():
            print(f"  Material cargado: {a.name} ({len(txt):,} caracteres)")
            bloques.append(f"=== ARCHIVO: {a.name} ===\n{txt.strip()}")
    texto = "\n\n".join(bloques)
    if not texto:
        sys.exit("--material: no se encontró texto legible (.pptx/.pdf/.txt/.md).")
    if len(texto) > MAX_CHARS_MATERIAL:
        print(f"  AVISO: material muy largo; se recorta a {MAX_CHARS_MATERIAL:,} caracteres.")
        texto = texto[:MAX_CHARS_MATERIAL]
    return texto


def construir_system() -> str:
    if not CONTEXTO_MATERIAL:
        return SYSTEM_PROMPT
    return (
        f"{SYSTEM_PROMPT}\n\nTu fuente principal es el MATERIAL DE REFERENCIA de abajo. "
        "Basa tus respuestas en él; si la respuesta no aparece ahí, usa tu conocimiento general.\n\n"
        f"<material_de_referencia>\n{CONTEXTO_MATERIAL}\n</material_de_referencia>"
    )


# ════════════════════════════════════════════════════════════════════════════
# b) CONSULTAR A CLAUDE
# ════════════════════════════════════════════════════════════════════════════
def consultar_claude(prompt: str, backend: str = "api") -> str:
    """Envía `prompt` a Claude y devuelve el texto de la respuesta."""
    if backend == "api":
        import anthropic  # import perezoso: solo si se usa

        cliente = anthropic.Anthropic()  # lee ANTHROPIC_API_KEY del entorno
        resp = cliente.messages.create(
            model=MODELO,
            max_tokens=1024,
            system=construir_system(),
            messages=[{"role": "user", "content": prompt}],
        )
        return "".join(b.text for b in resp.content if b.type == "text").strip()

    if backend == "cli":
        # Requiere tener instalado y autenticado el CLI de Claude Code.
        # El prompt va por stdin (sin límite de longitud del comando, importante
        # si usas --material) y se desactivan las herramientas del agente.
        res = subprocess.run(
            ["claude", "-p", "--tools", ""],
            input=f"{construir_system()}\n\n---\n\n{prompt}",
            capture_output=True, text=True, encoding="utf-8", timeout=300,
        )
        if res.returncode != 0:
            raise RuntimeError(f"claude CLI falló: {res.stderr.strip()}")
        return res.stdout.strip()

    if backend == "mock":
        # Respuestas falsas SOLO para probar el flujo sin gastar API.
        if "SOLO el número" in prompt:
            return "2"
        if "arreglo JSON" in prompt:
            return "[1, 3]"
        return "Respuesta de prueba (mock)."

    raise ValueError(f"backend desconocido: {backend}")


def pedir_abierta(enunciado: str, backend: str) -> str:
    prompt = (
        f"Responde de forma concisa y directa (máximo 2-3 frases) a:\n\n{enunciado}\n\n"
        "Devuelve SOLO la respuesta."
    )
    return consultar_claude(prompt, backend)


def _lista_numerada(opciones: list[str]) -> str:
    return "\n".join(f"{i}. {t}" for i, t in enumerate(opciones, 1))


def pedir_unica(enunciado: str, opciones: list[str], backend: str) -> int | None:
    """Devuelve el índice (0-based) de la opción elegida, o None si no se pudo."""
    prompt = (
        f"Pregunta:\n{enunciado}\n\nOpciones:\n{_lista_numerada(opciones)}\n\n"
        "Elige la ÚNICA opción correcta. Devuelve SOLO el número de la opción "
        "(por ejemplo: 2). Nada más."
    )
    texto = consultar_claude(prompt, backend)
    m = re.search(r"\d+", texto)
    if m and 1 <= int(m.group()) <= len(opciones):
        return int(m.group()) - 1
    # Respaldo: Claude devolvió el texto de la opción en vez del número.
    t = texto.strip().lower()
    for i, op in enumerate(opciones):
        if op.strip().lower() == t:
            return i
    return None


def pedir_multiple(enunciado: str, opciones: list[str], backend: str) -> list[int]:
    """Devuelve la lista de índices (0-based) de las opciones correctas."""
    prompt = (
        f"Pregunta:\n{enunciado}\n\nOpciones:\n{_lista_numerada(opciones)}\n\n"
        "Selecciona TODAS las opciones correctas (una o más). Devuelve SOLO un "
        "arreglo JSON con los números, por ejemplo: [1, 3]. Nada más."
    )
    texto = consultar_claude(prompt, backend)
    try:
        m = re.search(r"\[.*?\]", texto, re.S)
        nums = json.loads(m.group()) if m else []
    except json.JSONDecodeError:
        nums = []
    if not nums:  # respaldo: separado por comas / cualquier número suelto
        nums = [int(n) for n in re.findall(r"\d+", texto)]
    return sorted({int(n) - 1 for n in nums if 1 <= int(n) <= len(opciones)})


# ════════════════════════════════════════════════════════════════════════════
# c) DETECTAR EL TIPO DE PREGUNTA E INTERACTUAR CON LA PÁGINA
# ════════════════════════════════════════════════════════════════════════════
_JS_ETIQUETA = """
el => {
  // Texto asociado a un radio/checkbox: <label for>, label padre, aria-label,
  // o el texto del contenedor más cercano que no incluya otros inputs.
  if (el.labels && el.labels.length) return el.labels[0].innerText;
  const lab = el.closest('label'); if (lab) return lab.innerText;
  const aria = el.getAttribute('aria-label'); if (aria) return aria;
  let n = el;
  for (let i = 0; i < 3 && n.parentElement; i++) {
    n = n.parentElement;
    const inputs = n.querySelectorAll('input[type=radio],input[type=checkbox],[role=radio],[role=checkbox]');
    if (inputs.length === 1 && n.innerText.trim()) return n.innerText;
  }
  return el.innerText || el.value || '';
}
"""


def _limpiar(t: str) -> str:
    return re.sub(r"\s+", " ", t or "").strip()


def clasificar(pregunta: Locator) -> str | None:
    """Devuelve 'multiple', 'unica', 'abierta' o None si no se reconoce."""
    # El orden importa: checkbox y radio tienen prioridad sobre texto libre.
    if pregunta.locator(SELECTOR_CHECKBOX).count() > 0:
        return "multiple"
    if pregunta.locator(SELECTOR_RADIO).count() > 0:
        return "unica"
    if pregunta.locator(SELECTOR_ABIERTA).count() > 0:
        return "abierta"
    return None


def extraer_enunciado(pregunta: Locator) -> str:
    try:
        el = pregunta.locator(SELECTOR_ENUNCIADO).first
        if el.count() > 0:
            return _limpiar(el.inner_text(timeout=TIMEOUT_ACCION_MS))
    except Exception:
        pass
    return _limpiar(pregunta.inner_text(timeout=TIMEOUT_ACCION_MS))  # respaldo


def extraer_opciones(pregunta: Locator, selector_campo: str) -> tuple[list[str], list[Locator]]:
    """Devuelve (textos, campos) de las opciones, en el mismo orden."""
    campos = pregunta.locator(selector_campo)
    textos, locs = [], []
    for i in range(campos.count()):
        campo = campos.nth(i)
        try:
            texto = _limpiar(campo.evaluate(_JS_ETIQUETA))
        except Exception:
            texto = ""
        textos.append(texto or f"Opción {i + 1}")
        locs.append(campo)
    return textos, locs


def marcar(campo: Locator) -> None:
    """Marca un radio/checkbox; si el input está oculto por CSS, hace clic en su etiqueta."""
    try:
        campo.check(timeout=TIMEOUT_ACCION_MS)
        return
    except Exception:
        pass
    try:  # inputs ocultos con estilo personalizado
        campo.locator("xpath=ancestor-or-self::label[1]").first.click(timeout=TIMEOUT_ACCION_MS)
        return
    except Exception:
        pass
    try:  # roles ARIA (div role="radio")
        campo.click(timeout=TIMEOUT_ACCION_MS)
    except Exception:
        campo.check(force=True, timeout=TIMEOUT_ACCION_MS)


def responder_pregunta(pregunta: Locator, n: int, backend: str) -> None:
    tipo = clasificar(pregunta)
    if tipo is None:
        print(f"[{n}] Tipo no reconocido; se omite.")
        return

    enunciado = extraer_enunciado(pregunta)
    print(f"[{n}] ({tipo.upper()}) {enunciado[:90]}")

    if tipo == "abierta":
        respuesta = pedir_abierta(enunciado, backend)
        pregunta.locator(SELECTOR_ABIERTA).first.fill(respuesta, timeout=TIMEOUT_ACCION_MS)
        print(f"     -> {respuesta[:90]}")

    elif tipo == "unica":
        textos, campos = extraer_opciones(pregunta, SELECTOR_RADIO)
        idx = pedir_unica(enunciado, textos, backend)
        if idx is None:
            print("     -> Claude no devolvió una opción válida; se omite.")
            return
        marcar(campos[idx])
        print(f"     -> {textos[idx][:90]}")

    elif tipo == "multiple":
        textos, campos = extraer_opciones(pregunta, SELECTOR_CHECKBOX)
        indices = pedir_multiple(enunciado, textos, backend)
        if not indices:
            print("     -> Claude no devolvió opciones válidas; se omite.")
            return
        for idx in indices:
            marcar(campos[idx])
        print(f"     -> {[textos[i][:40] for i in indices]}")


def procesar_pagina(page: Page, backend: str) -> None:
    preguntas = page.locator(SELECTOR_PREGUNTA)
    total = preguntas.count()
    if total == 0:
        print(f"No se encontró ninguna pregunta con el selector '{SELECTOR_PREGUNTA}'.")
        print("Ajusta SELECTOR_PREGUNTA al inicio del script.")
        return
    print(f"\nSe encontraron {total} preguntas.\n")

    for i in range(total):
        try:
            responder_pregunta(preguntas.nth(i), i + 1, backend)
        except Exception as e:  # un fallo en una pregunta no detiene las demás
            print(f"[{i + 1}] ERROR: {type(e).__name__}: {str(e)[:150]}")

    if SELECTOR_BOTON_ENVIAR:
        try:
            page.locator(SELECTOR_BOTON_ENVIAR).first.click(timeout=TIMEOUT_ACCION_MS)
            print("Formulario enviado.")
        except Exception as e:
            print(f"No se pudo enviar: {e}")
    else:
        print("\nListo. No se envió el formulario: revisa las respuestas y envíalo tú.")


# ════════════════════════════════════════════════════════════════════════════
# Formulario de demostración local (para probar sin ningún sitio real)
# ════════════════════════════════════════════════════════════════════════════
DEMO_HTML = """<!doctype html><meta charset=utf-8><title>Demo</title>
<div class=pregunta><p class=enunciado>1. ¿Cuál es la capital de Francia?</p>
  <textarea name=q1></textarea></div>
<div class=pregunta><p class=enunciado>2. ¿Cuál es el planeta más grande?</p>
  <div class=opcion><label><input type=radio name=q2> Marte</label></div>
  <div class=opcion><label><input type=radio name=q2> Júpiter</label></div>
  <div class=opcion><label><input type=radio name=q2> Venus</label></div></div>
<div class=pregunta><p class=enunciado>3. Selecciona los números primos</p>
  <div class=opcion><input type=checkbox id=a><label for=a>2</label></div>
  <div class=opcion><input type=checkbox id=b><label for=b>4</label></div>
  <div class=opcion><input type=checkbox id=c><label for=c>5</label></div>
  <div class=opcion><input type=checkbox id=d><label for=d>9</label></div></div>
<button type=submit>Enviar</button>"""


def crear_demo() -> str:
    ruta = Path(tempfile.mkdtemp()) / "demo.html"
    ruta.write_text(DEMO_HTML, encoding="utf-8")
    return ruta.as_uri()


# ════════════════════════════════════════════════════════════════════════════
# main
# ════════════════════════════════════════════════════════════════════════════
def elegir_backend(pedido: str) -> str:
    if pedido != "auto":
        return pedido
    if os.getenv("ANTHROPIC_API_KEY"):
        return "api"
    if shutil.which("claude"):
        return "cli"
    sys.exit("Define ANTHROPIC_API_KEY o instala el CLI `claude` (o usa --backend mock).")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--url", help="URL inicial a abrir (opcional)")
    ap.add_argument("--backend", choices=["auto", "api", "cli", "mock"], default="auto")
    ap.add_argument("--material", help="archivo o carpeta (.pptx/.pdf/.txt/.md) que Claude usará como base")
    ap.add_argument("--demo", action="store_true", help="usa un formulario local de ejemplo")
    ap.add_argument("--headless", action="store_true", help="sin ventana (solo pruebas)")
    ap.add_argument("--auto-start", action="store_true", help="no espera ENTER (solo pruebas)")
    args = ap.parse_args()

    backend = elegir_backend(args.backend)
    url = crear_demo() if args.demo else args.url
    print(f"Backend de Claude: {backend}")
    if args.material:
        global CONTEXTO_MATERIAL
        print("Cargando material de consulta...")
        CONTEXTO_MATERIAL = cargar_material(args.material)

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=args.headless)
        except Exception:
            # Respaldo: usar un Chromium ya instalado en el sistema (variable opcional).
            ruta = os.getenv("CHROMIUM_PATH")
            if not ruta:
                raise
            browser = p.chromium.launch(headless=args.headless, executable_path=ruta)
        page = browser.new_page()
        if url:
            page.goto(url)

        if not args.auto_start:
            esperar_activador()
        procesar_pagina(page, backend)

        if not args.headless:
            input("\nPresiona ENTER para cerrar el navegador... ")
        else:
            for etiqueta, selector in [("textarea", "textarea"), ("radio", "input[type=radio]:checked"),
                                       ("checkbox", "input[type=checkbox]:checked")]:
                print(f"[verificación] {etiqueta}:", page.locator(selector).count())
        browser.close()


if __name__ == "__main__":
    main()

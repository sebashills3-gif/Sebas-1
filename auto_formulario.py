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
    python auto_formulario.py --demo-material --backend cli --material ./material   # verifica que usa TU material

Modelos: las preguntas fáciles (recordar un dato) usan Haiku y las difíciles (cálculos,
desarrollo matemático) usan Sonnet automáticamente. Cámbialo con --model-facil / --model-dificil
(ej. --model-dificil claude-opus-5-5) o fuerza un solo modelo con --model.

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

# Enrutamiento por dificultad: una consulta barata a Haiku (sin el material) decide si la
# pregunta es FÁCIL (recordar/ubicar un dato) o DIFÍCIL (cálculo, desarrollo, varios pasos).
MODELO_FACIL = os.getenv("CLAUDE_MODEL_FACIL", "claude-haiku-5-5")
MODELO_DIFICIL = os.getenv("CLAUDE_MODEL_DIFICIL", "claude-sonnet-5-5")  # o "claude-opus-5-5"
ENRUTAR = True  # pasa a False con --model (un solo modelo para todo)

# Si el enunciado contiene estas palabras se considera DIFÍCIL sin gastar la consulta de clasificación.
PATRON_DIFICIL = re.compile(
    r"calcul|resuelv|resuelve|determina|demuestr|deduc|deriv|despej|desarroll|paso a paso|"
    r"procedimiento|encuentra el valor|obt[eé]n", re.I)
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
def consultar_claude(prompt: str, backend: str = "api", modelo: str | None = None,
                     usar_material: bool = True, max_tokens: int = 1024) -> str:
    """Envía `prompt` a Claude y devuelve el texto de la respuesta."""
    modelo = modelo or MODELO_FACIL
    system = construir_system() if usar_material else SYSTEM_PROMPT
    if backend == "api":
        import anthropic  # import perezoso: solo si se usa

        cliente = anthropic.Anthropic()  # lee ANTHROPIC_API_KEY del entorno
        resp = cliente.messages.create(
            model=modelo,
            max_tokens=max_tokens,
            system=system,
            messages=[{"role": "user", "content": prompt}],
        )
        return "".join(b.text for b in resp.content if b.type == "text").strip()

    if backend == "cli":
        # Requiere tener instalado y autenticado el CLI de Claude Code.
        # El prompt va por stdin (sin límite de longitud del comando, importante
        # si usas --material) y se desactivan las herramientas del agente.
        error = ""
        for intento in (1, 2):  # un reintento si el CLI se cuelga o devuelve vacío
            try:
                res = subprocess.run(
                    ["claude", "-p", "--tools", "", "--model", modelo],
                    input=f"{system}\n\n---\n\n{prompt}",
                    capture_output=True, text=True, encoding="utf-8", timeout=240,
                )
                if res.returncode == 0 and res.stdout.strip():
                    return res.stdout.strip()
                error = res.stderr.strip() or "respuesta vacía"
            except subprocess.TimeoutExpired:
                error = "tiempo de espera agotado"
            if intento == 1:
                print(f"     (reintentando: {error[:80]})")
        raise RuntimeError(f"claude CLI falló: {error}")

    if backend == "mock":
        # Respuestas falsas SOLO para probar el flujo sin gastar API.
        if "SOLO la palabra FACIL o DIFICIL" in prompt:
            return "FACIL"
        if "arreglo JSON" in prompt:
            return '["A", "C"]'
        if "SOLO la letra" in prompt:
            return "B"
        return "Respuesta de prueba (mock)."

    raise ValueError(f"backend desconocido: {backend}")


def clasificar_dificultad(enunciado: str, opciones: list[str], backend: str) -> bool:
    """True si la pregunta es DIFÍCIL (cálculo/desarrollo). Si algo falla, asume difícil (más seguro)."""
    if not ENRUTAR:
        return False
    if PATRON_DIFICIL.search(enunciado) or len(enunciado) > 700:
        return True
    prompt = (
        "Clasifica la dificultad de esta pregunta de examen. Responde SOLO la palabra FACIL o DIFICIL.\n"
        "FACIL: basta recordar o ubicar un dato, definición, concepto o lista.\n"
        "DIFICIL: requiere cálculos, desarrollo matemático o algebraico de varios pasos, aplicar un "
        "procedimiento o razonar con varias condiciones.\n\n"
        f"Pregunta: {enunciado}\n"
        + (f"Opciones:\n{_lista_letras(opciones)}\n" if opciones else "")
    )
    try:
        r = consultar_claude(prompt, backend, modelo=MODELO_FACIL, usar_material=False, max_tokens=10)
        return "DIFICIL" in r.upper().replace("Í", "I")
    except Exception:
        return True


def pedir_abierta(enunciado: str, backend: str, modelo: str, dificil: bool = False) -> str:
    if dificil:
        prompt = (
            f"Resuelve este ejercicio:\n\n{enunciado}\n\n"
            "Desarrolla el procedimiento paso a paso de forma ordenada y compacta (fórmulas y cálculos "
            "clave, sin explicaciones largas) y termina con una línea 'Respuesta final: ...'. "
            "Texto plano, sin markdown ni LaTeX."
        )
        return consultar_claude(prompt, backend, modelo, max_tokens=2048)
    prompt = (
        f"Responde de forma concisa y directa (máximo 2-3 frases) a:\n\n{enunciado}\n\n"
        "Devuelve SOLO la respuesta."
    )
    return consultar_claude(prompt, backend, modelo)


def _letra(i: int) -> str:
    return chr(ord("A") + i) if i < 26 else str(i + 1)


def _lista_letras(opciones: list[str]) -> str:
    # Letras (A, B, C...) en vez de números: así no se confunden con opciones numéricas.
    return "\n".join(f"{_letra(i)}) {t}" for i, t in enumerate(opciones))


def _letras_a_indices(texto: str, opciones: list[str]) -> list[int]:
    """Convierte la respuesta de Claude (letras, o el texto de las opciones) en índices 0-based."""
    n = len(opciones)
    # Se quitan comillas/corchetes y se buscan letras sueltas: A, B, "C", [A, C], A) ...
    limpio = texto.replace("\n", " ")
    for patron in (r"(?<![A-Za-zÁ-ú0-9])([A-Z])(?![A-Za-zÁ-ú0-9])",   # primero mayúsculas
                   r"(?<![A-Za-zÁ-ú0-9])([a-z])(?![A-Za-zÁ-ú0-9])"):  # luego minúsculas ("b.")
        idx = sorted({ord(l.upper()) - 65 for l in re.findall(patron, limpio)
                      if 0 <= ord(l.upper()) - 65 < n})
        if idx:
            return idx
    # Respaldo: Claude devolvió el texto exacto de las opciones.
    partes = {x.strip().strip("\"'[]").lower() for x in re.split(r"[,\n;]", texto)}
    return [i for i, op in enumerate(opciones) if op.strip().lower() in partes]


def pedir_unica(enunciado: str, opciones: list[str], backend: str, modelo: str) -> int | None:
    """Devuelve el índice (0-based) de la opción elegida, o None si no se pudo."""
    prompt = (
        f"Pregunta:\n{enunciado}\n\nOpciones:\n{_lista_letras(opciones)}\n\n"
        "Elige la ÚNICA opción correcta. Devuelve SOLO la letra de la opción "
        "(por ejemplo: B). Nada más."
    )
    idx = _letras_a_indices(consultar_claude(prompt, backend, modelo), opciones)
    return idx[0] if idx else None


def pedir_multiple(enunciado: str, opciones: list[str], backend: str, modelo: str) -> list[int]:
    """Devuelve la lista de índices (0-based) de las opciones correctas."""
    prompt = (
        f"Pregunta:\n{enunciado}\n\nOpciones:\n{_lista_letras(opciones)}\n\n"
        "Selecciona TODAS las opciones correctas (una o más). Devuelve SOLO un "
        'arreglo JSON con las letras, por ejemplo: ["A", "C"]. Nada más.'
    )
    return _letras_a_indices(consultar_claude(prompt, backend, modelo), opciones)


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
    textos, campos = [], []
    if tipo == "unica":
        textos, campos = extraer_opciones(pregunta, SELECTOR_RADIO)
    elif tipo == "multiple":
        textos, campos = extraer_opciones(pregunta, SELECTOR_CHECKBOX)

    dificil = clasificar_dificultad(enunciado, textos, backend)
    modelo = MODELO_DIFICIL if dificil else MODELO_FACIL
    etiqueta = ("DIFÍCIL" if dificil else "fácil") if ENRUTAR else "modelo único"
    print(f"[{n}] ({tipo.upper()} | {etiqueta} | {modelo}) {enunciado[:80]}")

    if tipo == "abierta":
        respuesta = pedir_abierta(enunciado, backend, modelo, dificil)
        pregunta.locator(SELECTOR_ABIERTA).first.fill(respuesta, timeout=TIMEOUT_ACCION_MS)
        print(f"     -> {respuesta[:90]}")

    elif tipo == "unica":
        idx = pedir_unica(enunciado, textos, backend, modelo)
        if idx is None:
            print("     -> Claude no devolvió una opción válida; se omite.")
            return
        marcar(campos[idx])
        print(f"     -> {textos[idx][:90]}")

    elif tipo == "multiple":
        indices = pedir_multiple(enunciado, textos, backend, modelo)
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


# Demo que VERIFICA el uso del material: preguntas cuyas respuestas salen de los .txt de
# Teoría de Juegos / Economía Administrativa (valores de las matrices, láminas, etc.).
# "ok" = palabras que debe contener la respuesta abierta, o los índices (0-based) correctos.
DEMO_MATERIAL = [
    {"tipo": "abierta", "dif": "fácil",
     "q": "¿Cómo se llama el curso al que pertenecen las presentaciones de Teoría de Juegos con información incompleta?",
     "ok": ["decisiones estrat"]},
    {"tipo": "unica", "dif": "fácil",
     "q": "En el dilema del prisionero con información incompleta (presentación TJ-9), ¿qué pago recibe el "
          "sospechoso 1 si ambos hablan y el sospechoso 2 es de tipo Compa?",
     "ops": ["−6", "−9", "−3", "2"], "ok": [1]},
    {"tipo": "unica", "dif": "difícil",
     "q": "En el ejemplo de selección adversa de la presentación TJ-III, si el comprador ofrece p = 20, "
          "¿cuánto vale E[v2 | venta]?",
     "ops": ["24", "20", "19", "14"], "ok": [2]},
    {"tipo": "unica", "dif": "fácil",
     "q": "En la batalla de los sexos bayesiana de TJ-9, si Cris elige Cine con certeza y π ≥ 1/4, ¿qué perfil "
          "(λ, µG, µS) es un equilibrio bayesiano de Nash en estrategias puras?",
     "ops": ["(0, 0, 1)", "(3/4, 1/4, 1/4)", "(1, 1, 0)", "(1, 0, 1)"], "ok": [2]},
    {"tipo": "abierta", "dif": "difícil",
     "q": "En el dilema del prisionero con información incompleta de la presentación TJ-9, el sospechoso 1 cree "
          "que el sospechoso 2 es Judas con probabilidad ρ. Calcula el valor de ρ que hace indiferente al "
          "sospechoso 1 entre Hablar y Callar. Muestra el desarrollo.",
     "ok": ["5/8", "0.625", "0,625"]},
    {"tipo": "abierta", "dif": "fácil",
     "q": "En el caso FríoTech México del Apunte 2, ¿cuántos controladores electrónicos necesita la empresa al año?",
     "ok": ["15,000", "15000", "15 000", "15.000"]},
    {"tipo": "unica", "dif": "fácil",
     "q": "En el ejemplo de integración atemperada (Make AND Buy) del Apunte 5, con una demanda total de "
          "100,000 unidades, ¿cuántas se compran al proveedor externo?",
     "ops": ["40,000", "70,000", "60,000", "100,000"], "ok": [2]},
    {"tipo": "multiple", "dif": "fácil",
     "q": "Según el Apunte 1, ¿qué dimensiones de la transacción destaca Williamson?",
     "ops": ["Tamaño de la empresa", "Frecuencia", "Reputación del proveedor",
             "Especificidad de activos", "Incertidumbre"], "ok": [1, 3, 4]},
    {"tipo": "multiple", "dif": "fácil",
     "q": "Según las láminas de TJ-IV (dos jurados, regla de unanimidad), ¿qué afirmaciones son correctas?",
     "ops": ["(CA, CA) es un equilibrio bayesiano de Nash.",
             "Un jurado solo es pivotal si el otro jurado vota C.",
             "Sin señales y con q > 1/2, votar C es una estrategia débilmente dominante.",
             "La persona acusada es condenada si al menos un jurado vota C."], "ok": [1, 2]},
]


def crear_demo_material() -> str:
    bloques = []
    for n, q in enumerate(DEMO_MATERIAL, 1):
        if q["tipo"] == "abierta":
            campo = "<textarea rows=4 cols=60></textarea>"
        else:
            kind = "radio" if q["tipo"] == "unica" else "checkbox"
            campo = "".join(
                f"<div class=opcion><label><input type={kind} name=q{n}> {op}</label></div>"
                for op in q["ops"])
        bloques.append(f"<div class=pregunta><p class=enunciado>{n}. {q['q']}</p>{campo}</div>")
    html = "<!doctype html><meta charset=utf-8><title>Demo material</title>" + "\n".join(bloques)
    ruta = Path(tempfile.mkdtemp()) / "demo_material.html"
    ruta.write_text(html, encoding="utf-8")
    return ruta.as_uri()


def evaluar_demo_material(page: Page) -> None:
    """Califica las respuestas de la demo contra las respuestas esperadas."""
    preguntas = page.locator(SELECTOR_PREGUNTA)
    aciertos = 0
    print("\n" + "=" * 60 + "\n RESULTADO DE LA DEMO (¿usó tu material?)\n" + "=" * 60)
    for i, q in enumerate(DEMO_MATERIAL):
        bloque = preguntas.nth(i)
        if q["tipo"] == "abierta":
            valor = bloque.locator("textarea").first.input_value().lower()
            ok = any(k in valor for k in q["ok"])
            esperado = " / ".join(q["ok"])
        else:
            campos = bloque.locator("input")
            marcados = [j for j in range(campos.count()) if campos.nth(j).is_checked()]
            ok = marcados == q["ok"]
            esperado = ", ".join(q["ops"][j] for j in q["ok"])
        aciertos += ok
        print(f" {'✔' if ok else '✘'} P{i + 1} ({q['tipo']}, {q['dif']}) esperado: {esperado}")
    print(f"\n Aciertos: {aciertos}/{len(DEMO_MATERIAL)}")
    if not CONTEXTO_MATERIAL:
        print(" (Corriste SIN --material: esto es la línea base con conocimiento general.)")


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
    global MODELO_FACIL, MODELO_DIFICIL, ENRUTAR, CONTEXTO_MATERIAL
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("--url", help="URL inicial a abrir (opcional)")
    ap.add_argument("--backend", choices=["auto", "api", "cli", "mock"], default="auto")
    ap.add_argument("--material", help="archivo o carpeta (.pptx/.pdf/.txt/.md) que Claude usará como base")
    ap.add_argument("--model", help="un solo modelo para TODAS las preguntas (desactiva el enrutamiento)")
    ap.add_argument("--model-facil", help=f"modelo para preguntas fáciles (por defecto {MODELO_FACIL})")
    ap.add_argument("--model-dificil", help=f"modelo para preguntas difíciles (por defecto {MODELO_DIFICIL})")
    ap.add_argument("--demo-material", action="store_true",
                    help="formulario local con preguntas que solo se contestan con tu material, y lo califica")
    ap.add_argument("--demo", action="store_true", help="usa un formulario local de ejemplo")
    ap.add_argument("--headless", action="store_true", help="sin ventana (solo pruebas)")
    ap.add_argument("--auto-start", action="store_true", help="no espera ENTER (solo pruebas)")
    args = ap.parse_args()

    if args.model:
        MODELO_FACIL = MODELO_DIFICIL = args.model
        ENRUTAR = False
    if args.model_facil:
        MODELO_FACIL = args.model_facil
    if args.model_dificil:
        MODELO_DIFICIL = args.model_dificil
    backend = elegir_backend(args.backend)
    url = crear_demo_material() if args.demo_material else (crear_demo() if args.demo else args.url)
    modelos = MODELO_FACIL if not ENRUTAR else f"fácil={MODELO_FACIL}, difícil={MODELO_DIFICIL}"
    print(f"Backend de Claude: {backend} | modelos: {modelos}")
    if args.material:
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
            try:
                page.goto(url)
            except Exception as e:  # URL mal escrita o sin conexión: no abortar, navegas tú a mano
                print(f"AVISO: no se pudo abrir '{url}' ({str(e).splitlines()[0][:80]}).")
                print("Navega manualmente en el navegador a la página correcta.")

        if not args.auto_start:
            esperar_activador()
        procesar_pagina(page, backend)
        if args.demo_material:
            evaluar_demo_material(page)

        if not args.headless:
            input("\nPresiona ENTER para cerrar el navegador... ")
        else:
            for etiqueta, selector in [("textarea", "textarea"), ("radio", "input[type=radio]:checked"),
                                       ("checkbox", "input[type=checkbox]:checked")]:
                print(f"[verificación] {etiqueta}:", page.locator(selector).count())
        browser.close()


if __name__ == "__main__":
    main()

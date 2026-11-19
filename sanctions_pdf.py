"""Revisa el PDF del Sanctions Check antes de crear securities en IMPACT.

El PDF tiene la estructura del PreStaging File: un bloque por identificador
(By ISIN, By SEDOL, By CUSIP, By COMMON, By FIGI ID), cada uno con las
columnas No / identificador / Sanctions Check, y un bloque "Combine Status".
La fila "No" k del PDF corresponde a la k-ésima fila pegada en el PreStaging
File (la fila 5 del Excel es la No 1).

Un security solo pasa si TODAS sus filas dicen PASS. Ante cualquier duda
(PDF ilegible, fila que no aparece, resultado distinto de PASS) el security
se bloquea: la regla es no tocar nada que no se haya podido comprobar.

Requiere: pip install pdfplumber
"""

import re
import statistics
from pathlib import Path

from prestaging import construir_filas, duenios_de_filas, extraer_securities

IDENTIFICADORES = ("ISIN", "SEDOL", "CUSIP", "COMMON", "FIGI_ID")

_TITULO = re.compile(
    r"^(ISIN|SEDOL|CUSIP|COMMON|FIGI(?:_ID)?)S?$", re.IGNORECASE
)
_TITULO_PEGADO = re.compile(
    r"^BY(ISIN|SEDOL|CUSIP|COMMON|FIGI(?:_ID)?)S?$", re.IGNORECASE
)
_CLAVE = {
    "ISIN": "ISIN", "SEDOL": "SEDOL", "CUSIP": "CUSIP",
    "COMMON": "COMMON", "FIGI": "FIGI_ID", "FIGI_ID": "FIGI_ID",
}
_ENCABEZADOS = {"NO", "SANCTIONS", "CHECK", "STATUS", "COMMENTS", "ID"}
_SIN_DATO = ("#N/A", "N.A.", "N/A")


class PdfSanctionsError(Exception):
    pass


# ============================================================
# LECTURA DEL PDF
# ============================================================

def _palabras_por_pagina(ruta):
    import pdfplumber

    paginas = []
    with pdfplumber.open(str(ruta)) as pdf:
        for pagina in pdf.pages:
            palabras = pagina.extract_words(use_text_flow=False)
            paginas.append(sorted(palabras, key=lambda w: (round(w["top"]), w["x0"])))
    return paginas


def _cx(w):
    return (w["x0"] + w["x1"]) / 2


def _cy(w):
    return (w["top"] + w["bottom"]) / 2


def _titulos(palabras):
    """Títulos de bloque de la página: {clave: {centro, top, bottom}}."""
    titulos = {}

    for i, w in enumerate(palabras):
        texto = w["text"].strip()

        pegado = _TITULO_PEGADO.match(texto.upper())
        if pegado:
            clave = _CLAVE[pegado.group(1).upper()]
            titulos[clave] = {"centro": _cx(w), "top": w["top"], "bottom": w["bottom"]}
            continue

        if texto.upper() == "BY":
            for sig in palabras[i + 1:]:
                if abs(sig["top"] - w["top"]) <= 3 and sig["x0"] >= w["x1"] - 1:
                    coincide = _TITULO.match(sig["text"].strip())
                    if coincide:
                        clave = _CLAVE[coincide.group(1).upper()]
                        titulos[clave] = {
                            "centro": (w["x0"] + sig["x1"]) / 2,
                            "top": w["top"],
                            "bottom": w["bottom"],
                        }
                    break

        if texto.upper() == "COMBINE":
            for sig in palabras[i + 1:]:
                if abs(sig["top"] - w["top"]) <= 3 and sig["x0"] >= w["x1"] - 1:
                    if sig["text"].strip().upper() == "STATUS":
                        titulos["COMBINE"] = {
                            "centro": (w["x0"] + sig["x1"]) / 2,
                            "top": w["top"],
                            "bottom": w["bottom"],
                        }
                    break

    return titulos


def _bloque_de(w, titulos):
    return min(titulos, key=lambda k: abs(_cx(w) - titulos[k]["centro"]))


def _geometria(palabras, titulos):
    """Columnas de estado (con sus límites), columna No y borde inferior del
    encabezado.

    Los resultados quedan pegados al borde izquierdo de su celda (no
    centrados como el encabezado), así que la columna de estado de cada
    bloque se delimita desde justo antes de su encabezado "Sanctions" hasta
    donde empieza el siguiente bloque.
    """
    if not any(k in titulos for k in IDENTIFICADORES):
        raise PdfSanctionsError("no se encontraron los títulos 'By ...'.")

    fondo_titulos = max(t["bottom"] for t in titulos.values())
    alto = fondo_titulos - min(t["top"] for t in titulos.values())
    limite = fondo_titulos + 3 * max(alto, 8) + 6

    orden = sorted(titulos, key=lambda k: titulos[k]["centro"])
    sub = {k: [] for k in titulos}
    fondo_encabezado = fondo_titulos

    for w in palabras:
        if fondo_titulos - 0.5 <= w["top"] <= limite:
            sub[_bloque_de(w, titulos)].append(w)
            texto = w["text"].strip().upper()
            if texto in _ENCABEZADOS or _TITULO.match(texto):
                fondo_encabezado = max(fondo_encabezado, w["bottom"])

    izquierda, col_no, no_izq = {}, None, {}
    combine = None

    for bloque in orden:
        palabras_b = sorted(sub[bloque], key=lambda w: w["x0"])

        if bloque == "COMBINE":
            estado = [w for w in palabras_b if w["text"].strip().upper() == "STATUS"
                      and w["top"] > fondo_titulos - 0.5]
            comentarios = [w for w in palabras_b
                           if w["text"].strip().upper() == "COMMENTS"]
            if estado:
                sc = _cx(estado[0])
                if comentarios:
                    cc = _cx(comentarios[0])
                    combine = (sc - (cc - sc) / 2, (sc + cc) / 2)
                else:
                    combine = (estado[0]["x0"] - 15, estado[0]["x1"] + 40)
            continue

        no = next((w for w in palabras_b if w["text"].strip().upper() == "NO"), None)
        sanciones = [w for w in palabras_b
                     if w["text"].strip().upper() in ("SANCTIONS", "CHECK")]
        if no is not None:
            no_izq[bloque] = no["x0"] - 12
            if col_no is None:
                col_no = (no["x0"] - 12, no["x1"] + 12)
        if sanciones:
            izquierda[bloque] = min(w["x0"] for w in sanciones) - 20

    if col_no is None:
        raise PdfSanctionsError("no se encontró la columna 'No'.")
    if not izquierda:
        raise PdfSanctionsError("no se encontraron las columnas 'Sanctions Check'.")

    estado = {}
    for pos, bloque in enumerate(orden):
        if bloque == "COMBINE" or bloque not in izquierda:
            continue
        siguiente = orden[pos + 1] if pos + 1 < len(orden) else None
        if siguiente is None:
            derecha = izquierda[bloque] + 120
        elif siguiente == "COMBINE":
            derecha = (combine[0] if combine else izquierda[bloque] + 120) - 1
        else:
            derecha = no_izq.get(siguiente, titulos[siguiente]["centro"] - 60) - 1
        estado[bloque] = (izquierda[bloque], derecha)

    if combine:
        estado["COMBINE"] = combine

    return {"estado": estado, "no": col_no, "fondo": fondo_encabezado}


def leer_estados(ruta):
    """Lee el PDF y devuelve (estados, hay_combine).

    estados: {numero_de_fila: {"ISIN": "PASS", ..., "COMBINE": "PASS"}}
    """
    paginas = _palabras_por_pagina(ruta)
    if not any(paginas):
        raise PdfSanctionsError("el PDF no tiene texto (¿es una imagen?).")

    estados = {}
    hay_combine = False
    geometria = None

    for palabras in paginas:
        titulos = _titulos(palabras)
        if titulos:
            try:
                geometria = _geometria(palabras, titulos)
                fondo = geometria["fondo"]
            except PdfSanctionsError:
                if geometria is None:
                    continue
                fondo = 0
        elif geometria is not None:
            fondo = 0  # página de continuación, sin encabezado
        else:
            continue

        x0, x1 = geometria["no"]
        anclas = [
            (int(w["text"]), _cy(w))
            for w in palabras
            if w["text"].strip().isdigit() and w["top"] >= fondo - 1
            and x0 <= _cx(w) <= x1
        ]
        if not anclas:
            continue

        pasos = [b[1] - a[1] for a, b in zip(anclas, anclas[1:]) if b[1] > a[1]]
        paso = statistics.median(pasos) if pasos else 12
        tolerancia = paso * 0.45

        hay_combine = hay_combine or "COMBINE" in geometria["estado"]

        for numero, cy in anclas:
            if numero in estados:
                continue
            fila = {}
            for bloque, (bx0, bx1) in geometria["estado"].items():
                textos = [
                    w["text"] for w in palabras
                    if w["top"] >= fondo - 1 and bx0 <= _cx(w) <= bx1
                    and abs(_cy(w) - cy) <= tolerancia
                ]
                fila[bloque] = " ".join(textos).strip()
            estados[numero] = fila

    if not estados:
        raise PdfSanctionsError("no se pudieron leer filas de resultados.")

    return estados, hay_combine


def _volcar_diagnostico(ruta, carpeta):
    """Guarda las palabras del PDF con sus coordenadas para poder revisarlas."""
    destino = Path(carpeta) / "sanctions_diagnostico.txt"
    lineas = []
    try:
        for n, palabras in enumerate(_palabras_por_pagina(ruta), start=1):
            lineas.append(f"--- página {n}: {len(palabras)} palabras")
            for w in palabras[:400]:
                lineas.append(
                    f"{w['text']!r} x0={w['x0']:.0f} x1={w['x1']:.0f} "
                    f"top={w['top']:.0f}"
                )
    except Exception as error:
        lineas.append(f"(no se pudo leer el PDF: {error})")
    destino.write_text("\n".join(lineas), encoding="utf-8")
    return destino


# ============================================================
# VERIFICACIÓN
# ============================================================

def _tiene_dato(valor):
    return valor is not None and not str(valor).strip().upper().startswith(_SIN_DATO)


def verificar_filas(estados, hay_combine, filas, duenios, cantidad):
    """Motivos de bloqueo por security: {indice: [motivos]} (lista vacía = pasa)."""
    motivos = {i: [] for i in range(cantidad)}

    for numero, (fila, dueno) in enumerate(zip(filas, duenios), start=1):
        estado = estados.get(numero)
        if estado is None:
            motivos[dueno].append(f"fila {numero}: no aparece en el PDF")
            continue

        for clave in IDENTIFICADORES:
            resultado = estado.get(clave, "").strip().upper()
            valor = fila.get(clave)

            if _tiene_dato(valor):
                if resultado != "PASS":
                    motivos[dueno].append(
                        f"fila {numero} {clave} {valor}: "
                        f"{resultado or 'sin resultado'}"
                    )
            elif resultado and resultado != "PASS":
                motivos[dueno].append(
                    f"fila {numero} {clave} (vacío en el PreStaging): {resultado}"
                )

        if hay_combine:
            combinado = estado.get("COMBINE", "").strip().upper()
            if combinado != "PASS":
                motivos[dueno].append(
                    f"fila {numero} Combine Status: {combinado or 'sin resultado'}"
                )

    return motivos


def _alerta(lineas):
    barra = "!" * 70
    print()
    print(barra)
    for linea in lineas:
        print(f"!!  {linea}")
    print(barra)


def verificar_securities(pdfs, matriz, securities, carpeta):
    """Revisa el PDF y marca `Security.bloqueo` en los que no pasan.

    Si el PDF no se puede leer o interpretar, se bloquean TODOS. Devuelve
    cuántos securities quedaron bloqueados.
    """
    def bloquear_todos(motivo):
        for sec in securities:
            sec.bloqueo = motivo
        _alerta([
            "NO SE PUDO VERIFICAR EL SANCTIONS CHECK",
            motivo,
            "NINGÚN SECURITY SE CREARÁ EN IMPACT.",
        ])
        return len(securities)

    if not pdfs:
        return bloquear_todos("no hay PDF de Sanctions Check que revisar")

    base = extraer_securities(matriz)
    if len(base) != len(securities):
        return bloquear_todos(
            f"el PreStaging tiene {len(base)} securities y la clasificación "
            f"{len(securities)}; no se puede relacionar el PDF"
        )

    filas = construir_filas(base)
    duenios = duenios_de_filas(base)
    motivos = {i: [] for i in range(len(securities))}

    for pdf in pdfs:
        try:
            estados, hay_combine = leer_estados(pdf)
        except Exception as error:
            diag = _volcar_diagnostico(pdf, carpeta)
            return bloquear_todos(
                f"no se pudo interpretar '{Path(pdf).name}': {error} "
                f"(diagnóstico en {diag})"
            )

        if len(pdfs) > 1:
            print(f"Revisando {Path(pdf).name}...")
        for i, lista in verificar_filas(
            estados, hay_combine, filas, duenios, len(securities)
        ).items():
            motivos[i].extend(lista)

    bloqueados = 0
    for i, sec in enumerate(securities):
        if motivos[i]:
            sec.bloqueo = "Sanctions Check NO PASS: " + "; ".join(motivos[i])
            bloqueados += 1

    if bloqueados:
        lineas = [f"SANCTIONS CHECK: {bloqueados} SECURITY(S) SIN PASS - NO SE TOCAN"]
        for i, sec in enumerate(securities):
            if sec.bloqueo and motivos[i]:
                lineas.append(f"- {sec.filas.get(21) or 'columna ' + str(sec.columna)}")
                lineas.extend(f"    {m}" for m in motivos[i])
        _alerta(lineas)
    else:
        print()
        print(f"Sanctions Check: los {len(securities)} securities dicen PASS.")

    return bloqueados

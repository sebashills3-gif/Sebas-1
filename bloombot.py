"""Etiqueta los códigos (CUSIP/ISIN) del Excel de salida y los envía al
archivo BLR00441 - GDS BLOOMBOT (hoja "Request File") para lanzar la macro
desde la hoja "Request Builder".

Requiere Windows con Excel instalado y: pip install pandas openpyxl xlwings
"""

import pandas as pd

NOMBRE_LIBRO = "BLR00441 - GDS BLOOMBOT"
HOJA_REQUEST_FILE = "Request File"
HOJA_REQUEST_BUILDER = "Request Builder"


def _texto(valor):
    """Devuelve el valor como texto limpio, o '' si está vacío."""
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def _columna(df, nombre):
    """Busca una columna sin distinguir mayúsculas ni espacios."""
    for col in df.columns:
        if str(col).strip().lower() == nombre.lower():
            return col
    raise KeyError(f"No se encontró la columna '{nombre}' en el Excel.")


def etiquetar_codigo(codigo):
    """8 o 9 caracteres -> 'X |CUSIP|'; 12 caracteres -> 'X |ISIN|'.

    Cualquier otra longitud devuelve None (no se puede identificar).
    """
    if len(codigo) in (8, 9):
        return f"{codigo} |CUSIP|"
    if len(codigo) == 12:
        return f"{codigo} |ISIN|"
    return None


def obtener_codigos(ruta_excel):
    """Lee el Excel de salida y devuelve (codigos_etiquetados, omitidos).

    Toma la columna Name y, solo si la celda está vacía, la columna Cusip.
    """
    df = pd.read_excel(ruta_excel, dtype=str)
    col_name = _columna(df, "Name")
    col_cusip = _columna(df, "Cusip")

    etiquetados, omitidos = [], []
    for nombre, cusip in zip(df[col_name], df[col_cusip]):
        codigo = _texto(nombre) or _texto(cusip)
        if not codigo:
            continue
        etiquetado = etiquetar_codigo(codigo)
        if etiquetado is None:
            omitidos.append(codigo)
        else:
            etiquetados.append(etiquetado)
    return etiquetados, omitidos


def _buscar_hoja(libro, nombre, indice):
    """Hoja por nombre; si no existe, por posición (1 = primera)."""
    for hoja in libro.sheets:
        if hoja.name.strip().lower() == nombre.lower():
            return hoja
    return libro.sheets[indice - 1]


def _hacer_click_submit(hoja, macro=None):
    """Ejecuta la macro del botón Submit de la hoja.

    Si se da `macro` se ejecuta por nombre; si no, se busca el botón cuyo
    texto contenga 'submit' y se ejecuta la macro que tiene asignada.
    """
    libro = hoja.book
    if macro:
        libro.macro(macro)()
        return

    for forma in hoja.api.Shapes:
        try:
            texto = forma.TextFrame2.TextRange.Text
        except Exception:
            continue
        if "submit" in texto.lower() and forma.OnAction:
            libro.app.api.Run(forma.OnAction.strip("'"))
            return
    raise RuntimeError(
        "No se encontró el botón Submit. Indica el nombre de la macro "
        "con el parámetro `macro`."
    )


def enviar_a_bloombot(codigos, macro=None, limpiar=True):
    """Pega los códigos en Request File!A y pulsa Submit en Request Builder.

    El libro BLR00441 - GDS BLOOMBOT debe estar ya abierto en Excel.
    """
    import xlwings as xw

    libro = None
    for app in xw.apps:
        for candidato in app.books:
            if candidato.name.lower().startswith(NOMBRE_LIBRO.lower()):
                libro = candidato
                break
        if libro:
            break
    if libro is None:
        raise RuntimeError(f"El archivo '{NOMBRE_LIBRO}' no está abierto.")

    request_file = _buscar_hoja(libro, HOJA_REQUEST_FILE, 3)
    request_builder = _buscar_hoja(libro, HOJA_REQUEST_BUILDER, 1)

    if limpiar:
        request_file.range("A:A").clear_contents()

    rango = request_file.range(f"A1:A{len(codigos)}")
    rango.number_format = "@"  # texto, para que Excel no altere los códigos
    rango.value = [[codigo] for codigo in codigos]

    request_builder.activate()
    _hacer_click_submit(request_builder, macro)


def procesar_y_enviar(ruta_excel, macro=None):
    """Flujo completo a partir del Excel que genera la extracción."""
    codigos, omitidos = obtener_codigos(ruta_excel)

    if omitidos:
        print(f"Códigos sin 8, 9 o 12 caracteres (no enviados): {omitidos}")
    if not codigos:
        print("No hay códigos válidos que enviar.")
        return

    enviar_a_bloombot(codigos, macro=macro)
    print(f"{len(codigos)} códigos enviados a {NOMBRE_LIBRO}.")

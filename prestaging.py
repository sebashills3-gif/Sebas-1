"""Después de correr la macro de BLOOMBOT:

1. Espera los archivos GIMP*.xls y GIMP*.xps que se generan en la carpeta
   DUMP y los mueve a la carpeta de evidencia.
2. Lee del XLS los identificadores de cada security.
3. Los pega en el libro abierto "PreStaging File" y lo guarda.

Requiere Windows con Excel y: pip install xlwings
"""

import re
import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path

CARPETA_DUMP = Path(r"C:\Users\45502358\Desktop\DUMP")

# Fila del XLS (1 = primera) donde está cada identificador. La columna A
# trae los mnemonics y de la B en adelante hay un security por columna.
FILAS_XLS = {
    "ISIN": 21,
    "CUSIP": 22,
    "COMMON": 23,
    "FIGI_ID": 91,
}
# El SEDOL puede aparecer en cualquiera de estas filas (puede haber varios).
FILAS_SEDOL = range(82, 87)
# Valores que Bloomberg devuelve cuando no hay dato; no cuentan como SEDOL.
SIN_DATO = ("#N/A", "N.A.", "N/A", "#NAME", "#VALUE", "#REF")
ULTIMA_FILA_XLS = 102

# Columna del PreStaging File donde se pega cada identificador.
COLUMNAS_PRESTAGING = {
    "ISIN": "C",
    "SEDOL": "G",
    "CUSIP": "K",
    "COMMON": "O",
    "FIGI_ID": "S",
}
FILA_INICIAL_PRESTAGING = 5
NOMBRE_PRESTAGING = "PreStaging File"


# ============================================================
# 1. ARCHIVOS DE DUMP
# ============================================================

def buscar_archivos(cantidad, desde, dump=CARPETA_DUMP):
    """Busca GIMPddmmyyhhmmss_<cantidad>.xls/.xps creados desde `desde`.

    Devuelve {extensión: ruta} del grupo más reciente que tenga XLS.
    """
    patron = re.compile(rf"^GIMP(\d{{12}})_{cantidad}$", re.IGNORECASE)
    grupos = {}

    if not Path(dump).exists():
        return {}

    for ruta in Path(dump).iterdir():
        ext = ruta.suffix.lower()
        coincide = patron.match(ruta.stem)
        if ext not in (".xls", ".xps") or not coincide:
            continue
        try:
            marca = datetime.strptime(coincide.group(1), "%d%m%y%H%M%S")
        except ValueError:
            continue
        if marca >= desde - timedelta(seconds=2):
            grupos.setdefault((marca, ruta.stem), {})[ext] = ruta

    con_xls = [clave for clave, grupo in grupos.items() if ".xls" in grupo]
    if not con_xls:
        return {}
    return grupos[max(con_xls)]


def _tamanos(archivos):
    return {ext: ruta.stat().st_size for ext, ruta in archivos.items()}


def esperar_archivos(cantidad, desde, dump=CARPETA_DUMP,
                     timeout_s=300, gracia_xps_s=30, pausa_s=2):
    """Espera a que aparezcan (y dejen de crecer) el XLS y el XPS.

    Si el XLS aparece pero el XPS no llega tras `gracia_xps_s`, sigue solo
    con el XLS. Devuelve (xls, xps_o_None).
    """
    inicio = time.time()
    visto_xls = None
    tamanos_previos = None

    while time.time() - inicio < timeout_s:
        archivos = buscar_archivos(cantidad, desde, dump)

        if ".xls" in archivos:
            if visto_xls is None:
                visto_xls = time.time()

            tamanos = _tamanos(archivos)
            estables = tamanos == tamanos_previos
            tamanos_previos = tamanos

            if ".xps" in archivos and estables:
                return archivos[".xls"], archivos[".xps"]

            if ".xps" not in archivos and estables \
                    and time.time() - visto_xls > gracia_xps_s:
                print("Aviso: no apareció el archivo XPS.")
                return archivos[".xls"], None

        time.sleep(pausa_s)

    raise TimeoutError(
        f"No aparecieron los archivos GIMP*_{cantidad} en {dump} "
        f"tras {timeout_s} segundos."
    )


def mover_archivos(rutas, destino, intentos=10, pausa_s=2):
    """Mueve los archivos a `destino`, reintentando si siguen en uso."""
    destino = Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    nuevas = []

    for ruta in rutas:
        if ruta is None:
            continue
        for intento in range(1, intentos + 1):
            try:
                nuevas.append(Path(shutil.move(str(ruta), str(destino / ruta.name))))
                break
            except OSError:
                if intento == intentos:
                    raise
                time.sleep(pausa_s)

    return nuevas


# ============================================================
# 2. LECTURA DEL XLS
# ============================================================

def _a_texto(valor):
    """Valor de celda como texto limpio, o None si está vacío."""
    if valor is None:
        return None
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).strip()
    return texto or None


def _es_sedol(valor):
    return valor is not None and not valor.upper().startswith(SIN_DATO)


def extraer_securities(matriz):
    """De la matriz de celdas (fila 1 = matriz[0]) saca un dict por security.

    Cada columna desde la B es un security. "SEDOL" es una lista con los
    SEDOL distintos de las filas 82 a 86, en el orden en que aparecen. Las
    columnas vacías al final se descartan; las vacías intermedias se
    conservan para no desalinear filas.
    """
    columnas = len(matriz[0]) if matriz else 0
    securities = []

    for col in range(1, columnas):
        security = {
            clave: _a_texto(matriz[fila - 1][col])
            for clave, fila in FILAS_XLS.items()
        }
        sedols = [_a_texto(matriz[fila - 1][col]) for fila in FILAS_SEDOL]
        security["SEDOL"] = list(dict.fromkeys(v for v in sedols if _es_sedol(v)))
        securities.append(security)

    while securities and not any(securities[-1].values()):
        securities.pop()

    return securities


def construir_filas(securities):
    """Convierte los securities en las filas del PreStaging File.

    Primera fila de cada security: todos los identificadores y el primer
    SEDOL. Por cada SEDOL adicional, una fila con el mismo ISIN, ese SEDOL
    y el resto vacío.
    """
    filas = []

    for security in securities:
        sedols = security["SEDOL"]
        filas.append({**security, "SEDOL": sedols[0] if sedols else None})

        for sedol in sedols[1:]:
            filas.append({
                "ISIN": security["ISIN"],
                "SEDOL": sedol,
                "CUSIP": None,
                "COMMON": None,
                "FIGI_ID": None,
            })

    return filas


def leer_matriz_xls(ruta):
    """Abre el XLS en una instancia oculta de Excel y devuelve sus celdas
    (filas 1 a ULTIMA_FILA_XLS; la fila 1 es matriz[0])."""
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        libro = app.books.open(str(ruta), read_only=True, update_links=False)
        hoja = libro.sheets[0]
        ultima_col = max(hoja.used_range.last_cell.column, 2)
        matriz = hoja.range((1, 1), (ULTIMA_FILA_XLS, ultima_col)) \
            .options(ndim=2).value
        libro.close()
    finally:
        app.quit()

    print("Mnemonics leídos del XLS:")
    for clave, fila in FILAS_XLS.items():
        print(f"  {clave} (fila {fila}): {matriz[fila - 1][0]}")
    for fila in FILAS_SEDOL:
        print(f"  SEDOL (fila {fila}): {matriz[fila - 1][0]}")

    return matriz


# ============================================================
# 3. PRESTAGING FILE
# ============================================================

def _para_excel(valor):
    """None limpia la celda; los códigos solo con dígitos se guardan como
    texto (apóstrofo) para no perder ceros a la izquierda."""
    if valor is None:
        return None
    return "'" + valor if valor.isdigit() else valor


def pegar_en_prestaging(securities, hoja=None):
    """Pega los identificadores desde la fila 5 y guarda el libro.

    Un security con varios SEDOL ocupa varias filas (ver construir_filas).

    Sobrescribe las celdas destino (si el valor viene vacío, la celda queda
    vacía) y borra el contenido que hubiera más abajo en esas mismas
    columnas, de corridas anteriores. `hoja` es el nombre de la hoja; por
    defecto, la hoja activa.
    """
    import xlwings as xw

    libro = None
    for app in xw.apps:
        for candidato in app.books:
            if NOMBRE_PRESTAGING.lower() in candidato.name.lower():
                libro = candidato
                break
        if libro:
            break
    if libro is None:
        raise RuntimeError(f"El archivo '{NOMBRE_PRESTAGING}' no está abierto.")

    destino = libro.sheets[hoja] if hoja else libro.sheets.active

    filas = construir_filas(securities)

    for i, valores in enumerate(filas):
        fila = FILA_INICIAL_PRESTAGING + i
        for clave, columna in COLUMNAS_PRESTAGING.items():
            destino.range(f"{columna}{fila}").value = _para_excel(valores[clave])

    # Limpia los restos de corridas anteriores (solo las 5 columnas de
    # identificadores, y solo el contenido, no el formato).
    primera_sobrante = FILA_INICIAL_PRESTAGING + len(filas)
    ultima_fila = destino.used_range.last_cell.row
    if ultima_fila >= primera_sobrante:
        for columna in COLUMNAS_PRESTAGING.values():
            destino.range(
                f"{columna}{primera_sobrante}:{columna}{ultima_fila}"
            ).clear_contents()

    libro.save()
    print(
        f"{len(securities)} securities ({len(filas)} filas) pegados en "
        f"'{libro.name}' (hoja '{destino.name}') y archivo guardado."
    )


# ============================================================
# FLUJO COMPLETO
# ============================================================

def procesar_gimp(cantidad, desde, carpeta_evidencia, dump=CARPETA_DUMP,
                  hoja=None):
    """Mueve los GIMP de DUMP a la evidencia y llena el PreStaging File.

    Devuelve las celdas del XLS (para clasificar los securities), o None si
    el XLS no trajo securities.
    """
    print()
    print(f"Esperando archivos GIMP*_{cantidad} en {dump}...", flush=True)
    xls, xps = esperar_archivos(cantidad, desde, dump)

    movidos = mover_archivos([xls, xps], carpeta_evidencia)
    for ruta in movidos:
        print(f"Archivo movido: {ruta}")

    xls_movido = next(r for r in movidos if r.suffix.lower() == ".xls")
    matriz = leer_matriz_xls(xls_movido)
    securities = extraer_securities(matriz)

    if len(securities) != cantidad:
        print(
            f"Aviso: se enviaron {cantidad} securities a BLOOMBOT pero el "
            f"XLS trae {len(securities)} columnas con datos."
        )
    if not securities:
        print("El XLS no trae securities; no se pega nada.")
        return None

    pegar_en_prestaging(securities, hoja=hoja)
    return matriz

"""Clasifica cada security del XLS de BLOOMBOT en un skeleton (plantilla) y
limpia sus datos antes de darlos de alta en IMPACT.

Estructura (para agregar skeletons nuevos solo se toca la sección SKELETONS):

* TRANSFORMACIONES GENERALES: aplican a todos los skeletons.
* Cada skeleton define su FILTRO (fila del XLS -> valor esperado) y sus
  TRANSFORMACIONES ESPECÍFICAS.
* El formato final de celdas (fechas sin diagonales y decimales limpios) se
  aplica siempre al final, para que las transformaciones anteriores trabajen
  con los valores originales.

Las filas son las del XLS de BLOOMBOT (1 = primera); la columna A trae el
mnemonic y de la B en adelante hay un security por columna.
"""

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

# ============================================================
# FILAS DEL XLS
# ============================================================

IS_UNIT_TRADED = 4
CPN_TYP = 6
CRNCY = 7
MARKET_SECTOR_DES = 12
SECURITY_DES = 15
SECURITY_TYP = 8
COLLAT_TYP = 9
TYPE_OF_BOND = 10
MATURITY_FILA = 19
DAY_COUNT_FILA = 20
ID_ISIN = 21
ID_CUSIP = 22
FIRST_CPN_DT = 27
IS_PERPETUAL = 52
SINKABLE = 56

# Mnemonic esperado en la columna A de cada fila usada (para avisar si el
# layout del XLS cambia).
MNEMONICS_ESPERADOS = {
    IS_UNIT_TRADED: "IS_UNIT_TRADED",
    CPN_TYP: "CPN_TYP",
    CRNCY: "CRNCY",
    MARKET_SECTOR_DES: "MARKET_SECTOR_DES",
    SECURITY_TYP: "SECURITY_TYP",
    COLLAT_TYP: "COLLAT_TYP",
    TYPE_OF_BOND: "TYPE_OF_BOND",
    SECURITY_DES: "SECURITY_DES",
    MATURITY_FILA: "MATURITY",
    DAY_COUNT_FILA: "DAY_CNT_DES",
    ID_ISIN: "ID_ISIN",
    ID_CUSIP: "ID_CUSIP",
    FIRST_CPN_DT: "FIRST_CPN_DT",
    IS_PERPETUAL: "IS_PERPETUAL",
    SINKABLE: "SINKABLE",
}

# Filas de identificadores: nunca se les "limpian decimales" (un CUSIP o un
# SEDOL numérico no debe volverse 12345.0).
FILAS_IDENTIFICADORES = {21, 22, 23, 82, 83, 84, 85, 86, 91}

ANIOS_PERPETUO = 30


# ============================================================
# MODELO
# ============================================================

@dataclass
class Security:
    columna: int                     # 2 = columna B
    filas: dict                      # fila del XLS -> valor
    mnemonics: dict                  # fila del XLS -> mnemonic (columna A)
    skeleton: str = None
    entidades: set = field(default_factory=set)   # columna Entity del TUR
    bloqueo: str = None              # motivo por el que NO se debe crear
    derivados: dict = field(default_factory=dict)
    advertencias: list = field(default_factory=list)

    def texto(self, fila):
        return _texto(self.filas.get(fila))


# ============================================================
# AYUDAS
# ============================================================

def _texto(valor):
    """Valor como texto limpio ('' si está vacío o es fecha/otro objeto)."""
    if valor is None or isinstance(valor, (date, datetime)):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def _letra(columna):
    letras = ""
    while columna:
        columna, resto = divmod(columna - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


_FECHA_SOLA = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$")


def a_fecha(valor):
    """date si el valor ES una fecha sola (objeto fecha o 'MM/DD/YYYY');
    None si no (por ejemplo 'BXT 12/05/2030' no cuenta)."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, str):
        coincide = _FECHA_SOLA.match(valor)
        if coincide:
            mes, dia, anio = (int(g) for g in coincide.groups())
            try:
                return date(anio, mes, dia)
            except ValueError:
                return None
    return None


def formato_fecha_sin_diagonales(fecha):
    """12/15/2030 -> 12152030."""
    return fecha.strftime("%m%d%Y")


def formato_decimal(valor):
    """Quita ceros sobrantes: 5.87000 -> 5.87, 5 -> 5.0, .78 -> 0.78."""
    decimal = Decimal(str(valor).strip())
    texto = format(decimal.normalize(), "f")
    if "." not in texto:
        texto += ".0"
    return texto


_NUMERO_CON_PUNTO = re.compile(r"^[+-]?(\d+\.\d*|\.\d+)$")


def _numero_a_limpiar(valor):
    """True si el valor es un número (celda numérica o texto con punto
    decimal). Los textos de solo dígitos no se tocan: pueden ser códigos."""
    if isinstance(valor, bool):
        return False
    if isinstance(valor, (int, float)):
        return True
    return isinstance(valor, str) and bool(_NUMERO_CON_PUNTO.match(valor.strip()))


def sumar_anios(fecha, anios):
    try:
        return fecha.replace(year=fecha.year + anios)
    except ValueError:  # 29 de febrero
        return fecha.replace(year=fecha.year + anios, day=28)


# ============================================================
# TRANSFORMACIONES GENERALES (todos los skeletons)
# ============================================================

def derivar_cusip_desde_isin(sec):
    """CRNCY = USD e ISIN que empieza por US: el CUSIP es el ISIN sin los 2
    primeros caracteres ni el último (US1234567890 -> 123456789)."""
    isin = sec.texto(ID_ISIN)

    if sec.texto(CRNCY).upper() == "USD" and isin.upper().startswith("US"):
        if len(isin) != 12:
            sec.advertencias.append(
                f"ISIN '{isin}' no tiene 12 caracteres; no se derivó el CUSIP."
            )
            return
        sec.derivados["CUSIP_ORIGINAL"] = sec.texto(ID_CUSIP)
        sec.filas[ID_CUSIP] = isin[2:-1]


_CODIGOS_DAY_COUNT = {
    "ACT/ACT": ("A", "A"),
    "ACT/360": ("A", "360"),
    "ACT/365": ("A", "365"),
    "30/360": ("30", "360"),
    "30/365": ("30", "365"),
    "ISMA-30/360": ("3E", "360"),
    "ISMA-30/365": ("3E", "365"),
}

# ISMA-30/36x va primero y los lookarounds evitan que 30/360 se tome de
# dentro de ISMA-30/360.
_DAY_COUNT_RE = re.compile(
    r"(?<![\w/-])(ISMA-30/36[05]|30/36[05]|ACT/ACT|ACT/36[05])(?![\w/])",
    re.IGNORECASE,
)


def separar_day_count(sec):
    """Fila 20 (sola o dentro de otro texto, p. ej. 'BXT 4V ACT/ACT 57') ->
    DAY COUNT y YEAR BASIS."""
    texto = sec.texto(DAY_COUNT_FILA)
    if not texto:
        return

    codigos = {m.upper() for m in _DAY_COUNT_RE.findall(texto)}

    if not codigos:
        sec.advertencias.append(
            f"Fila {DAY_COUNT_FILA}: no se reconoció el day count en '{texto}'."
        )
        return
    if len(codigos) > 1:
        sec.advertencias.append(
            f"Fila {DAY_COUNT_FILA}: varios day count en '{texto}' "
            f"({sorted(codigos)}); se usa el primero."
        )

    primero = _DAY_COUNT_RE.search(texto).group(1).upper()
    sec.derivados["DAY COUNT"], sec.derivados["YEAR BASIS"] = \
        _CODIGOS_DAY_COUNT[primero]


TRANSFORMACIONES_GENERALES = [
    derivar_cusip_desde_isin,
    separar_day_count,
]


def formatear_celdas(sec):
    """Siempre al final: fechas solas -> MMDDYYYY y decimales sin ceros."""
    for fila, valor in list(sec.filas.items()):
        fecha = a_fecha(valor)

        if fecha:
            sec.filas[fila] = formato_fecha_sin_diagonales(fecha)
        elif fila not in FILAS_IDENTIFICADORES and _numero_a_limpiar(valor):
            try:
                sec.filas[fila] = formato_decimal(valor)
            except InvalidOperation:
                pass
        elif valor is not None:
            sec.filas[fila] = _texto(valor)


# ============================================================
# SKELETON 5/80
# ============================================================

_FRACCION_MIXTA = re.compile(r"(?<![\d/.])(\d+)\s+(\d+)/(\d+)(?![\d/])")
_FRACCION_SOLA = re.compile(r"(?<![\d/.])(\d+)/(\d+)(?![\d/])")


def _decimal_de_fraccion(entero, numerador, denominador):
    return Decimal(entero) + Decimal(numerador) / Decimal(denominador)


def cupon_fraccion_a_decimal(sec):
    """SECURITY_DES 'ISSUER 5 1/2 12/15/2030' -> 'ISSUER 5.5 12/15/2030'.
    Si el cupón ya viene en decimal no se toca."""
    descripcion = sec.texto(SECURITY_DES)

    mixta = _FRACCION_MIXTA.search(descripcion)
    sola = None if mixta else _FRACCION_SOLA.search(descripcion)
    coincide = mixta or sola
    if not coincide:
        return

    grupos = coincide.groups()
    entero, numerador, denominador = (
        grupos if mixta else ("0",) + grupos
    )
    if int(denominador) == 0:
        return

    cupon = formato_decimal(
        _decimal_de_fraccion(entero, numerador, denominador)
    )
    sec.filas[SECURITY_DES] = (
        descripcion[:coincide.start()] + cupon + descripcion[coincide.end():]
    )


def perpetuo_a_vencimiento(sec):
    """IS_PERPETUAL = Y: MATURITY (fila 19) = FIRST_CPN_DT + 30 años, y si
    SECURITY_DES termina en PERP se cambia por la fecha en MM/DD/YYYY."""
    if sec.texto(IS_PERPETUAL).upper() != "Y":
        return

    primer_cupon = a_fecha(sec.filas.get(FIRST_CPN_DT))
    if primer_cupon is None:
        sec.advertencias.append(
            "IS_PERPETUAL = Y pero FIRST_CPN_DT no es una fecha válida; "
            "no se calculó MATURITY."
        )
        return

    vencimiento = sumar_anios(primer_cupon, ANIOS_PERPETUO)
    sec.filas[MATURITY_FILA] = vencimiento  # se formatea al final

    descripcion = sec.texto(SECURITY_DES)
    if re.search(r"(?i)(?<!\S)PERP$", descripcion):
        sec.filas[SECURITY_DES] = re.sub(
            r"(?i)(?<!\S)PERP$", vencimiento.strftime("%m/%d/%Y"), descripcion
        )


@dataclass
class Skeleton:
    nombre: str
    filtro: dict                     # fila del XLS -> valor esperado
    transformaciones: list
    # fila -> valor (o lista de valores) que NO debe tener
    excluye: dict = field(default_factory=dict)
    # Combinaciones prohibidas: cada dict (fila -> valor) excluye el security
    # solo si cumple TODAS sus condiciones a la vez.
    excluye_combinaciones: list = field(default_factory=list)

    def coincide(self, sec):
        cumple = all(
            sec.texto(fila).upper() == esperado.upper()
            for fila, esperado in self.filtro.items()
        )
        no_excluido = all(
            sec.texto(fila).upper() not in {
                v.upper() for v in ([prohibidos] if isinstance(prohibidos, str)
                                    else prohibidos)
            }
            for fila, prohibidos in self.excluye.items()
        )
        sin_combinacion = not any(
            all(
                sec.texto(fila).upper() == valor.upper()
                for fila, valor in combinacion.items()
            )
            for combinacion in self.excluye_combinaciones
        )
        return cumple and no_excluido and sin_combinacion


SKELETONS = [
    Skeleton(
        nombre="5/80",
        filtro={
            IS_UNIT_TRADED: "N",
            CPN_TYP: "FIXED",
            CRNCY: "USD",
            MARKET_SECTOR_DES: "Corp",
            SINKABLE: "N",
        },
        # Con cualquiera de estos valores (aunque venga solo) no es 5/80. Un
        # valor vacío sí pasa.
        excluye={
            SECURITY_TYP: "US DOMESTIC",
            COLLAT_TYP: ["CERT OF DEPOSIT", "COVERED"],
            TYPE_OF_BOND: ["Convertible", "Covered"],
        },
        transformaciones=[
            cupon_fraccion_a_decimal,
            perpetuo_a_vencimiento,
        ],
    ),
]


# ============================================================
# FLUJO
# ============================================================

def clasificar(sec):
    """Primer skeleton cuyo filtro cumple, o None."""
    for skeleton in SKELETONS:
        if skeleton.coincide(sec):
            return skeleton
    return None


def limpiar(sec):
    """Clasifica el security y le aplica sus transformaciones."""
    skeleton = clasificar(sec)
    sec.skeleton = skeleton.nombre if skeleton else None

    for transformacion in TRANSFORMACIONES_GENERALES:
        transformacion(sec)

    if skeleton:
        for transformacion in skeleton.transformaciones:
            transformacion(sec)

    formatear_celdas(sec)
    return sec


def verificar_mnemonics(matriz):
    """Avisos si la columna A no trae los mnemonics que se esperan."""
    avisos = []
    for fila, esperado in MNEMONICS_ESPERADOS.items():
        real = _texto(matriz[fila - 1][0]) if len(matriz) >= fila else ""
        if real.upper() != esperado:
            avisos.append(
                f"Fila {fila}: se esperaba '{esperado}' y el XLS trae '{real}'."
            )
    return avisos


def construir_securities(matriz):
    """Un Security por columna (desde la B); descarta columnas vacías al
    final."""
    columnas = len(matriz[0]) if matriz else 0
    mnemonics = {i + 1: _texto(fila[0]) for i, fila in enumerate(matriz)}
    securities = []

    for col in range(1, columnas):
        valores = {i + 1: fila[col] for i, fila in enumerate(matriz)}
        securities.append(
            Security(columna=col + 1, filas=valores, mnemonics=mnemonics)
        )

    while securities and all(
        v is None for v in securities[-1].filas.values()
    ):
        securities.pop()

    return securities


def exportar_excel(securities, ruta):
    """Un renglón por security con el resultado de la limpieza."""
    import pandas as pd

    registros = []
    for sec in securities:
        registro = {
            "Columna XLS": _letra(sec.columna),
            "Skeleton": sec.skeleton or "SIN SKELETON",
            "Entidad": ", ".join(sorted(sec.entidades)),
            "ISIN": sec.filas.get(ID_ISIN),
            "CUSIP": sec.filas.get(ID_CUSIP),
            "DAY COUNT": sec.derivados.get("DAY COUNT"),
            "YEAR BASIS": sec.derivados.get("YEAR BASIS"),
            "MATURITY": sec.filas.get(MATURITY_FILA),
            "Advertencias": " | ".join(sec.advertencias),
        }
        for fila, valor in sec.filas.items():
            registro[f"{fila:03d} {sec.mnemonics.get(fila, '')}"] = valor
        registros.append(registro)

    pd.DataFrame(registros).to_excel(ruta, index=False)


def procesar_matriz(matriz, carpeta_destino, entidades=None):
    """Clasifica y limpia todos los securities del XLS y guarda un Excel de
    revisión en `carpeta_destino`. Devuelve la lista de Security.

    `entidades` (opcional) es la lista de entidades de cada security en el
    orden de columnas del XLS; se guarda en `Security.entidades`.
    """
    print()
    print("Clasificando securities por skeleton...")

    for aviso in verificar_mnemonics(matriz):
        print(f"Aviso: {aviso}")

    securities = [limpiar(sec) for sec in construir_securities(matriz)]

    if entidades is not None:
        if len(entidades) != len(securities):
            print(
                f"Aviso: {len(entidades)} códigos enviados a BLOOMBOT pero el "
                f"XLS trae {len(securities)} securities; no se asignan entidades."
            )
        else:
            for sec, entidad in zip(securities, entidades):
                sec.entidades = set(entidad)

    for sec in securities:
        print(
            f"  Columna {_letra(sec.columna)}: "
            f"{sec.filas.get(ID_ISIN) or '(sin ISIN)'} -> "
            f"{sec.skeleton or 'SIN SKELETON'}"
            + (f" [{', '.join(sorted(sec.entidades))}]" if sec.entidades else "")
        )
        for advertencia in sec.advertencias:
            print(f"      Aviso: {advertencia}")

    ruta = Path(carpeta_destino) / "Securities_clasificados.xlsx"
    exportar_excel(securities, ruta)
    print(f"Excel de revisión: {ruta}")

    return securities

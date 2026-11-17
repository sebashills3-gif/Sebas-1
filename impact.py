"""Alta de securities del skeleton 5/80 en IMPACT (pantallas NSEX02 a NSEX05).

Para cada security: se navega con changeProgram, se elige mercado 5 e
instrumento 80, se llenan las pantallas con los datos ya limpios (ver
skeletons.py), se resuelven los warnings y errores conocidos y se deja la
pantalla NSEX05 lista. El guardado lo hace la persona a mano.
"""

import calendar
import re
import time
from dataclasses import dataclass
from datetime import date, datetime

PROGRAMA = ("NSEX", "HBSI", "")
MERCADO = "5"
INSTRUMENTO = "80"
SKELETON = "5/80"

PAUSA_TRAS_ENTER_S = 1.5
MAX_ENTER_POR_PANTALLA = 12
MAX_REINTENTOS_POR_ERROR = 2

# Campo que identifica cada pantalla.
HUELLAS_PANTALLA = (
    ("NSEX05", "ici_trailerCode"),
    ("NSEX03", "ici_informationCommonStrikePrice"),
    ("NSEX02", "ici_securityName"),
    ("SELECCION", "ici_selectMarket"),
)

FRECUENCIAS = {1: "A", 2: "S", 4: "Q", 12: "M"}
MESES_POR_FRECUENCIA = {"A": 12, "S": 6, "Q": 3, "M": 1}

# Alias de mnemonic en la columna A del XLS.
ALIAS = {
    "NAME": ("NAME",),
    "ID_BB": ("ID_BB", "ID_BB_UNIQUE"),
    "ID_COMMON": ("ID_COMMON", "COMMON"),
}


class ImpactError(Exception):
    pass


# ============================================================
# DATOS DEL SECURITY
# ============================================================

def _es_na(valor):
    """Valor que Bloomberg devuelve cuando no hay dato."""
    if valor is None:
        return True
    return str(valor).strip().upper().startswith(("#N/A", "N.A.", "N/A"))


@dataclass
class Datos:
    """Valores ya limpios de un security, tal como se capturan en IMPACT."""
    security_des: str = None
    name: str = None
    isin: str = None
    cusip: str = None
    cpn: str = None
    maturity: str = None
    day_count: str = None
    year_basis: str = None
    commodity: str = None
    nickname: str = None
    series: str = None
    country: str = None
    int_acc: str = None
    issue_dt: str = None
    first_cpn: str = None
    call_dt: str = None
    call_px: str = None
    frecuencia: str = None
    crncy: str = None
    avisos: list = None


def construir_datos(sec):
    """Datos de IMPACT a partir de un Security ya limpiado (skeletons.py)."""
    fila_de = {}
    for fila, mnemonic in sec.mnemonics.items():
        fila_de.setdefault(mnemonic.strip().upper(), fila)

    avisos = []

    def valor(*nombres, conservar_na=False):
        for nombre in nombres:
            fila = fila_de.get(nombre)
            if fila is not None:
                v = sec.filas.get(fila)
                if v is None:
                    return None
                v = str(v).strip()
                if not v or (_es_na(v) and not conservar_na):
                    return None
                return v
        avisos.append(f"No se encontró el mnemonic {nombres[0]} en el XLS.")
        return None

    security_des = valor("SECURITY_DES")
    cusip = valor("ID_CUSIP")
    cusip_original = sec.derivados.get("CUSIP_ORIGINAL") or cusip

    bb = valor(*ALIAS["ID_BB"])
    common = valor(*ALIAS["ID_COMMON"])
    # ID_COMMON solo cuando ID_BB y el ID_CUSIP del XLS son iguales.
    nickname = common if (bb is not None and bb == cusip_original) else bb

    series = (valor("SERIES") or "").upper()

    frecuencia = None
    cpn_freq = valor("CPN_FREQ")
    if cpn_freq is not None:
        try:
            frecuencia = FRECUENCIAS.get(int(float(cpn_freq)))
        except ValueError:
            pass
        if frecuencia is None:
            avisos.append(f"CPN_FREQ '{cpn_freq}' no es 1, 2, 4 ni 12.")

    return Datos(
        security_des=security_des,
        name=valor(*ALIAS["NAME"]),
        isin=valor("ID_ISIN"),
        cusip=cusip,
        cpn=valor("CPN", conservar_na=True),
        maturity=valor("MATURITY"),
        day_count=sec.derivados.get("DAY COUNT"),
        year_basis=sec.derivados.get("YEAR BASIS"),
        commodity=(security_des or "")[:4] or None,
        nickname=nickname,
        series=series,
        country=valor("CNTRY_OF_INCORPORATION"),
        int_acc=valor("INT_ACC_DT"),
        issue_dt=valor("ISSUE_DT"),
        first_cpn=valor("FIRST_CPN_DT"),
        call_dt=valor("NXT_CALL_DT"),
        call_px=valor("NXT_CALL_PX"),
        frecuencia=frecuencia,
        crncy=(valor("CRNCY") or "").upper(),
        avisos=avisos,
    )


def campos_nsex02(d):
    return {
        "ici_securityName": d.security_des,
        "ici_description1": d.name,
        "ici_description4": d.isin,
        "ici_iisinCode": d.isin,
        "ici_customerIPNumber": d.cusip,
        "ici_sortCouponStrikePrice": d.cpn,
        "ici_sortMaturity": d.maturity,
        "ici_dayCount": d.day_count,
        "ici_yearBasis": d.year_basis,
        "ici_sortIssueCommodityCodePart1": d.commodity,
        "ici_nickname": d.nickname,
        "ici_description2": d.series if d.series in ("144A", "REGS") else None,
    }


def campos_nsex03(d):
    return {
        "ici_informationCommonStrikePrice": d.cpn,
        "ici_informationMaturity": d.maturity,
        "ici_countryofOrigin": d.country,
        "ici_datedDate": d.int_acc,
        "ici_issueDate": d.issue_dt,
        "ici_firstCouponDate": d.first_cpn,
        "ici_callDate": d.call_dt,
        "ici_callPrice": d.call_px,
        "ici_informationIssueCommodityCode": d.commodity,
        "ici_elig144a": "Y" if d.series == "144A" else None,
        "ici_RegS": "Y" if d.series == "REGS" else None,
        "ici_paymentFrequencyCode": d.frecuencia,
    }


def campos_nsex05(d):
    """Solo los campos que aplican; los demás no se tocan."""
    campos = {}
    if "HSBC" in (d.security_des or "").upper():
        campos["ici_trailerCode"] = "REP"
    if d.crncy == "USD" and (d.isin or "").upper().startswith("XS"):
        campos["ici_depositoryAccountPointer"] = "FORG"
    return campos


# ============================================================
# ACCESO A LA PÁGINA
# ============================================================

def _selector(nombre):
    return f'[name="{nombre}"], [id="{nombre}"]'


def _buscar_campo(page, nombre):
    """Campo por name o id, en la página y en todos sus iframes."""
    for frame in page.frames:
        try:
            localizador = frame.locator(_selector(nombre))
            if localizador.count():
                return localizador.first
        except Exception:
            continue
    return None


def _esperar_campo(page, nombre, timeout_s=15):
    inicio = time.time()
    while True:
        campo = _buscar_campo(page, nombre)
        if campo is not None:
            return campo
        if time.time() - inicio >= timeout_s:
            return None
        time.sleep(0.5)


def poner_valor(page, nombre, valor):
    """Escribe el valor en el campo (None lo deja vacío)."""
    campo = _esperar_campo(page, nombre)
    if campo is None:
        raise ImpactError(f"No se encontró el campo {nombre} en la página.")

    texto = "" if valor is None else str(valor)
    if campo.evaluate("e => e.tagName") == "SELECT":
        campo.select_option(texto)
    else:
        campo.fill(texto)


def llenar(page, campos):
    for nombre, valor in campos.items():
        poner_valor(page, nombre, valor)


def _esperar_pagina_lista(page, timeout_ms=15000):
    for estado in ("domcontentloaded", "load"):
        try:
            page.wait_for_load_state(estado, timeout=timeout_ms)
        except Exception:
            pass


def enter(page, campo_ancla=None):
    """Da Enter (desde un campo de la pantalla, si se indica) y espera."""
    campo = _buscar_campo(page, campo_ancla) if campo_ancla else None
    if campo is not None:
        campo.press("Enter")
    else:
        page.keyboard.press("Enter")
    time.sleep(PAUSA_TRAS_ENTER_S)
    _esperar_pagina_lista(page)


def leer_texto(page):
    partes = []
    for frame in page.frames:
        try:
            partes.append(frame.locator("body").inner_text(timeout=5000))
        except Exception:
            continue
    return "\n".join(partes)


def pantalla_actual(page):
    """NSEX02, NSEX03, NSEX05, SELECCION; NSEX04 solo si su nombre aparece
    en el texto; None si no se reconoce."""
    for nombre, campo in HUELLAS_PANTALLA:
        if _buscar_campo(page, campo) is not None:
            return nombre

    coincide = re.search(r"NSEX0[1-5]", leer_texto(page).upper())
    return coincide.group(0) if coincide else None


def esperar_pantalla(page, nombre, timeout_s=30):
    inicio = time.time()
    while time.time() - inicio < timeout_s:
        if pantalla_actual(page) == nombre:
            return
        time.sleep(0.5)
    raise ImpactError(
        f"No se llegó a la pantalla {nombre} "
        f"(pantalla actual: {pantalla_actual(page)})."
    )


# ============================================================
# NAVEGACIÓN
# ============================================================

def ir_a_seleccion(page):
    """Ejecuta changeProgram en el frame que lo define y espera la pantalla
    de selección de mercado e instrumento."""
    for frame in page.frames:
        try:
            definido = frame.evaluate("typeof changeProgram") == "function"
        except Exception:
            continue
        if definido:
            try:
                frame.evaluate(
                    "([a, b, c]) => changeProgram(a, b, c)", list(PROGRAMA)
                )
            except Exception as error:  # la página puede navegar al llamarlo
                if "context was destroyed" not in str(error).lower() \
                        and "navigation" not in str(error).lower():
                    raise
            break
    else:
        raise ImpactError("No se encontró changeProgram en la página de IMPACT.")

    if _esperar_campo(page, "ici_selectMarket", timeout_s=30) is None:
        raise ImpactError("No apareció la pantalla de selección de mercado.")


def _anio_cuatro_digitos(texto):
    """ISSUER 5.0 12/03/26 -> ISSUER 5.0 12/03/2026."""
    return re.sub(r"(\d{1,2}/\d{1,2}/)(\d{2})(?!\d)", r"\g<1>20\2", texto or "")


def _restar_meses(fecha, meses):
    indice = fecha.year * 12 + (fecha.month - 1) - meses
    anio, mes = divmod(indice, 12)
    mes += 1
    dia = min(fecha.day, calendar.monthrange(anio, mes)[1])
    return date(anio, mes, dia)


# --- Errores conocidos: texto del error -> corrección -------------------

def _corregir_sort_duplicado(page, d):
    poner_valor(page, "ici_sortSuffix", (d.cusip or "")[2:-1])
    d.name = _anio_cuatro_digitos(d.name)
    poner_valor(page, "ici_description1", d.name)


def _corregir_commodity(page, d):
    d.commodity = (d.name or "")[:4]
    poner_valor(page, "ici_sortIssueCommodityCodePart1", d.commodity)


def _corregir_nickname(page, d):
    d.nickname = None
    poner_valor(page, "ici_nickname", None)


def _corregir_cupon(page, d):
    d.cpn = "0"
    poner_valor(page, "ici_sortCouponStrikePrice", "0")


def _corregir_penultimo_cupon(page, d):
    meses = MESES_POR_FRECUENCIA.get(d.frecuencia)
    if meses is None or not d.maturity:
        raise ImpactError(
            "No se puede calcular NEXT TO LAST COUPON DATE "
            "(falta MATURITY o CPN_FREQ válido)."
        )
    vencimiento = datetime.strptime(d.maturity, "%m%d%Y").date()
    poner_valor(
        page,
        "ici_nextToLastCouponDate",
        _restar_meses(vencimiento, meses).strftime("%m%d%Y"),
    )


ERRORES_NSEX02 = {
    "DUPLICATE SORT SEQUENCE NOT ALLOWED": _corregir_sort_duplicado,
    "INVALID ISSUE/COMMODITY CODE ENTRY": _corregir_commodity,
    "DUPLICATE NICKNAME NOT ALLOWED": _corregir_nickname,
    "INVALID COUPON/STRIKE PRICE": _corregir_cupon,
}

ERRORES_NSEX03 = {
    "NEXT TO LAST COUPON DATE ENTRY REQUIRED": _corregir_penultimo_cupon,
}


def avanzar(page, origen, destino, errores, datos, campo_ancla=None):
    """Da Enter hasta llegar a `destino`: repite ante warnings y corrige los
    errores conocidos. Un error desconocido (o que no se corrige) detiene
    el security con ImpactError."""
    repeticiones = {}

    for _ in range(MAX_ENTER_POR_PANTALLA):
        enter(page, campo_ancla)
        pantalla = pantalla_actual(page)

        if pantalla == destino:
            return
        if destino == "NSEX04" and pantalla not in (origen, "NSEX05"):
            return  # NSEX04 no tiene campos propios con los que reconocerla

        mensajes = re.findall(r"ERROR:?\s*([^\n\r]+)", leer_texto(page).upper())
        if not mensajes:
            continue  # warning: se vuelve a dar Enter

        for mensaje in mensajes:
            for patron, corregir in errores.items():
                if patron in mensaje:
                    repeticiones[patron] = repeticiones.get(patron, 0) + 1
                    if repeticiones[patron] > MAX_REINTENTOS_POR_ERROR:
                        raise ImpactError(
                            f"El error '{patron}' sigue apareciendo tras "
                            "corregirlo."
                        )
                    corregir(page, datos)
                    break
            else:
                raise ImpactError(f"Error no manejado en {origen}: {mensaje}")

    raise ImpactError(f"No se pudo avanzar de {origen} a {destino}.")


# ============================================================
# ALTA DE UN SECURITY
# ============================================================

def alta_5_80(page, sec):
    """Llena NSEX02 a NSEX05 para un security 5/80 y deja la última pantalla
    sin guardar. Devuelve los avisos de datos faltantes (lista)."""
    d = construir_datos(sec)

    ir_a_seleccion(page)
    poner_valor(page, "ici_selectMarket", MERCADO)
    poner_valor(page, "ici_selectInstrument", INSTRUMENTO)
    enter(page, "ici_selectInstrument")
    esperar_pantalla(page, "NSEX02")

    llenar(page, campos_nsex02(d))
    avanzar(page, "NSEX02", "NSEX03", ERRORES_NSEX02, d,
            campo_ancla="ici_securityName")

    llenar(page, campos_nsex03(d))
    avanzar(page, "NSEX03", "NSEX04", ERRORES_NSEX03, d,
            campo_ancla="ici_informationCommonStrikePrice")

    avanzar(page, "NSEX04", "NSEX05", {}, d)

    esperar_pantalla(page, "NSEX05")
    llenar(page, campos_nsex05(d))

    return d.avisos


# ============================================================
# FLUJO
# ============================================================

def _etiqueta(sec):
    isin = sec.filas.get(21)
    return str(isin) if isin else f"columna {sec.columna}"


def procesar_securities_impact(page, securities):
    """Da de alta (sin guardar) los securities 5/80 en IMPACT, uno por uno.

    Tras dejar cada uno listo en NSEX05 espera a que la persona lo guarde a
    mano. Al final imprime qué securities se tomaron y cuáles no.
    """
    resultados = []   # (estado, etiqueta, detalle)
    pendientes = [s for s in securities if s.skeleton == SKELETON]

    for sec in securities:
        if sec.skeleton != SKELETON:
            resultados.append(
                ("No tomado", _etiqueta(sec), f"no cumple el skeleton {SKELETON}")
            )

    terminar = False
    for i, sec in enumerate(pendientes, start=1):
        etiqueta = _etiqueta(sec)

        if terminar:
            resultados.append(("No tomado", etiqueta, "proceso interrumpido"))
            continue

        print()
        print("=" * 60)
        print(f"IMPACT: security {i} de {len(pendientes)} - {etiqueta}")
        print("=" * 60)

        try:
            avisos = alta_5_80(page, sec)

        except Exception as error:
            print(f"No se pudo preparar el security: {error}")
            resultados.append(("Con error", etiqueta, str(error)))
            continue

        for aviso in avisos:
            print(f"Aviso: {aviso}")

        print()
        print("Security listo en NSEX05. Revísalo y guárdalo a mano en IMPACT.")
        respuesta = input(
            "Enter = ya lo guardé | o = no lo guardé (omitir) | "
            "q = terminar: "
        ).strip().lower()

        if respuesta == "q":
            resultados.append(("No tomado", etiqueta, "interrumpido por ti"))
            terminar = True
        elif respuesta == "o":
            resultados.append(("Preparado, no guardado", etiqueta, ""))
        else:
            resultados.append(("Guardado (confirmado por ti)", etiqueta, ""))

    print()
    print("=" * 60)
    print("RESUMEN IMPACT")
    print("=" * 60)
    for estado in (
        "Guardado (confirmado por ti)",
        "Preparado, no guardado",
        "Con error",
        "No tomado",
    ):
        grupo = [r for r in resultados if r[0] == estado]
        if not grupo:
            continue
        print(f"{estado}: {len(grupo)}")
        for _, etiqueta, detalle in grupo:
            print(f"  - {etiqueta}" + (f"  ({detalle})" if detalle else ""))

    return resultados

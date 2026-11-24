from pathlib import Path
from datetime import datetime
import shutil
import time

import pandas as pd
from playwright.sync_api import sync_playwright

from bloombot import procesar_y_enviar
from prestaging import procesar_gimp
from genie import hacer_click_sanction_check
from sanctions import procesar_sanctions
from sanctions_pdf import verificar_securities
from skeletons import procesar_matriz
from impact import procesar_securities_impact
from setup_securities import (
    elegir_archivo,
    nueva_carpeta_bloombot,
    preparar_setup,
)


# ============================================================
# CONFIGURACIÓN
# ============================================================

URL_SISTEMA = (
    "https://prd.impact.spsfi.broadridge.com/"
    "impact_hsbpro/controller"
)

PERFIL = (Path.cwd() / "perfil_chrome_sistema").resolve()

COLUMNAS = [
    "Entity",
    "External Trade ID",
    "Trader",
    "Cusip",
    "Name",
    "Date",
    "Status",
    "Security Id",
    "Amount_External",
    "Depository",
    "Amount_Normalized",
    "Cur",
]

ERRORES_OBJETIVO = {
    "SECURITY NOT FOUND",
    "SECURITY DOES NOT EXIST IN MSD, BROADCAST SENT",
}

MENSAJE_FIN = "LAST QUALIFYING TRANSACTION ALREADY SHOWN"

MAX_PAGINAS = 100

# Reintentos de detección y de extracción
INTENTOS_DETECCION = 5
PAUSA_DETECCION_S = 2
INTENTOS_EXTRACCION = 3


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def limpiar_texto(texto):
    return " ".join(texto.split()).strip()


def texto_visible(page):
    """
    Texto del body de la página y de todos sus iframes.
    """

    partes = []

    for frame in page.frames:
        try:
            partes.append(
                frame.locator("body").inner_text(timeout=30000)
            )
        except Exception:
            continue

    return "\n".join(partes)


def esperar_pagina_lista(page, timeout_ms=15000):
    """
    Espera a que la página termine de cargar.
    No falla si algún estado no se alcanza.
    """

    for estado in ("domcontentloaded", "load", "networkidle"):
        try:
            page.wait_for_load_state(estado, timeout=timeout_ms)
        except Exception:
            pass


def obtener_tabla_objetivo(page):
    """
    Busca la tabla que contiene 'External Trade ID'
    en la página y en todos sus iframes.
    """

    for numero_frame, frame in enumerate(page.frames):
        try:
            tablas = frame.locator("table")
            total = tablas.count()
        except Exception:
            continue

        for i in range(total):
            try:
                tabla = tablas.nth(i)
                contenido = limpiar_texto(tabla.inner_text())

                if "External Trade ID" in contenido:
                    return tabla, f"{numero_frame}:{i}"

            except Exception:
                continue

    return None, None


def tabla_visible(
    page,
    intentos=INTENTOS_DETECCION,
    pausa_s=PAUSA_DETECCION_S,
):
    """
    Comprueba si la tabla principal está visible.

    Reintenta varias veces esperando la carga de la página,
    porque la tabla puede tardar en pintarse.
    """

    for intento in range(1, intentos + 1):
        esperar_pagina_lista(page)

        tabla, _ = obtener_tabla_objetivo(page)

        if tabla is not None:
            return True

        if intento < intentos:
            print(
                f"Tabla no detectada "
                f"(intento {intento}/{intentos}); reintentando...",
                flush=True,
            )
            time.sleep(pausa_s)

    return False


def extraer_pagina(page, numero_pagina):
    """
    Extrae las filas de la tabla de la página actual.

    Devuelve:
        registros: lista de diccionarios.
        firma: texto que identifica el contenido de la página.
    """

    tabla, indice_tabla = obtener_tabla_objetivo(page)

    # Durante la carga de una página la tabla puede desaparecer
    # un instante; se reintenta antes de darla por perdida.
    if tabla is None and numero_pagina > 0:
        for _ in range(3):
            time.sleep(2)
            tabla, indice_tabla = obtener_tabla_objetivo(page)

            if tabla is not None:
                break

    if tabla is None:
        if numero_pagina > 0:
            print(
                f"[Página {numero_pagina}] "
                "No se encontró la tabla objetivo.",
                flush=True,
            )
        return [], ""

    if numero_pagina > 0:
        print(
            f"[Página {numero_pagina}] "
            f"Tabla objetivo encontrada en índice {indice_tabla}.",
            flush=True,
        )

    filas = tabla.locator("tr")
    registros = []

    for i in range(filas.count()):
        fila = filas.nth(i)
        celdas = fila.locator("th, td")
        valores = []

        for j in range(celdas.count()):
            try:
                valor = limpiar_texto(celdas.nth(j).inner_text())
            except Exception:
                valor = ""

            valores.append(valor)

        # Ignorar encabezados y filas auxiliares
        if len(valores) != len(COLUMNAS):
            continue

        if valores[0] == "Entity":
            continue

        registro = dict(zip(COLUMNAS, valores))
        registro["Pagina"] = numero_pagina
        registros.append(registro)

    firma = "|".join(
        f"{registro['External Trade ID']}::"
        f"{registro['Status']}::"
        f"{registro['Cusip']}"
        for registro in registros
    )

    if numero_pagina > 0:
        print(
            f"[Página {numero_pagina}] "
            f"Filas extraídas: {len(registros)}",
            flush=True,
        )

    return registros, firma


def obtener_enlace_next(page):
    """
    Busca el enlace cuyo texto visible sea Next
    en la página y en todos sus iframes.
    """

    for frame in page.frames:
        try:
            enlaces = frame.locator("a")
            total = enlaces.count()
        except Exception:
            continue

        for i in range(total):
            enlace = enlaces.nth(i)

            try:
                texto = limpiar_texto(enlace.inner_text())

                if texto.upper() == "NEXT":
                    return enlace

            except Exception:
                continue

    return None


def guardar_captura(page, carpeta, nombre_archivo):
    """
    Guarda una captura completa de la página.
    """

    ruta = carpeta / nombre_archivo

    page.screenshot(
        path=str(ruta),
        full_page=True,
    )

    print(
        f"Captura guardada: {ruta.name}",
        flush=True,
    )

    return ruta


def esperar_cambio_de_pagina(
    page,
    firma_anterior,
    segundos_maximos=30,
):
    """
    Espera a que la tabla cambie después de pulsar Next.

    Devuelve:
        CAMBIO     si la tabla cambió.
        FIN        si apareció el mensaje de última página.
        SIN_CAMBIO si agotó el tiempo.
    """

    inicio = time.time()

    while time.time() - inicio < segundos_maximos:
        time.sleep(0.5)

        try:
            texto_actual = texto_visible(page).upper()

            if MENSAJE_FIN in texto_actual:
                return "FIN"

            _, firma_actual = extraer_pagina(page, -1)

            if firma_actual and firma_actual != firma_anterior:
                return "CAMBIO"

        except Exception:
            pass

    return "SIN_CAMBIO"


def deduplicar_registros(df):
    """
    Elimina duplicados cuando existe External Trade ID.

    Las filas sin External Trade ID se conservan.
    """

    columna_id = "External Trade ID"

    con_id = df[
        df[columna_id]
        .fillna("")
        .astype(str)
        .str.strip()
        != ""
    ].copy()

    sin_id = df[
        df[columna_id]
        .fillna("")
        .astype(str)
        .str.strip()
        == ""
    ].copy()

    con_id = con_id.drop_duplicates(
        subset=[columna_id],
        keep="first",
    )

    return pd.concat(
        [con_id, sin_id],
        ignore_index=True,
    )


# ============================================================
# EXTRACCIÓN COMPLETA DE UNA EJECUCIÓN
# ============================================================

def ejecutar_extraccion(page):
    """
    Procesa todas las páginas desde la posición actual
    y genera una carpeta de evidencia independiente.

    Devuelve la ruta del Excel filtrado, o None si no se
    extrajo nada (en ese caso no deja carpeta vacía).
    """

    fecha_ejecucion = datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    evidencia = (
        Path.cwd() / f"TUR Evidence_{fecha_ejecucion}"
    )

    print()
    print("=" * 60)
    print("INICIANDO NUEVA EXTRACCIÓN")
    print("=" * 60)
    print(f"Carpeta de evidencia: {evidencia}")
    print()

    if not tabla_visible(page):
        print(
            "No se detectó la tabla con 'External Trade ID'."
        )
        print(
            "Ve manualmente a la primera página de la tabla "
            "y vuelve a intentarlo."
        )
        return None

    evidencia.mkdir(
        parents=True,
        exist_ok=True,
    )

    registros_totales = []
    firmas_vistas = set()
    paginas_procesadas = 0

    for numero_pagina in range(1, MAX_PAGINAS + 1):
        print()
        print("-" * 60)
        print(f"Procesando página {numero_pagina}")
        print("-" * 60)

        # Tiempo para que termine de renderizarse la tabla
        page.wait_for_timeout(2000)

        registros, firma_actual = extraer_pagina(
            page,
            numero_pagina,
        )

        if not registros:
            print(
                "No se encontraron registros. "
                "Se detiene esta extracción."
            )
            break

        if firma_actual in firmas_vistas:
            print(
                "La página ya se había procesado. "
                "Se detiene para evitar duplicados."
            )
            break

        firmas_vistas.add(firma_actual)
        registros_totales.extend(registros)
        paginas_procesadas = numero_pagina

        # Captura de la página actual
        guardar_captura(
            page,
            evidencia,
            f"TUR_page_{numero_pagina:03d}.png",
        )

        texto_actual = texto_visible(page).upper()

        # Si el aviso ya aparece, se captura y termina
        if MENSAJE_FIN in texto_actual:
            print(
                "Se detectó el aviso de última página."
            )

            guardar_captura(
                page,
                evidencia,
                f"TUR_last_page_notice_{numero_pagina:03d}.png",
            )

            break

        enlace_next = obtener_enlace_next(page)

        if enlace_next is None:
            print(
                "No se encontró el enlace Next. "
                "Se detiene esta extracción."
            )
            break

        firma_anterior = firma_actual

        print("Pulsando Next...", flush=True)

        try:
            enlace_next.click(timeout=30000)

        except Exception as error:
            print(
                f"No se pudo pulsar Next: {error}"
            )
            break

        resultado = esperar_cambio_de_pagina(
            page,
            firma_anterior,
            segundos_maximos=30,
        )

        if resultado == "FIN":
            print(
                "Se detectó: "
                "Last Qualifying Transaction Already Shown"
            )

            guardar_captura(
                page,
                evidencia,
                f"TUR_last_page_notice_{numero_pagina:03d}.png",
            )

            break

        if resultado == "SIN_CAMBIO":
            print(
                "La tabla no cambió después de pulsar Next. "
                "Se detiene por seguridad."
            )
            break

        print(
            "La nueva página se cargó correctamente.",
            flush=True,
        )

    if not registros_totales:
        print(
            "No se extrajeron registros en esta ejecución."
        )
        # No dejar una carpeta de evidencia vacía
        shutil.rmtree(evidencia, ignore_errors=True)
        return None

    # ========================================================
    # GENERAR EXCEL
    # ========================================================

    df = pd.DataFrame(
        registros_totales,
        columns=COLUMNAS + ["Pagina"],
    )

    df = deduplicar_registros(df)

    df["Status_normalizado"] = (
        df["Status"]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    df_filtrado = df[
        df["Status_normalizado"].isin(ERRORES_OBJETIVO)
    ].copy()

    df.drop(
        columns=["Status_normalizado"],
        inplace=True,
    )

    df_filtrado.drop(
        columns=["Status_normalizado"],
        inplace=True,
    )

    archivo_todos = (
        evidencia / "TUR_requests_todas_paginas.xlsx"
    )

    archivo_filtrado = (
        evidencia / "TUR_requests_filtrados.xlsx"
    )

    df.to_excel(
        archivo_todos,
        index=False,
    )

    df_filtrado.to_excel(
        archivo_filtrado,
        index=False,
    )

    # ========================================================
    # RESUMEN
    # ========================================================

    print()
    print("=" * 60)
    print("EXTRACCIÓN TERMINADA")
    print("=" * 60)
    print(f"Páginas procesadas: {paginas_procesadas}")
    print(f"Filas totales: {len(df)}")
    print(f"Filas filtradas: {len(df_filtrado)}")
    print()
    print(f"Carpeta de evidencia:")
    print(evidencia)
    print()
    print(f"Excel completo:")
    print(archivo_todos)
    print()
    print(f"Excel filtrado:")
    print(archivo_filtrado)
    print()

    print("Estados encontrados en el Excel filtrado:")

    if len(df_filtrado) > 0:
        print(
            df_filtrado["Status"]
            .value_counts()
            .to_string()
        )
    else:
        print("Ninguno")

    return archivo_filtrado


def ejecutar_extraccion_con_reintentos(
    page,
    intentos=INTENTOS_EXTRACCION,
):
    """
    Corre la extracción y, si no obtiene nada, vuelve a
    detectar la tabla y la repite.

    Solo se reintenta cuando no se extrajo ninguna fila, es
    decir, cuando el navegador sigue en la primera página.
    Devuelve la ruta del Excel filtrado, o None.
    """

    for intento in range(1, intentos + 1):
        archivo_filtrado = ejecutar_extraccion(page)

        if archivo_filtrado is not None:
            return archivo_filtrado

        if intento < intentos:
            print()
            print(
                f"Extracción sin resultados "
                f"(intento {intento}/{intentos}). "
                "Volviendo a detectar la tabla...",
                flush=True,
            )
            time.sleep(3)

    print()
    print(
        "No se obtuvo ningún resultado tras "
        f"{intentos} intentos."
    )
    return None


def enviar_a_bloombot(
    archivo_filtrado,
    page,
    carpeta_evidencia=None,
    guardado_automatico=False,
):
    """
    Pasa los códigos del Excel filtrado a BLOOMBOT y, cuando la macro
    termina, mueve los archivos GIMP a la carpeta de evidencia, llena
    el PreStaging File, clasifica los securities, pulsa "Perform
    Sanctions Check" en Genie, mueve el PDF resultante a la carpeta de
    evidencia y prepara el alta de los securities en IMPACT.

    `carpeta_evidencia` es donde se guardan GIMP, Sanctions y revisiones
    (por defecto, la carpeta del Excel filtrado). Con `guardado_automatico`
    IMPACT guarda cada security sin esperar la revisión manual.
    Un error aquí no interrumpe el programa.
    """

    inicio = datetime.now()

    try:
        resultado = procesar_y_enviar(archivo_filtrado)

    except Exception as error:
        print()
        print(f"No se pudo enviar a BLOOMBOT: {error}")
        return

    if not resultado:
        return

    cantidad, entidades = resultado

    if carpeta_evidencia is None:
        carpeta_evidencia = Path(archivo_filtrado).parent

    try:
        matriz = procesar_gimp(
            cantidad,
            inicio,
            carpeta_evidencia,
        )

    except Exception as error:
        print()
        print(f"No se pudo completar el PreStaging File: {error}")
        return

    # Clasificación por skeleton y limpieza de datos (no detiene el flujo)
    securities = None

    if matriz:
        try:
            securities = procesar_matriz(
                matriz,
                carpeta_evidencia,
                entidades,
            )

        except Exception as error:
            print()
            print(f"No se pudo clasificar los securities: {error}")

    inicio_sanctions = datetime.now()
    pdfs = []

    try:
        hacer_click_sanction_check()

        pdfs = procesar_sanctions(
            inicio_sanctions,
            carpeta_evidencia,
        )

    except Exception as error:
        print()
        print(f"El Sanctions Check no se completó: {error}")

    if not securities:
        return

    # Revisión del PDF: solo los securities con PASS pueden crearse. Si no
    # hay PDF o no se puede leer, no se crea ninguno.
    try:
        verificar_securities(pdfs, matriz, securities, carpeta_evidencia)

    except Exception as error:
        print()
        print(f"No se pudo revisar el Sanctions Check: {error}")
        for sec in securities:
            sec.bloqueo = f"no se pudo revisar el Sanctions Check: {error}"

    try:
        procesar_securities_impact(
            page,
            securities,
            guardado_automatico=guardado_automatico,
        )

    except Exception as error:
        print()
        print(f"No se pudo completar el trabajo en IMPACT: {error}")


# ============================================================
# OPCIONES DEL MENÚ
# ============================================================

def preparar_tabla(page):
    """Pide dejar Chrome en la primera página de la tabla y la comprueba.
    Devuelve True si se puede empezar."""

    print()
    print("Deja Chrome en la primera página de la tabla.")
    print("Cuando esté listo, vuelve aquí y pulsa Enter.")
    input()

    while not tabla_visible(page):
        print()
        print(
            "No se detectó la tabla en la pestaña "
            "controlada por Playwright."
        )

        volver = input(
            "¿Quieres volver a comprobarla? (s/n): "
        ).strip().lower()

        if volver != "s":
            print("No se inició el proceso.")
            return False

    return True


def preguntar_guardado_automatico():
    """True si IMPACT debe guardar cada security sin revisión manual."""

    respuesta = input(
        "¿Guardar automáticamente cada security en IMPACT? (s/N): "
    ).strip().lower()

    return respuesta == "s"


def proceso_completo(page, guardado_automatico=False):
    """Extracción, BLOOMBOT, Sanctions e IMPACT."""

    archivo_filtrado = ejecutar_extraccion_con_reintentos(page)

    if archivo_filtrado is not None:
        enviar_a_bloombot(
            archivo_filtrado,
            page,
            guardado_automatico=guardado_automatico,
        )


def solo_extraccion(page):
    """Solo capturas de pantalla y Excel filtrados (antes de BLOOMBOT)."""

    ejecutar_extraccion_con_reintentos(page)


def desde_bloombot(page):
    """BLOOMBOT, Sanctions e IMPACT a partir de un Excel filtrado ya
    existente; todo se guarda en una carpeta nueva."""

    excel = elegir_archivo(
        "Elige el Excel filtrado (TUR_requests_filtrados.xlsx)"
    )

    if excel is None:
        print("No se eligió ningún archivo.")
        return

    guardado_automatico = preguntar_guardado_automatico()
    carpeta, copia = nueva_carpeta_bloombot(excel)

    print()
    print(f"Carpeta nueva: {carpeta}")

    enviar_a_bloombot(
        copia,
        page,
        carpeta_evidencia=carpeta,
        guardado_automatico=guardado_automatico,
    )


def solo_setup(page):
    """Solo el set up de securities en IMPACT, a partir de una carpeta de
    evidencia ya existente."""

    excel = elegir_archivo(
        "Elige el Excel base (XLS de BLOOMBOT GIMP*.xls o el Excel filtrado)"
    )

    if excel is None:
        print("No se eligió ningún archivo.")
        return

    guardado_automatico = preguntar_guardado_automatico()

    try:
        securities, _ = preparar_setup(excel)

    except Exception as error:
        print()
        print(f"No se pudo preparar el set up: {error}")
        return

    try:
        procesar_securities_impact(
            page,
            securities,
            guardado_automatico=guardado_automatico,
        )

    except Exception as error:
        print()
        print(f"No se pudo completar el trabajo en IMPACT: {error}")


def mostrar_menu():
    print()
    print("=" * 60)
    print("¿QUÉ QUIERES HACER?")
    print("=" * 60)
    print("1 - Todo el proceso")
    print("2 - Todo el proceso sin chequeo manual (IMPACT guarda solo)")
    print("3 - Solo extracción: capturas y Excel filtrados")
    print("4 - Desde BLOOMBOT: Sanctions y creación (elige el Excel filtrado)")
    print("5 - Solo set up de securities en IMPACT (elige el Excel base)")
    print("6 - Cerrar el proceso")


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():
    with sync_playwright() as p:
        print("Abriendo Chrome y el sistema...", flush=True)

        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PERFIL),
            channel="chrome",
            headless=False,
            timeout=30000,
        )

        page = (
            context.pages[0]
            if context.pages
            else context.new_page()
        )

        page.goto(
            URL_SISTEMA,
            wait_until="commit",
            timeout=60000,
        )

        print()
        print("1. Introduce tus credenciales manualmente en Chrome.")
        print("2. Completa MFA si el sistema lo solicita.")
        print("3. Regresa a esta consola y pulsa Enter para ver el menú.")
        input()

        while True:
            mostrar_menu()

            opcion = input("Selecciona una opción (1-6): ").strip()

            if opcion == "1":
                if preparar_tabla(page):
                    proceso_completo(page)

            elif opcion == "2":
                if preparar_tabla(page):
                    proceso_completo(page, guardado_automatico=True)

            elif opcion == "3":
                if preparar_tabla(page):
                    solo_extraccion(page)

            elif opcion == "4":
                desde_bloombot(page)

            elif opcion == "5":
                solo_setup(page)

            elif opcion == "6":
                cerrar = input(
                    "¿Cerrar también Chrome? (s/n): "
                ).strip().lower()

                if cerrar == "s":
                    print()
                    print("Cerrando Chrome...")
                    context.close()

                else:
                    print()
                    print(
                        "El script terminará y se intentará dejar "
                        "Chrome abierto."
                    )
                    # No se llama a context.close().
                    # Esto evita cerrar Chrome explícitamente.

                break

            else:
                print()
                print("Opción no válida. Selecciona un número del 1 al 6.")


if __name__ == "__main__":
    main()

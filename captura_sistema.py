        channel="chrome",
        headless=False,
        timeout=30000,
    )

    page = context.pages[0] if context.pages else context.new_page()

    page.goto(
        URL_SISTEMA,
        wait_until="commit",
        timeout=60000,
    )

    print()
    print("1. Introduce tus credenciales manualmente en Chrome.")
    print("2. Completa MFA si el sistema lo solicita.")
    print("3. Navega hasta la primera página de la tabla.")
    print("4. Asegúrate de que la tabla esté visible.")
    print("5. Regresa a esta consola y pulsa Enter.")
    input()

    if not tabla_visible(page):
        print()
        print(
            "No se detectó la tabla con "
            "'External Trade ID'."
        )
        print(
            "Comprueba que estás en la pantalla correcta."
        )

        respuesta = input(
            "¿Quieres continuar de todos modos? (s/n): "
        ).strip().lower()

        if respuesta != "s":
            print("Proceso cancelado.")
            context.close()
            raise SystemExit(0)

    while True:
        ejecutar_extraccion(page)

        print()
        print("=" * 60)
        print("¿QUÉ QUIERES HACER AHORA?")
        print("=" * 60)
        print("1 - Repetir la extracción")
        print("2 - Dejar Chrome abierto y terminar")
        print("3 - Cerrar Chrome y terminar")

        opcion = input(
            "Selecciona 1, 2 o 3: "
        ).strip()

        if opcion == "1":
            print()
            print(
                "Deja Chrome en la primera página de la tabla."
            )
            print(
                "Cuando esté listo, vuelve aquí y pulsa Enter."
            )
            input()

            if not tabla_visible(page):
                print()
                print(
                    "No se detectó la tabla en la pestaña "
                    "controlada por Playwright."
                )

                continuar = input(
                    "¿Quieres volver a comprobarla? (s/n): "
                ).strip().lower()

                if continuar != "s":
                    print(
                        "No se inició la nueva extracción."
                    )
                    continue

            continue

        if opcion == "2":
            print()
            print(
                "El script terminará y se intentará dejar "
                "Chrome abierto."
            )
            print(
                "La sesión y la ventana quedarán disponibles "
                "para continuar manualmente."
            )

            # No se llama a context.close().
            # Esto evita cerrar Chrome explícitamente.
            break

        if opcion == "3":
            print()
            print("Cerrando Chrome...")
            context.close()
            break

        print()
        print("Opción no válida. Selecciona 1, 2 o 3.")

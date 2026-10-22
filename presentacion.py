import pandas as pd
import matplotlib.pyplot as plt
import time


# ============================================================
# CONFIGURACIÓN
# ============================================================

ARCHIVO = "galletitas_mariano_max.xlsx"
PRECIO_MINIMO_VALIDO = 100

PORCENTAJES = {
    "DULCES": 0.40,
    "SALADAS": 0.30,
    "RELLENAS": 0.20,
    "OBLEAS": 0.10
}


# ============================================================
# FUNCIONES GENERALES
# ============================================================

def pesos(valor):
    """
    Convierte un número a formato de pesos argentinos.
    Ejemplo: 150000 -> $150.000,00
    """

    texto = f"{valor:,.2f}"

    texto = (
        texto
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    return "$" + texto


def pausa():
    time.sleep(0.5)


# ============================================================
# CLASIFICACIÓN DE PRODUCTOS
# ============================================================

def clasificar_producto(descripcion):

    descripcion = str(descripcion).upper()

    saladas = [
        "CRACKER",
        "CRAKERS",
        "CRIOLLITAS",
        "MARINERA",
        "SALVADO",
        "CLUB SOCIAL",
        "PORTEÑITAS",
        "KESITAS",
        "SEMOLADAS",
        "DORADITAS",
        "HOGAREÑAS",
        "GRANAGUA",
        "SIN SAL",
        "S/SAL",
        "SESAMO",
        "AMARAN"
    ]

    rellenas = [
        "RELLENA",
        "RELLENAS",
        "RELL",
        "SANDWICH",
        "SONRISAS",
        "BRIGITTE",
        "DALE BLACK"
    ]

    obleas = [
        "OBLEA",
        "OPERA",
        "BLOCK"
    ]

    if any(palabra in descripcion for palabra in saladas):
        return "SALADAS"

    if any(palabra in descripcion for palabra in rellenas):
        return "RELLENAS"

    if any(palabra in descripcion for palabra in obleas):
        return "OBLEAS"

    return "DULCES"


# ============================================================
# CARGA Y LIMPIEZA
# ============================================================

def cargar_datos():

    print("\n" + "=" * 55)
    print("       SISTEMA DE COMPRA DE GALLETITAS")
    print("=" * 55)

    print("\n[1/5] Buscando archivo...")
    pausa()

    try:

        datos = pd.read_excel(ARCHIVO)

    except FileNotFoundError:

        print("\nERROR:")
        print(f"No se encontró el archivo {ARCHIVO}")
        exit()

    print("[2/5] Cargando base de datos...")
    pausa()

    columnas_necesarias = [
        "sucursal",
        "descripcion",
        "marca",
        "precio"
    ]

    for columna in columnas_necesarias:

        if columna not in datos.columns:

            print(
                f"\nERROR: falta la columna '{columna}'"
            )

            exit()

    filas_originales = len(datos)

    print("[3/5] Limpiando datos...")
    pausa()

    # Eliminar duplicados
    datos = datos.drop_duplicates(
        subset=[
            "sucursal",
            "descripcion",
            "marca",
            "precio"
        ]
    ).copy()

    # Convertir precio a número
    datos["precio"] = pd.to_numeric(
        datos["precio"],
        errors="coerce"
    )

    # Eliminar valores vacíos
    datos = datos.dropna(
        subset=[
            "descripcion",
            "marca",
            "precio"
        ]
    )

    # Eliminar precios sospechosos
    datos = datos[
        datos["precio"] >= PRECIO_MINIMO_VALIDO
    ].copy()

    print("[4/5] Clasificando productos...")
    pausa()

    datos["categoria"] = (
        datos["descripcion"]
        .apply(clasificar_producto)
    )

    print("[5/5] Comparando precios entre sucursales...")
    pausa()

    # ========================================================
    # SELECCIONAR LA SUCURSAL MÁS BARATA PARA CADA PRODUCTO
    # ========================================================

    indices = (
        datos
        .groupby(
            ["descripcion", "marca"]
        )["precio"]
        .idxmin()
    )

    catalogo = (
        datos
        .loc[indices]
        .copy()
        .reset_index(drop=True)
    )

    catalogo = catalogo.sort_values(
        [
            "categoria",
            "precio"
        ]
    )

    catalogo = catalogo.reset_index(drop=True)

    # ========================================================
    # INFORMACIÓN DE CARGA
    # ========================================================

    print("\n" + "=" * 55)
    print("                  DATOS LISTOS")
    print("=" * 55)

    print(
        f"\nFilas originales: {filas_originales}"
    )

    print(
        f"Productos válidos encontrados: {len(catalogo)}"
    )

    print("\nProductos por categoría:")

    conteo = (
        catalogo["categoria"]
        .value_counts()
    )

    for categoria in PORCENTAJES:

        cantidad = conteo.get(
            categoria,
            0
        )

        print(
            f"  {categoria:<10}: "
            f"{cantidad}"
        )

    print(
        "\nLa base ya fue procesada correctamente."
    )

    return catalogo


# ============================================================
# PEDIR PRESUPUESTO
# ============================================================

def pedir_presupuesto():

    print("\n" + "-" * 55)

    while True:

        try:

            presupuesto = float(
                input(
                    "\nIngrese el presupuesto disponible: $"
                )
            )

            if presupuesto <= 0:

                print(
                    "El presupuesto debe ser mayor a 0."
                )

                continue

            return presupuesto

        except ValueError:

            print(
                "Ingrese solamente números."
            )


# ============================================================
# COMPRA INICIAL POR CATEGORÍA
# ============================================================

def comprar_categoria(
    catalogo,
    categoria,
    presupuesto_categoria,
    cantidades
):

    productos = catalogo[
        catalogo["categoria"] == categoria
    ]

    gasto = 0

    if productos.empty:
        return gasto

    while True:

        disponible = (
            presupuesto_categoria - gasto
        )

        candidatos = productos[
            productos["precio"] <= disponible
        ].copy()

        if candidatos.empty:
            break

        # Cantidad ya comprada de cada producto
        candidatos["cantidad_comprada"] = (
            candidatos.index.map(
                lambda i: cantidades.get(i, 0)
            )
        )

        # Primero variedad.
        # Después precio.
        candidatos = candidatos.sort_values(
            [
                "cantidad_comprada",
                "precio"
            ]
        )

        elegido = candidatos.index[0]

        precio = catalogo.loc[
            elegido,
            "precio"
        ]

        cantidades[elegido] = (
            cantidades.get(elegido, 0) + 1
        )

        gasto += precio

    return gasto


# ============================================================
# APROVECHAR EL PRESUPUESTO SOBRANTE
# ============================================================

def gastar_sobrante(
    catalogo,
    presupuesto,
    cantidades
):

    gasto = sum(

        catalogo.loc[i, "precio"]
        * cantidad

        for i, cantidad
        in cantidades.items()

    )

    while True:

        sobrante = presupuesto - gasto

        candidatos = catalogo[
            catalogo["precio"] <= sobrante
        ].copy()

        if candidatos.empty:
            break

        # ====================================================
        # BUSCAMOS EL PRODUCTO QUE MEJOR APROVECHE
        # EL DINERO QUE QUEDA
        #
        # Elegimos el producto más caro que todavía entra.
        # ====================================================

        candidatos = candidatos.sort_values(
            "precio",
            ascending=False
        )

        elegido = candidatos.index[0]

        precio = catalogo.loc[
            elegido,
            "precio"
        ]

        cantidades[elegido] = (
            cantidades.get(elegido, 0) + 1
        )

        gasto += precio

    return gasto


# ============================================================
# GENERAR LISTA COMPLETA
# ============================================================

def generar_lista(
    catalogo,
    presupuesto
):

    cantidades = {}

    print("\n" + "=" * 55)
    print("             GENERANDO LISTA")
    print("=" * 55)

    # ========================================================
    # ETAPA 1:
    # DISTRIBUIR SEGÚN PORCENTAJES
    # ========================================================

    for categoria, porcentaje in PORCENTAJES.items():

        monto_categoria = (
            presupuesto * porcentaje
        )

        print(
            f"\n{categoria}: "
            f"{porcentaje * 100:.0f}% "
            f"→ {pesos(monto_categoria)}"
        )

        comprar_categoria(
            catalogo,
            categoria,
            monto_categoria,
            cantidades
        )

        pausa()

    # ========================================================
    # ETAPA 2:
    # APROVECHAR SOBRANTE
    # ========================================================

    print(
        "\nBuscando aprovechar el presupuesto sobrante..."
    )

    pausa()

    gasto_total = gastar_sobrante(
        catalogo,
        presupuesto,
        cantidades
    )

    # ========================================================
    # ARMAR DATAFRAME FINAL
    # ========================================================

    filas = []

    for indice, cantidad in cantidades.items():

        producto = catalogo.loc[
            indice
        ]

        precio = producto["precio"]

        subtotal = (
            precio * cantidad
        )

        filas.append(
            {
                "categoria":
                    producto["categoria"],

                "descripcion":
                    producto["descripcion"],

                "marca":
                    producto["marca"],

                "sucursal":
                    producto["sucursal"],

                "cantidad":
                    cantidad,

                "precio_unitario":
                    precio,

                "subtotal":
                    subtotal
            }
        )

    lista = pd.DataFrame(filas)

    if not lista.empty:

        orden = {
            "DULCES": 1,
            "SALADAS": 2,
            "RELLENAS": 3,
            "OBLEAS": 4
        }

        lista["orden"] = (
            lista["categoria"]
            .map(orden)
        )

        lista = (
            lista
            .sort_values(
                [
                    "orden",
                    "subtotal"
                ],
                ascending=[
                    True,
                    False
                ]
            )
            .drop(
                columns="orden"
            )
            .reset_index(
                drop=True
            )
        )

    return lista, gasto_total


# ============================================================
# MOSTRAR LISTA EN CONSOLA
# ============================================================

def mostrar_lista(
    lista,
    presupuesto,
    gasto_total
):

    if lista.empty:

        print(
            "\nEl presupuesto no alcanza "
            "para ningún producto."
        )

        return

    print("\n\n" + "=" * 75)
    print("                    LISTA DE COMPRA")
    print("=" * 75)

    for categoria in PORCENTAJES:

        productos = lista[
            lista["categoria"] == categoria
        ]

        if productos.empty:
            continue

        gasto_categoria = (
            productos["subtotal"]
            .sum()
        )

        porcentaje_real = (
            gasto_categoria
            / gasto_total
            * 100
        )

        print(
            f"\n\n--- {categoria} ---"
        )

        print(
            f"Gasto: {pesos(gasto_categoria)}"
        )

        print(
            f"Porcentaje real: "
            f"{porcentaje_real:.1f}%"
        )

        tabla = productos[
            [
                "descripcion",
                "marca",
                "sucursal",
                "cantidad",
                "precio_unitario",
                "subtotal"
            ]
        ].copy()

        tabla["precio_unitario"] = (
            tabla["precio_unitario"]
            .apply(pesos)
        )

        tabla["subtotal"] = (
            tabla["subtotal"]
            .apply(pesos)
        )

        print(
            "\n",
            tabla.to_string(
                index=False
            )
        )

    # ========================================================
    # RESUMEN
    # ========================================================

    sobrante = (
        presupuesto - gasto_total
    )

    porcentaje_usado = (
        gasto_total
        / presupuesto
        * 100
    )

    print("\n\n" + "=" * 55)
    print("                    RESUMEN")
    print("=" * 55)

    print(
        f"\nPresupuesto: "
        f"{pesos(presupuesto)}"
    )

    print(
        f"Gasto total: "
        f"{pesos(gasto_total)}"
    )

    print(
        f"Sobrante: "
        f"{pesos(sobrante)}"
    )

    print(
        f"Presupuesto utilizado: "
        f"{porcentaje_usado:.2f}%"
    )

    print(
        f"Productos diferentes: "
        f"{len(lista)}"
    )

    print(
        f"Paquetes totales: "
        f"{lista['cantidad'].sum()}"
    )

    print(
        f"Sucursales a visitar: "
        f"{lista['sucursal'].nunique()}"
    )

    print("\nDistribución:")

    for categoria, objetivo in PORCENTAJES.items():

        gasto_categoria = (
            lista.loc[
                lista["categoria"] == categoria,
                "subtotal"
            ]
            .sum()
        )

        real = (
            gasto_categoria
            / gasto_total
            * 100
        )

        print(
            f"{categoria:<10} | "
            f"Objetivo {objetivo * 100:>4.0f}% | "
            f"Real {real:>5.1f}% | "
            f"{pesos(gasto_categoria)}"
        )


# ============================================================
# CREAR RESUMEN PARA EXCEL Y GRÁFICOS
# ============================================================

def crear_resumen(
    lista,
    presupuesto,
    gasto_total
):

    datos = []

    for categoria, porcentaje in PORCENTAJES.items():

        filtro = (
            lista["categoria"] == categoria
        )

        gasto = (
            lista.loc[
                filtro,
                "subtotal"
            ]
            .sum()
        )

        unidades = (
            lista.loc[
                filtro,
                "cantidad"
            ]
            .sum()
        )

        real = (
            gasto / gasto_total * 100
            if gasto_total > 0
            else 0
        )

        datos.append(
            {
                "categoria":
                    categoria,

                "porcentaje_objetivo":
                    porcentaje * 100,

                "porcentaje_real":
                    real,

                "gasto":
                    gasto,

                "unidades":
                    unidades
            }
        )

    return pd.DataFrame(datos)


# ============================================================
# DASHBOARD
# ============================================================

def crear_dashboard(
    lista,
    resumen,
    presupuesto,
    gasto_total
):

    sobrante = (
        presupuesto - gasto_total
    )

    fig, graficos = plt.subplots(
        2,
        2,
        figsize=(14, 9)
    )

    fig.suptitle(
        "Dashboard - Compra de Galletitas",
        fontsize=18,
        fontweight="bold"
    )

    # ========================================================
    # GRÁFICO 1
    # OBJETIVO VS RESULTADO REAL
    # ========================================================

    ax1 = graficos[0, 0]

    posiciones = range(
        len(resumen)
    )

    ancho = 0.35

    ax1.bar(
        [
            x - ancho / 2
            for x in posiciones
        ],
        resumen["porcentaje_objetivo"],
        width=ancho,
        label="Objetivo"
    )

    ax1.bar(
        [
            x + ancho / 2
            for x in posiciones
        ],
        resumen["porcentaje_real"],
        width=ancho,
        label="Real"
    )

    ax1.set_title(
        "Distribución objetivo vs real"
    )

    ax1.set_ylabel(
        "Porcentaje (%)"
    )

    ax1.set_xticks(
        list(posiciones)
    )

    ax1.set_xticklabels(
        resumen["categoria"]
    )

    ax1.legend()

    # ========================================================
    # GRÁFICO 2
    # DISTRIBUCIÓN DEL GASTO
    # ========================================================

    ax2 = graficos[0, 1]

    gastos_validos = resumen[
        resumen["gasto"] > 0
    ]

    ax2.pie(
        gastos_validos["gasto"],
        labels=gastos_validos["categoria"],
        autopct="%1.1f%%",
        startangle=90
    )

    ax2.set_title(
        "Distribución del gasto"
    )

    # ========================================================
    # GRÁFICO 3
    # GASTO POR SUCURSAL
    # ========================================================

    ax3 = graficos[1, 0]

    gasto_sucursal = (
        lista
        .groupby("sucursal")["subtotal"]
        .sum()
        .sort_values(
            ascending=False
        )
    )

    gasto_sucursal.plot(
        kind="bar",
        ax=ax3
    )

    ax3.set_title(
        "Gasto por sucursal"
    )

    ax3.set_xlabel(
        "Sucursal"
    )

    ax3.set_ylabel(
        "Gasto ($)"
    )

    ax3.tick_params(
        axis="x",
        rotation=0
    )

    # ========================================================
    # GRÁFICO 4
    # PRESUPUESTO UTILIZADO
    # ========================================================

    ax4 = graficos[1, 1]

    ax4.bar(
        [
            "Gastado",
            "Sobrante"
        ],
        [
            gasto_total,
            sobrante
        ]
    )

    ax4.set_title(
        "Uso del presupuesto"
    )

    ax4.set_ylabel(
        "Pesos ($)"
    )

    porcentaje = (
        gasto_total
        / presupuesto
        * 100
    )

    ax4.text(
        0,
        gasto_total / 2,
        f"{porcentaje:.1f}%",
        ha="center",
        fontsize=12
    )

    # ========================================================
    # AJUSTAR Y GUARDAR
    # ========================================================

    plt.tight_layout(
        rect=[
            0,
            0,
            1,
            0.95
        ]
    )

    plt.savefig(
        "dashboard_galletitas.png",
        dpi=150
    )

    print(
        "\nDashboard guardado como: "
        "dashboard_galletitas.png"
    )

    plt.show()


# ============================================================
# GUARDAR RESULTADOS EN EXCEL
# ============================================================

def guardar_excel(
    lista,
    resumen
):

    archivo_salida = (
        "lista_compra_galletitas.xlsx"
    )

    with pd.ExcelWriter(
        archivo_salida
    ) as writer:

        lista.to_excel(
            writer,
            sheet_name="Lista de compra",
            index=False
        )

        resumen.to_excel(
            writer,
            sheet_name="Resumen",
            index=False
        )

    print(
        f"\nExcel generado: "
        f"{archivo_salida}"
    )


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():

    # 1. Cargar datos
    catalogo = cargar_datos()

    # 2. Pedir presupuesto
    presupuesto = pedir_presupuesto()

    # 3. Crear lista
    lista, gasto_total = generar_lista(
        catalogo,
        presupuesto
    )

    # 4. Mostrar resultados
    mostrar_lista(
        lista,
        presupuesto,
        gasto_total
    )

    if lista.empty:
        return

    # 5. Crear resumen
    resumen = crear_resumen(
        lista,
        presupuesto,
        gasto_total
    )

    # 6. Guardar Excel
    guardar_excel(
        lista,
        resumen
    )

    # 7. Crear dashboard
    crear_dashboard(
        lista,
        resumen,
        presupuesto,
        gasto_total
    )

    print("\n" + "=" * 55)
    print("            PROCESO FINALIZADO")
    print("=" * 55)


# ============================================================
# EJECUTAR PROGRAMA
# ============================================================

if __name__ == "__main__":
    main()
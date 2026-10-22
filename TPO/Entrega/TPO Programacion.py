import pandas as pd
import matplotlib.pyplot as plt
import time

ARCHIVO = "galletitas_mariano_max.xlsx"
PRECIO_MINIMO = 100
PORCENTAJES = {
    "DULCES": 0.40,
    "SALADAS": 0.30,
    "RELLENAS": 0.20,
    "OBLEAS": 0.10
}

def pesos(valor):
    texto = f"{valor:,.2f}"
    return "$" + texto.replace(",", "X").replace(".", ",").replace("X", ".")

def clasificar(descripcion):
    texto = str(descripcion).upper()
    saladas = [
        "CRACKER", "CRAKERS", "CRIOLLITAS", "MARINERA", "SALVADO",
        "CLUB SOCIAL", "PORTEÑITAS", "KESITAS", "SEMOLADAS", "DORADITAS",
        "HOGAREÑAS", "GRANAGUA", "SIN SAL", "S/SAL", "SESAMO", "AMARAN"
    ]
    rellenas = [
        "RELLENA", "RELLENAS", "RELL", "SANDWICH",
        "SONRISAS", "BRIGITTE", "DALE BLACK"
    ]
    obleas = ["OBLEA", "OPERA", "BLOCK"]
    if any(p in texto for p in saladas):
        return "SALADAS"
    if any(p in texto for p in rellenas):
        return "RELLENAS"
    if any(p in texto for p in obleas):
        return "OBLEAS"
    return "DULCES"

def cargar_datos():
    print("\n" + "=" * 58)
    print("        SISTEMA DE COMPRA DE GALLETITAS")
    print("=" * 58)
    pasos = [
        "Buscando archivo...",
        "Cargando datos...",
        "Limpiando registros...",
        "Clasificando productos...",
        "Comparando precios entre sucursales..."
    ]
    for i, paso in enumerate(pasos, 1):
        print(f"[{i}/5] {paso}")
        time.sleep(0.35)
    try:
        datos = pd.read_excel(ARCHIVO)
    except FileNotFoundError:
        print(f"\nERROR: no se encontro {ARCHIVO}")
        raise SystemExit
    columnas = ["sucursal", "descripcion", "marca", "precio"]
    faltantes = [c for c in columnas if c not in datos.columns]
    if faltantes:
        print("ERROR. Faltan columnas:", faltantes)
        raise SystemExit
    filas_originales = len(datos)
    datos = datos.drop_duplicates(
        subset=["sucursal", "descripcion", "marca", "precio"]
    ).copy()
    datos["precio"] = pd.to_numeric(datos["precio"], errors="coerce")
    datos = datos.dropna(subset=columnas)
    datos = datos[datos["precio"] >= PRECIO_MINIMO].copy()
    datos["categoria"] = datos["descripcion"].apply(clasificar)
    indices = datos.groupby(
        ["descripcion", "marca"]
    )["precio"].idxmin()
    catalogo = datos.loc[indices].copy()
    catalogo = catalogo.sort_values(
        ["categoria", "precio"]
    ).reset_index(drop=True)
    print("\n" + "=" * 58)
    print("                    DATOS LISTOS")
    print("=" * 58)
    print(f"Filas originales: {filas_originales}")
    print(f"Productos validos: {len(catalogo)}")
    print("\nProductos por categoria:")
    conteo = catalogo["categoria"].value_counts()
    for categoria in PORCENTAJES:
        print(f"{categoria:<10}: {conteo.get(categoria, 0)}")
    return catalogo

def pedir_presupuesto():
    while True:
        try:
            monto = float(
                input(
                    "\nIngrese el presupuesto disponible: $"
                )
            )
            if monto > 0:
                return monto
            print("El presupuesto debe ser mayor que cero.")
        except ValueError:
            print("Ingrese solamente numeros.")

def comprar_categoria(
    catalogo,
    categoria,
    limite,
    cantidades
):
    productos = catalogo[
        catalogo["categoria"] == categoria
    ].copy()
    gasto = 0
    while True:
        disponible = limite - gasto
        candidatos = productos[
            productos["precio"] <= disponible
        ].copy()
        if candidatos.empty:
            break
        candidatos["cantidad_actual"] = (
            candidatos.index.map(
                lambda i: cantidades.get(i, 0)
            )
        )
        candidatos = candidatos.sort_values(
            ["cantidad_actual", "precio"]
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

def completar_presupuesto(
    catalogo,
    presupuesto,
    cantidades
):
    gasto = sum(
        catalogo.loc[i, "precio"] * cantidad
        for i, cantidad in cantidades.items()
    )
    while True:
        sobrante = presupuesto - gasto
        candidatos = catalogo[
            catalogo["precio"] <= sobrante
        ]
        if candidatos.empty:
            break
        elegido = candidatos[
            "precio"
        ].idxmax()
        precio = catalogo.loc[
            elegido,
            "precio"
        ]
        cantidades[elegido] = (
            cantidades.get(elegido, 0) + 1
        )
        gasto += precio
    return gasto

def generar_lista(
    catalogo,
    presupuesto
):
    cantidades = {}
    print("\n" + "=" * 58)
    print("                 GENERANDO COMPRA")
    print("=" * 58)
    for categoria, porcentaje in PORCENTAJES.items():
        limite = presupuesto * porcentaje
        print(
            f"{categoria:<10} "
            f"{porcentaje * 100:.0f}% "
            f"= {pesos(limite)}"
        )
        comprar_categoria(
            catalogo,
            categoria,
            limite,
            cantidades
        )
        time.sleep(0.2)
    print(
        "\nAprovechando el dinero restante..."
    )
    gasto_total = completar_presupuesto(
        catalogo,
        presupuesto,
        cantidades
    )
    filas = []
    for indice, cantidad in cantidades.items():
        producto = catalogo.loc[indice]
        precio = producto["precio"]
        filas.append({
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
                precio * cantidad
        })
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
        lista = lista.sort_values(
            ["orden", "subtotal"],
            ascending=[True, False]
        )
        lista = lista.drop(
            columns="orden"
        ).reset_index(drop=True)
    return lista, gasto_total

def mostrar_resultados(
    lista,
    presupuesto,
    gasto_total
):
    if lista.empty:
        print(
            "\nEl presupuesto no alcanza "
            "para ningun producto."
        )
        return
    print("\n" + "=" * 78)
    print("                     LISTA DE COMPRA")
    print("=" * 78)
    for categoria in PORCENTAJES:
        productos = lista[
            lista["categoria"] == categoria
        ].copy()
        if productos.empty:
            continue
        gasto = productos[
            "subtotal"
        ].sum()
        porcentaje = (
            gasto
            / gasto_total
            * 100
        )
        print(
            f"\n--- {categoria} ---"
        )
        print(
            f"Gasto: {pesos(gasto)} | "
            f"Participacion real: "
            f"{porcentaje:.1f}%"
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
            tabla.to_string(
                index=False
            )
        )
    sobrante = (
        presupuesto
        - gasto_total
    )
    total_paquetes = (
        lista["cantidad"]
        .sum()
    )
    por_sucursal = (
        lista
        .groupby("sucursal")["cantidad"]
        .sum()
        .sort_values(
            ascending=False
        )
    )
    print("\n" + "=" * 58)
    print("                      RESUMEN")
    print("=" * 58)
    print(
        f"Presupuesto: "
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
        f"{gasto_total / presupuesto * 100:.2f}%"
    )
    print(
        f"Productos diferentes: "
        f"{len(lista)}"
    )
    print(
        f"Paquetes totales: "
        f"{total_paquetes}"
    )
    print(
        f"Sucursales utilizadas: "
        f"{lista['sucursal'].nunique()}"
    )
    print(
        "\nPAQUETES POR SUCURSAL"
    )
    for sucursal, cantidad in por_sucursal.items():
        porcentaje = (
            cantidad
            / total_paquetes
            * 100
        )
        print(
            f"Sucursal {sucursal}: "
            f"{cantidad} paquetes "
            f"({porcentaje:.1f}%)"
        )
    print(
        f"\nSucursal con mayor cantidad "
        f"asignada: {por_sucursal.idxmax()}"
    )

def crear_resumen(
    lista,
    gasto_total
):
    filas = []
    for categoria, objetivo in PORCENTAJES.items():
        productos = lista[
            lista["categoria"] == categoria
        ]
        gasto = productos[
            "subtotal"
        ].sum()
        unidades = productos[
            "cantidad"
        ].sum()
        real = (
            gasto / gasto_total * 100
            if gasto_total
            else 0
        )
        filas.append({
            "categoria":
                categoria,
            "objetivo":
                objetivo * 100,
            "real":
                real,
            "gasto":
                gasto,
            "unidades":
                unidades
        })
    return pd.DataFrame(
        filas
    )

def resumen_sucursales(lista):
    return (
        lista
        .groupby("sucursal")
        .agg(
            paquetes=(
                "cantidad",
                "sum"
            ),
            gasto=(
                "subtotal",
                "sum"
            ),
            productos=(
                "descripcion",
                "nunique"
            )
        )
        .reset_index()
    )

def guardar_excel(
    lista,
    resumen,
    sucursales
):
    with pd.ExcelWriter(
        "lista_compra_galletitas.xlsx"
    ) as writer:
        lista.to_excel(
            writer,
            sheet_name="Lista",
            index=False
        )
        resumen.to_excel(
            writer,
            sheet_name="Categorias",
            index=False
        )
        sucursales.to_excel(
            writer,
            sheet_name="Sucursales",
            index=False
        )
    print(
        "\nExcel generado: "
        "lista_compra_galletitas.xlsx"
    )

def etiquetas_barras(
    eje,
    decimales=0,
    sufijo=""
):
    for contenedor in eje.containers:
        valores = [
            f"{barra.get_height():.{decimales}f}"
            f"{sufijo}"
            for barra in contenedor
        ]
        eje.bar_label(
            contenedor,
            labels=valores,
            padding=3,
            fontsize=9
        )

def crear_dashboard(
    lista,
    resumen,
    sucursales,
    presupuesto,
    gasto_total
):
    plt.rcParams.update({
        "font.size": 10,
        "axes.titlesize": 12,
        "axes.labelsize": 10
    })
    sobrante = (
        presupuesto
        - gasto_total
    )
    fig, ax = plt.subplots(
        2,
        3,
        figsize=(17, 9)
    )
    fig.suptitle(
        "Dashboard de Compra de Galletitas",
        fontsize=20,
        fontweight="bold"
    )

    # GRAFICO 1
    # Objetivo vs real con barras lado a lado

    posiciones = list(
        range(len(resumen))
    )
    ancho = 0.36

    ax[0, 0].bar(
        [
            x - ancho / 2
            for x in posiciones
        ],
        resumen["objetivo"],
        width=ancho,
        label="Objetivo"
    )

    ax[0, 0].bar(
        [
            x + ancho / 2
            for x in posiciones
        ],
        resumen["real"],
        width=ancho,
        label="Real"
    )

    ax[0, 0].set_title(
        "Distribucion objetivo vs real",
        fontweight="bold"
    )

    ax[0, 0].set_ylabel(
        "Porcentaje del gasto (%)"
    )

    ax[0, 0].set_xticks(
        posiciones
    )

    ax[0, 0].set_xticklabels(
        resumen["categoria"]
    )

    ax[0, 0].grid(
        axis="y",
        alpha=0.25
    )

    ax[0, 0].legend()

    etiquetas_barras(
        ax[0, 0],
        1,
        "%"
    )

    # GRAFICO 2
    # Distribucion del gasto

    pie = resumen[
        resumen["gasto"] > 0
    ]

    ax[0, 1].pie(
        pie["gasto"],
        labels=pie["categoria"],
        autopct="%1.1f%%",
        startangle=90
    )

    ax[0, 1].set_title(
        "Distribucion del gasto",
        fontweight="bold"
    )

    # GRAFICO 3
    # Cantidad de paquetes por categoria

    ax[0, 2].bar(
        resumen["categoria"],
        resumen["unidades"]
    )

    ax[0, 2].set_title(
        "Paquetes por categoria",
        fontweight="bold"
    )

    ax[0, 2].set_ylabel(
        "Cantidad de paquetes"
    )

    ax[0, 2].grid(
        axis="y",
        alpha=0.25
    )

    etiquetas_barras(
        ax[0, 2]
    )

    # GRAFICO 4
    # Gasto total por sucursal

    suc_gasto = (
        sucursales
        .sort_values(
            "gasto",
            ascending=False
        )
    )

    ax[1, 0].bar(
        suc_gasto["sucursal"]
        .astype(str),
        suc_gasto["gasto"]
    )

    ax[1, 0].set_title(
        "Gasto por sucursal",
        fontweight="bold"
    )

    ax[1, 0].set_xlabel(
        "Sucursal"
    )

    ax[1, 0].set_ylabel(
        "Pesos"
    )

    ax[1, 0].grid(
        axis="y",
        alpha=0.25
    )

    for contenedor in ax[1, 0].containers:
        valores = [
            pesos(
                barra.get_height()
            ).replace(",00", "")
            for barra in contenedor
        ]

        ax[1, 0].bar_label(
            contenedor,
            labels=valores,
            padding=3,
            fontsize=8,
            rotation=15
        )

    # GRAFICO 5
    # Cantidad de paquetes por sucursal

    suc_paquetes = (
        sucursales
        .sort_values(
            "paquetes",
            ascending=False
        )
    )

    ax[1, 1].bar(
        suc_paquetes["sucursal"]
        .astype(str),
        suc_paquetes["paquetes"]
    )

    ax[1, 1].set_title(
        "Paquetes asignados por sucursal",
        fontweight="bold"
    )

    ax[1, 1].set_xlabel(
        "Sucursal"
    )

    ax[1, 1].set_ylabel(
        "Cantidad de paquetes"
    )

    ax[1, 1].grid(
        axis="y",
        alpha=0.25
    )

    total = (
        suc_paquetes["paquetes"]
        .sum()
    )

    for contenedor in ax[1, 1].containers:
        valores = []

        for barra in contenedor:
            cantidad = (
                barra.get_height()
            )

            porcentaje = (
                cantidad
                / total
                * 100
            )

            valores.append(
                f"{int(cantidad)} "
                f"({porcentaje:.1f}%)"
            )

        ax[1, 1].bar_label(
            contenedor,
            labels=valores,
            padding=3,
            fontsize=9
        )

    # GRAFICO 6
    # Aprovechamiento del presupuesto

    ax[1, 2].pie(
        [
            gasto_total,
            sobrante
        ],
        labels=[
            "Gastado",
            "Sobrante"
        ],
        autopct="%1.1f%%",
        startangle=90,
        wedgeprops={
            "width": 0.42
        }
    )

    ax[1, 2].set_title(
        "Aprovechamiento del presupuesto",
        fontweight="bold"
    )

    ax[1, 2].text(
        0,
        0,
        f"{gasto_total / presupuesto * 100:.1f}%\n"
        f"usado",
        ha="center",
        va="center",
        fontsize=13,
        fontweight="bold"
    )

    plt.tight_layout(
        rect=[
            0,
            0,
            1,
            0.94
        ]
    )

    plt.savefig(
        "dashboard_galletitas.png",
        dpi=180,
        bbox_inches="tight"
    )

    print(
        "Dashboard generado: "
        "dashboard_galletitas.png"
    )

    plt.show()

def main():
    catalogo = cargar_datos()

    presupuesto = (
        pedir_presupuesto()
    )

    lista, gasto_total = generar_lista(
        catalogo,
        presupuesto
    )

    mostrar_resultados(
        lista,
        presupuesto,
        gasto_total
    )

    if lista.empty:
        return

    resumen = crear_resumen(
        lista,
        gasto_total
    )

    sucursales = resumen_sucursales(
        lista
    )

    guardar_excel(
        lista,
        resumen,
        sucursales
    )

    crear_dashboard(
        lista,
        resumen,
        sucursales,
        presupuesto,
        gasto_total
    )

    print("\n" + "=" * 58)
    print("                 PROCESO TERMINADO")
    print("=" * 58)

if __name__ == "__main__":
    main()
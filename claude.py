import re
from collections import defaultdict
import numpy as np
import pandas as pd

pd.set_option("display.max_rows", None, "display.width", 200)

# ============================================================
# PARÁMETROS Y SUPUESTOS (todos modificables, aclararlos en el informe)
# ============================================================
ARCHIVO = "galletitas_mariano_max.xlsx"
PRECIO_SOSPECHOSO = 100      # precios menores se excluyen (Nora Skills, Persiotto)
MARKUP = 0.40                # solo informativo (estimación de ganancia)
PRESUPUESTO = 300000
PERFIL = "kiosco"            # "hogar" o "kiosco"

PERFILES = {
    "hogar": {
        "paquetes_semana": 5,                       # consumo total estimado
        "share": {"DULCE": 0.40, "SALADA": 0.35, "RELLENA": 0.15, "OBLEA": 0.10},
        "max_unidades": 3,                          # por producto
        "max_marca": 0.30,                          # % del presupuesto por marca
        "cap_atributos": {"sin_gluten": 0.05, "light": 0.15},
        "min_prod_categoria": 2,
    },
    "kiosco": {
        "paquetes_semana": 60,
        "share": {"DULCE": 0.45, "SALADA": 0.25, "RELLENA": 0.20, "OBLEA": 0.10},
        "max_unidades": 6,
        "max_marca": 0.35,
        "cap_atributos": {"sin_gluten": 0.05, "light": 0.10},
        "min_prod_categoria": 2,
    },
}

# Correcciones manuales de clasificación (se chequean primero)
CORRECCIONES = {"BRIGITTE": "RELLENA", "DALE BLACK": "RELLENA",
                "SESAMO Y AMARAN": "SALADA"}

# ============================================================
# FUNCIONES DE DATOS
# ============================================================
def normalizar(descripcion):
    d = descripcion.replace("*", "").upper()
    return re.sub(r"\s+", " ", d).strip()

def clasificar_categoria(d):
    """d: descripción ya normalizada."""
    for clave, cat in CORRECCIONES.items():
        if clave in d:
            return cat
    if any(k in d for k in ["OPERA", "BLOCK"]):
        return "OBLEA"
    if any(k in d for k in ["RELL", "SANDWICH", "SONRISAS"]):
        return "RELLENA"
    if any(k in d for k in ["CRACKER", "CRAKERS", "MARINERA", "CRIOLLITAS",
                            "SALVADO", "CLUB SOCIAL", "PORTEÑITAS", "KESITAS",
                            "SEMOLADAS", "DORADITAS", "MELBA", "HOGAREÑAS",
                            "GRANAGUA", "SIN SAL", "S/SAL"]):
        return "SALADA"
    return "DULCE"

def extraer_gramos(d):
    """Gramos totales del envase; NaN si la descripción viene cortada."""
    m = re.search(r"(\d+)\s*(?:UN)?\s*X\s*(\d+)", d)      # packs: 3 X 108
    if m:
        return float(m.group(1)) * float(m.group(2))
    m = re.search(r"(\d+)\s*(?:GRS?|G)\b", d)
    return float(m.group(1)) if m else np.nan

def cargar_datos(archivo):
    datos = pd.read_excel(archivo)                       # original intacto
    limpios = datos.drop_duplicates(
        subset=["sucursal", "descripcion", "marca", "precio"]).copy()
    return datos, limpios

def construir_catalogo(limpios):
    d = limpios[limpios["precio"] >= PRECIO_SOSPECHOSO].copy()
    d["descripcion_norm"] = d["descripcion"].map(normalizar)
    d["producto"] = d["descripcion_norm"]

    # Descripciones que en una misma sucursal tienen precios distintos = productos
    # distintos con la descripción cortada -> se separan en variantes por precio
    conf = d.groupby(["sucursal", "producto"])["precio"].nunique()
    ambiguos = conf[conf > 1].index.get_level_values("producto").unique()
    for p in ambiguos:
        m = d["producto"] == p
        precios = sorted(d.loc[m, "precio"].unique())
        mapa = {pr: f"{p} [VAR.{i + 1}]" for i, pr in enumerate(precios)}
        d.loc[m, "producto"] = d.loc[m, "precio"].map(mapa)
        print(f"Variantes creadas para: {p} -> {list(mapa.values())}")

    tabla = d.pivot_table(index="producto", columns="sucursal",
                          values="precio", aggfunc="min")

    cat = (d[["producto", "descripcion_norm", "marca"]]
           .drop_duplicates("producto").set_index("producto"))
    cat["categoria"] = cat["descripcion_norm"].map(clasificar_categoria)
    cat["light"] = cat["descripcion_norm"].str.contains("LIGHT")
    cat["sin_gluten"] = cat["descripcion_norm"].str.contains("S/GLUTEN")
    cat["gramos"] = cat["descripcion_norm"].map(extraer_gramos)
    cat["costo"] = tabla.min(axis=1)
    cat["precio_100g"] = cat["costo"] / cat["gramos"] * 100
    mediana = cat.groupby("categoria")["precio_100g"].transform("median")
    cat["valor"] = cat["precio_100g"].fillna(mediana).fillna(cat["costo"])
    return cat, tabla

# ============================================================
# ALGORITMO: LISTA EQUILIBRADA (greedy)
# ============================================================
def asignar_sucursales(productos, tabla, catalogo):
    """Cada producto se compra a su precio mínimo; entre las sucursales empatadas
    se elige la que cubre más productos pendientes (menos viajes)."""
    pendientes = set(productos)
    asignacion = {}
    while pendientes:
        ganados = {s: [p for p in pendientes
                       if tabla.at[p, s] == catalogo.at[p, "costo"]]
                   for s in tabla.columns}
        mejor = max(ganados, key=lambda s: (len(ganados[s]), -s))
        for p in ganados[mejor]:
            asignacion[p] = mejor
            pendientes.discard(p)
    return asignacion

def generar_lista(catalogo, tabla, presupuesto, perfil):
    share = perfil["share"]
    obj_cat = {c: presupuesto * s for c, s in share.items()}
    cap_marca = perfil["max_marca"] * presupuesto
    cap_attr = {a: f * presupuesto for a, f in perfil["cap_atributos"].items()}

    filas = catalogo.to_dict("index")
    por_cat = defaultdict(list)
    for p, r in filas.items():
        if r["categoria"] in share:
            por_cat[r["categoria"]].append(p)

    unidades = defaultdict(int)
    gasto_cat = defaultdict(float)
    gasto_marca = defaultdict(float)
    gasto_attr = defaultdict(float)

    def mejor_producto(cat, tope_cat):
        total = sum(gasto_cat.values())
        mejor, mejor_clave = None, None
        for p in por_cat[cat]:
            r = filas[p]
            c = r["costo"]
            if (unidades[p] >= perfil["max_unidades"]
                    or total + c > presupuesto
                    or gasto_cat[cat] + c > tope_cat
                    or gasto_marca[r["marca"]] + c > cap_marca
                    or any(r[a] and gasto_attr[a] + c > cap_attr[a]
                           for a in cap_attr)):
                continue
            clave = (unidades[p], r["valor"], p)   # 1° variedad, 2° mejor precio/100g
            if mejor_clave is None or clave < mejor_clave:
                mejor, mejor_clave = p, clave
        return mejor

    # Pasada 1: respeta el reparto por categoría. Pasada 2: usa el sobrante
    # permitiendo hasta 25% más en cada categoría.
    for factor in (1.0, 1.25):
        while True:
            orden = sorted(obj_cat, key=lambda c: gasto_cat[c] / obj_cat[c])
            elegido = None
            for cat in orden:                      # la categoría más atrasada primero
                p = mejor_producto(cat, obj_cat[cat] * factor)
                if p is not None:
                    elegido = (cat, p)
                    break
            if elegido is None:
                break
            cat, p = elegido
            r, c = filas[p], filas[p]["costo"]
            unidades[p] += 1
            gasto_cat[cat] += c
            gasto_marca[r["marca"]] += c
            for a in cap_attr:
                if r[a]:
                    gasto_attr[a] += c

    comprados = [p for p in unidades if unidades[p] > 0]
    if not comprados:
        raise ValueError("Presupuesto insuficiente para comprar algo.")

    lista = catalogo.loc[comprados, ["marca", "categoria"]].copy()
    lista["sucursal"] = pd.Series(asignar_sucursales(comprados, tabla, catalogo))
    lista["unidades"] = pd.Series(dict(unidades))
    lista["costo_unit"] = catalogo.loc[comprados, "costo"]
    lista["subtotal"] = lista["costo_unit"] * lista["unidades"]
    lista["precio_venta_unit"] = lista["costo_unit"] * (1 + MARKUP)
    lista["ganancia_est"] = (lista["precio_venta_unit"] - lista["costo_unit"]) * lista["unidades"]
    lista = lista.reset_index().sort_values(["sucursal", "categoria", "producto"])
    return lista.reset_index(drop=True)

# ============================================================
# VALIDACIÓN Y RESÚMENES
# ============================================================
def validar_lista(lista, presupuesto, perfil):
    problemas = []
    if lista["subtotal"].sum() > presupuesto:
        problemas.append("Se superó el presupuesto.")
    if lista["unidades"].max() > perfil["max_unidades"]:
        problemas.append("Hay productos por encima del máximo de unidades.")
    marcas = lista.groupby("marca")["subtotal"].sum() / presupuesto
    for m, v in marcas[marcas > perfil["max_marca"] + 1e-9].items():
        problemas.append(f"Marca {m} supera el máximo ({v:.1%}).")
    distintos = lista.groupby("categoria")["producto"].nunique()
    for c in perfil["share"]:
        if distintos.get(c, 0) < perfil["min_prod_categoria"]:
            problemas.append(f"Categoría {c} con menos de "
                             f"{perfil['min_prod_categoria']} productos distintos.")
    return problemas

def resumen_categoria(lista, perfil):
    share = perfil["share"]
    g = (lista.groupby("categoria")
              .agg(gasto=("subtotal", "sum"), productos=("producto", "nunique"),
                   unidades=("unidades", "sum"))
              .reindex(list(share)).fillna(0))
    g["pct_real"] = (g["gasto"] / g["gasto"].sum() * 100).round(1)
    g["pct_objetivo"] = [share[c] * 100 for c in g.index]
    g["semanas_cobertura"] = (g["unidades"] /
                              (perfil["paquetes_semana"] * pd.Series(share))).round(1)
    return g

def mostrar(lista, presupuesto, perfil):
    gasto = lista["subtotal"].sum()
    print(f"\nPresupuesto: ${presupuesto:,.0f} | Gasto: ${gasto:,.0f} "
          f"| Sobrante: ${presupuesto - gasto:,.0f}")
    print(f"Productos distintos: {len(lista)} | Unidades: {lista['unidades'].sum()} "
          f"| Sucursales a visitar: {lista['sucursal'].nunique()}")
    print(f"Ganancia estimada (markup {MARKUP:.0%}, informativa): "
          f"${lista['ganancia_est'].sum():,.0f}")
    print("\nEquilibrio por categoría (semanas_cobertura similares = lista balanceada):")
    print(resumen_categoria(lista, perfil))
    print("\nGasto por marca (% del presupuesto):")
    print((lista.groupby("marca")["subtotal"].sum() / presupuesto * 100)
          .round(1).sort_values(ascending=False))
    print("\nGasto por sucursal:")
    print(lista.groupby("sucursal").agg(productos=("producto", "count"),
                                        gasto=("subtotal", "sum")))
    problemas = validar_lista(lista, presupuesto, perfil)
    print("\nValidación:", "OK, cumple todas las restricciones." if not problemas
          else problemas)

# ============================================================
# EJECUCIÓN
# ============================================================
if __name__ == "__main__":
    datos, limpios = cargar_datos(ARCHIVO)
    catalogo, tabla = construir_catalogo(limpios)
    perfil = PERFILES[PERFIL]

    print("Productos válidos:", len(catalogo))
    print(catalogo["categoria"].value_counts())

    lista = generar_lista(catalogo, tabla, PRESUPUESTO, perfil)
    print(lista[["sucursal", "categoria", "producto", "unidades",
                 "costo_unit", "subtotal"]])
    mostrar(lista, PRESUPUESTO, perfil)

    # Cómo cambia la lista con el presupuesto
    print("\nSensibilidad al presupuesto:")
    for pres in (100000, 300000, 600000):
        l = generar_lista(catalogo, tabla, pres, perfil)
        cob = resumen_categoria(l, perfil)["semanas_cobertura"]
        print(f"${pres:>8,.0f}: {len(l):3d} productos, {l['unidades'].sum():3d} u., "
              f"{l['sucursal'].nunique()} sucursales, cobertura por categoría "
              f"{cob.min():.1f}-{cob.max():.1f} semanas")

    with pd.ExcelWriter("lista_de_compras.xlsx") as w:
        lista.to_excel(w, sheet_name="Lista", index=False)
        resumen_categoria(lista, perfil).to_excel(w, sheet_name="Por categoria")
    print("\nGenerado: lista_de_compras.xlsx")
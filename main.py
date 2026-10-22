import pandas as pd

datos = pd.read_excel("galletitas_mariano_max.xlsx")

print(datos.head())
print(datos.shape)
print(datos.columns)

datos.info()

print(datos.isnull().sum())

print(datos["sucursal"].unique())
print(datos["sucursal"].nunique())

print(datos["marca"].unique())
print(datos["marca"].nunique())

print("Precio mínimo:", datos["precio"].min())
print("Precio máximo:", datos["precio"].max())
print("Precio promedio:", datos["precio"].mean())

print("Cantidad de EAN diferentes:", datos["ean"].nunique())
print(datos["ean"].unique())

print("Cantidad de descripciones diferentes:", datos["descripcion"].nunique())

print(datos["descripcion"].value_counts().head(10))

print(datos["precio"].sort_values().head(10))
print(datos["precio"].sort_values(ascending=False).head(10))

print(datos.loc[datos["precio"].idxmin()])
print(datos.loc[datos["precio"].idxmax()])

print(datos["sucursal"].value_counts().sort_index()) #Paso 19 — ¿Cuántos productos hay por sucursal?

print(datos.groupby("sucursal")["descripcion"].nunique()) #Paso 20 — ¿Qué productos tenemos en cada sucursal?

print(datos.groupby("sucursal")["precio"].mean()) #Paso 21 — Comparar precios entre sucursales

print(datos.groupby("sucursal")["precio"].mean().sort_values()) #Paso 22 — Ordenar las sucursales por precio promedio

productos_marcas = datos.groupby("descripcion")["marca"].nunique()
print(productos_marcas.sort_values(ascending=False).head(20)) #Paso 23 — Comprobar descripción vs. marca

print(productos_marcas[productos_marcas > 1]) #Paso 24 — Ver si existen esos casos

print("Filas duplicadas:", datos.duplicated().sum())
print(datos[datos.duplicated()])

print(
    datos.duplicated(
        subset=["sucursal", "descripcion", "marca", "precio"]
    ).sum()
)

repetidos = datos.groupby(
    ["sucursal", "descripcion"]
).size()

print(repetidos.sort_values(ascending=False).head(20))

datos_limpios = datos.drop_duplicates(
    subset=["sucursal", "descripcion", "marca", "precio"]
) 

print("Filas originales:", len(datos))
print("Filas después de limpiar:", len(datos_limpios))

print(datos_limpios.duplicated(
    subset=["sucursal", "descripcion", "marca", "precio"]
).sum())

productos = datos_limpios["descripcion"].drop_duplicates()

for producto in productos:
    print(producto)
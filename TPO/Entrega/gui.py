"""Prototipo inicial de interfaz para explorar precios de galletitas con Tkinter.

Si está disponible, carga el archivo de referencia ubicado en sources/.
Para ejecutar: python prototipo_tkinter.py
"""
from pathlib import Path
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

BG = "#F4F6F1"
WHITE = "#FFFFFF"
INK = "#202820"
MUTED = "#788176"
GREEN = "#3D694D"
PALE_GREEN = "#E7F0E7"
LINE = "#E5E9E1"

SAMPLE = [
    (1, "GALLETITAS GRANIX LIGHT SIN SAL 360 GRS.", "Granix", 2459),
    (1, "GALLETITAS SONRISAS FRAMBUESA 3 X 108 GR", "Bagley", 3699),
    (1, "GALLETITAS FRUTIGRAN CHOCOLATE 255 GRS.", "Granix", 2669),
    (1, "GALLETITAS GRANIX SALVADO SIN SAL 230 GR", "Granix", 1459),
    (1, "GALLETITAS FRUTIGRAN CON CHIA Y LINO 240", "Granix", 2199),
    (1, "GALLETITAS LA PROVIDENCIA CRACKERS 505 G", "San Salvador", 2290),
    (1, "GALLETITAS PASEO CRACKER S/SAL 300 GR", "Molino Cañuelas", 1399),
    (1, "GALLETITAS MINI PAMELA LIMON 180GR.", "Firenze", 1149),
    (2, "GALLETITAS MINI PAMELA COCO 180GR.", "Firenze", 1149),
    (2, "GALLETITAS GRANAGUA SANDWICH 600GR.", "Granix", 3399),
    (3, "GALLETITAS CLASICAS DULCES 400 GRS.", "Bagley", 1899),
    (4, "GALLETITAS CRACKERS CLÁSICAS 300 GRS.", "Molino Cañuelas", 1599),
    (5, "GALLETITAS MINI COCO 180 GR.", "Firenze", 1299),
    (6, "GALLETITAS SURTIDAS 400 GRS.", "Wellington Food", 2099),
    (7, "GALLETITAS DE AGUA 3 X 100 GR.", "Bagley", 1799),
    (8, "GALLETITAS SALVADO 230 GR.", "Granix", 1549),
    (9, "GALLETITAS DULCES RELLENAS 300 GR.", "Grandiet", 2499),
    (10, "GALLETITAS CRACKERS SIN SAL 300 GR.", "Molino Cañuelas", 1699),
    (11, "GALLETITAS MINI VAINILLA 180 GR.", "Firenze", 1149),
]


def cargar_productos():
    """Lee la planilla sin modificarla; si falta openpyxl, usa registros de muestra."""
    archivo = Path(__file__).parent / "sources" / "galletitas_mariano_max.xlsx"
    if archivo.exists():
        try:
            from openpyxl import load_workbook
            libro = load_workbook(archivo, read_only=True, data_only=True)
            hoja = libro["galletitas"]
            filas = []
            for sucursal, _ean, descripcion, marca, precio in hoja.iter_rows(min_row=2, values_only=True):
                if descripcion and precio is not None:
                    filas.append({
                        "sucursal": int(sucursal),
                        "producto": str(descripcion).strip().title(),
                        "marca": str(marca or "Sin marca").strip(),
                        "precio": float(precio),
                    })
            libro.close()
            if filas:
                return filas, "Planilla cargada · galletitas_mariano_max.xlsx", True
        except Exception:
            pass
    filas = [dict(sucursal=s, producto=p.title(), marca=m, precio=float(pr)) for s, p, m, pr in SAMPLE]
    return filas, "Datos de muestra · no se encontró openpyxl o la planilla", False


def moneda(valor):
    return "$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".").removesuffix(",00")


class ExploradorPrecios:
    def __init__(self, raiz):
        self.raiz = raiz
        self.raiz.title("Miga | Explorador de precios")
        self.raiz.geometry("1220x790")
        self.raiz.minsize(980, 650)
        self.raiz.configure(bg=BG)
        self.productos, estado, self.datos_reales = cargar_productos()
        self.estado_archivo = estado
        self.resultados = list(self.productos)
        self.crear_estilos()
        self.crear_interfaz()
        self.aplicar_filtros()

    def crear_estilos(self):
        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass
        estilo.configure("Treeview", background=WHITE, fieldbackground=WHITE,
                         foreground=INK, rowheight=36, borderwidth=0,
                         font=("Segoe UI", 9))
        estilo.configure("Treeview.Heading", background="#F8F9F6", foreground=MUTED,
                         font=("Segoe UI", 8, "bold"), relief="flat", padding=(10, 11))
        estilo.map("Treeview", background=[("selected", PALE_GREEN)],
                   foreground=[("selected", INK)])
        estilo.configure("Vertical.TScrollbar", background="#E5E9E1", troughcolor=WHITE,
                         bordercolor=WHITE, arrowsize=12)

    def crear_interfaz(self):
        contenedor = tk.Frame(self.raiz, bg=BG)
        contenedor.pack(fill="both", expand=True)
        self.crear_barra_lateral(contenedor)
        contenido = tk.Frame(contenedor, bg=BG)
        contenido.pack(side="left", fill="both", expand=True, padx=(28, 32), pady=22)
        self.crear_encabezado(contenido)
        self.crear_filtros(contenido)
        self.crear_tarjetas(contenido)
        self.crear_paneles(contenido)
        self.crear_tabla(contenido)
        self.crear_aviso(contenido)

    def crear_barra_lateral(self, padre):
        barra = tk.Frame(padre, bg=WHITE, width=210, highlightbackground=LINE, highlightthickness=1)
        barra.pack(side="left", fill="y")
        barra.pack_propagate(False)
        logo = tk.Frame(barra, bg=WHITE)
        logo.pack(fill="x", padx=18, pady=(23, 32))
        tk.Label(logo, text="m", bg=GREEN, fg=WHITE, font=("Segoe UI", 16, "bold"),
                 width=2, height=1, pady=3).pack(side="left")
        tk.Label(logo, text="miga", bg=WHITE, fg=INK, font=("Segoe UI", 19, "bold")).pack(side="left", padx=(9, 4))
        tk.Label(logo, text="lab", bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).pack(side="left", pady=(9, 0))
        tk.Label(barra, text="ESPACIO DE TRABAJO", bg=WHITE, fg="#9AA197",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=20, pady=(0, 9))
        for texto, icono, activo in [("Resumen", "▦", True), ("Productos", "◷", False),
                                      ("Sucursales", "▤", False), ("Marcas", "⌁", False),
                                      ("Recolecciones", "↻", False)]:
            fila = tk.Frame(barra, bg=PALE_GREEN if activo else WHITE)
            fila.pack(fill="x", padx=10, pady=2)
            tk.Label(fila, text=icono, bg=fila["bg"], fg=GREEN if activo else MUTED,
                     font=("Segoe UI", 12), width=3, pady=10).pack(side="left")
            tk.Label(fila, text=texto, bg=fila["bg"], fg=GREEN if activo else "#697167",
                     font=("Segoe UI", 9, "bold" if activo else "normal")).pack(side="left")
        pie = tk.Frame(barra, bg=WHITE)
        pie.pack(side="bottom", fill="x", padx=15, pady=18)
        tk.Label(pie, text="DATOS", bg=WHITE, fg="#9AA197", font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(0, 8))
        tk.Label(pie, text="●  Datos cargados", bg=WHITE, fg=GREEN, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(pie, text=("●  Planilla cargada" if self.datos_reales else "●  Usando muestra de ejemplo"), bg=WHITE, fg=GREEN, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(pie, text="galletitas_mariano_max.xlsx", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8), wraplength=175).pack(anchor="w", padx=(13, 0), pady=(4, 12))
        tk.Label(pie, text="⚙  Configuración", bg=WHITE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w")

    def crear_encabezado(self, padre):
        top = tk.Frame(padre, bg=BG)
        top.pack(fill="x", pady=(0, 20))
        tk.Label(top, text="Inicio   /   Resumen", bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
        tk.Label(top, text="Vista general · Octubre 2026     ◉  M", bg=BG, fg=MUTED,
                 font=("Segoe UI", 9)).pack(side="right")
        tk.Label(padre, text="PANEL DE ANÁLISIS", bg=BG, fg=GREEN,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        linea = tk.Frame(padre, bg=BG)
        linea.pack(fill="x", pady=(4, 16))
        izquierda = tk.Frame(linea, bg=BG)
        izquierda.pack(side="left")
        tk.Label(izquierda, text="Precios de galletitas", bg=BG, fg=INK,
                 font=("Segoe UI", 22, "bold")).pack(anchor="w")
        tk.Label(izquierda, text="Una mirada simple a los productos relevados por sucursal.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
        tk.Button(linea, text="↓  Exportar CSV", command=self.exportar_csv,
                  bg=GREEN, fg=WHITE, activebackground="#31583F", activeforeground=WHITE,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=13, pady=9,
                  cursor="hand2").pack(side="right", anchor="s")

    def crear_filtros(self, padre):
        barra = tk.Frame(padre, bg=BG)
        barra.pack(fill="x", pady=(0, 13))
        busqueda = tk.Frame(barra, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        busqueda.pack(side="left", fill="x", expand=True, padx=(0, 8))
        tk.Label(busqueda, text="⌕", bg=WHITE, fg=MUTED, font=("Segoe UI", 14)).pack(side="left", padx=(10, 3))
        self.buscar_var = tk.StringVar()
        entrada = tk.Entry(busqueda, textvariable=self.buscar_var, relief="flat", bg=WHITE,
                           fg=INK, insertbackground=INK, font=("Segoe UI", 9))
        entrada.pack(side="left", fill="x", expand=True, ipady=10, padx=(2, 8))
        entrada.insert(0, "Buscar producto o marca…")
        entrada.configure(fg="#9AA197")
        entrada.bind("<FocusIn>", self.limpiar_placeholder)
        entrada.bind("<KeyRelease>", lambda _e: self.aplicar_filtros())
        self.sucursal_var = tk.StringVar(value="Todas las sucursales")
        self.marca_var = tk.StringVar(value="Todas las marcas")
        self.orden_var = tk.StringVar(value="Relevancia")
        self.combo_sucursal = ttk.Combobox(barra, textvariable=self.sucursal_var, state="readonly", width=19,
                                          values=["Todas las sucursales"] + [f"Sucursal {i}" for i in range(1, 12)])
        self.combo_sucursal.pack(side="left", padx=4, ipady=5)
        marcas = sorted({p["marca"] for p in self.productos})
        self.combo_marca = ttk.Combobox(barra, textvariable=self.marca_var, state="readonly", width=18,
                                       values=["Todas las marcas"] + marcas)
        self.combo_marca.pack(side="left", padx=4, ipady=5)
        self.combo_orden = ttk.Combobox(barra, textvariable=self.orden_var, state="readonly", width=18,
                                       values=["Relevancia", "Precio: menor a mayor", "Precio: mayor a menor"])
        self.combo_orden.pack(side="left", padx=(4, 0), ipady=5)
        for combo in (self.combo_sucursal, self.combo_marca, self.combo_orden):
            combo.bind("<<ComboboxSelected>>", lambda _e: self.aplicar_filtros())

    def limpiar_placeholder(self, _evento):
        entrada = _evento.widget
        if entrada.get() == "Buscar producto o marca…":
            entrada.delete(0, "end")
            entrada.configure(fg=INK)

    def crear_tarjetas(self, padre):
        marco = tk.Frame(padre, bg=BG)
        marco.pack(fill="x", pady=(0, 14))
        total = len(self.productos)
        promedio = sum(p["precio"] for p in self.productos) / total if total else 0
        minimo = min((p["precio"] for p in self.productos), default=0)
        marcas = len({p["marca"] for p in self.productos})
        sucursales = len({p["sucursal"] for p in self.productos})
        datos = [("Productos relevados", f"{total:,}".replace(",", "."), f"en {sucursales} sucursales", "▦"),
                 ("Precio promedio", moneda(promedio), "por producto relevado", "$"),
                 ("Precio más bajo", moneda(minimo), "entre los datos cargados", "↓"),
                 ("Marcas distintas", str(marcas), "identificadas en la planilla", "◉")]
        self.valores_tarjeta = []
        for titulo, valor, detalle, icono in datos:
            tarjeta = tk.Frame(marco, bg=WHITE, highlightbackground=LINE, highlightthickness=1, padx=13, pady=11)
            tarjeta.pack(side="left", fill="both", expand=True, padx=(0, 9))
            fila = tk.Frame(tarjeta, bg=WHITE)
            fila.pack(fill="x")
            tk.Label(fila, text=titulo, bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(side="left")
            tk.Label(fila, text=icono, bg=PALE_GREEN, fg=GREEN, font=("Segoe UI", 10, "bold"),
                     width=3, pady=3).pack(side="right")
            valor_label = tk.Label(tarjeta, text=valor, bg=WHITE, fg=INK, font=("Segoe UI", 17, "bold"))
            valor_label.pack(anchor="w", pady=(8, 0))
            tk.Label(tarjeta, text=detalle, bg=WHITE, fg="#9AA197", font=("Segoe UI", 8)).pack(anchor="w")
            self.valores_tarjeta.append(valor_label)

    def crear_paneles(self, padre):
        mitad = tk.Frame(padre, bg=BG)
        mitad.pack(fill="x", pady=(0, 14))
        panel_chart = tk.Frame(mitad, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        panel_chart.pack(side="left", fill="both", expand=True, padx=(0, 8))
        panel_brand = tk.Frame(mitad, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        panel_brand.pack(side="left", fill="both", expand=True, padx=(8, 0))
        for panel, titulo, subtitulo in [(panel_chart, "Registros por sucursal", "Cantidad de productos relevados"),
                                         (panel_brand, "Marcas con más productos", "Participación sobre el total")]:
            tk.Label(panel, text=titulo, bg=WHITE, fg=INK, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(12, 0))
            tk.Label(panel, text=subtitulo, bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=14, pady=(2, 0))
        self.grafico = tk.Canvas(panel_chart, height=135, bg=WHITE, highlightthickness=0)
        self.grafico.pack(fill="x", padx=10, pady=(2, 0))
        self.grafico.bind("<Configure>", lambda _e: self.dibujar_grafico())
        self.lista_marcas = tk.Frame(panel_brand, bg=WHITE)
        self.lista_marcas.pack(fill="both", expand=True, padx=14, pady=(7, 10))
        self.dibujar_marcas()

    def dibujar_grafico(self):
        if not hasattr(self, "grafico"):
            return
        c = self.grafico
        c.delete("all")
        ancho = max(c.winfo_width(), 350)
        alto = 135
        conteos = [sum(1 for p in self.productos if p["sucursal"] == i) for i in range(1, 12)]
        maximo = max(conteos, default=1) or 1
        for proporcion in (0.32, 0.65):
            y = int(alto * proporcion)
            c.create_line(16, y, ancho - 8, y, fill="#EEF0EB", dash=(2, 3))
        margen = 22
        espacio = (ancho - margen * 2) / 11
        for i, valor in enumerate(conteos):
            x0 = margen + i * espacio + 4
            x1 = margen + (i + 1) * espacio - 4
            altura = max(8, (valor / maximo) * 75)
            color = "#85A98A" if i == 0 else "#DDE9DC"
            c.create_rectangle(x0, 96 - altura, x1, 96, fill=color, outline="")
            c.create_text((x0 + x1) / 2, 112, text=str(i + 1), fill=MUTED,
                          font=("Segoe UI", 7))
        c.create_text(17, 13, text=f"{maximo} reg.", fill=MUTED, anchor="w", font=("Segoe UI", 7))

    def dibujar_marcas(self):
        for widget in self.lista_marcas.winfo_children():
            widget.destroy()
        conteos = {}
        for p in self.productos:
            conteos[p["marca"]] = conteos.get(p["marca"], 0) + 1
        principales = sorted(conteos.items(), key=lambda x: x[1], reverse=True)[:5]
        maximo = max((n for _, n in principales), default=1)
        for nombre, cantidad in principales:
            fila = tk.Frame(self.lista_marcas, bg=WHITE)
            fila.pack(fill="x", pady=5)
            tk.Label(fila, text=nombre[:20], bg=WHITE, fg="#596055", width=19, anchor="w",
                     font=("Segoe UI", 8)).pack(side="left")
            base = tk.Frame(fila, bg="#F0F2ED", height=6)
            base.pack(side="left", fill="x", expand=True, padx=7)
            base.pack_propagate(False)
            tk.Frame(base, bg="#91AE8B", width=max(4, int(95 * cantidad / maximo)).__int__()).pack(side="left", fill="y")
            tk.Label(fila, text=str(cantidad), bg=WHITE, fg=MUTED, width=4,
                     font=("Segoe UI", 8)).pack(side="right")

    def crear_tabla(self, padre):
        panel = tk.Frame(padre, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        panel.pack(fill="both", expand=True)
        cabeza = tk.Frame(panel, bg=WHITE)
        cabeza.pack(fill="x", padx=14, pady=(11, 9))
        tk.Label(cabeza, text="Productos relevados", bg=WHITE, fg=INK,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(cabeza, text="Explorá los precios cargados en cada sucursal", bg=WHITE,
                 fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))
        columnas = ("producto", "marca", "sucursal", "precio")
        self.tabla = ttk.Treeview(panel, columns=columnas, show="headings", height=6)
        self.tabla.heading("producto", text="PRODUCTO")
        self.tabla.heading("marca", text="MARCA")
        self.tabla.heading("sucursal", text="SUCURSAL")
        self.tabla.heading("precio", text="PRECIO")
        self.tabla.column("producto", width=460, minwidth=220, anchor="w")
        self.tabla.column("marca", width=180, minwidth=110, anchor="w")
        self.tabla.column("sucursal", width=130, minwidth=100, anchor="w")
        self.tabla.column("precio", width=130, minwidth=100, anchor="e")
        self.tabla.pack(fill="both", expand=True, padx=8)
        pie = tk.Frame(panel, bg=WHITE)
        pie.pack(fill="x", padx=14, pady=9)
        self.resultado_label = tk.Label(pie, text="", bg=WHITE, fg=MUTED, font=("Segoe UI", 8))
        self.resultado_label.pack(side="left")
        tk.Button(pie, text="Limpiar filtros", command=self.limpiar_filtros,
                  bg=WHITE, fg="#596055", relief="solid", bd=1, font=("Segoe UI", 8),
                  padx=9, pady=5, cursor="hand2").pack(side="right")

    def crear_aviso(self, padre):
        aviso = tk.Frame(padre, bg="#FFF8EC", highlightbackground="#F0DFC7", highlightthickness=1)
        aviso.pack(fill="x", pady=(10, 0))
        tk.Label(aviso, text="ⓘ", bg="#FFF8EC", fg="#80633C", font=("Segoe UI", 11)).pack(side="left", padx=(10, 5), pady=7)
        tk.Label(aviso, text="Revisar calidad de datos: el campo EAN aparece con el mismo valor en todos los registros.",
                 bg="#FFF8EC", fg="#80633C", font=("Segoe UI", 8)).pack(side="left", pady=8)

    def aplicar_filtros(self):
        busqueda = self.buscar_var.get().strip().lower()
        if busqueda == "buscar producto o marca…":
            busqueda = ""
        sucursal = self.sucursal_var.get()
        marca = self.marca_var.get()
        orden = self.orden_var.get()
        encontrados = [p for p in self.productos
                       if (not busqueda or busqueda in p["producto"].lower() or busqueda in p["marca"].lower())
                       and (sucursal == "Todas las sucursales" or sucursal == f"Sucursal {p['sucursal']}")
                       and (marca == "Todas las marcas" or marca == p["marca"])]
        if orden == "Precio: menor a mayor":
            encontrados.sort(key=lambda p: p["precio"])
        elif orden == "Precio: mayor a menor":
            encontrados.sort(key=lambda p: p["precio"], reverse=True)
        self.resultados = encontrados
        for iid in self.tabla.get_children():
            self.tabla.delete(iid)
        for p in encontrados[:100]:
            self.tabla.insert("", "end", values=(p["producto"], p["marca"],
                              f"Sucursal {p['sucursal']}", moneda(p["precio"])))
        texto = f"Mostrando {min(100, len(encontrados))} de {len(encontrados)} productos"
        if not self.datos_reales:
            texto += " · muestra de ejemplo"
        self.resultado_label.configure(text=texto)

    def limpiar_filtros(self):
        self.buscar_var.set("")
        self.sucursal_var.set("Todas las sucursales")
        self.marca_var.set("Todas las marcas")
        self.orden_var.set("Relevancia")
        self.aplicar_filtros()

    def exportar_csv(self):
        ruta = filedialog.asksaveasfilename(title="Guardar productos", defaultextension=".csv",
                                            filetypes=[("Archivo CSV", "*.csv")],
                                            initialfile="productos_galletitas.csv")
        if not ruta:
            return
        try:
            with open(ruta, "w", newline="", encoding="utf-8-sig") as archivo:
                escritor = csv.writer(archivo)
                escritor.writerow(["Producto", "Marca", "Sucursal", "Precio"])
                for p in self.resultados:
                    escritor.writerow([p["producto"], p["marca"], p["sucursal"], p["precio"]])
            messagebox.showinfo("Exportación lista", f"Se guardaron {len(self.resultados)} productos.")
        except OSError as error:
            messagebox.showerror("No se pudo exportar", str(error))


if __name__ == "__main__":
    ventana = tk.Tk()
    aplicacion = ExploradorPrecios(ventana)
    ventana.mainloop()
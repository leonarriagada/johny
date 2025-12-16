import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime
import glob
import sys

# Configurar el tema del sistema
try:
    from ttkthemes import ThemedStyle

    USE_CUSTOM_THEME = True
except ImportError:
    USE_CUSTOM_THEME = False


class FiltradorMultiArchivosGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(
            "🦷 VidaSalud Dental - Filtrador Multi-Archivos | © Jonathan Fuentes Toledo"
        )
        self.root.geometry("1000x750")

        # Configurar tema y estilo personalizado
        if USE_CUSTOM_THEME:
            self.style = ThemedStyle(self.root)
            self.style.set_theme("arc")
        else:
            self.style = ttk.Style()

        # Configurar estilos personalizados
        self.style.configure(
            "Custom.TButton", padding=10, relief="flat", font=("Arial", 10, "bold")
        )
        self.style.map(
            "Custom.TButton",
            background=[("active", "#5BA0B4"), ("!active", "#4A90A4")],
            foreground=[("active", "#FFFFFF"), ("!active", "#FFFFFF")],
        )

        # Colores corporativos de Vidasalud Dental
        self.colors = {
            "primary": "#4A90A4",  # Verde agua principal
            "secondary": "#7FB3C7",  # Verde agua claro
            "primary_dark": "#2E5A6B",  # Verde agua oscuro para botones principales
            "accent": "#8B7355",  # Café claro
            "light": "#F5F5F5",  # Blanco grisáceo
            "white": "#FFFFFF",  # Blanco puro
            "dark": "#1A1A1A",  # Negro suave para texto principal
            "dark_gray": "#2C2C2C",  # Gris oscuro para texto secundario
            "success": "#27AE60",  # Verde para éxito
            "warning": "#F39C12",  # Naranja para advertencias
            "error": "#E74C3C",  # Rojo para errores
            "gradient_start": "#E8F4F8",  # Inicio del gradiente
            "gradient_end": "#F0F8FA",  # Fin del gradiente
            "border": "#D1E7DD",  # Color de bordes suaves
            "hover": "#5BA0B4",  # Color hover para botones
            "shadow": "#E0E0E0",  # Color de sombras
            "button_gradient_start": "#4A90A4",  # Gradiente para botones
            "button_gradient_end": "#2E5A6B",  # Gradiente para botones
        }

        self.root.configure(bg=self.colors["light"])

        # Variables
        self.archivos_excel = []
        self.archivo_seleccionado = None
        self.df = None
        self.prestaciones = []
        self.resultado_filtrado = None
        self.columnas_seleccionadas = []  # Columnas que se mostrarán en los resultados

        # Cargar archivos disponibles
        self.cargar_archivos_disponibles()

        # Crear interfaz
        self.crear_interfaz()

        # Habilitar botón de guardar desde el inicio
        if hasattr(self, "save_button"):
            self.save_button.config(state="normal")

    def get_base_path(self):
        """Retorna la ruta base correcta tanto para script como para ejecutable"""
        if getattr(sys, "frozen", False):
            # Si es un ejecutable (PyInstaller)
            # En macOS .app, sys.executable está dentro del bundle
            return os.path.dirname(sys.executable)
        else:
            # Si es script normal
            return os.path.dirname(os.path.abspath(__file__))

    def cargar_archivos_disponibles(self):
        """Carga todos los archivos Excel de la carpeta archivos_excel"""
        base_path = self.get_base_path()
        carpeta_archivos = os.path.join(base_path, "archivos_excel")

        if not os.path.exists(carpeta_archivos):
            try:
                os.makedirs(carpeta_archivos)
                messagebox.showinfo(
                    "📁 Carpeta Creada",
                    f"Se creó la carpeta '{carpeta_archivos}'\nColoca aquí todos tus archivos Excel de datos.",
                )
            except OSError as e:
                messagebox.showerror(
                    "❌ Error", f"No se pudo crear la carpeta:\n{str(e)}"
                )
            return

        # Buscar todos los archivos Excel
        patron = os.path.join(carpeta_archivos, "*.xlsx")
        self.archivos_excel = glob.glob(patron)

        if not self.archivos_excel:
            messagebox.showwarning(
                "⚠️ Sin Archivos",
                f"No se encontraron archivos Excel en la carpeta '{carpeta_archivos}'\nColoca aquí tus archivos .xlsx",
            )

    def _interpolate_color(self, color1, color2, factor):
        """Interpola entre dos colores hexadecimales."""

        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip("#")
            return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))

        def rgb_to_hex(rgb):
            return "#{:02x}{:02x}{:02x}".format(*rgb)

        c1 = hex_to_rgb(color1)
        c2 = hex_to_rgb(color2)

        rgb = tuple(int(c1[i] + (c2[i] - c1[i]) * factor) for i in range(3))
        return rgb_to_hex(rgb)

    def crear_interfaz(self):
        """Crea la interfaz gráfica con selector de archivos"""
        # Configurar el estilo de la ventana
        self.root.configure(bg=self.colors["gradient_start"])

        # Header principal con gradiente y sombra
        header_frame = tk.Frame(self.root, height=120)
        header_frame.pack(fill="x", padx=0, pady=0)
        header_frame.pack_propagate(False)

        # Canvas para crear el gradiente en el header
        header_canvas = tk.Canvas(header_frame, height=120, highlightthickness=0)
        header_canvas.pack(fill="x", expand=True)

        # Crear gradiente
        header_canvas.create_rectangle(
            0, 0, 2000, 120, fill=self.colors["primary_dark"], outline=""
        )
        for i in range(60):
            color = self._interpolate_color(
                self.colors["primary_dark"], self.colors["primary"], i / 60
            )
            header_canvas.create_line(0, i * 2, 2000, i * 2, fill=color)

        # Título principal con efecto de sombra
        title_shadow = tk.Label(
            header_frame,
            text="🦷 Vidasalud Dental",
            font=("Arial", 24, "bold"),
            fg=self.colors["shadow"],
            bg=self.colors["primary_dark"],
        )
        title_shadow.place(relx=0.5, rely=0.3, anchor="center", x=2, y=2)

        title_label = tk.Label(
            header_frame,
            text="🦷 Vidasalud Dental",
            font=("Arial", 24, "bold"),
            fg=self.colors["white"],
            bg=self.colors["primary_dark"],
        )
        title_label.place(relx=0.5, rely=0.3, anchor="center")

        # Subtítulo con efecto de transparencia
        subtitle_label = tk.Label(
            header_frame,
            text="Sistema de Filtrado Multi-Archivos",
            font=("Arial", 12),
            fg=self.colors["light"],
            bg=self.colors["primary_dark"],
        )
        subtitle_label.place(relx=0.5, rely=0.6, anchor="center")

        # Marca registrada sutil en el header
        copyright_header = tk.Label(
            header_frame,
            text="© Jonathan Fuentes Toledo",
            font=("Arial", 8),
            bg=self.colors["primary"],
            fg=self.colors["light"],
        )
        copyright_header.pack(pady=2)

        # Frame principal con fondo degradado y sombra
        main_frame = tk.Frame(self.root, bg=self.colors["gradient_start"])
        main_frame.pack(padx=30, pady=20, fill="both", expand=True)

        # Estilo personalizado para los frames
        def create_rounded_frame(parent, title):
            frame = tk.Frame(parent, bg=self.colors["white"])
            frame.pack(fill="x", pady=(0, 15), padx=2)

            # Efecto de sombra
            shadow_frame = tk.Frame(parent, bg=self.colors["shadow"])
            shadow_frame.place(
                in_=frame, relx=0.002, rely=0.002, relwidth=1, relheight=1
            )
            frame.lift()

            # Título del frame
            title_label = tk.Label(
                frame,
                text=title,
                font=("Arial", 13, "bold"),
                bg=self.colors["primary"],
                fg=self.colors["white"],
                padx=15,
                pady=8,
            )
            title_label.pack(anchor="w", pady=(0, 10))

            # Contenedor interno
            inner_frame = tk.Frame(frame, bg=self.colors["white"], padx=20, pady=15)
            inner_frame.pack(fill="x", expand=True)

            return inner_frame

        # Frame para selección de archivo con nuevo estilo
        file_frame = create_rounded_frame(main_frame, "📁 Selección de Archivo Excel")

        # Botón para cargar archivo con estilo moderno y efecto hover
        self.load_button = tk.Button(
            file_frame,
            text="📂 Cargar Archivo Excel",
            command=self.cargar_archivo,
            font=("Arial", 12, "bold"),
            bg=self.colors["primary_dark"],
            fg=self.colors["white"],
            relief="flat",
            bd=0,
            padx=25,
            pady=12,
            cursor="hand2",
        )
        self.load_button.pack(side="left", padx=(0, 15))

        # Configurar efectos hover para el botón de cargar
        def on_enter_load(e):
            self.load_button.config(bg=self.colors["hover"])

        def on_leave_load(e):
            self.load_button.config(bg=self.colors["primary_dark"])

        self.load_button.bind("<Enter>", on_enter_load)
        self.load_button.bind("<Leave>", on_leave_load)

        # Botón para refrescar lista con estilo moderno y efecto hover
        self.refresh_button = tk.Button(
            file_frame,
            text="🔄 Refrescar Lista",
            command=self.refrescar_archivos,
            font=("Arial", 11),
            bg=self.colors["secondary"],
            fg=self.colors["dark"],
            relief="flat",
            bd=0,
            padx=20,
            pady=10,
            cursor="hand2",
        )
        self.refresh_button.pack(side="left")

        # Configurar efectos hover para el botón de refrescar
        def on_enter_refresh(e):
            self.refresh_button.config(bg=self.colors["hover"], fg=self.colors["white"])

        def on_leave_refresh(e):
            self.refresh_button.config(
                bg=self.colors["secondary"], fg=self.colors["white"]
            )

        self.refresh_button.bind("<Enter>", on_enter_refresh)
        self.refresh_button.bind("<Leave>", on_leave_refresh)

        # Label para mostrar archivo seleccionado con estilo mejorado
        self.file_label = tk.Label(
            file_frame,
            text="Ningún archivo seleccionado",
            font=("Arial", 11),
            bg=self.colors["white"],
            fg=self.colors["dark_gray"],
            wraplength=400,
            padx=15,
            pady=8,
            relief="sunken",
            bd=1,
        )
        self.file_label.pack(side="right", padx=(15, 0), fill="x", expand=True)

        # Frame para filtros con nuevo estilo
        filter_frame = create_rounded_frame(main_frame, "🔍 Filtros de Búsqueda")

        # Frame para prestaciones con estilo moderno
        prestacion_frame = tk.Frame(filter_frame, bg=self.colors["white"])
        prestacion_frame.pack(fill="x", pady=(0, 8))

        # Frame para búsqueda con diseño mejorado
        search_prestacion_frame = tk.Frame(prestacion_frame, bg=self.colors["white"])
        search_prestacion_frame.pack(fill="x", pady=(0, 3))

        # Estilo personalizado para entry
        entry_frame = tk.Frame(
            search_prestacion_frame, bg=self.colors["light"], bd=1, relief="flat"
        )
        entry_frame.pack(side="left", fill="x", expand=True, padx=(0, 8))

        # Label y campo de búsqueda con diseño moderno
        search_label = tk.Label(
            search_prestacion_frame,
            text="🔍 Buscar:",
            font=("Arial", 10, "bold"),
            bg=self.colors["white"],
            fg=self.colors["primary"],
        )
        search_label.pack(side="left", padx=(0, 8))

        # Campo de búsqueda con estilo mejorado
        self.search_prestacion_var = tk.StringVar()
        self.search_prestacion_entry = tk.Entry(
            entry_frame,
            textvariable=self.search_prestacion_var,
            font=("Arial", 10),
            bg=self.colors["light"],
            fg=self.colors["dark"],
            relief="flat",
            bd=0,
            width=25,
            insertbackground=self.colors["primary"],
        )
        self.search_prestacion_entry.pack(
            side="left", fill="x", expand=True, padx=8, pady=6
        )

        # Botón para limpiar búsqueda con estilo moderno
        self.clear_search_button = tk.Button(
            search_prestacion_frame,
            text="❌",
            command=self.limpiar_busqueda_prestacion,
            font=("Arial", 8),
            bg=self.colors["error"],
            fg=self.colors["white"],
            relief="flat",
            bd=0,
            width=2,
            height=1,
            cursor="hand2",
        )
        self.clear_search_button.pack(side="right")

        # Efectos hover para el botón de limpiar
        def on_enter_clear(e):
            self.clear_search_button.config(bg=self.colors["hover"])

        def on_leave_clear(e):
            self.clear_search_button.config(bg=self.colors["error"])

        self.clear_search_button.bind("<Enter>", on_enter_clear)
        self.clear_search_button.bind("<Leave>", on_leave_clear)

        # Frame para selección de prestación con estilo moderno
        select_prestacion_frame = tk.Frame(prestacion_frame, bg=self.colors["white"])
        select_prestacion_frame.pack(fill="x", pady=(10, 0))

        # Label para prestación con diseño mejorado
        prestacion_label = tk.Label(
            select_prestacion_frame,
            text="📋 Prestación:",
            font=("Arial", 10, "bold"),
            bg=self.colors["white"],
            fg=self.colors["primary"],
        )
        prestacion_label.pack(side="left", padx=(0, 8))

        # Frame para el combobox con estilo moderno
        combo_frame = tk.Frame(
            select_prestacion_frame, bg=self.colors["light"], bd=1, relief="flat"
        )
        combo_frame.pack(side="left", fill="x", expand=True)

        self.prestacion_var = tk.StringVar()
        self.prestacion_combo = ttk.Combobox(
            combo_frame,
            textvariable=self.prestacion_var,
            font=("Arial", 10),
            state="readonly",
            width=35,
        )
        self.prestacion_combo.pack(side="left", fill="x", expand=True, padx=8, pady=6)

        # Configurar estilo del combobox
        combo_style = ttk.Style()
        combo_style.configure(
            "TCombobox",
            background=self.colors["light"],
            fieldbackground=self.colors["light"],
            foreground=self.colors["dark"],
            arrowcolor=self.colors["primary"],
            relief="flat",
        )

        # Vincular eventos de búsqueda
        self.search_prestacion_var.trace("w", self.filtrar_prestaciones)
        self.prestacion_combo.bind("<<ComboboxSelected>>", self.on_prestacion_selected)

        # Frame para botones con diseño moderno
        buttons_frame = tk.Frame(filter_frame, bg=self.colors["white"])
        buttons_frame.pack(pady=(15, 0), anchor="center")

        # Botón para buscar prestación con estilo moderno
        self.search_button = tk.Button(
            buttons_frame,
            text="🔍 Buscar",
            command=self.buscar_prestacion,
            font=("Arial", 11, "bold"),
            bg=self.colors["success"],
            fg=self.colors["white"],
            relief="flat",
            bd=0,
            padx=20,
            pady=8,
            cursor="hand2",
        )
        self.search_button.pack(side="left", padx=5)

        # Efectos hover para el botón de búsqueda
        def on_enter_search(e):
            self.search_button.config(bg=self.colors["hover"])

        def on_leave_search(e):
            self.search_button.config(bg=self.colors["success"])

        self.search_button.bind("<Enter>", on_enter_search)
        self.search_button.bind("<Leave>", on_leave_search)

        # Botón para configurar columnas con estilo moderno
        self.columns_button = tk.Button(
            buttons_frame,
            text="⚙️ Columnas",
            command=self.configurar_columnas,
            font=("Arial", 11),
            bg=self.colors["secondary"],
            fg=self.colors["dark"],
            relief="flat",
            bd=0,
            padx=15,
            pady=8,
            cursor="hand2",
        )
        self.columns_button.pack(side="left", padx=5)

        # Efectos hover para el botón de columnas
        def on_enter_columns(e):
            self.columns_button.config(bg=self.colors["hover"], fg=self.colors["white"])

        def on_leave_columns(e):
            self.columns_button.config(
                bg=self.colors["secondary"], fg=self.colors["dark"]
            )

        self.columns_button.bind("<Enter>", on_enter_columns)
        self.columns_button.bind("<Leave>", on_leave_columns)
        self.columns_button.pack(pady=(5, 0))

        # Frame para botones de acción en la parte superior más compacto
        top_buttons_frame = tk.Frame(filter_frame, bg=self.colors["white"])
        top_buttons_frame.pack(fill="x", pady=(8, 0))

        # Botón para guardar resultados más compacto
        self.save_button = tk.Button(
            top_buttons_frame,
            text="💾 GUARDAR",
            command=self.guardar_resultado,
            font=("Arial", 9, "bold"),
            bg="#27AE60",  # Verde brillante
            fg="white",
            activebackground="#2ECC71",
            activeforeground="white",
            relief="raised",
            bd=1,
            padx=10,
            pady=4,
            cursor="hand2",
        )
        self.save_button.pack(side="right")

        # Label informativo más compacto
        info_label = tk.Label(
            top_buttons_frame,
            text="📊 Haz clic en el botón verde para guardar",
            font=("Arial", 8),
            bg=self.colors["white"],
            fg=self.colors["dark_gray"],
        )
        info_label.pack(side="left")

        # Frame para resultados con más espacio
        results_frame = tk.LabelFrame(
            main_frame,
            text="📊 Resultados del Filtrado",
            font=("Arial", 13, "bold"),
            bg=self.colors["white"],
            fg=self.colors["primary"],
            padx=15,
            pady=10,
            relief="groove",
            bd=2,
        )
        results_frame.pack(fill="both", expand=True, pady=(0, 10))

        # Información de resultados más compacta
        self.results_info = tk.Label(
            results_frame,
            text="No hay datos cargados",
            font=("Arial", 10),
            bg=self.colors["white"],
            fg=self.colors["dark_gray"],
            padx=10,
            pady=6,
            relief="sunken",
            bd=1,
        )
        self.results_info.pack(anchor="w", pady=(0, 8), fill="x")

        # Treeview para mostrar resultados con más altura
        tree_frame = tk.Frame(results_frame, bg=self.colors["white"])
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, show="headings", height=18)

        # Scrollbars para el treeview con estilo mejorado
        tree_scroll_y = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.tree.yview
        )
        tree_scroll_x = ttk.Scrollbar(
            tree_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(
            yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set
        )

        # Empaquetar treeview y scrollbars
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x.pack(side="bottom", fill="x")

        # Frame para botones de acción con estilo mejorado
        action_frame = tk.Frame(main_frame, bg=self.colors["gradient_start"])
        action_frame.pack(fill="x", pady=(0, 15))

        # Botón para limpiar resultados con estilo mejorado
        self.clear_button = tk.Button(
            action_frame,
            text="🧹 Limpiar Resultados",
            command=self.limpiar_resultados,
            font=("Arial", 11),
            bg=self.colors["warning"],
            fg=self.colors["white"],
            activebackground=self.colors["hover"],
            activeforeground=self.colors["white"],
            relief="raised",
            bd=2,
            padx=20,
            pady=8,
            cursor="hand2",
        )
        self.clear_button.pack(side="left", padx=(0, 15))

        # Botón para mostrar todas las prestaciones con estilo mejorado
        self.show_all_button = tk.Button(
            action_frame,
            text="📋 Mostrar Todas las Prestaciones",
            command=self.mostrar_todas_prestaciones,
            font=("Arial", 11),
            bg=self.colors["secondary"],
            fg=self.colors["dark"],
            activebackground=self.colors["hover"],
            activeforeground=self.colors["white"],
            relief="raised",
            bd=2,
            padx=20,
            pady=8,
            cursor="hand2",
        )
        self.show_all_button.pack(side="right")

        # Status bar con estilo mejorado
        self.status_var = tk.StringVar()
        self.status_var.set("Sistema listo - Seleccione un archivo Excel")
        self.status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w",
            font=("Arial", 10),
            bg=self.colors["primary"],
            fg=self.colors["white"],
            padx=15,
            pady=8,
        )
        self.status_bar.pack(side="bottom", fill="x")

        # Footer con marca registrada
        footer_frame = tk.Frame(self.root, bg=self.colors["primary_dark"], height=25)
        footer_frame.pack(side="bottom", fill="x")
        footer_frame.pack_propagate(False)

        # Marca registrada sutil
        copyright_label = tk.Label(
            footer_frame,
            text="© 2024 Jonathan Fuentes Toledo - Vidasalud Dental",
            font=("Arial", 8),
            bg=self.colors["primary_dark"],
            fg=self.colors["light"],
        )
        copyright_label.pack(side="right", padx=10, pady=5)

    def refrescar_archivos(self):
        """Refresca la lista de archivos disponibles"""
        self.cargar_archivos_disponibles()

        # Actualizar información de archivos
        if self.archivos_excel:
            self.file_label.config(
                text=f"{len(self.archivos_excel)} archivos disponibles"
            )
        else:
            self.file_label.config(text="Ningún archivo encontrado")

        # Limpiar prestaciones
        self.prestacion_combo["values"] = []
        self.prestacion_var.set("")
        self.search_button.config(state="disabled")

        self.status_var.set(
            f"✅ Lista actualizada - {len(self.archivos_excel)} archivos encontrados"
        )

    def on_archivo_selected(self, event):
        """Cuando se selecciona un archivo"""
        # Este método ya no es necesario con la nueva interfaz
        pass

    def cargar_archivo(self):
        """Carga el archivo Excel seleccionado"""
        # Mostrar diálogo para seleccionar archivo
        archivo_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
            initialdir=os.path.join(self.get_base_path(), "archivos_excel"),
        )

        if not archivo_path:
            return

        try:
            # Cargar datos
            self.df = pd.read_excel(archivo_path, engine="openpyxl", header=1)
            self.archivo_seleccionado = os.path.basename(archivo_path)

            # Obtener prestaciones
            if self.df is not None and "Prestación" in self.df.columns:
                self.prestaciones = sorted(self.df["Prestación"].unique())

                # Actualizar combo de prestaciones
                self.prestacion_combo["values"] = self.prestaciones
                self.prestacion_var.set("")

                # Configurar búsqueda de prestaciones
                self.search_prestacion_var.set("")

                # Habilitar búsqueda
                self.search_button.config(state="normal")

                # Habilitar botón de guardar (aunque no haya filtrado aún)
                self.save_button.config(state="normal")

                # Mostrar información
                info_text = f"✅ Archivo cargado: {self.archivo_seleccionado}\n📊 Registros: {len(self.df):,}\n📋 Prestaciones: {len(self.prestaciones)}"
                self.results_info.config(text=info_text)

                # Actualizar label del archivo
                self.file_label.config(text=self.archivo_seleccionado)

                self.status_var.set(
                    f"✅ Archivo cargado: {self.archivo_seleccionado} - {len(self.df):,} registros"
                )
            else:
                messagebox.showerror(
                    "❌ Error", "El archivo no contiene la columna 'Prestación'"
                )

        except Exception as e:
            messagebox.showerror("❌ Error", f"Error al cargar el archivo:\n{str(e)}")

    def on_prestacion_selected(self, event):
        """Cuando se selecciona una prestación"""
        prestacion = self.prestacion_var.get()
        if prestacion:
            self.status_var.set(f"📋 Prestación seleccionada: {prestacion}")

    def filtrar_prestaciones(self, *args):
        """Filtra las prestaciones según el texto de búsqueda"""
        texto_busqueda = self.search_prestacion_var.get().lower().strip()

        if not self.prestaciones:
            return

        if not texto_busqueda:
            # Si no hay texto de búsqueda, mostrar todas las prestaciones
            self.prestacion_combo["values"] = self.prestaciones
        else:
            # Filtrar prestaciones que contengan el texto de búsqueda
            prestaciones_filtradas = [
                prestacion
                for prestacion in self.prestaciones
                if texto_busqueda in prestacion.lower()
            ]
            self.prestacion_combo["values"] = prestaciones_filtradas

            # Si solo hay una prestación filtrada, seleccionarla automáticamente
            if len(prestaciones_filtradas) == 1:
                self.prestacion_var.set(prestaciones_filtradas[0])

        # Actualizar estado
        prestaciones_mostradas = len(self.prestacion_combo["values"])
        total_prestaciones = len(self.prestaciones)

        if texto_busqueda:
            self.status_var.set(
                f"🔍 Búsqueda: '{texto_busqueda}' - {prestaciones_mostradas} de {total_prestaciones} prestaciones"
            )
        else:
            self.status_var.set(
                f"📋 Mostrando todas las prestaciones ({total_prestaciones})"
            )

    def limpiar_busqueda_prestacion(self):
        """Limpia el campo de búsqueda de prestaciones"""
        self.search_prestacion_var.set("")
        self.prestacion_var.set("")

        if self.prestaciones:
            self.prestacion_combo["values"] = self.prestaciones
            self.status_var.set(
                f"✅ Búsqueda limpiada - Mostrando todas las prestaciones ({len(self.prestaciones)})"
            )
        else:
            self.status_var.set("✅ Búsqueda limpiada")

    def buscar_prestacion(self):
        """Busca la prestación seleccionada"""
        prestacion = self.prestacion_var.get()

        if not prestacion:
            messagebox.showwarning(
                "⚠️ Advertencia", "Por favor selecciona una prestación"
            )
            return

        if self.df is None:
            messagebox.showwarning("⚠️ Advertencia", "No hay datos cargados")
            return

        try:
            # Filtrar datos
            filtro = self.df["Prestación"] == prestacion
            self.resultado_filtrado = self.df[filtro]

            # Mostrar resultados en treeview
            self.mostrar_resultados_en_treeview()

            # Mostrar información
            total_registros = len(self.resultado_filtrado)
            total_original = len(self.df)
            porcentaje = (total_registros / total_original) * 100

            info_text = f"✅ Filtrado completado\n📊 Registros: {total_registros:,} de {total_original:,} ({porcentaje:.1f}%)\n🔍 Prestación: {prestacion}"
            self.results_info.config(text=info_text)

            # Habilitar botón de guardar
            self.save_button.config(state="normal")
            self.status_var.set(
                f"✅ Encontrados {total_registros} registros para '{prestacion}'"
            )

        except Exception as e:
            messagebox.showerror("❌ Error", f"Error al filtrar:\n{str(e)}")

    def mostrar_resultados_en_treeview(self):
        """Muestra los resultados en el treeview"""
        # Limpiar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.resultado_filtrado is None or len(self.resultado_filtrado) == 0:
            return

        # Usar columnas seleccionadas o todas si no hay selección
        if self.columnas_seleccionadas:
            columnas = [
                col
                for col in self.columnas_seleccionadas
                if col in self.resultado_filtrado.columns
            ]
            if columnas:
                df_mostrar = self.resultado_filtrado[columnas].copy()
            else:
                df_mostrar = self.resultado_filtrado.copy()
                columnas = list(self.resultado_filtrado.columns)
        else:
            columnas = list(self.resultado_filtrado.columns)
            df_mostrar = self.resultado_filtrado.copy()

        # Verificar que df_mostrar sea un DataFrame válido
        if not isinstance(df_mostrar, pd.DataFrame):
            df_mostrar = pd.DataFrame(df_mostrar)

        # Configurar columnas
        self.tree["columns"] = columnas

        # Configurar encabezados
        for col in columnas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, minwidth=50)

        # Insertar datos
        for idx, row in df_mostrar.iterrows():
            valores = [str(val) for val in row.values]
            self.tree.insert("", "end", values=valores)

    def obtener_ruta_guardado(self):
        """Obtiene la ruta segura para guardar archivos"""
        # Usar carpeta Documentos del usuario
        docs = os.path.join(os.path.expanduser("~"), "Documents")
        output_dir = os.path.join(docs, "Vidasalud_Export")

        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except:
                return docs  # Fallback a documentos si falla crear carpeta

        return output_dir

    def guardar_resultado(self):
        """Guarda el resultado filtrado"""
        # Si no hay datos filtrados, preguntar si quiere guardar todos los datos
        if self.resultado_filtrado is None or len(self.resultado_filtrado) == 0:
            if self.df is not None and len(self.df) > 0:
                respuesta = messagebox.askyesno(
                    "📊 Sin Filtros",
                    "No hay datos filtrados.\n\n¿Quieres guardar TODOS los datos del archivo?\n\nSi no, primero filtra una prestación.",
                )
                if respuesta:
                    # Guardar todos los datos
                    self.guardar_todos_datos()
                return
            else:
                messagebox.showwarning(
                    "⚠️ Advertencia",
                    "No hay datos para guardar. Primero carga un archivo Excel.",
                )
                return

        try:
            # Crear nombre del archivo
            prestacion = self.prestacion_var.get()
            if self.archivo_seleccionado:
                archivo_base = os.path.splitext(self.archivo_seleccionado)[0]
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"Vidasalud_{archivo_base}_{prestacion.replace(' ', '_').replace('/', '_')}_{timestamp}_JFT.xlsx"

                output_dir = self.obtener_ruta_guardado()
                ruta_completa = os.path.join(output_dir, nombre_archivo)

                # Guardar archivo
                self.resultado_filtrado.to_excel(ruta_completa, index=False)

                messagebox.showinfo(
                    "✅ Archivo Guardado",
                    f"El archivo se ha guardado exitosamente!\n\n📁 Nombre: {nombre_archivo}\n📊 Registros: {len(self.resultado_filtrado):,}\n📍 Ubicación: {output_dir}",
                )

                self.status_var.set(f"💾 Archivo guardado: {nombre_archivo}")
            else:
                messagebox.showerror("❌ Error", "No hay archivo seleccionado")

        except Exception as e:
            messagebox.showerror("❌ Error", f"Error al guardar:\n{str(e)}")

    def guardar_todos_datos(self):
        """Guarda todos los datos del archivo"""
        try:
            if self.df is not None and self.archivo_seleccionado:
                archivo_base = os.path.splitext(self.archivo_seleccionado)[0]
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = (
                    f"Vidasalud_{archivo_base}_TODOS_LOS_DATOS_{timestamp}_JFT.xlsx"
                )

                output_dir = self.obtener_ruta_guardado()
                ruta_completa = os.path.join(output_dir, nombre_archivo)

                # Guardar archivo
                self.df.to_excel(ruta_completa, index=False)

                messagebox.showinfo(
                    "✅ Archivo Guardado",
                    f"Se guardaron TODOS los datos exitosamente!\n\n📁 Nombre: {nombre_archivo}\n📊 Registros: {len(self.df):,}\n📍 Ubicación: {output_dir}",
                )

                self.status_var.set(f"💾 Archivo guardado: {nombre_archivo}")
            else:
                messagebox.showerror(
                    "❌ Error", "No hay archivo seleccionado o datos cargados"
                )

        except Exception as e:
            messagebox.showerror("❌ Error", f"Error al guardar:\n{str(e)}")

    def limpiar_resultados(self):
        """Limpia los resultados y la búsqueda"""
        # Limpiar variables de prestación
        self.prestacion_var.set("")
        self.search_prestacion_var.set("")
        self.resultado_filtrado = None

        # Limpiar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Limpiar información de resultados
        if self.df is not None:
            info_text = f"✅ Archivo cargado: {self.archivo_seleccionado}\n📊 Registros: {len(self.df):,}\n📋 Prestaciones: {len(self.prestaciones)}"
            self.results_info.config(text=info_text)
        else:
            self.results_info.config(text="No hay datos cargados")

        # Restaurar todas las prestaciones en el combo
        if self.prestaciones:
            self.prestacion_combo["values"] = self.prestaciones

        # Deshabilitar botón de guardar
        self.save_button.config(state="disabled")

        # Actualizar estado
        self.status_var.set("✅ Resultados limpiados - Listo para nueva búsqueda")

        # Mostrar mensaje de confirmación
        messagebox.showinfo(
            "🧹 Limpieza Completada",
            "Se han limpiado todos los resultados y la búsqueda.\n\nPuedes realizar una nueva búsqueda.",
        )

    def mostrar_todas_prestaciones(self):
        """Muestra todas las prestaciones disponibles"""
        if not self.prestaciones:
            messagebox.showinfo(
                "📋 Prestaciones",
                "No hay prestaciones cargadas. Primero carga un archivo Excel.",
            )
            return

        # Crear ventana de prestaciones con estilo mejorado
        prestaciones_window = tk.Toplevel(self.root)
        prestaciones_window.title("📋 Todas las Prestaciones Disponibles")
        prestaciones_window.geometry("700x500")
        prestaciones_window.configure(bg=self.colors["gradient_start"])

        # Header de la ventana
        header_prestaciones = tk.Frame(
            prestaciones_window, bg=self.colors["primary"], height=80
        )
        header_prestaciones.pack(fill="x", padx=0, pady=0)
        header_prestaciones.pack_propagate(False)

        # Título con estilo mejorado
        tk.Label(
            header_prestaciones,
            text="📋 Prestaciones Disponibles",
            font=("Arial", 16, "bold"),
            bg=self.colors["primary"],
            fg=self.colors["white"],
        ).pack(pady=15)

        # Frame principal
        main_prestaciones_frame = tk.Frame(prestaciones_window, bg=self.colors["white"])
        main_prestaciones_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Información de prestaciones
        info_label = tk.Label(
            main_prestaciones_frame,
            text=f"Total: {len(self.prestaciones)} prestaciones disponibles",
            font=("Arial", 11, "bold"),
            bg=self.colors["white"],
            fg=self.colors["primary"],
            pady=10,
        )
        info_label.pack()

        # Listbox con scrollbar mejorado
        frame_listbox = tk.Frame(main_prestaciones_frame, bg=self.colors["white"])
        frame_listbox.pack(fill="both", expand=True, padx=20, pady=10)

        scrollbar = tk.Scrollbar(frame_listbox, bg=self.colors["secondary"])
        scrollbar.pack(side="right", fill="y")

        listbox = tk.Listbox(
            frame_listbox,
            yscrollcommand=scrollbar.set,
            font=("Arial", 11),
            bg=self.colors["light"],
            fg=self.colors["dark"],
            selectmode="single",
            relief="sunken",
            bd=2,
            selectbackground=self.colors["primary"],
            selectforeground=self.colors["white"],
        )
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        # Insertar prestaciones con numeración mejorada
        for i, prestacion in enumerate(self.prestaciones, 1):
            listbox.insert(tk.END, f"{i:3d}. {prestacion}")

        # Frame para botones
        button_frame = tk.Frame(main_prestaciones_frame, bg=self.colors["white"])
        button_frame.pack(fill="x", padx=20, pady=15)

        # Botón cerrar con estilo mejorado
        tk.Button(
            button_frame,
            text="✅ Cerrar",
            command=prestaciones_window.destroy,
            font=("Arial", 11, "bold"),
            bg=self.colors["success"],
            fg=self.colors["white"],
            relief="raised",
            bd=2,
            padx=25,
            pady=8,
            cursor="hand2",
        ).pack(side="right")

    def configurar_columnas(self):
        """Abre una ventana para seleccionar qué columnas mostrar"""
        if self.df is None:
            messagebox.showinfo(
                "📋 Columnas",
                "Primero carga un archivo Excel para ver las columnas disponibles.",
            )
            return

        # Crear ventana de configuración de columnas
        columns_window = tk.Toplevel(self.root)
        columns_window.title("⚙️ Configurar Columnas de Resultados")
        columns_window.geometry("500x600")
        columns_window.configure(bg=self.colors["light"])

        # Título
        tk.Label(
            columns_window,
            text="⚙️ Seleccionar Columnas para Mostrar",
            font=("Arial", 14, "bold"),
            bg=self.colors["light"],
            fg=self.colors["dark"],
        ).pack(pady=10)

        # Instrucciones
        tk.Label(
            columns_window,
            text="Marca las columnas que quieres ver en los resultados:",
            font=("Arial", 10),
            bg=self.colors["light"],
            fg=self.colors["dark_gray"],
        ).pack(pady=5)

        # Frame para checkboxes
        checkbox_frame = tk.Frame(columns_window, bg=self.colors["light"])
        checkbox_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Scrollbar para checkboxes
        scrollbar = tk.Scrollbar(checkbox_frame)
        scrollbar.pack(side="right", fill="y")

        canvas = tk.Canvas(
            checkbox_frame, yscrollcommand=scrollbar.set, bg=self.colors["light"]
        )
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=canvas.yview)

        inner_frame = tk.Frame(canvas, bg=self.colors["light"])
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        # Variables para checkboxes
        column_vars = {}

        # Crear checkboxes para cada columna
        for i, columna in enumerate(self.df.columns):
            var = tk.BooleanVar()
            # Marcar como seleccionada si ya está en la lista o si es la primera vez
            if (
                not self.columnas_seleccionadas
                or columna in self.columnas_seleccionadas
            ):
                var.set(True)
            column_vars[columna] = var

            checkbox = tk.Checkbutton(
                inner_frame,
                text=columna,
                variable=var,
                font=("Arial", 10),
                bg=self.colors["light"],
                fg=self.colors["dark"],
                selectcolor=self.colors["secondary"],
                activebackground=self.colors["light"],
                activeforeground=self.colors["dark"],
            )
            checkbox.pack(anchor="w", pady=2)

        # Botones de acción
        button_frame = tk.Frame(columns_window, bg=self.colors["light"])
        button_frame.pack(fill="x", padx=20, pady=10)

        def aplicar_configuracion():
            # Obtener columnas seleccionadas
            self.columnas_seleccionadas = [
                col for col, var in column_vars.items() if var.get()
            ]

            # Si hay resultados filtrados, actualizar la vista
            if self.resultado_filtrado is not None:
                self.mostrar_resultados_en_treeview()

            columns_window.destroy()
            messagebox.showinfo(
                "✅ Configuración",
                f"Se configuraron {len(self.columnas_seleccionadas)} columnas para mostrar.",
            )

        def seleccionar_todas():
            for var in column_vars.values():
                var.set(True)

        def deseleccionar_todas():
            for var in column_vars.values():
                var.set(False)

        # Botones
        tk.Button(
            button_frame,
            text="✅ Aplicar",
            command=aplicar_configuracion,
            font=("Arial", 11, "bold"),
            bg=self.colors["success"],
            fg=self.colors["white"],
            relief="raised",
            bd=2,
            padx=20,
            pady=5,
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            button_frame,
            text="📋 Seleccionar Todas",
            command=seleccionar_todas,
            font=("Arial", 10),
            bg=self.colors["primary"],
            fg=self.colors["white"],
            relief="raised",
            bd=2,
            padx=15,
            pady=5,
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            button_frame,
            text="❌ Deseleccionar Todas",
            command=deseleccionar_todas,
            font=("Arial", 10),
            bg=self.colors["warning"],
            fg=self.colors["white"],
            relief="raised",
            bd=2,
            padx=15,
            pady=5,
        ).pack(side="left")

        # Configurar scroll
        inner_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))


def main():
    root = tk.Tk()
    app = FiltradorMultiArchivosGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

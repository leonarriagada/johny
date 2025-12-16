import pandas as pd
import tkinter as tk  # Keep for file dialogs and some constants if needed
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk  # NEW: Modern UI library
import os
from datetime import datetime
import glob
import sys
import threading  # For non-blocking operations if needed

# Configuration for High DPI (Windows) - Optional but good practice
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

# --- THEME CONFIGURATION ---
# "System" uses the OS mode (Dark/Light)
# "DarkBlue", "Blue", "Green" are built-in themes. We can use a custom color if needed for "Million Dollar" look.
ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")  # We will override specific colors for a premium look

class FiltradorMultiArchivosGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🦷 Vidasalud Dental - Filtrador Multi-Archivos | © Jonathan Fuentes Toledo")
        
        # Determine platform for maximized state
        if sys.platform == "win32":
            self.root.state("zoomed")
        elif sys.platform == "linux":
            self.root.attributes("-zoomed", True)
        else:
            # macOS: Start with a large size or fullscreen, user can maximize
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            self.root.geometry(f"{int(screen_width*0.9)}x{int(screen_height*0.9)}+{int(screen_width*0.05)}+{int(screen_height*0.05)}")

        # --- PREMIUM PALETTE (Subtle adjustments) ---
        self.colors = {
            "bg_color": ("#F9F9FA", "#1A1A1A"),  # Light/Dark background
            "card_bg": ("#FFFFFF", "#2B2B2B"),   # Card background
            "primary": "#3B8ED0",                # Modern Blue
            "primary_hover": "#36719F",
            "text_main": ("#1A1A1A", "#FFFFFF"),
            "text_sec": ("#666666", "#AAAAAA"),
            "success": "#2CC985",
            "danger": "#E74C3C"
        }

        # Root configuration
        self.root.configure(bg=self.colors["bg_color"][0] if ctk.get_appearance_mode()=="Light" else self.colors["bg_color"][1])

        # Variables
        self.archivos_excel = []
        self.archivo_seleccionado = None
        self.df = None
        self.prestaciones = []
        self.resultado_filtrado = None
        self.columnas_seleccionadas = []

        # Layout Setup
        self.setup_ui()
        
        # Load files initially
        self.cargar_archivos_disponibles()

    def get_base_path(self):
        """Returns executable/script directory"""
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        else:
            return os.path.dirname(os.path.abspath(__file__))

    def get_data_path(self):
        """Returns safe data directory (Documents on Mac)"""
        if sys.platform == "darwin" and getattr(sys, "frozen", False):
            docs = os.path.join(os.path.expanduser("~"), "Documents")
            data_dir = os.path.join(docs, "Vidasalud_Data")
            if not os.path.exists(data_dir):
                try: os.makedirs(data_dir)
                except: pass
            return data_dir
        else:
            return self.get_base_path()

    def setup_ui(self):
        """Construye la interfaz moderna"""
        # Configuración principal de la grilla
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        # --- 1. HEADER (Encabezado) ---
        self.header = ctk.CTkFrame(self.root, corner_radius=0, fg_color=self.colors["card_bg"])
        self.header.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        
        # Título principal
        self.title_label = ctk.CTkLabel(
            self.header, 
            text="🦷 VidaSalud Dental", 
            font=ctk.CTkFont(family="Arial", size=24, weight="bold"),
            text_color=self.colors["primary"]
        )
        self.title_label.pack(side="left", padx=30, pady=20)

        # Subtítulo y Créditos
        credits_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        credits_frame.pack(side="left", padx=10)
        
        self.subtitle_label = ctk.CTkLabel(
            credits_frame, 
            text="Sistema de Filtrado Profesional", 
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.colors["text_sec"]
        )
        self.subtitle_label.pack(anchor="w")
        
        self.copyright_label = ctk.CTkLabel(
            credits_frame, 
            text="© Jonathan Fuentes Toledo", 
            font=ctk.CTkFont(size=11),
            text_color=self.colors["text_sec"]
        )
        self.copyright_label.pack(anchor="w")

        # Switch de Tema
        self.theme_switch = ctk.CTkSwitch(self.header, text="Modo Oscuro", command=self.toggle_theme)
        self.theme_switch.pack(side="right", padx=30)
    
        # --- 2. CONTENIDO PRINCIPAL ---
        self.content_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=30, pady=30)
        self.content_frame.grid_columnconfigure(1, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        # --- BARRA LATERAL (Controles) ---
        self.sidebar = ctk.CTkFrame(self.content_frame, corner_radius=15, width=320, fg_color=self.colors["card_bg"])
        self.sidebar.grid(row=0, column=0, sticky="nsw", padx=(0, 20), pady=0)
        self.sidebar.grid_propagate(False)

        # Sección: Archivo
        self.lbl_file = ctk.CTkLabel(self.sidebar, text="FUENTE DE DATOS", font=ctk.CTkFont(size=12, weight="bold"), text_color=self.colors["text_sec"])
        self.lbl_file.pack(anchor="w", padx=20, pady=(25, 10))

        self.btn_load = ctk.CTkButton(
            self.sidebar, 
            text="📂 Cargar Archivo Excel", 
            command=self.cargar_archivo,
            height=45,
            corner_radius=8,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=self.colors["primary"],
            hover_color=self.colors["primary_hover"]
        )
        self.btn_load.pack(fill="x", padx=20, pady=5)

        self.lbl_current_file = ctk.CTkLabel(
            self.sidebar, 
            text="Ningún archivo seleccionado", 
            font=ctk.CTkFont(size=12), 
            text_color=self.colors["text_sec"], 
            wraplength=280,
            justify="left"
        )
        self.lbl_current_file.pack(anchor="w", padx=20, pady=10)

        # Separador
        ctk.CTkFrame(self.sidebar, height=2, fg_color=self.colors["bg_color"][0] if ctk.get_appearance_mode()=="Light" else "gray30").pack(fill="x", padx=20, pady=15)

        # Sección: Filtros
        self.lbl_filter = ctk.CTkLabel(self.sidebar, text="FILTRADO INTELIGENTE", font=ctk.CTkFont(size=12, weight="bold"), text_color=self.colors["text_sec"])
        self.lbl_filter.pack(anchor="w", padx=20, pady=(5, 10))

        self.txt_search = ctk.CTkEntry(self.sidebar, placeholder_text="🔍 Buscar prestación...", height=40)
        self.txt_search.pack(fill="x", padx=20, pady=5)
        self.txt_search.bind("<KeyRelease>", self.filtrar_prestaciones_evento)

        self.combo_prestacion = ctk.CTkComboBox(self.sidebar, values=[], height=40)
        self.combo_prestacion.pack(fill="x", padx=20, pady=10)

        self.btn_apply_filter = ctk.CTkButton(
            self.sidebar, 
            text="🔍 Aplicar Filtro", 
            command=self.buscar_prestacion,
            height=45, 
            font=ctk.CTkFont(weight="bold"),
            fg_color=self.colors["success"],
            hover_color="#26AF73"
        )
        self.btn_apply_filter.pack(fill="x", padx=20, pady=10)

        self.btn_clear = ctk.CTkButton(
            self.sidebar, 
            text="🧹 Limpiar Filtros", 
            command=self.limpiar_resultados,
            height=35,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray80")
        )
        self.btn_clear.pack(fill="x", padx=20, pady=5)

        # Footer Actions
        self.btn_save = ctk.CTkButton(
            self.sidebar, 
            text="💾 GUARDAR RESULTADOS", 
            command=self.guardar_resultado,
            height=50,
            font=ctk.CTkFont(size=14, weight="bold"),
            state="disabled"
        )
        self.btn_save.pack(side="bottom", fill="x", padx=20, pady=30)


        # --- ÁREA PRINCIPAL ---
        self.main_area_container = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.main_area_container.grid(row=0, column=1, sticky="nsew")
        self.main_area_container.grid_rowconfigure(2, weight=1) # Treeview expande
        self.main_area_container.grid_columnconfigure(0, weight=1)

        # 1. Panel de Estadísticas (Nuevo, restaurado y mejorado)
        self.stats_frame = ctk.CTkFrame(self.main_area_container, corner_radius=15, height=100, fg_color=self.colors["card_bg"])
        self.stats_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        self.stats_frame.pack_propagate(False) # Mantener altura fija

        # Widget de estadística auxiliar
        def create_stat_widget(parent, title, value_var, color):
            f = ctk.CTkFrame(parent, fg_color="transparent")
            f.pack(side="left", fill="y", padx=40, pady=15)
            ctk.CTkLabel(f, text=title, font=ctk.CTkFont(size=12, weight="bold"), text_color="gray").pack(anchor="w")
            ctk.CTkLabel(f, textvariable=value_var, font=ctk.CTkFont(size=24, weight="bold"), text_color=color).pack(anchor="w")
            return f

        self.stat_total_var = tk.StringVar(value="0")
        self.stat_filtro_var = tk.StringVar(value="0")
        self.stat_perc_var = tk.StringVar(value="0%")

        create_stat_widget(self.stats_frame, "REGISTROS TOTALES", self.stat_total_var, self.colors["text_main"][0])
        # Separador vertical
        ctk.CTkFrame(self.stats_frame, width=2, height=40, fg_color="gray80").pack(side="left", pady=20)
        create_stat_widget(self.stats_frame, "REGISTROS FILTRADOS", self.stat_filtro_var, self.colors["primary"])
        # Separador vertical
        ctk.CTkFrame(self.stats_frame, width=2, height=40, fg_color="gray80").pack(side="left", pady=20)
        create_stat_widget(self.stats_frame, "PORCENTAJE", self.stat_perc_var, self.colors["success"])


        # 2. Barra Superior de Tabla
        self.table_header = ctk.CTkFrame(self.main_area_container, height=50, fg_color="transparent")
        self.table_header.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        self.lbl_results_title = ctk.CTkLabel(self.table_header, text="Vista de Datos", font=ctk.CTkFont(size=18, weight="bold"))
        self.lbl_results_title.pack(side="left")

        self.btn_columns = ctk.CTkButton(
            self.table_header, 
            text="⚙️ Columnas", 
            width=120, 
            command=self.configurar_columnas,
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90")
        )
        self.btn_columns.pack(side="right")

        # 3. Tabla (Treeview)
        self.tree_container = ctk.CTkFrame(self.main_area_container, corner_radius=15, fg_color=self.colors["card_bg"])
        self.tree_container.grid(row=2, column=0, sticky="nsew")
        
        # Frame interno para padding
        self.tree_frame = ctk.CTkFrame(self.tree_container, fg_color="transparent")
        self.tree_frame.pack(fill="both", expand=True, padx=20, pady=20)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
            background="white",
            foreground="#333333",
            rowheight=35,
            fieldbackground="white",
            bordercolor="#E5E5E5",
            borderwidth=0,
            font=("Segoe UI", 10)
        )
        style.configure("Treeview.Heading", 
            background="#F1F3F4", 
            foreground="#555555", 
            relief="flat", 
            font=("Segoe UI", 10, "bold")
        )
        style.map("Treeview", background=[('selected', self.colors['primary'])], foreground=[('selected', 'white')])

        self.tree_scroll_y = ctk.CTkScrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side="right", fill="y")
        
        self.tree_scroll_x = ctk.CTkScrollbar(self.tree_frame, orientation="horizontal")
        self.tree_scroll_x.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(
            self.tree_frame, 
            show="headings", 
            yscrollcommand=self.tree_scroll_y.set, 
            xscrollcommand=self.tree_scroll_x.set
        )
        self.tree.pack(fill="both", expand=True)
        
        self.tree_scroll_y.configure(command=self.tree.yview)
        self.tree_scroll_x.configure(command=self.tree.xview)


    # --- LOGIC METHODS (Adatped from original) ---

    def toggle_theme(self):
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")

    def cargar_archivos_disponibles(self):
        base_path = self.get_data_path()
        carpeta_archivos = os.path.join(base_path, "archivos_excel")
        if not os.path.exists(carpeta_archivos):
            try: os.makedirs(carpeta_archivos)
            except: pass
            
        patron = os.path.join(carpeta_archivos, "*.xlsx")
        self.archivos_excel = glob.glob(patron)
        if self.archivos_excel:
            self.lbl_current_file.configure(text=f"{len(self.archivos_excel)} archivos encontrados en la biblioteca.")
        else:
            self.lbl_current_file.configure(text="No se encontraron archivos en la carpeta.")

    def cargar_archivo(self):
        archivo_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
            initialdir=os.path.join(self.get_data_path(), "archivos_excel"),
        )
        if not archivo_path: return

        try:
            self.df = pd.read_excel(archivo_path, engine="openpyxl", header=1)
            self.archivo_seleccionado = os.path.basename(archivo_path)
            
            if "Prestación" in self.df.columns:
                self.prestaciones = sorted(self.df["Prestación"].unique().astype(str))
                self.combo_prestacion.configure(values=self.prestaciones)
                self.combo_prestacion.set("")
                self.lbl_current_file.configure(text=f"Cargado: {self.archivo_seleccionado}")
                
                # Actualizar estadísticas iniciales
                self.stat_total_var.set(f"{len(self.df):,}")
                self.stat_filtro_var.set("0")
                self.stat_perc_var.set("0%")
                
                self.btn_save.configure(state="normal")
            else:
                 messagebox.showerror("Error", "Columna 'Prestación' no encontrada.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def filtrar_prestaciones_evento(self, event):
        texto = self.txt_search.get().lower()
        if not texto:
            self.combo_prestacion.configure(values=self.prestaciones)
        else:
            values = [p for p in self.prestaciones if texto in p.lower()]
            self.combo_prestacion.configure(values=values)
            if len(values) > 0:
                self.combo_prestacion.set(values[0])

    def buscar_prestacion(self):
        prestacion = self.combo_prestacion.get()
        if not prestacion or self.df is None: return

        filtro = self.df["Prestación"].astype(str) == prestacion
        self.resultado_filtrado = self.df[filtro]
        
        self.mostrar_resultados()
        
    def mostrar_resultados(self):
        # Clear
        for item in self.tree.get_children(): self.tree.delete(item)
        
        if self.resultado_filtrado is None: return

        # Columns
        if self.columnas_seleccionadas:
            cols = [c for c in self.columnas_seleccionadas if c in self.resultado_filtrado.columns]
            df_show = self.resultado_filtrado[cols].copy()
        else:
            cols = list(self.resultado_filtrado.columns)
            df_show = self.resultado_filtrado.copy()

        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for idx, row in df_show.iterrows():
            valores = [str(val) for val in row.values]
            self.tree.insert("", "end", values=valores)
            
        # Actualizar Estadísticas
        total = len(self.df) if self.df is not None else 0
        filtrados = len(df_show)
        porcentaje = (filtrados / total * 100) if total > 0 else 0
        
        self.stat_filtro_var.set(f"{filtrados:,}")
        self.stat_perc_var.set(f"{porcentaje:.1f}%")

    def limpiar_resultados(self):
        self.resultado_filtrado = None
        for item in self.tree.get_children(): self.tree.delete(item)
        self.combo_prestacion.set("")
        self.txt_search.delete(0, 'end')
        
        # Reset stats
        self.stat_filtro_var.set("0")
        self.stat_perc_var.set("0%")

    def guardar_resultado(self):
        if self.resultado_filtrado is None or len(self.resultado_filtrado) == 0:
             # Logic to save all if no filter
             if self.df is not None:
                 if messagebox.askyesno("¿Guardar Todo?", "No hay filtro aplicado. ¿Deseas guardar TODOS los datos?"):
                     self.guardar_df(self.df, "COMPLETO")
             return

        self.guardar_df(self.resultado_filtrado, self.combo_prestacion.get())

    def guardar_df(self, dataframe, suffix):
         try:
            archivo_base = os.path.splitext(self.archivo_seleccionado)[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            clean_suffix = suffix.replace(' ', '_').replace('/', '_')
            nombre = f"VS_{archivo_base}_{clean_suffix}_{timestamp}.xlsx"
            
            docs = os.path.join(os.path.expanduser("~"), "Documents")
            out_dir = os.path.join(docs, "Vidasalud_Export")
            if not os.path.exists(out_dir): os.makedirs(out_dir)
            
            path = os.path.join(out_dir, nombre)
            dataframe.to_excel(path, index=False)
            messagebox.showinfo("Guardado", f"Archivo guardado en:\n{path}")
         except Exception as e:
             messagebox.showerror("Error", str(e))

    def configurar_columnas(self):
        if self.df is None: return
        
        # Simple pop-up window using CTk
        pop = ctk.CTkToplevel(self.root)
        pop.title("Configurar Columnas")
        pop.geometry("400x500")
        
        lbl = ctk.CTkLabel(pop, text="Seleccionar Columnas", font=ctk.CTkFont(weight="bold"))
        lbl.pack(pady=10)
        
        scroll = ctk.CTkScrollableFrame(pop)
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.check_vars = {}
        for col in self.df.columns:
            var = ctk.BooleanVar(value=(col in self.columnas_seleccionadas or not self.columnas_seleccionadas))
            self.check_vars[col] = var
            chk = ctk.CTkCheckBox(scroll, text=col, variable=var)
            chk.pack(anchor="w", pady=2)
            
        def apply():
            self.columnas_seleccionadas = [col for col, var in self.check_vars.items() if var.get()]
            if self.resultado_filtrado is not None:
                self.mostrar_resultados()
            pop.destroy()
            
        ctk.CTkButton(pop, text="Aplicar Cambios", command=apply).pack(pady=10)

def main():
    app = ctk.CTk()
    gui = FiltradorMultiArchivosGUI(app)
    app.mainloop()

if __name__ == "__main__":
    main()

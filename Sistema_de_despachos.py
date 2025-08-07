import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, NamedStyle, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font as tkfont
from tkinter.simpledialog import Dialog
import json
import os
import sys
import subprocess
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Any
import traceback

def excepthook(exc_type, exc_value, exc_traceback):
    error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    print("Error detallado:", error_msg)  # Esto aparecerá en la consola
    with open("error_log.txt", "a") as f:
        f.write(error_msg)
    messagebox.showerror("Error crítico", f"Ocurrió un error:\n\n{error_msg}")

# Constantes para configuración
DEFAULT_CONFIG = {
    "ultimo_archivo": "",
    "tipo_articulos": ["Importado", "Nacional"],
    "columnas_clientes": ["Nombre", "RIF", "Teléfono", "Dirección"],
    "columnas_articulos": ["Código Importado", "Descripción Importado", 
                          "Código Nacional", "Descripción Nacional"],
    "formato_fecha": "%d/%m/%Y",
    "mostrar_pesos_individuales": True,
    "color_primario": "#0765a3",
    "color_secundario": "#1f924f",
    "color_fondo": "#f8f9fa"
}

# Modificaciones en la clase CalculadoraPesoDialog
class CalculadoraPesoDialog(tk.Toplevel):
    """Diálogo flexible para calcular peso de bultos con botones fijos"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Cálculo de Peso de Bultos")
        self.geometry("600x600")
        self.resizable(True, True)
        
        self.pesos = {}  # {número_bulto: peso}
        self.total = 0.0
        self.result = None
        self.entries = []
        self.max_bultos = 50  # Límite de 50 bultos
        
        # Configurar el diálogo para no cerrarse con Enter
        self.bind('<Return>', lambda e: None)
        
        # Crear la interfaz
        self._crear_interfaz()
        
        # Configurar comportamiento al cerrar
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.grab_set()
    
    def _validar_entrada_peso(self, valor: str) -> bool:
        """Valida que la entrada sea un número decimal positivo o vacío"""
        if not valor:  # Permitir campo vacío
            return True
        try:
            # Verificar que sea un número positivo
            num = float(valor)
            return num >= 0
        except ValueError:
            return False
    
    def _crear_interfaz(self):
        """Crea todos los elementos de la interfaz"""
        # Frame principal con configuración de grid
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame para el contenido con scroll
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame para botones (fijo a la derecha)
        button_frame = ttk.Frame(main_frame, width=120)
        button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        
        # Canvas para scroll
        canvas = tk.Canvas(content_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Empaquetar widgets de scroll
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Título e instrucciones
        ttk.Label(
            scrollable_frame,
            text="Ingrese pesos de bultos (kg)",
            font=('Segoe UI', 12, 'bold')
        ).pack(pady=(0, 10))
        
        ttk.Label(
            scrollable_frame,
            text="Complete los pesos necesarios (máximo 50 bultos)",
            font=('Segoe UI', 9),
            foreground='#666666'
        ).pack(pady=(0, 15))
        
        # Contenedor para los controles de peso
        entry_frame = ttk.Frame(scrollable_frame)
        entry_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear campos para hasta max_bultos bultos
        for i in range(1, self.max_bultos + 1):
            frame = ttk.Frame(entry_frame, padding=2)
            frame.pack(fill=tk.X, pady=1)
            
            ttk.Label(frame, text=f"Bulto {i}:", width=8).pack(side=tk.LEFT)
            
            entry_var = tk.StringVar()
            entry = ttk.Entry(
                frame,
                width=10,
                textvariable=entry_var,
                justify=tk.RIGHT,
                validate="key",
                validatecommand=(frame.register(self._validar_entrada_peso), '%P'))
            entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(frame, text="kg").pack(side=tk.LEFT)
            
            # Configurar eventos
            entry.bind("<KeyRelease>", lambda e, idx=i: self._actualizar_peso(idx))
            entry.bind("<Return>", self._manejar_enter)
            entry.bind("<Down>", self._manejar_flecha_abajo)
            entry.bind("<Up>", self._manejar_flecha_arriba)
            
            self.entries.append((entry, entry_var))
        
        # Sección de total
        ttk.Label(
            button_frame,
            text="Peso Total:",
            font=('Segoe UI', 10, 'bold')
        ).pack(pady=(10, 5))
        
        self.total_label = ttk.Label(
            button_frame,
            text="0.00 kg",
            font=('Segoe UI', 10)
        )
        self.total_label.pack(pady=(0, 20))
        
        # Botones
        ttk.Button(
            button_frame,
            text="Aceptar",
            command=self._on_accept,
            style='Accent.TButton',
            width=15
        ).pack(pady=5)
        
        ttk.Button(
            button_frame,
            text="Cancelar",
            command=self._on_cancel,
            style='Secondary.TButton',
            width=15
        ).pack(pady=5)
        
        # Ajustar canvas
        scrollable_frame.update_idletasks()
        canvas.config(width=scrollable_frame.winfo_reqwidth())
        
        # Enfocar primer campo
        if self.entries:
            self.entries[0][0].focus_set()
    
    def _actualizar_peso(self, bulto_num: int):
        """Actualiza el peso de un bulto específico, solo si tiene valor"""
        entry, entry_var = self.entries[bulto_num - 1]
        valor = entry_var.get().strip()
        
        if valor:  # Solo actualizar si hay un valor
            try:
                peso = float(valor)
                if peso >= 0:
                    self.pesos[bulto_num] = peso
                else:
                    self.pesos.pop(bulto_num, None)  # Eliminar si es negativo
            except ValueError:
                self.pesos.pop(bulto_num, None)  # Eliminar si no es número válido
        else:
            self.pesos.pop(bulto_num, None)  # Eliminar si está vacío
            
        self._calcular_total()
    
    def _calcular_total(self):
        """Calcula el peso total de todos los bultos y la cantidad real de bultos con peso"""
        self.total = sum(self.pesos.values())
        cantidad_bultos = sum(1 for peso in self.pesos.values() if peso > 0)  # Solo contar bultos con peso > 0
        self.total_label.config(text=f"{self.total:.2f} kg\nBultos: {cantidad_bultos}")
    def _manejar_enter(self, event):
        """Maneja la tecla Enter moviendo el foco al siguiente campo"""
        current = self.focus_get()
        for i, (entry, _) in enumerate(self.entries):
            if entry == current and i < len(self.entries) - 1:
                self.entries[i + 1][0].focus_set()
                break
    
    def _manejar_flecha_abajo(self, event):
        """Maneja la flecha hacia abajo"""
        self._manejar_enter(event)
    
    def _manejar_flecha_arriba(self, event):
        """Maneja la flecha hacia arriba"""
        current = self.focus_get()
        for i, (entry, _) in enumerate(self.entries):
            if entry == current and i > 0:
                self.entries[i - 1][0].focus_set()
                break
    
    def _on_accept(self):
        """Maneja el botón Aceptar, solo incluyendo bultos con peso"""
        # Filtrar solo bultos con peso > 0
        pesos_validos = {k: v for k, v in self.pesos.items() if v > 0}
        self.result = (pesos_validos, sum(pesos_validos.values()))
        self.destroy()
    
    def _on_cancel(self):
        """Maneja el botón Cancelar"""
        self.result = None
        self.destroy()
        

class AplicacionDespachos:
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.config = {}
        self._configurar_ventana_principal()
        self._inicializar_datos()
        self._inicializar_estilos()
        
        # Ahora que la ventana raíz está configurada, podemos crear variables Tkinter
        self.peso_total_var = tk.StringVar(value="Peso Total: 0.00 kg")
        self.tipo_var = tk.StringVar()
        
        self._crear_interfaz()
        self._cargar_ultimo_archivo()

        self.datos_sin_guardar = False
        self.cantidad_temp = ""
        self.peso_temp = ""
        self.obs_temp = ""
        self.articulo_seleccionado_temp = None

    def _configurar_ventana_principal(self):
        """Configura los parámetros iniciales de la ventana principal"""
        self.root.title("Sistema de Gestión de Despachos")
        self.root.geometry("1280x800")
        self.root.minsize(1200, 750)
        self.root.protocol("WM_DELETE_WINDOW", self._confirmar_salida)
        
        # Configurar icono si existe
        icon_path = os.path.join(os.path.dirname(__file__), "icono.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

    def _inicializar_datos(self):
        """Inicializa las variables de datos y configuración"""
        self.config_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), 
            'config_despachos.json')
        self.config = self._cargar_configuracion()
        
        # Variables de datos
        self.excel_path = ""
        self.clientes_df = pd.DataFrame()
        self.articulos_df = pd.DataFrame()
        self.cliente_actual: Optional[Dict[str, str]] = None
        self.bultos_data: Dict[str, Dict[str, float]] = {}
        
        # Inicializar widgets como None
        self._inicializar_widgets()

    def _inicializar_estilos(self):
        """Configura los estilos visuales de la aplicación"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configurar paleta de colores desde la configuración
        self.colores = {
            'primario': self._obtener_color_config('color_primario', '#0765a3'),
            'secundario': self._obtener_color_config('color_secundario', '#1f924f'),
            'fondo': self._obtener_color_config('color_fondo', '#f8f9fa'),
            'texto': '#2c3e50',
            'exito': '#27ae60',
            'error': '#e74c3c'
        }
        
        # Configurar estilos base
        style.configure('.', 
                      font=('Segoe UI', 10),
                      background=self.colores['fondo'])
        
        # Estilos para botones
        button_config = {
            'font': ('Segoe UI', 10, 'bold'),
            'padding': 6,
            'borderwidth': 1
        }
        
        style.configure('Accent.TButton',
                      foreground='white',
                      background=self.colores['primario'],
                      **button_config)
        
        style.configure('Secondary.TButton',
                      foreground='white',
                      background=self.colores['secundario'],
                      **button_config)
        
        # Estilo para campos inválidos
        style.configure('Error.TEntry',
                      foreground='black',
                      fieldbackground='#ffdddd',
                      borderwidth=1)
        
        # Estilo para Treeview
        style.configure('Custom.Treeview', 
                      rowheight=30,
                      fieldbackground='white',
                      background='white')
        
        style.configure('Custom.Treeview.Heading', 
                      font=('Segoe UI', 10, 'bold'),
                      background=self.colores['primario'],
                      foreground='white')
    
    def _confirmar_salida(self):
        """Cierra la aplicación con confirmación si hay datos sin guardar"""
        if hasattr(self, 'tree') and self.tree.get_children():
            respuesta = messagebox.askyesnocancel(
                "Confirmar Salida",
                "Tiene un despacho en progreso con artículos no guardados.\n\n"
                "¿Desea guardar antes de salir?",
                icon=messagebox.WARNING,
                default=messagebox.YES
            )
            
            if respuesta is None:  # Cancelar
                return
            elif respuesta:  # Sí, guardar
                if not self._guardar_despacho():
                    return  # No salir si el guardado falla
        
        # Guardar configuración antes de salir
        try:
            self._guardar_configuracion()
        except Exception as e:
            messagebox.showerror(
                "Error", 
                f"No se pudo guardar la configuración:\n{str(e)}\n"
                "La aplicación se cerrará de todos modos.")
        
        self.root.destroy()

    def _crear_seccion_tejido(self, parent):
        """Eliminada - La sección de tejido ya está incluida en la sección del cliente"""
        pass
        
    def _obtener_color_config(self, clave: str, default: str) -> str:
        """Obtiene un color de la configuración validando su formato"""
        if not hasattr(self, 'config'):
            return default
        color = self.config.get(clave, default)
        return color if self._es_color_valido(color) else default

    def _es_color_valido(self, color: Any) -> bool:
        """Valida que un valor sea un color hexadecimal válido"""
        if not hasattr(self, 'config'):
            return False
        if not isinstance(color, str):
            return False
        if color.startswith('#') and len(color) in (4, 7):
            try:
                int(color[1:], 16)
                return True
            except ValueError:
                pass
        return False

    def _validar_archivo_excel(self, filepath):
        """Valida que el archivo Excel tenga la estructura requerida"""
        try:
            with pd.ExcelFile(filepath) as xls:
                # Verificar existencia de hojas requeridas
                hojas_requeridas = ['Clientes', 'ARTICULOS']
                hojas_faltantes = [hoja for hoja in hojas_requeridas if hoja not in xls.sheet_names]
                
                if hojas_faltantes:
                    return False, f"Faltan hojas requeridas: {', '.join(hojas_faltantes)}"
                
                # Verificar que la hoja ARTICULOS tenga al menos 4 columnas
                df_articulos = pd.read_excel(xls, sheet_name='ARTICULOS', nrows=1)
                if len(df_articulos.columns) < 4:
                    return False, "La hoja 'ARTICULOS' debe tener al menos 4 columnas"
                
                return True, ""
        except Exception as e:
            return False, f"Error al validar archivo: {str(e)}"
    
    def _inicializar_widgets(self):
        """Inicializa todos los widgets como None"""
        widgets = [
            'cantidad_entry', 'tipo_combobox', 'peso_entry', 'obs_entry',
            'cliente_search', 'clientes_listbox', 'cliente_nombre',
            'cliente_rif', 'cliente_telefono', 'cliente_direccion',
            'tree', 'peso_total_var', 'btn_agregar', 'btn_guardar',
            'btn_exportar', 'tejido_entry', 'codigo_tejido_entry'
        ]
        for widget in widgets:
            setattr(self, widget, None) 
    def _cargar_configuracion(self) -> Dict[str, Any]:
        """Carga la configuración desde archivo JSON o crea una nueva con valores por defecto"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Asegurar que todas las claves necesarias existan
                    return {**DEFAULT_CONFIG, **config}
            
            # Crear archivo de configuración si no existe
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)
            return DEFAULT_CONFIG.copy()
            
        except Exception as e:
            self._mostrar_error(
                "Error de Configuración",
                f"No se pudo cargar la configuración:\n{str(e)}\nSe usarán valores por defecto.")
            self._registrar_error(e)
            return DEFAULT_CONFIG.copy()

    def _guardar_configuracion(self) -> bool:
        """Guarda la configuración actual en el archivo JSON"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            self._mostrar_error("Error", f"No se pudo guardar la configuración:\n{str(e)}")
            self._registrar_error(e)
            return False

    def _crear_interfaz(self):
        """Construye la interfaz gráfica principal usando pack()"""
        self._crear_menu_principal()
        
        # Frame principal con pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Pestaña de Despacho
        self.despacho_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.despacho_frame, text="Gestión de Despachos")
        
        # Secciones usando pack()
        self._crear_seccion_cliente(self.despacho_frame)
        self._crear_seccion_articulos(self.despacho_frame)
        self._crear_botones_accion(self.despacho_frame)
        
        # Barra de estado
        self.statusbar = ttk.Label(self.root, text="Listo", relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

    def _nuevo_despacho(self):
        """Limpia el formulario para un nuevo despacho"""
        if hasattr(self, 'tree') and self.tree.get_children():
            respuesta = messagebox.askyesnocancel(
                "Nuevo Despacho",
                "Tiene un despacho en progreso con artículos no guardados.\n\n"
                "¿Desea guardar antes de limpiar el formulario?",
                icon=messagebox.WARNING,
                default=messagebox.YES
            )
            
            if respuesta is None:  # Cancelar
                return
            elif respuesta:  # Sí, guardar
                if not self._guardar_despacho():
                    return  # No continuar si el guardado falla
        
        # Limpiar campos del cliente
        for widget in [self.cliente_search, self.cliente_nombre, 
                    self.cliente_rif, self.cliente_telefono, 
                    self.cliente_direccion]:
            if hasattr(self, widget):
                widget.delete(0, tk.END)
                
        if hasattr(self, 'clientes_listbox'):
            self.clientes_listbox.selection_clear(0, tk.END)
            
        self.cliente_actual = None
        
        # Limpiar artículos
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)
        
        if hasattr(self, 'bultos_data'):
            self.bultos_data.clear()
            
        # Restablecer totales
        if hasattr(self, 'peso_total_var'):
            self.peso_total_var.set("Peso Total: 0.00 kg")
        
        # Restablecer tejido a valores por defecto
        if hasattr(self, 'tejido_entry'):
            self.tejido_entry.delete(0, tk.END)
            self.tejido_entry.insert(0, self.config.get("tejido_default", ""))
        
        if hasattr(self, 'codigo_tejido_entry'):
            self.codigo_tejido_entry.delete(0, tk.END)
            self.codigo_tejido_entry.insert(0, self.config.get("codigo_tejido_default", ""))
        
        # Enfocar búsqueda de cliente
        if hasattr(self, 'cliente_search'):
            self.cliente_search.focus()
        
        self._actualizar_estado("Listo para nuevo despacho")

    def _crear_menu_principal(self):
        """Crea la barra de menú principal"""
        menubar = tk.Menu(self.root)
        
        # Menú Archivo
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Nuevo Despacho", command=self._nuevo_despacho)
        file_menu.add_command(label="Abrir Archivo...", command=self._abrir_archivo_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Salir", command=self._confirmar_salida)
        menubar.add_cascade(label="Archivo", menu=file_menu)
        
        # Menú Despacho
        despacho_menu = tk.Menu(menubar, tearoff=0)
        despacho_menu.add_command(label="Agregar Artículo", command=self._agregar_articulo)
        despacho_menu.add_command(label="Eliminar Artículo", command=self._eliminar_articulo)
        despacho_menu.add_separator()
        despacho_menu.add_command(label="Calcular Peso Total", command=self._actualizar_peso_total)
        menubar.add_cascade(label="Despacho", menu=despacho_menu)
        
        # Menú Herramientas
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Calculadora de Peso", command=self._mostrar_calculadora_peso)
        menubar.add_cascade(label="Herramientas", menu=tools_menu)
        
        # Menú Ayuda
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Documentación", command=self._mostrar_documentacion)
        help_menu.add_command(label="Acerca de...", command=self._mostrar_acerca_de)
        menubar.add_cascade(label="Ayuda", menu=help_menu)
        
        self.root.config(menu=menubar)

    def _crear_seccion_cliente(self, parent):
        """Crea la sección de datos del cliente (versión simplificada)"""
        cliente_frame = ttk.LabelFrame(
            parent,
            text=" Datos del Cliente ",
            padding=(15, 10))
        cliente_frame.pack(fill=tk.BOTH, pady=5)
        
        # Búsqueda de cliente
        search_frame = ttk.Frame(cliente_frame)
        search_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(
            search_frame,
            text="Buscar Cliente:",
            font=('Segoe UI', 10, 'bold')
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        self.cliente_search = ttk.Entry(
            search_frame,
            width=40)
        self.cliente_search.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.cliente_search.bind('<Return>', lambda e: self._buscar_cliente())
        
        ttk.Button(
            search_frame,
            text="Buscar",
            command=self._buscar_cliente,
            style='Accent.TButton',
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        # Contenedor principal
        content_frame = ttk.Frame(cliente_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Lista de clientes
        list_frame = ttk.Frame(content_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self.clientes_listbox = tk.Listbox(
            list_frame,
            height=6,
            width=50,
            yscrollcommand=scroll_y.set,
            font=('Segoe UI', 10),
            selectbackground=self.colores['primario'])
        scroll_y.config(command=self.clientes_listbox.yview)
        
        self.clientes_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.clientes_listbox.bind('<<ListboxSelect>>', self._seleccionar_cliente)
        
        # Datos del cliente seleccionado
        data_frame = ttk.LabelFrame(
            content_frame,
            text=" Información del Cliente Seleccionado ",
            padding=15)
        data_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Campos de información del cliente
        campos = [
            ("Nombre:", 'cliente_nombre', 30),
            ("RIF:", 'cliente_rif', 15),
            ("Teléfono:", 'cliente_telefono', 15),
            ("Dirección:", 'cliente_direccion', 40)
        ]
        
        for i, (texto, attr, ancho) in enumerate(campos):
            ttk.Label(
                data_frame,
                text=texto,
                font=('Segoe UI', 10, 'bold')
            ).grid(row=i, column=0, sticky=tk.W, pady=5, padx=5)
            
            entry = ttk.Entry(
                data_frame,
                width=ancho)
            entry.grid(row=i, column=1, sticky=tk.W, pady=5, padx=5)
            setattr(self, attr, entry)
    def _crear_seccion_articulos(self, parent):
        """Crea la sección de artículos usando pack() consistentemente"""
        articulos_frame = ttk.LabelFrame(
            parent,
            text=" Artículos del Despacho ",
            padding=(15, 10))
        articulos_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Treeview y controles - todos con pack()
        self._crear_treeview_articulos(articulos_frame)
        self._crear_controles_articulos(articulos_frame)

    def _crear_treeview_articulos(self, parent):
        """Crea el Treeview con columnas para los bultos (hasta 50 bultos) usando pack()"""
        # Frame contenedor principal
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Frame interno para el treeview y scrollbar vertical
        tree_container = ttk.Frame(tree_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar vertical
        scroll_y = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        
        # Columnas principales + suficientes para bultos
        columns = ['Cantidad', 'Descripción', 'Peso Total (kg)'] + [f'Bulto {i} (kg)' for i in range(1, 51)]
        
        # Treeview
        self.tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show='headings',
            height=8,
            selectmode='extended',
            yscrollcommand=scroll_y.set,
            style='Custom.Treeview'
        )
        
        # Configurar columnas
        self.tree.heading('Cantidad', text='Cantidad', anchor=tk.CENTER)
        self.tree.column('Cantidad', width=80, anchor=tk.CENTER, stretch=False)
        
        self.tree.heading('Descripción', text='Descripción', anchor=tk.W)
        self.tree.column('Descripción', width=300, anchor=tk.W, stretch=False)
        
        self.tree.heading('Peso Total (kg)', text='Peso Total (kg)', anchor=tk.CENTER)
        self.tree.column('Peso Total (kg)', width=100, anchor=tk.CENTER, stretch=False)
        
        # Columnas para bultos
        for i in range(1, 51):
            col_name = f'Bulto {i} (kg)'
            self.tree.heading(col_name, text=col_name, anchor=tk.CENTER)
            self.tree.column(col_name, width=80, anchor=tk.CENTER, stretch=False)
        
        # Scrollbar horizontal
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(xscrollcommand=scroll_x.set)
        
        # Posicionamiento con pack()
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_x.pack(fill=tk.X)
        
        # Configurar eventos de edición
        self.tree.bind('<Double-1>', self._iniciar_edicion_celda)
        self.tree.bind('<Return>', self._iniciar_edicion_celda)
        
        # Variable para controlar la edición
        self.editing_cell = None


    def _iniciar_edicion_celda(self, event):
        """Inicia la edición de una celda (cantidad o bultos)"""
        region = self.tree.identify("region", event.x, event.y)
        if region not in ("cell", "tree"):
            return
        
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        # Permitir editar cantidad (columna 0) o bultos (columnas 3+)
        col_num = int(column[1:])
        if col_num != 0 and (col_num < 3 or col_num > 53):  # Solo cantidad y bultos
            return
        
        x, y, width, height = self.tree.bbox(item, column)
        value = self.tree.item(item, 'values')[col_num-1]
        
        # Crear Entry para edición
        self.editing_cell = {
            'item': item,
            'column': column,
            'col_num': col_num,
            'entry': ttk.Entry(self.tree)
        }
        
        entry = self.editing_cell['entry']
        entry.place(x=x, y=y, width=width, height=height)
        
        # Limpiar " kg" si es columna de peso total
        display_value = value.replace(' kg', '') if col_num == 2 else value
        entry.insert(0, display_value)
        entry.select_range(0, tk.END)
        entry.focus()
        
        # Bind para finalizar edición
        entry.bind('<FocusOut>', lambda e: self._finalizar_edicion_celda())
        entry.bind('<Return>', lambda e: self._finalizar_edicion_celda())
        entry.bind('<Escape>', lambda e: self._cancelar_edicion_celda())

    def _finalizar_edicion_celda(self, event=None):
        """Finaliza la edición con reorganización completa de bultos"""
        if not self.editing_cell:
            return
        
        entry = self.editing_cell['entry']
        item = self.editing_cell['item']
        col_num = self.editing_cell['col_num']
        new_value = entry.get().strip()
        
        values = list(self.tree.item(item, 'values'))
        descripcion = values[1]
        codigo = descripcion.split(' - ')[0] if ' - ' in descripcion else ""

        try:
            # Edición de BULTO INDIVIDUAL (columnas 4+)
            if col_num >= 4:
                bulto_num = col_num - 3
                
                if codigo:
                    # Procesar el nuevo valor (0 para eliminar)
                    if new_value:
                        peso = float(new_value)
                        
                        # Eliminar si es 0, actualizar si tiene valor
                        if peso == 0:
                            if codigo in self.bultos_data and str(bulto_num) in self.bultos_data[codigo]:
                                del self.bultos_data[codigo][str(bulto_num)]
                                
                                # Reorganizar los bultos restantes para mantener secuencia continua
                                if codigo in self.bultos_data:
                                    # 1. Obtener todos los bultos existentes ordenados por número
                                    bultos_existentes = sorted(
                                        [(int(k), v) for k, v in self.bultos_data[codigo].items()],
                                        key=lambda x: x[0]
                                    )
                                    
                                    # 2. Crear nueva estructura con numeración secuencial desde 1
                                    nuevos_bultos = {}
                                    for new_pos, (_, peso) in enumerate(bultos_existentes, 1):
                                        nuevos_bultos[str(new_pos)] = peso
                                    
                                    # 3. Actualizar estructura de datos
                                    self.bultos_data[codigo] = nuevos_bultos
                        else:
                            if codigo not in self.bultos_data:
                                self.bultos_data[codigo] = {}
                            self.bultos_data[codigo][str(bulto_num)] = peso
                    
                    # Actualizar visualización completa
                    if codigo in self.bultos_data:
                        self._actualizar_fila_completa(item, codigo, values)
            
            # Actualizar Treeview
            self.tree.item(item, values=values)
            self._actualizar_peso_total()
            
        except ValueError:
            pass
        
        entry.destroy()
        self.editing_cell = None

    def _actualizar_fila_completa(self, item, codigo, values):
        """Actualiza toda la fila con la nueva estructura de bultos"""
        # Limpiar todos los bultos en la visualización
        for i in range(4, len(values)):
            values[i] = ""
        
        if codigo in self.bultos_data and self.bultos_data[codigo]:
            # Calcular nuevos totales
            pesos_validos = list(self.bultos_data[codigo].values())
            total_peso = sum(pesos_validos)
            cantidad_bultos = len(pesos_validos)
            
            values[0] = str(cantidad_bultos)
            values[2] = f"{total_peso:.2f}"
            
            # Mostrar bultos en sus nuevas posiciones (1, 2, 3...)
            for bulto, peso in sorted(self.bultos_data[codigo].items(), key=lambda x: int(x[0])):
                col_pos = 3 + int(bulto)
                if col_pos < len(values):
                    values[col_pos] = f"{peso:.2f}"
        else:
            # Si no quedan bultos, resetear valores
            values[0] = "0"
            values[2] = "0.00"

    def _actualizar_visualizacion_bultos(self, item, codigo, values):
        """Actualiza la visualización de los bultos en el Treeview"""
        # Limpiar todos los bultos en la visualización
        for i in range(4, len(values)):
            values[i] = ""
        
        # Si hay bultos, mostrarlos en orden
        if codigo in self.bultos_data and self.bultos_data[codigo]:
            # Calcular totales
            pesos_validos = list(self.bultos_data[codigo].values())
            total_peso = sum(pesos_validos)
            cantidad_bultos = len(pesos_validos)
            
            values[2] = f"{total_peso:.2f}"
            values[0] = str(cantidad_bultos)
            
            # Mostrar bultos en sus nuevas posiciones
            for bulto, peso in sorted(self.bultos_data[codigo].items(), key=lambda x: int(x[0])):
                col_pos = 3 + int(bulto)
                if col_pos < len(values):
                    values[col_pos] = f"{peso:.2f}"
        else:
            # Si no hay bultos, limpiar todo
            values[0] = "0"
            values[2] = "0.00"

    def _cancelar_edicion_celda(self):
        """Cancela la edición sin guardar cambios"""
        if self.editing_cell:
            self.editing_cell['entry'].destroy()
            self.editing_cell = None

    def _validar_entrada_peso(self, valor: str) -> bool:
        """Valida que la entrada sea un número decimal positivo o vacío"""
        if not valor:  # Permitir campo vacío
            return True
        try:
            # Verificar que sea un número positivo
            num = float(valor)
            return num >= 0
        except ValueError:
            return False
        
    def _editar_celda(self, event):
        """Permite editar el contenido de una celda con doble click (solo bultos)"""
        # Finalizar cualquier edición en curso
        if self.editing_cell:
            self._finalizar_edicion()
        
        # Identificar la celda clickeada
        region = self.tree.identify("region", event.x, event.y)
        if region not in ("cell", "tree"):
            return
        
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        # Solo permitir edición en columnas de bultos (columnas 4-53)
        if not column.startswith('#') or not column[1:].isdigit():
            return
        
        col_num = int(column[1:])
        if col_num < 4 or col_num > 53:  # Columnas de bultos (4-53)
            return
        
        # Obtener coordenadas y valor actual
        x, y, width, height = self.tree.bbox(item, column)
        values = self.tree.item(item, 'values')
        if not values or len(values) < col_num:
            return
        
        value = values[col_num-1]
        
        # Crear Entry para edición
        self.editing_cell = {
            'item': item,
            'column': column,
            'col_num': col_num,
            'entry': ttk.Entry(self.tree, 
                            justify='center',
                            validate='key',
                            validatecommand=(self.tree.register(self._validar_entrada_peso), '%P'))
        }
        
        entry = self.editing_cell['entry']
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, value.replace(' kg', '') if value else '')
        entry.select_range(0, tk.END)
        entry.focus()
        
        # Bind para finalizar edición
        entry.bind('<FocusOut>', lambda e: self._finalizar_edicion())
        entry.bind('<Return>', lambda e: self._finalizar_edicion())
        entry.bind('<Escape>', lambda e: self._cancelar_edicion())

    def _finalizar_edicion(self, event=None):
        """Finaliza la edición y actualiza los datos"""
        if not self.editing_cell:
            return
        
        entry = self.editing_cell['entry']
        item = self.editing_cell['item']
        column = self.editing_cell['column']
        col_num = self.editing_cell['col_num']
        
        # Obtener el nuevo valor
        new_value = entry.get().strip()
        
        # Validar que sea un número válido
        try:
            if new_value:  # Permitir vacío para eliminar un bulto
                float(new_value)
        except ValueError:
            messagebox.showerror("Error", "Ingrese un valor numérico válido", parent=self.root)
            entry.destroy()
            self.editing_cell = None
            return
        
        # Obtener valores actuales del item
        values = list(self.tree.item(item, 'values'))
        
        # Actualizar el valor en el Treeview
        values[col_num-1] = f"{float(new_value):.2f}" if new_value else ""
        self.tree.item(item, values=values)
        
        # Actualizar los datos internos del bulto
        bulto_num = col_num - 3  # 1-50 (columnas 4-53 son bultos 1-50)
        descripcion = values[1]
        
        # Extraer código del artículo (primera parte de la descripción)
        codigo = descripcion.split(' - ')[0] if ' - ' in descripcion else ""
        
        if codigo:
            # Inicializar diccionario si no existe
            if codigo not in self.bultos_data:
                self.bultos_data[codigo] = {}
            
            if new_value:
                self.bultos_data[codigo][str(bulto_num)] = float(new_value)
            else:
                self.bultos_data[codigo].pop(str(bulto_num), None)
        
        # Recalcular peso total del artículo
        self._actualizar_peso_articulo(item)
        
        # Limpiar
        entry.destroy()
        self.editing_cell = None

    def _actualizar_peso_articulo(self, item):
        """Actualiza peso total y cantidad de bultos para un artículo"""
        values = list(self.tree.item(item, 'values'))
        descripcion = values[1]
        codigo = descripcion.split(' - ')[0] if ' - ' in descripcion else ""
        
        if codigo and codigo in self.bultos_data:
            # Calcular peso total
            total = sum(self.bultos_data[codigo].values())
            values[2] = f"{total:.2f}"
            
            # Actualizar cantidad automáticamente
            values[0] = str(len(self.bultos_data[codigo]))
            
            self.tree.item(item, values=values)
            self._actualizar_peso_total()

    def _cancelar_edicion(self):
        """Cancela la edición sin guardar cambios"""
        if self.editing_cell:
            self.editing_cell['entry'].destroy()
            self.editing_cell = None

    def _actualizar_peso_articulo(self, item):
        """Recalcula el peso total de un artículo basado en sus bultos"""
        values = list(self.tree.item(item, 'values'))
        descripcion = values[1]
        
        # Extraer código del artículo
        codigo = descripcion.split(' - ')[0] if ' - ' in descripcion else ""
        
        if codigo and codigo in self.bultos_data:
            # Sumar pesos de todos los bultos
            total = sum(self.bultos_data[codigo].values())
            values[2] = f"{total:.2f}"
            self.tree.item(item, values=values)

    def _crear_controles_articulos(self, parent):
        """Crea los controles usando pack()"""
        control_frame = ttk.Frame(parent, padding=(10, 15, 10, 10))
        control_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Frame para tipo de artículo
        tipo_frame = ttk.Frame(control_frame)
        tipo_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
        ttk.Label(
            tipo_frame,
            text="Tipo de Artículo:",
            font=('Segoe UI', 10, 'bold')
        ).pack(side=tk.LEFT, padx=(0, 10))
            
        self.tipo_var = tk.StringVar()
        self.tipo_combobox = ttk.Combobox(
            tipo_frame,
            width=15,
            textvariable=self.tipo_var,
            values=["Importado", "Nacional"],  # Valores fijos según la imagen
            state="readonly")
        self.tipo_combobox.pack(side=tk.LEFT, padx=5)
        self.tipo_combobox.current(0)
            
        # Frame para botones
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side=tk.RIGHT)
            
        ttk.Button(
            button_frame,
            text="Agregar Artículo",
            command=self._agregar_articulo,
            style='Accent.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)
            
        ttk.Button(
            button_frame,
            text="Eliminar Seleccionado",
            command=self._eliminar_articulo,
            style='Secondary.TButton',
            width=18
        ).pack(side=tk.LEFT, padx=5)
        
        # Peso total (simplificado)
        self.peso_total_var = tk.StringVar(value="Peso Total: 0.00 kg")
        ttk.Label(
            control_frame,
            textvariable=self.peso_total_var,
            font=('Segoe UI', 11, 'bold'),
            foreground=self.colores['primario']
        ).pack(side=tk.RIGHT, padx=10)

    def _crear_botones_accion(self, parent):
        """Crea los botones de acción principales con la nueva opción de exportación detallada"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Botones izquierdos
        left_frame = ttk.Frame(button_frame)
        left_frame.pack(side=tk.LEFT, expand=True)
        
        ttk.Button(
            left_frame,
            text="Guardar Despacho",
            command=self._guardar_despacho,
            style='Accent.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            left_frame,
            text="Nuevo Despacho",
            command=self._nuevo_despacho,
            style='Secondary.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        # Botones derechos
        right_frame = ttk.Frame(button_frame)
        right_frame.pack(side=tk.RIGHT)
        
        ttk.Button(
            right_frame,
            text="Exportar Normal",
            command=self._exportar_excel,
            style='Accent.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            right_frame,
            text="Exportar Detallado",
            command=self._exportar_despacho_detallado,
            style='Accent.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            right_frame,
            text="Salir",
            command=self._confirmar_salida,
            style='Secondary.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)

    def _habilitar_interfaz(self, habilitar: bool):
        """Habilita o deshabilita los controles según el estado"""
        widgets = [
            self.tipo_combobox, self.btn_agregar,
            self.btn_guardar, self.btn_exportar,
            self.tejido_entry, self.codigo_tejido_entry
        ]
        
        state = tk.NORMAL if habilitar else tk.DISABLED
        for widget in widgets:
            if widget:
                widget['state'] = state
    
    def _cargar_ultimo_archivo(self):
        """Carga el último archivo Excel usado si existe en la configuración"""
        ultimo_archivo = self.config.get("ultimo_archivo", "")
        
        if ultimo_archivo and os.path.exists(ultimo_archivo):
            try:
                # Verificar que el archivo sea válido
                valido, mensaje = self._validar_archivo_excel(ultimo_archivo)
                if not valido:
                    raise ValueError(mensaje)
                
                self.excel_path = ultimo_archivo
                self._cargar_datos_excel()
                self._actualizar_estado("Archivo cargado: " + os.path.basename(ultimo_archivo))
                
            except Exception as e:
                self._mostrar_error(
                    "Error al cargar archivo",
                    f"No se pudo cargar el último archivo usado:\n{str(e)}")
                self._registrar_error(e)
                self.excel_path = ""
                self.config["ultimo_archivo"] = ""
                self._guardar_configuracion()

    def _cargar_datos_excel(self):
        """Carga los datos de clientes y artículos desde el archivo Excel con un rango ampliado"""
        if not self.excel_path:
            return
            
        try:
            # Cargar hoja de clientes (sin cambios)
            self.clientes_df = pd.read_excel(
                self.excel_path,
                sheet_name='Clientes',
                dtype=str,
                usecols=lambda x: x in ['Nombre', 'RIF', 'Teléfono', 'Dirección']
            ).fillna('')
            
            if self.clientes_df.empty:
                raise ValueError("La hoja 'Clientes' está vacía o no contiene datos válidos")
                
            # Cargar hoja de artículos con rango ampliado (hasta 800 filas)
            articulos_df = pd.read_excel(
                self.excel_path,
                sheet_name='ARTICULOS',
                dtype=str,
                header=1,  # Saltar la primera fila que son encabezados generales
                nrows=800  # Ampliar el rango a 800 filas
            ).fillna('')
                
            if articulos_df.empty:
                raise ValueError("La hoja 'ARTICULOS' está vacía o no contiene datos válidos")
                    
            # Procesar los artículos para crear un DataFrame unificado con más columnas
            articulos_list = []
            
            # Mapeo de columnas a tipos de artículos
            column_mapping = [
                # (columnas, tipo)
                ([0, 1], 'Hilado Nacional'),
                ([2, 3], 'Hilado Importado'),
                ([4, 5], 'Hilado Especial'),  # Nueva categoría si existe
                ([6, 7], 'Tejido Importado'),
                ([8, 9], 'Tejido Nacional'),
                ([10, 11], 'Tejido Especial')  # Nueva categoría si existe
            ]
            
            for cols, tipo in column_mapping:
                # Verificar que las columnas existan en el DataFrame
                if max(cols) < len(articulos_df.columns):
                    df_temp = articulos_df.iloc[:, cols].copy()
                    df_temp.columns = ['Codigo', 'Descripcion']
                    df_temp['Tipo'] = tipo
                    df_temp = df_temp[df_temp['Codigo'].str.strip().astype(bool)]  # Filtrar filas vacías
                    articulos_list.append(df_temp)
            
            # Combinar todos los artículos en un solo DataFrame
            self.articulos_df = pd.concat(articulos_list, ignore_index=True).drop_duplicates()
            
            # Actualizar interfaz
            self._actualizar_lista_clientes()
            self._habilitar_interfaz(True)
            
        except Exception as e:
            self._mostrar_error("Error al cargar datos", f"No se pudieron cargar los datos:\n{str(e)}")
            self._registrar_error(e)
            self._habilitar_interfaz(False)
            raise

    def _filtrar_articulos(self, busqueda: str, search_tree=None):
        """Filtra artículos basado en el texto de búsqueda y tipo seleccionado"""
        if search_tree is None:
            search_tree = self.search_tree
        if not search_tree or self.articulos_df.empty:
            return
        
        search_tree.delete(*search_tree.get_children())
        
        tipo = self.tipo_var.get()
        busqueda = busqueda.lower().strip()
        
        # Filtrar por tipo (Nacional o Importado)
        if tipo == "Nacional":
            mask = self.articulos_df['Tipo'].isin(['Hilado Nacional', 'Tejido Nacional'])
        else:  # "Importado"
            mask = self.articulos_df['Tipo'].isin(['Hilado Importado', 'Tejido Importado'])
        
        df_filtrado = self.articulos_df[mask].copy()
        
        # Aplicar filtro de búsqueda si existe
        if busqueda:
            mask = (
                df_filtrado['Codigo'].str.lower().str.contains(busqueda) | 
                df_filtrado['Descripcion'].str.lower().str.contains(busqueda)
            )
            df_filtrado = df_filtrado[mask]
        
        # Limitar resultados para mejor rendimiento
        df_filtrado = df_filtrado.head(200)
        
        # Insertar datos en el Treeview
        for i, (_, row) in enumerate(df_filtrado.iterrows(), 1):
            self.search_tree.insert(
                '', 
                tk.END,
                values=(row['Codigo'], row['Descripcion'], ""),
                tags=('oddrow' if i % 2 == 1 else 'evenrow')
            )

    def _actualizar_lista_clientes(self):
        """Actualiza la lista de clientes en el Listbox"""
        if not hasattr(self, 'clientes_listbox') or self.clientes_df.empty:
            return
            
        self.clientes_listbox.delete(0, tk.END)
        
        nombre_col = self.config["columnas_clientes"][0]
        rif_col = self.config["columnas_clientes"][1]
        
        # Ordenar y mostrar clientes
        clientes_ordenados = self.clientes_df.sort_values(by=[nombre_col, rif_col])
        
        for _, row in clientes_ordenados.iterrows():
            nombre = str(row[nombre_col]).strip()
            rif = str(row[rif_col]).strip()
            
            if nombre and rif:
                self.clientes_listbox.insert(tk.END, f"{nombre} - {rif}")
        
        self._actualizar_estado(f"Clientes cargados: {len(clientes_ordenados)}")

    def _buscar_cliente(self):
        """Busca clientes según el texto ingresado"""
        if not hasattr(self, 'cliente_search'):
            return
            
        busqueda = self.cliente_search.get().strip().lower()
        self.clientes_listbox.delete(0, tk.END)
        
        if not busqueda:
            self._actualizar_lista_clientes()
            return
            
        if not self.clientes_df.empty:
            nombre_col = self.config["columnas_clientes"][0]
            rif_col = self.config["columnas_clientes"][1]
            
            # Búsqueda flexible (coincidencias parciales)
            mask = (
                self.clientes_df[nombre_col].astype(str).str.lower().str.contains(busqueda) | 
                self.clientes_df[rif_col].astype(str).str.lower().str.contains(busqueda)
            )  # <-- Paréntesis de cierre correctamente ubicado

            resultados = self.clientes_df[mask].sort_values(by=[nombre_col, rif_col])

            for _, row in resultados.iterrows():
                nombre = str(row[nombre_col]).strip()
                rif = str(row[rif_col]).strip()
                self.clientes_listbox.insert(tk.END, f"{nombre} - {rif}")
                self._actualizar_estado(f"Clientes encontrados: {len(resultados)}")

    def _seleccionar_cliente(self, event):
        """Selecciona un cliente de la lista y muestra sus datos (versión simplificada)"""
        if not self.clientes_listbox.curselection():
            return
            
        seleccion = self.clientes_listbox.get(self.clientes_listbox.curselection())
        rif = seleccion.split(" - ")[-1].strip()
        
        if not self.clientes_df.empty:
            rif_col = self.config["columnas_clientes"][1]
            
            try:
                cliente = self.clientes_df[self.clientes_df[rif_col] == rif].iloc[0].to_dict()
                
                self.cliente_actual = {
                    "Nombre": cliente.get(self.config["columnas_clientes"][0], ""),
                    "RIF": cliente.get(self.config["columnas_clientes"][1], ""),
                    "Teléfono": cliente.get(self.config["columnas_clientes"][2], ""),
                    "Dirección": cliente.get(self.config["columnas_clientes"][3], "")
                }
                
                # Actualizar campos
                for attr, key in [
                    ('cliente_nombre', 'Nombre'),
                    ('cliente_rif', 'RIF'),
                    ('cliente_telefono', 'Teléfono'),
                    ('cliente_direccion', 'Dirección')
                ]:
                    widget = getattr(self, attr)
                    widget.delete(0, tk.END)
                    widget.insert(0, self.cliente_actual[key])
                
                self._habilitar_interfaz(True)
                self._actualizar_estado(f"Cliente seleccionado: {self.cliente_actual['Nombre']}")
                
            except IndexError:
                self._mostrar_error("Error", "No se encontraron datos para el cliente seleccionado")
            except Exception as e:
                self._mostrar_error("Error", f"Error al cargar datos del cliente:\n{str(e)}")
                self._registrar_error(e)

    
    def _seleccion_articulo_busqueda(self, peso_var: tk.StringVar):
        """Actualiza el campo de peso cuando se selecciona un artículo en la búsqueda"""
        try:
            selected = self._obtener_articulo_seleccionado_busqueda()
            if selected:
                # Aquí puedes agregar lógica para actualizar el peso_var si es necesario
                # Por ejemplo, si los artículos tienen pesos predefinidos:
                # peso_var.set("1.00")  # Peso por defecto
                pass
        except Exception as e:
            self._mostrar_error("Error", f"No se pudo procesar la selección:\n{str(e)}")
            
    def _agregar_articulo(self):
        """Muestra diálogo para buscar y agregar artículos con interfaz mejorada"""
        if not self.cliente_actual:
            self._mostrar_advertencia("Seleccione un cliente primero")
            return
            
        if self.articulos_df.empty:
            self._mostrar_advertencia("No hay datos de artículos cargados. Verifique el archivo Excel.")
            return
        
        # Variables para controlar si hay datos sin guardar
        self.datos_sin_guardar = False
        self.cantidad_temp = "1"
        self.peso_temp = ""
        self.obs_temp = ""
        self.articulo_seleccionado_temp = None
        
        # Crear diálogo de búsqueda
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Agregar Artículo ({self.tipo_var.get()})")
        dialog.geometry("1000x700")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Configurar grid principal
        dialog.grid_rowconfigure(1, weight=1)
        dialog.grid_columnconfigure(0, weight=1)

        # Frame de búsqueda
        search_frame = ttk.Frame(dialog, padding=10)
        search_frame.grid(row=0, column=0, sticky='ew')
        search_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(
            search_frame,
            text="Buscar Artículo:",
            font=('Segoe UI', 10, 'bold')
        ).grid(row=0, column=0, padx=(0, 10), sticky='w')

        search_var = tk.StringVar()
        search_entry = ttk.Entry(
            search_frame,
            textvariable=search_var,
            width=50
        )
        search_entry.grid(row=0, column=1, padx=5, sticky='ew')
        search_entry.focus()
        
        # Frame para resultados
        results_frame = ttk.Frame(dialog)
        results_frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        # Treeview para resultados
        columns = ["Código", "Descripción", "Peso Unitario"]
        self.search_tree = ttk.Treeview(
            results_frame,
            columns=columns,
            show='headings',
            height=20,
            selectmode='browse',
            style='Custom.Treeview'
        )
        self.search_tree.tag_configure('selected', background='#0078d7', foreground='white')

        # Configurar columnas
        col_widths = {'Código': 200, 'Descripción': 500, 'Peso Unitario': 100}
        for col in columns:
            self.search_tree.heading(col, text=col)
            self.search_tree.column(col, width=col_widths.get(col, 150), 
                                anchor=tk.W if col != 'Peso Unitario' else tk.CENTER)

        # Scrollbars
        y_scroll = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.search_tree.yview)
        x_scroll = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=self.search_tree.xview)
        self.search_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        # Posicionamiento
        self.search_tree.grid(row=0, column=0, sticky='nsew')
        y_scroll.grid(row=0, column=1, sticky='ns')
        x_scroll.grid(row=1, column=0, sticky='ew')

        # Frame de controles
        control_frame = ttk.Frame(dialog, padding=10)
        control_frame.grid(row=2, column=0, sticky='ew')
        control_frame.grid_columnconfigure(1, weight=1)
        control_frame.grid_columnconfigure(3, weight=1)
        control_frame.grid_columnconfigure(5, weight=1)

        # Campos de entrada
        cantidad_var = tk.StringVar(value="1")
        peso_var = tk.StringVar()
        obs_var = tk.StringVar()

        # Configuración de widgets
        ttk.Label(control_frame, text="Cantidad:").grid(row=0, column=0, padx=(0, 5), sticky='e')
        cantidad_entry = ttk.Entry(
            control_frame,
            textvariable=cantidad_var,
            width=8,
            validate="key",
            validatecommand=(dialog.register(self._validar_entero_positivo), '%P')
        )
        cantidad_entry.grid(row=0, column=1, padx=5, sticky='w')

        ttk.Label(control_frame, text="Peso Total (kg):").grid(row=0, column=2, padx=(10, 5), sticky='e')
        peso_entry = ttk.Entry(
            control_frame,
            textvariable=peso_var,
            width=10,
            validate="key",
            validatecommand=(dialog.register(self._validar_decimal_positivo), '%P')
        )
        peso_entry.grid(row=0, column=3, padx=5, sticky='w')

        ttk.Button(
            control_frame,
            text="Calcular Peso",
            command=lambda: self._calcular_peso_bultos(cantidad_var, peso_var),
            style='Secondary.TButton',
            width=12
        ).grid(row=0, column=4, padx=5)

        ttk.Label(control_frame, text="Observaciones:").grid(row=0, column=5, padx=(10, 5), sticky='e')
        obs_entry = ttk.Entry(
            control_frame,
            textvariable=obs_var,
            width=30
        )
        obs_entry.grid(row=0, column=6, padx=5, sticky='ew')

        # Frame de botones del diálogo
        dialog_button_frame = ttk.Frame(dialog, padding=10)
        dialog_button_frame.grid(row=3, column=0, sticky='e')

        def agregar_y_limpiar():
            self._agregar_desde_busqueda(
                cantidad_var.get(),
                peso_var.get(),
                obs_var.get())
            # Limpiar campos después de agregar
            cantidad_var.set("1")
            peso_var.set("")
            obs_var.set("")
            for item in self.search_tree.get_children():
                self.search_tree.item(item, tags=())
            self.datos_sin_guardar = False

        def marcar_datos_sin_guardar(*args):
            self.datos_sin_guardar = True
            self.cantidad_temp = cantidad_var.get()
            self.peso_temp = peso_var.get()
            self.obs_temp = obs_var.get()
            selected = self.search_tree.selection()
            if selected:
                self.articulo_seleccionado_temp = self.search_tree.item(selected[0])['values'][:2]

        ttk.Button(
            dialog_button_frame,
            text="Agregar al Despacho",
            command=agregar_y_limpiar,
            style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            dialog_button_frame,
            text="Cancelar",
            command=dialog.destroy,
            style='Secondary.TButton'
        ).pack(side=tk.LEFT)

        # Configuración de eventos - MODIFICACIÓN PRINCIPAL AQUÍ
        def on_tree_select(event=None):
            selected = self.search_tree.selection()
            if selected:
                item = self.search_tree.item(selected[0])
                self.search_tree.tag_configure('selected', background='#0078d7', foreground='white')
                for i in self.search_tree.get_children():
                    self.search_tree.item(i, tags=())
                self.search_tree.item(selected[0], tags=('selected',))
                cantidad_entry.focus()

        self.search_tree.bind('<<TreeviewSelect>>', on_tree_select)
        # Solo permitir selección con click simple, no agregar automáticamente
        self.search_tree.bind('<Button-1>', lambda e: on_tree_select(e))
        
        search_entry.bind('<Return>', lambda e: self._filtrar_articulos(search_var.get(), self.search_tree))
        cantidad_entry.bind('<Return>', lambda e: peso_entry.focus())
        peso_entry.bind('<Return>', lambda e: obs_entry.focus())
        obs_entry.bind('<Return>', lambda e: agregar_y_limpiar())

        # Cargar datos iniciales
        self._filtrar_articulos("", self.search_tree)

        def on_close():
            if self.datos_sin_guardar:
                respuesta = messagebox.askyesno(
                    "Artículo no agregado",
                    "Tiene un artículo listo para agregar pero no lo ha guardado.\n\n"
                    "¿Desea agregarlo ahora?",
                    parent=dialog)
                if respuesta:
                    agregar_y_limpiar()
            dialog.destroy()
        
        dialog.protocol("WM_DELETE_WINDOW", on_close)

        # Configurar seguimiento de cambios
        cantidad_var.trace_add('write', marcar_datos_sin_guardar)
        peso_var.trace_add('write', marcar_datos_sin_guardar)
        obs_var.trace_add('write', marcar_datos_sin_guardar)
    
    def _validar_entero_positivo(self, valor: str) -> bool:
        """Valida que la entrada sea un entero positivo"""
        if not valor:
            return True
        return valor.isdigit() and int(valor) > 0

    def _validar_decimal_positivo(self, valor: str) -> bool:
        """Valida que la entrada sea un decimal positivo"""
        if not valor:
            return True
        try:
            return float(valor) > 0
        except ValueError:
            return False

    def _calcular_peso_bultos(self, cantidad_var: tk.StringVar, peso_var: tk.StringVar):
        """Calcula el peso total para múltiples bultos (versión optimizada)"""
        try:
            dialog = CalculadoraPesoDialog(self.root)
            self.root.wait_window(dialog)
            
            if dialog.result is not None:
                bultos_data, total = dialog.result
                
                # Actualizar el campo de peso total
                peso_var.set(f"{total:.2f}")
                
                # Actualizar la cantidad con el número REAL de bultos con peso
                cantidad_var.set(str(len(bultos_data)))
                
                # Guardar los pesos de bultos
                selected = self._obtener_articulo_seleccionado_busqueda()
                if selected:
                    codigo_articulo = selected[0]
                    self.bultos_data[codigo_articulo] = {str(k): v for k, v in bultos_data.items()}
                    
        except Exception as e:
            self._mostrar_error("Error", f"No se pudo calcular pesos:\n{str(e)}")
    def _obtener_articulo_seleccionado_busqueda(self) -> Optional[Tuple[str, str]]:
        """Obtiene el artículo seleccionado en el diálogo de búsqueda"""
        if not hasattr(self, 'search_tree'):
            return None
        
        seleccion = self.search_tree.selection()
        if not seleccion:
            self._mostrar_advertencia("Seleccione un artículo de la lista")
            return None
            
        item = self.search_tree.item(seleccion[0])
        return tuple(item['values'][:2])  # (Código, Descripción)

    def _agregar_desde_busqueda(self, cantidad: str, peso_total: str, observaciones: str):
        """Agrega un artículo con los pesos de los bultos al Treeview principal"""
        try:
            selected = self._obtener_articulo_seleccionado_busqueda()
            if not selected:
                return
                
            codigo_articulo, descripcion = selected
            
            # Obtener datos de bultos si existen
            bultos_data = self.bultos_data.get(codigo_articulo, {})
            
            # Crear lista de valores para el Treeview
            valores = [
                cantidad,
                f"{codigo_articulo} - {descripcion}",
                peso_total.replace(' kg', '') if peso_total else "0.00"
            ]
            
            # Agregar pesos individuales de bultos (hasta 50 bultos)
            for i in range(1, 51):  # Mostrar hasta 50 bultos
                peso = bultos_data.get(str(i), 0.0)  # Los bultos se guardan como strings
                valores.append(f"{peso:.2f}" if peso > 0 else "")
            
            # Insertar en el Treeview
            tags = ('evenrow',) if len(self.tree.get_children()) % 2 == 0 else ('oddrow',)
            self.tree.insert('', tk.END, values=valores, tags=tags)
            
            # Actualizar peso total
            self._actualizar_peso_total()
            
        except Exception as e:
            self._mostrar_error("Error", f"No se pudo agregar el artículo:\n{str(e)}")
            
    def _mostrar_detalle_articulo(self, event):
        """Muestra un diálogo con los detalles del artículo seleccionado"""
        seleccion = self.tree.selection()
        if not seleccion:
            return
            
        item = self.tree.item(seleccion[0])
        valores = item['values']
        
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Detalles del Artículo: {valores[2]}")
        dialog.geometry("500x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Frame principal
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Mostrar información básica
        campos = [
            ("Cantidad:", valores[0]),
            ("Tipo:", valores[1]),
            ("Código:", valores[2]),
            ("Descripción:", valores[3]),
            ("Peso Total (kg):", valores[4]),
            ("Observaciones:", valores[5])
        ]
        
        for i, (label, value) in enumerate(campos):
            ttk.Label(main_frame, text=label, font=('Segoe UI', 10, 'bold')).grid(
                row=i, column=0, sticky=tk.W, pady=5, padx=5)
            ttk.Label(main_frame, text=value, font=('Segoe UI', 10)).grid(
                row=i, column=1, sticky=tk.W, pady=5, padx=5)
        
        # Mostrar detalles de bultos si existen
        codigo_articulo = valores[2]
        if codigo_articulo in self.bultos_data:
            ttk.Separator(main_frame).grid(
                row=len(campos), column=0, columnspan=2, pady=10, sticky=tk.EW)
            
            ttk.Label(main_frame, text="Detalles de Bultos:", 
                     font=('Segoe UI', 10, 'bold')).grid(
                row=len(campos)+1, column=0, columnspan=2, sticky=tk.W, pady=5)
            
            for i, (bulto, peso) in enumerate(self.bultos_data[codigo_articulo].items(), start=len(campos)+2):
                ttk.Label(main_frame, text=f"Bulto {bulto}:", font=('Segoe UI', 9)).grid(
                    row=i, column=0, sticky=tk.E, padx=5)
                ttk.Label(main_frame, text=f"{peso} kg", font=('Segoe UI', 9)).grid(
                    row=i, column=1, sticky=tk.W, padx=5)
        
        # Botón para cerrar
        ttk.Button(main_frame, text="Cerrar", command=dialog.destroy,
                  style='Accent.TButton').grid(
            row=len(campos)+4, column=0, columnspan=2, pady=10)

    def _mostrar_calculadora_peso(self):
        """Muestra el diálogo para calcular peso de múltiples bultos"""
        try:
            # Verificar si tenemos los widgets necesarios
            if not hasattr(self, 'cantidad_entry'):
                self._mostrar_advertencia("No se puede acceder al campo de cantidad")
                return
                
            cantidad_str = self.cantidad_entry.get().strip()
            
            if not cantidad_str or not cantidad_str.isdigit():
                self._mostrar_advertencia("Ingrese una cantidad válida de bultos primero")
                return
                
            cantidad = int(cantidad_str)
            
            if cantidad <= 0:
                self._mostrar_advertencia("La cantidad debe ser mayor a cero")
                return
                
            # Mostrar el diálogo de cálculo de peso
            dialog = CalculadoraPesoDialog(self.root, cantidad)
            
            # Si se ingresaron valores, actualizar el campo de peso
            if dialog.result:
                pesos, total = dialog.result
                if hasattr(self, 'peso_entry'):
                    self.peso_entry.delete(0, tk.END)
                    self.peso_entry.insert(0, f"{total:.2f}")
                    
        except Exception as e:
            self._mostrar_error("Error en calculadora", f"Ocurrió un error:\n{str(e)}")
            self._registrar_error(e)

    def _validar_entero_positivo(self, valor: str) -> bool:
        """Valida que la entrada sea un entero positivo"""
        if not valor:
            return True
        return valor.isdigit() and int(valor) > 0

    def _actualizar_peso_total(self):
        """Calcula y muestra el peso total de todos los artículos"""
        try:
            total = 0.0
            
            for item in self.tree.get_children():
                try:
                    peso_str = self.tree.item(item)['values'][2].replace(' kg', '')  # Eliminar 'kg'
                    peso = float(peso_str)
                    total += peso
                except (ValueError, IndexError):
                    continue  # Ignorar artículos con peso inválido
            
            self.peso_total_var.set(f"Peso Total: {total:.2f} kg")
            
        except Exception as e:
            self._mostrar_error("Error", f"No se pudo calcular el peso total:\n{str(e)}")
            self._registrar_error(e)

    def _eliminar_articulo(self, event=None):
        """Elimina el artículo seleccionado del despacho"""
        seleccion = self.tree.selection()
        if not seleccion:
            self._mostrar_advertencia("Seleccione al menos un artículo para eliminar")
            return
            
        # Obtener descripciones de los artículos a eliminar
        articulos = [self.tree.item(item)['values'][3] for item in seleccion]
        mensaje = (f"¿Eliminar los siguientes {len(seleccion)} artículos?\n\n" +
                  "\n".join(f"- {art[:50]}{'...' if len(art) > 50 else ''}" for art in articulos))
        
        if not messagebox.askyesno("Confirmar Eliminación", mensaje):
            return
            
        # Eliminar los items seleccionados
        for item in seleccion:
            # Eliminar datos de bultos asociados si existen
            codigo = self.tree.item(item)['values'][2]
            if codigo in self.bultos_data:
                del self.bultos_data[codigo]
            self.tree.delete(item)
        
        # Reaplicar colores alternados
        for i, item in enumerate(self.tree.get_children()):
            tags = ('evenrow',) if i % 2 == 0 else ('oddrow',)
            self.tree.item(item, tags=tags)
        
        self._actualizar_peso_total()
        self._mostrar_info("Éxito", f"Se eliminaron {len(seleccion)} artículos")

    def _abrir_archivo_excel(self):
        """Permite al usuario seleccionar un archivo Excel con datos"""
        filetypes = [
            ("Archivos Excel", "*.xlsx *.xls"),
            ("Todos los archivos", "*.*")
        ]
        
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel de datos",
            initialdir=os.path.dirname(self.excel_path) if self.excel_path else os.path.expanduser("~"),
            filetypes=filetypes
        )
        
        if filepath:
            try:
                # Validar el archivo antes de cargarlo
                valido, mensaje = self._validar_archivo_excel(filepath)
                if not valido:
                    raise ValueError(mensaje)
                
                # Si pasa la validación, proceder con la carga
                self.excel_path = filepath
                self.config["ultimo_archivo"] = filepath
                    
                if not self._guardar_configuracion():
                    raise ValueError("No se pudo guardar la configuración")
                    
                self._cargar_datos_excel()
                self._habilitar_interfaz(True)
                    
                self._mostrar_info(
                    "Archivo cargado",
                    f"Se cargó correctamente:\n{os.path.basename(filepath)}")
                    
            except Exception as e:
                self._mostrar_error("Error", f"No se pudo cargar el archivo:\n{str(e)}")
                self._registrar_error(e)

    def _validar_despacho(self) -> bool:
        """Valida que el despacho esté completo antes de guardar"""
        errores = []
        
        if not self.cliente_actual:
            errores.append("- Seleccione un cliente primero")
            
        if not hasattr(self, 'tree') or not self.tree.get_children():
            errores.append("- Agregue al menos un artículo al despacho")
        
        # Validar pesos de artículos (ignorando 'kg')
        articulos_invalidos = []
        for item in self.tree.get_children():
            try:
                peso_str = self.tree.item(item)['values'][2].replace(' kg', '')  # Eliminar 'kg'
                peso = float(peso_str)
                if peso <= 0:
                    articulos_invalidos.append(self.tree.item(item)['values'][1])  # Usar descripción
            except (ValueError, IndexError):
                articulos_invalidos.append(self.tree.item(item)['values'][1])  # Usar descripción
        
        if articulos_invalidos:
            errores.append(
                f"- Los siguientes artículos tienen pesos inválidos:\n  "
                f"{', '.join(articulos_invalidos)}")
        
        # Resto de validaciones...
        if errores:
            self._mostrar_error(
                "Error en el despacho",
                "Corrija los siguientes errores antes de guardar:\n\n" +
                "\n".join(errores))
            return False
                
        return True

    def _guardar_despacho(self) -> bool:
        """Guarda el despacho en el archivo Excel principal con el nuevo formato"""
        if not self._validar_despacho():
            return False

        try:
            # Cargar o crear archivo Excel
            try:
                wb = load_workbook(self.excel_path) if os.path.exists(self.excel_path) else Workbook()
                if 'Despachos' not in wb.sheetnames:
                    wb.create_sheet('Despachos')
                ws = wb['Despachos']
            except Exception as e:
                raise ValueError(f"No se pudo abrir el archivo Excel: {str(e)}")

            # Encontrar primera fila vacía
            fila = self._encontrar_fila_vacia(ws)

            # ----------------------------
            # CABECERA (ESTILO IDÉNTICO AL FORMATO)
            # ----------------------------
            hoy = datetime.now()
            
            # Información fija de la empresa
            ws[f'B{fila}'] = "Domicilio Fiscal: Zona Industrial Los Naranjos, Av Maturin, Edif, Centro Industrial Ferro"
            ws[f'B{fila+1}'] = "Piso 1, Guarenas - Edo Miranda, Telfs.:(0212) 361.42.54 - Fax (0212) 361.20.31"
            ws[f'B{fila+2}'] = "e-mail: evertex@cantv.net / tejidosevertex@gmail.com"
            
            # Fecha de emisión (formato DD/MM/AAAA)
            ws[f'J{fila}'] = "FECHA DE EMISION"
            ws[f'L{fila}'] = hoy.day       # Día
            ws[f'M{fila}'] = hoy.month     # Mes
            ws[f'N{fila}'] = hoy.year      # Año
            ws[f'J{fila+1}'] = "GUARENAS"  # Texto fijo

            # Información del cliente
            ws[f'B{fila+6}'] = "DESTINATARIO/Nombre o Razon Social"
            ws[f'B{fila+7}'] = self.cliente_actual.get('Nombre', '')
            ws[f'I{fila+7}'] = f"RIF: {self.cliente_actual.get('RIF', '')}"
            ws[f'L{fila+7}'] = f"Telefonos(s): {self.cliente_actual.get('Teléfono', '')}"
            ws[f'B{fila+8}'] = "DIRECCION:"
            ws[f'B{fila+9}'] = self.cliente_actual.get('Dirección', '')

            # ----------------------------
            # TABLA DE PRODUCTOS (CON ESTILO)
            # ----------------------------
            encabezados = ["DESCRIPCION DEL PRODUCTO", "BULTOS", "CANTI.KILOS/METROS"]
            ws[f'B{fila+11}'] = encabezados[0]
            ws[f'I{fila+11}'] = encabezados[1]
            ws[f'J{fila+11}'] = encabezados[2]

            # Aplicar estilo a encabezados
            for col, header in zip(['B', 'I', 'J'], encabezados):
                cell = ws[f'{col}{fila+11}']
                cell.font = Font(bold=True, size=11, name='Arial')
                cell.alignment = Alignment(horizontal='center')
                cell.border = Border(
                    left=Side(style='medium'),
                    right=Side(style='medium'),
                    top=Side(style='medium'),
                    bottom=Side(style='medium')
                )

            # Llenar datos de productos
            fila_actual = fila + 12
            total_bultos = 0
            total_peso = 0.0

            for item in self.tree.get_children():
                valores = self.tree.item(item, 'values')
                descripcion = valores[1] if len(valores) > 1 else "Desconocido"
                cantidad = int(valores[0]) if valores[0] else 0
                peso = float(valores[2].replace(' kg', '')) if len(valores) > 2 else 0.0

                ws[f'B{fila_actual}'] = descripcion
                ws[f'I{fila_actual}'] = cantidad
                ws[f'J{fila_actual}'] = peso

                # Aplicar estilo a celdas de datos
                for col in ['B', 'I', 'J']:
                    cell = ws[f'{col}{fila_actual}']
                    cell.font = Font(size=11, name='Arial')
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    if col == 'J':
                        cell.number_format = '0.00'

                total_bultos += cantidad
                total_peso += peso
                fila_actual += 1

            # ----------------------------
            # TOTALES Y PIE DE PÁGINA
            # ----------------------------
            # Totales
            ws[f'I{fila_actual}'] = "TOTALES"
            ws[f'J{fila_actual}'] = total_peso
            ws[f'J{fila_actual}'].number_format = '0.00'

            # Estilo para totales
            for col in ['I', 'J']:
                cell = ws[f'{col}{fila_actual}']
                cell.font = Font(bold=True, size=11, name='Arial')
                cell.border = Border(
                    top=Side(style='double'),
                    bottom=Side(style='double')
                )

            # Pie de página (campos para llenar manualmente)
            ws[f'B{fila_actual+2}'] = "MOTIVO DEL TRASLADO :"
            ws[f'B{fila_actual+4}'] = "CHOFER :"
            ws[f'I{fila_actual+4}'] = "C.I"
            ws[f'B{fila_actual+5}'] = "VEHICULO:"
            ws[f'I{fila_actual+5}'] = "PLACA:"

            # Guardar cambios
            wb.save(self.excel_path)
            self._mostrar_info("Despacho guardado", "El despacho se registró correctamente")
            return True

        except Exception as e:
            self._mostrar_error("Error al guardar", f"No se pudo guardar el despacho:\n{str(e)}")
            self._registrar_error(e)
            return False

    def _encontrar_fila_vacia(self, sheet) -> int:
        """Encuentra la primera fila vacía en la hoja"""
        fila = 1
        while sheet.cell(row=fila, column=1).value is not None:
            fila += 1
            if fila > 1000:  # Límite para evitar bucles infinitos
                raise ValueError("No se encontró fila vacía en las primeras 1000 filas")
        return fila

    def _escribir_info_cliente(self, sheet, fila_inicial: int):
        """Escribe la información del cliente en la hoja Excel"""
        if not self.cliente_actual:
            return
            
        # Escribir encabezado
        sheet.cell(row=fila_inicial, column=1, value="INFORMACIÓN DEL CLIENTE").font = Font(
            bold=True, size=12, color="FFFFFF")
        sheet.cell(row=fila_inicial, column=1).fill = PatternFill(
            "solid", fgColor=self.colores['primario'])
        
        # Combinar celdas
        sheet.merge_cells(start_row=fila_inicial, start_column=1, end_row=fila_inicial, end_column=2)
        
        # Escribir datos del cliente
        campos = [
            ("Nombre", self.cliente_actual.get("Nombre", "")),
            ("RIF", self.cliente_actual.get("RIF", "")),
            ("Teléfono", self.cliente_actual.get("Teléfono", "")),
            ("Dirección", self.cliente_actual.get("Dirección", ""))
        ]
        
        for i, (campo, valor) in enumerate(campos, start=1):
            # Etiqueta
            sheet.cell(row=fila_inicial+i, column=1, value=f"{campo}:").font = Font(
                bold=True, color="44546A")
            
            # Valor
            sheet.cell(row=fila_inicial+i, column=2, value=valor).border = Border(
                bottom=Side(style='thin', color="D9D9D9"))
            
            # Aplicar bordes
            for col in [1, 2]:
                sheet.cell(row=fila_inicial+i, column=col).border = Border(
                    left=Side(style='thin', color="D9D9D9"),
                    right=Side(style='thin', color="D9D9D9"),
                    bottom=Side(style='thin', color="D9D9D9"))
                
    def _exportar_excel(self):
        """Exporta el despacho en formato profesional idéntico al ejemplo proporcionado"""
        try:
            if not self._validar_despacho():
                return False

            # Configurar nombre del archivo
            hoy = datetime.now()
            fecha_str = hoy.strftime("%d-%m-%Y")
            cliente_nombre = self.cliente_actual.get('Nombre', 'DESCONOCIDO').replace(' ', '_')
            default_filename = f"DESPACHO_{cliente_nombre}_{fecha_str}.xlsx"
            
            filepath = filedialog.asksaveasfilename(
                title="Guardar Despacho como...",
                defaultextension=".xlsx",
                filetypes=[("Archivo Excel", "*.xlsx")],
                initialfile=default_filename
            )
            
            if not filepath:
                return False

            # Crear libro de Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "GUIA DE DESPACHO"

            # ----------------------------
            # CONFIGURACIÓN DE FORMATO GENERAL
            # ----------------------------
            # 1. Anchura de columnas (manteniendo las originales)
            ws.column_dimensions['A'].width = 5       # Margen izquierdo mínimo
            ws.column_dimensions['B'].width = 60      # Descripción 
            ws.column_dimensions['C'].width = 12      # BULTOS 
            ws.column_dimensions['D'].width = 12      # KILOS 

            # 2. Altura de filas
            for row in range(1, 51):  # Asegurando que llegue hasta la fila 50
                ws.row_dimensions[row].height = 15
            
            # Ajustes específicos de altura para filas clave
            ws.row_dimensions[8].height = 18    # Nombre o Razón Social
            ws.row_dimensions[9].height = 18    # Nombre del cliente
            ws.row_dimensions[10].height = 18   # RIF
            ws.row_dimensions[11].height = 18   # Dirección
            ws.row_dimensions[12].height = 18   # Teléfono
            ws.row_dimensions[13].height = 20   # Encabezado tabla
            ws.row_dimensions[14].height = 20   # Primera fila de productos

            # ----------------------------
            # CABECERA
            # ----------------------------
            # "GUARENAS" en C2 con fuente de 11pt (reducido de 12pt)
            ws['C2'] = "GUARENAS"
            ws['C2'].font = Font(name='Arial', size=11, bold=True)
            
            # Fecha en D2-D3 (estructura de dos filas)
            ws['D2'] = "FECHA:"
            ws['D2'].font = Font(name='Arial', size=10, bold=True)
            ws['D2'].alignment = Alignment(horizontal='right')
            
            ws['D3'] = hoy.strftime("%d/%m/%Y")
            ws['D3'].font = Font(name='Arial', size=10)
            ws['D3'].alignment = Alignment(horizontal='right')

            # ----------------------------
            # INFORMACIÓN DEL CLIENTE
            # ----------------------------
            # "Nombre o Razon Social:" en B8 (fila 8)
            ws['B8'] = "Nombre o Razon Social:"
            ws['B8'].font = Font(name='Arial', size=10, bold=True)
            
            # Nombre del cliente en B9 (fila 9)
            ws['B9'] = f"{self.cliente_actual.get('Nombre', '')}"
            ws['B9'].font = Font(name='Arial', size=10)
            
            # RIF en B10 (fila 10) con "RIF:" en negrita
            rif_text = f"RIF: {self.cliente_actual.get('RIF', '')}"
            ws['B10'] = rif_text
            ws['B10'].font = Font(name='Arial', size=10, bold=True)
            
            # Dirección en B11 (fila 11)
            ws['B11'] = "DIRECCION:"
            ws['B11'].font = Font(name='Arial', size=10, bold=True)
            
            if self.cliente_actual.get('Dirección', ''):
                ws['B12'] = self.cliente_actual.get('Dirección', '')
                ws['B12'].font = Font(name='Arial', size=10)
                ws.row_dimensions[12].height = 18

            # Teléfono en C12 (fila 12) con "TEL:" en negrita
            tel_text = f"TEL: {self.cliente_actual.get('Teléfono', '')}"
            ws['C12'] = tel_text
            ws['C12'].font = Font(name='Arial', size=10, bold=True)

            # ----------------------------
            # TABLA DE PRODUCTOS (INICIANDO EN FILA 13)
            # ----------------------------
            # Encabezados de tabla en fila 13
            encabezados = ["DESCRIPCIÓN", "BULTOS", "KILOS"]
            ws['B13'] = encabezados[0]
            ws['C13'] = encabezados[1]
            ws['D13'] = encabezados[2]
            
            # Estilo para encabezados (bordes completos y negrita)
            for col in ['B', 'C', 'D']:
                cell = ws[f'{col}13']
                cell.font = Font(name='Arial', size=10, bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.border = Border(
                    left=Side(style='medium', color='000000'),
                    right=Side(style='medium', color='000000'),
                    top=Side(style='medium', color='000000'),
                    bottom=Side(style='medium', color='000000')
                )

            # Datos de productos (comenzando en fila 14)
            fila_actual = 14
            total_bultos = 0
            total_peso = 0.0

            for item in self.tree.get_children():
                valores = self.tree.item(item, 'values')
                
                # Descripción (con borde completo negro)
                ws[f'B{fila_actual}'] = valores[1] if len(valores) > 1 else "Desconocido"
                ws[f'B{fila_actual}'].font = Font(name='Arial', size=10)
                ws[f'B{fila_actual}'].border = Border(
                    left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000')
                )
                
                # Bultos (centrado, con borde completo negro)
                try:
                    bultos = int(float(valores[0])) if valores[0] else 0
                except:
                    bultos = 0
                ws[f'C{fila_actual}'] = bultos
                ws[f'C{fila_actual}'].font = Font(name='Arial', size=10)
                ws[f'C{fila_actual}'].alignment = Alignment(horizontal='center')
                ws[f'C{fila_actual}'].border = Border(
                    left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000')
                )
                
                # Peso (centrado, con borde completo negro, formato numérico)
                try:
                    peso = float(str(valores[2]).replace(' kg', '').replace(',', '.')) if len(valores) > 2 else 0.0
                except:
                    peso = 0.0
                ws[f'D{fila_actual}'] = peso
                ws[f'D{fila_actual}'].font = Font(name='Arial', size=10)
                ws[f'D{fila_actual}'].number_format = '0.00'
                ws[f'D{fila_actual}'].alignment = Alignment(horizontal='center')
                ws[f'D{fila_actual}'].border = Border(
                    left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000')
                )
                
                total_bultos += bultos
                total_peso += peso
                fila_actual += 1

            # ----------------------------
            # TOTALES (SIN ESPACIO, CON BORDES DOBLES NEGROS)
            # ----------------------------
            # Nota: fila_totales = fila_actual (sin +1 para que quede pegado)
            fila_totales = fila_actual
            
            ws[f'B{fila_totales}'] = "TOTALES:"
            ws[f'B{fila_totales}'].font = Font(name='Arial', size=10, bold=True)
            ws[f'B{fila_totales}'].border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='double', color='000000'),
                bottom=Side(style='double', color='000000')
            )
            
            ws[f'C{fila_totales}'] = total_bultos
            ws[f'C{fila_totales}'].font = Font(name='Arial', size=10, bold=True)
            ws[f'C{fila_totales}'].alignment = Alignment(horizontal='center')
            ws[f'C{fila_totales}'].border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='double', color='000000'),
                bottom=Side(style='double', color='000000')
            )
            
            ws[f'D{fila_totales}'] = total_peso
            ws[f'D{fila_totales}'].font = Font(name='Arial', size=10, bold=True)
            ws[f'D{fila_totales}'].number_format = '0.00'
            ws[f'D{fila_totales}'].alignment = Alignment(horizontal='center')
            ws[f'D{fila_totales}'].border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='double', color='000000'),
                bottom=Side(style='double', color='000000')
            )

            # ----------------------------
            # PIE DE PÁGINA FIJO (HASTA FILA 50)
            # ----------------------------
            # Motivo del traslado (fila 48)
            ws['B48'] = "MOTIVO DEL TRASLADO:"
            ws['B48'].font = Font(name='Arial', size=10, bold=True)
            
            # Chofer y C.I. (fila 49)
            ws['B49'] = "CHOFER: ____________________"
            ws['B49'].font = Font(name='Arial', size=10)
            ws['C49'] = "C.I: ____________________"
            ws['C49'].font = Font(name='Arial', size=10)
            
            # Vehículo y Placa (fila 50)
            ws['B50'] = "VEHICULO: ____________________"
            ws['B50'].font = Font(name='Arial', size=10)
            ws['C50'] = "PLACA: ____________________"
            ws['C50'].font = Font(name='Arial', size=10)

            # ----------------------------
            # CONFIGURACIÓN DE IMPRESIÓN (HASTA FILA 50)
            # ----------------------------
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
            
            # Márgenes mínimos posibles (en pulgadas)
            ws.page_margins.left = 0.12    # 3.0 mm (reducido de 3.8)
            ws.page_margins.right = 0.12   # 3.0 mm (reducido de 3.8)
            ws.page_margins.top = 0.16     # 4.0 mm (reducido de 5.0)
            ws.page_margins.bottom = 0.12  # 3.0 mm (reducido de 4.5) - Ajuste clave
            ws.page_margins.header = 0.08  # 2.0 mm 
            ws.page_margins.footer = 0.08  # 2.0 mm

            # Reducción general de alturas de fila
            for row in range(1, 51):
                if ws.row_dimensions[row].height > 14:  # Reducimos todas las filas
                    ws.row_dimensions[row].height = 14  # 0.93 cm (antes 1.0 cm)
            
            # Ajustamos filas clave específicamente
            ws.row_dimensions[13].height = 18  # Encabezado tabla
            ws.row_dimensions[14].height = 18  # Primera fila productos
            
            # Configuración de escala forzada
            ws.page_setup.fitToHeight = 1  # Estrictamente 1 página
            ws.page_setup.fitToWidth = 1
            ws.page_setup.scale = 98       # Reducción del 2% si es necesario

            # Área de impresión forzada
            ws.print_area = f"A1:D50"

            # ----------------------------
            # AJUSTE ALTERNATIVO AUTOMÁTICO (si aún no cabe)
            # ----------------------------
            if fila_actual > 45:  # Si hay demasiados artículos
                # Reducir fuente de descripción de productos a 9.5pt
                for row in range(14, fila_actual + 1):
                    ws[f'B{row}'].font = Font(name='Arial', size=9.5)
                    ws.row_dimensions[row].height = 17  # Reducción adicional

                # Mover pie de página 1 fila arriba (a 47-49)
                ws['B47'] = "MOTIVO DEL TRASLADO:"
                ws['B48'] = "CHOFER: ____________________"
                ws['C48'] = "C.I: ____________________"
                ws['B49'] = "VEHICULO: ____________________"
                ws['C49'] = "PLACA: ____________________"
                
                # Limpiar filas originales
                for row in [48, 49, 50]:
                    ws[f'B{row}'] = ws[f'C{row}'] = None
                
                ws.print_area = f"A1:D49"  # Nueva área de impresión

            # ----------------------------
            # GUARDAR ARCHIVO
            # ----------------------------
            wb.save(filepath)
            
            # Preguntar si abrir el archivo
            respuesta = messagebox.askyesno(
                "Exportación exitosa",
                f"El despacho se exportó correctamente a:\n{os.path.basename(filepath)}\n\n"
                "¿Desea abrir el archivo ahora?")
            
            if respuesta:
                self._abrir_archivo(filepath)
            
            return True

        except Exception as e:
            error_msg = f"No se pudo exportar el despacho:\n{str(e)}"
            self._mostrar_error("Error al exportar", error_msg)
            self._registrar_error(e)
            return False
        

    def _exportar_despacho_detallado(self):
        """Exporta el despacho con cada familia en cuadros separados verticalmente"""
        try:
            if not self._validar_despacho():
                return False

            # Configuración del archivo
            hoy = datetime.now()
            fecha_str = hoy.strftime("%d-%m-%Y")
            cliente_nombre = self.cliente_actual.get('Nombre', 'DESCONOCIDO').replace(' ', '_')
            default_filename = f"DESPACHO_DETALLADO_{cliente_nombre}_{fecha_str}.xlsx"
            
            filepath = filedialog.asksaveasfilename(
                title="Guardar Despacho Detallado como...",
                defaultextension=".xlsx",
                filetypes=[("Archivo Excel", "*.xlsx")],
                initialfile=default_filename
            )
            
            if not filepath:
                return False

            wb = Workbook()
            ws = wb.active
            ws.title = "DESPACHO"[:31]  # Limitar a 31 caracteres

            # ------------------------------------------
            # PROCESAMIENTO AVANZADO DE ARTÍCULOS (VERSIÓN FINAL)
            # ------------------------------------------
            productos = {}  # {nombre_producto: {'codigos': {color: codigo}, 'pesos': {color: [pesos]}, 'tipo': tipo}
            
            for item in self.tree.get_children():
                valores = self.tree.item(item, 'values')
                if not valores or len(valores) < 2:
                    continue

                descripcion = valores[1]
                
                # ALGORITMO MEJORADO DE EXTRACCIÓN DE NOMBRE Y COLOR
                partes = [p.strip() for p in descripcion.split(' - ') if p.strip()]
                
                # 1. Extraer código (si existe)
                codigo = partes[0] if len(partes) > 0 and ' ' not in partes[0] else ""
                
                # 2. Determinar nombre base (ignorando "TEJIDO" y términos de color)
                nombre_base = ""
                if len(partes) >= 2:
                    nombre_candidato = partes[1].upper()
                    nombre_base = nombre_candidato.replace("TEJIDO", "").strip()
                    
                    # Si después de quitar "TEJIDO" queda vacío, tomar el siguiente
                    if not nombre_base and len(partes) >= 3:
                        nombre_base = partes[2].upper()
                else:
                    nombre_base = " ".join(partes[1:]).upper() if len(partes) > 1 else descripcion.upper()
                
                # 3. Determinar color (buscando en las partes)
                colores_conocidos = ["BLANCO", "OSCURO", "PASTEL", "ESPECIAL", "MELANGE", "NEGRO", 
                                "ROJO", "AZUL", "VERDE", "GRIS", "BEIGE", "CREMA", "AMARILLO"]
                color = "BLANCO"  # Por defecto
                
                for parte in reversed(partes):
                    parte_upper = parte.upper()
                    if any(c in parte_upper for c in colores_conocidos):
                        # Encontrar el color exacto que coincide
                        for c in colores_conocidos:
                            if c in parte_upper:
                                color = c
                                break
                        break
                
                # 4. Limpieza final del nombre base (quitar color si está al final)
                for color_term in colores_conocidos:
                    if nombre_base.endswith(color_term):
                        nombre_base = nombre_base[:-len(color_term)].strip()
                        break
                
                # 5. Determinar tipo
                tipo = "TEJIDO" if codigo.startswith(('T', 'TN', 'TI')) or "TEJIDO" in descripcion.upper() else "HILADO"
                
                # Agrupar por nombre base del producto
                if nombre_base not in productos:
                    productos[nombre_base] = {'codigos': {}, 'pesos': {}, 'tipo': tipo}
                
                # Agregar código y pesos para cada color
                if color not in productos[nombre_base]['codigos']:
                    productos[nombre_base]['codigos'][color] = codigo
                    productos[nombre_base]['pesos'][color] = []
                
                # Procesar pesos individuales de bultos
                if codigo and codigo in self.bultos_data:
                    for peso in self.bultos_data[codigo].values():
                        try:
                            productos[nombre_base]['pesos'][color].append(float(peso))
                        except (ValueError, TypeError):
                            continue
                else:
                    try:
                        peso_total = float(valores[2].replace(' kg', '')) if len(valores) > 2 else 0.0
                        if peso_total > 0:
                            productos[nombre_base]['pesos'][color].append(peso_total)
                    except (ValueError, TypeError):
                        continue

            # Ordenar productos alfabéticamente por nombre base
            productos_ordenados = sorted(productos.items(), key=lambda x: x[0])

            # ------------------------------------------
            # ESCRITURA EN EXCEL CON FORMATO FINAL MEJORADO
            # ------------------------------------------
            current_row = 1
            
            # Estilos de bordes
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                top=Side(style='thin'), bottom=Side(style='thin'))
            
            thick_border = Border(left=Side(style='medium'), right=Side(style='medium'),
                                top=Side(style='medium'), bottom=Side(style='medium'))
            
            bottom_double_border = Border(bottom=Side(style='double'))
            
            # Fuentes
            header_font = Font(name='Arial', size=11, bold=True)
            normal_font = Font(name='Arial', size=10)
            bold_font = Font(name='Arial', size=10, bold=True)
            title_font = Font(name='Arial', size=12, bold=True)
            
            # Cabecera con información del cliente
            ws.cell(row=current_row, column=1, value="FECHA:").font = bold_font
            ws.cell(row=current_row, column=2, value=hoy.strftime("%d/%m/%Y")).font = normal_font
            current_row += 1
            
            ws.cell(row=current_row, column=1, value="NOMBRE DE LA EMPRESA:").font = bold_font
            ws.cell(row=current_row, column=2, value=self.cliente_actual.get('Nombre', '')).font = normal_font
            current_row += 3  # Más espacio entre secciones

            # Variables para totales globales
            total_global_bultos = 0
            total_global_peso = 0.0

            # Para cada producto
            for nombre_base, datos in productos_ordenados:
                colors = sorted(datos['pesos'].keys())
                max_bultos = max(len(datos['pesos'][color]) for color in colors) if colors else 0
                
                # ----------------------------
                # NOMBRE DEL TEJIDO (con borde completo grueso)
                # ----------------------------
                # Aplicar borde grueso completo a toda la fila
                for col in range(1, len(colors)+2):
                    cell = ws.cell(row=current_row, column=col)
                    cell.border = thick_border
                
                ws.cell(row=current_row, column=1, value="NOMBRE DEL TEJIDO:").font = bold_font
                ws.cell(row=current_row, column=2, value=nombre_base).font = bold_font
                current_row += 1
                
                # ----------------------------
                # SECCIÓN DE COLORES (con bordes completos gruesos)
                # ----------------------------
                # Fila superior: "COLORES:" en A y nombres de colores en B, C, etc.
                for col in range(1, len(colors)+2):
                    cell = ws.cell(row=current_row, column=col)
                    cell.border = thick_border
                
                ws.cell(row=current_row, column=1, value="COLORES:").font = bold_font
                
                # Escribir cada color como encabezado de columna (fila superior)
                for i, color in enumerate(colors, start=2):
                    cell = ws.cell(row=current_row, column=i, value=color)
                    cell.font = bold_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.fill = PatternFill("solid", fgColor="D9D9D9")  # Fondo gris claro
                
                current_row += 1
                
                # Fila inferior: "CÓDIGOS:" en A y códigos en B, C, etc.
                for col in range(1, len(colors)+2):
                    cell = ws.cell(row=current_row, column=col)
                    cell.border = thick_border
                
                ws.cell(row=current_row, column=1, value="CÓDIGOS:").font = bold_font
                
                # Escribir códigos de cada color
                for i, color in enumerate(colors, start=2):
                    cell = ws.cell(row=current_row, column=i, value=datos['codigos'].get(color, ''))
                    cell.font = bold_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                current_row += 1
                
                # ----------------------------
                # ENCABEZADO DE TABLA (ITEM y Peso kg)
                # ----------------------------
                # Encabezado de ITEM (con borde grueso completo)
                for col in range(1, len(colors)+2):
                    cell = ws.cell(row=current_row, column=col)
                    cell.border = thick_border
                    if col > 1:
                        cell.fill = PatternFill("solid", fgColor="F2F2F2")  # Fondo gris muy claro
                
                cell = ws.cell(row=current_row, column=1, value="ITEM")
                cell.font = bold_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.fill = PatternFill("solid", fgColor="F2F2F2")  # Fondo gris muy claro
                
                # Escribir "Peso kg" en cada columna de color
                for i in range(2, len(colors)+2):
                    cell = ws.cell(row=current_row, column=i, value="Peso kg")
                    cell.font = normal_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                current_row += 1
                
                # ----------------------------
                # TABLA DE BULTOS (con bordes delgados completos)
                # ----------------------------
                # Escribir cada bulto en su columna correspondiente
                for bulto_idx in range(max_bultos):
                    # Celda de ITEM (número entero, centrado, con borde delgado completo)
                    cell = ws.cell(row=current_row, column=1, value=bulto_idx+1)
                    cell.number_format = '0'  # Formato entero sin decimales
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border
                    
                    # Escribir el peso en la columna del color correspondiente
                    for i, color in enumerate(colors, start=2):
                        pesos = datos['pesos'][color]
                        if bulto_idx < len(pesos):
                            peso = pesos[bulto_idx]
                            cell = ws.cell(row=current_row, column=i, value=peso)
                            cell.number_format = '0.00'  # Dos decimales para pesos
                        else:
                            cell = ws.cell(row=current_row, column=i, value="")
                        
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    current_row += 1
                
                # ----------------------------
                # TOTALES POR COLOR (con borde doble inferior y laterales)
                # ----------------------------
                # Aplicar borde doble inferior y laterales a toda la fila
                for col in range(1, len(colors)+2):
                    cell = ws.cell(row=current_row, column=col)
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    bottom=Side(style='double'))
                
                ws.cell(row=current_row, column=1, value="TOTAL:").font = bold_font
                
                # Calcular y escribir totales por color en sus columnas
                totales_color = {}
                for i, color in enumerate(colors, start=2):
                    pesos = datos['pesos'][color]
                    total_color = sum(pesos)
                    totales_color[color] = total_color
                    cell = ws.cell(row=current_row, column=i, value=total_color)
                    cell.font = bold_font
                    cell.number_format = '0.00'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                current_row += 1
                
                # ----------------------------
                # RESUMEN FINAL DEL PRODUCTO (con bordes completos)
                # ----------------------------
                # Total KG (suma de todos los colores)
                total_producto = sum(totales_color.values())
                total_global_peso += total_producto
                
                # Aplicar bordes a las celdas de totales
                ws.cell(row=current_row, column=1, value="TOTAL KG:").font = bold_font
                ws.cell(row=current_row, column=1).border = thin_border
                ws.cell(row=current_row, column=2, value=total_producto).font = bold_font
                ws.cell(row=current_row, column=2).border = thin_border
                
                # Total de bultos (suma de bultos por color)
                bultos_producto = sum(len(datos['pesos'][color]) for color in colors)
                total_global_bultos += bultos_producto
                
                ws.cell(row=current_row+1, column=1, value="TOTAL BULTOS:").font = bold_font
                ws.cell(row=current_row+1, column=1).border = thin_border
                ws.cell(row=current_row+1, column=2, value=bultos_producto).font = bold_font
                ws.cell(row=current_row+1, column=2).border = thin_border
                
                current_row += 4  # Más espacio entre productos

            # ------------------------------------------
            # TOTAL GENERAL FINAL (con bordes mejorados)
            # ------------------------------------------
            # Aplicar bordes gruesos completos a los totales generales
            thick_border_all = Border(left=Side(style='medium'), right=Side(style='medium'),
                                    top=Side(style='medium'), bottom=Side(style='medium'))
            
            # Total General KG
            for col in range(1, 3):
                cell = ws.cell(row=current_row, column=col)
                cell.border = thick_border_all
            
            ws.cell(row=current_row, column=1, value="TOTAL GENERAL KG:").font = title_font
            ws.cell(row=current_row, column=2, value=total_global_peso).font = title_font
            
            current_row += 1
            
            # Total General Bultos
            for col in range(1, 3):
                cell = ws.cell(row=current_row, column=col)
                cell.border = thick_border_all
            
            ws.cell(row=current_row, column=1, value="TOTAL GENERAL BULTOS:").font = title_font
            ws.cell(row=current_row, column=2, value=total_global_bultos).font = title_font
            
            # ------------------------------------------
            # AJUSTES FINALES DE FORMATO
            # ------------------------------------------
            # Ajustar anchos de columnas (ampliar columna A significativamente)
            ws.column_dimensions['A'].width = 30  # Ampliada a 30 (antes era 22)
            for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
                ws.column_dimensions[col].width = 15
            
            # Ajustar alineación de columnas
            for row in ws.iter_rows():
                for cell in row:
                    if cell.column_letter == 'A' and cell.row > 1 and not cell.value in ["ITEM", "TOTAL:", "TOTAL KG:", "TOTAL BULTOS:"]:
                        cell.alignment = Alignment(horizontal='center')
            
            # Configurar márgenes de página
            ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)
            
            # Guardar archivo
            wb.save(filepath)
            
            # Preguntar si abrir el archivo
            respuesta = messagebox.askyesno(
                "Exportación exitosa",
                f"El despacho detallado se exportó correctamente a:\n{os.path.basename(filepath)}\n\n"
                "¿Desea abrir el archivo ahora?")
            
            if respuesta:
                self._abrir_archivo(filepath)
            
            return True

        except Exception as e:
            error_msg = f"No se pudo exportar el despacho detallado:\n{str(e)}"
            self._mostrar_error("Error al exportar", error_msg)
            self._registrar_error(e)
            return False
        



    
    
    def _add_border(self, ws, cell_range, border_style='thin'):
        """Añade bordes a un rango de celdas"""
        rows = ws[cell_range]
        side = Side(style=border_style)
        
        for row in rows:
            for cell in row:
                cell.border = Border(
                    left=side, right=side,
                    top=side, bottom=side
                )


    def _escribir_info_exportacion(self, sheet):
        """Escribe la información básica del despacho en la hoja Excel"""
        sheet['A1'] = "FECHA:"
        sheet['A1'].style = 'InfoLabel'
        sheet['B1'] = datetime.now().strftime(self.config.get("formato_fecha", "%d/%m/%Y %H:%M"))
        sheet['B1'].style = 'InfoValue'
        
        sheet['A2'] = "NOMBRE DEL CLIENTE:"
        sheet['A2'].style = 'InfoLabel'
        sheet['B2'] = self.cliente_actual.get('Nombre', '')
        sheet['B2'].style = 'InfoValue'
        
        sheet['A3'] = "RIF:"
        sheet['A3'].style = 'InfoLabel'
        sheet['B3'] = self.cliente_actual.get('RIF', '')
        sheet['B3'].style = 'InfoValue'
        
        

    def _escribir_totales_exportacion(self, sheet):
        """Escribe los totales del despacho en la hoja Excel"""
        total_row = 8 + len(self.tree.get_children())
        
        # Escribir línea de totales
        sheet.cell(row=total_row, column=2, value="TOTAL:").style = 'InfoLabel'
        
        # Sumar columna de peso (C)
        peso_cell = sheet.cell(row=total_row, column=3, value=f"=SUM(C8:C{total_row-1})")
        peso_cell.style = 'Total'

        # Escribir resumen final (2 líneas después del total)
        summary_data = [
            ("TOTAL KG:", f"=C{total_row}"),
            ("TOTAL ARTÍCULOS:", f"=COUNT(A8:A{total_row-1})")
        ]
        
        for i, (label, value) in enumerate(summary_data, start=1):
            sheet.cell(row=total_row + 1 + i, column=2, value=label).style = 'InfoLabel'
            sheet.cell(row=total_row + 1 + i, column=3, value=value).style = 'InfoValue'

    def _configurar_estilos_excel(self, wb):
        """Configura los estilos para el archivo Excel exportado"""
        # Estilo para títulos
        title_style = NamedStyle(name="title_style")
        title_style.font = Font(bold=True, size=12, color="FFFFFF")
        title_style.fill = PatternFill("solid", fgColor="4472C4")
        title_style.alignment = Alignment(horizontal="left")
        wb.add_named_style(title_style)

        # Estilo para encabezados de tabla
        header_style = NamedStyle(name="header_style")
        header_style.font = Font(bold=True, color="000000")
        header_style.fill = PatternFill("solid", fgColor="D9E1F2")
        header_style.alignment = Alignment(horizontal="center")
        header_style.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"))
        wb.add_named_style(header_style)

        # Estilo para datos
        data_style = NamedStyle(name="data_style")
        data_style.font = Font(size=11)
        data_style.alignment = Alignment(horizontal="left")
        data_style.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"))
        wb.add_named_style(data_style)

        # Estilo para totales
        total_style = NamedStyle(name="total_style")
        total_style.font = Font(bold=True)
        total_style.fill = PatternFill("solid", fgColor="F2F2F2")
        total_style.border = Border(
            top=Side(style="double"),
            bottom=Side(style="double"))
        wb.add_named_style(total_style)

    def _aplicar_formato_tabla(self, ws):
        """Aplica formato a la tabla de artículos en el Excel"""
        # Aplicar estilos a los títulos
        for row in ws.iter_rows(min_row=1, max_row=1):
            for cell in row:
                cell.style = "title_style"
        
        # Aplicar estilos a los encabezados de la tabla de artículos
        for row in ws.iter_rows(min_row=15, max_row=15):  # Ajusta estos números según tu estructura
            for cell in row:
                cell.style = "header_style"
        
        # Ajustar anchos de columnas
        ws.column_dimensions['A'].width = 50  # Descripción
        ws.column_dimensions['B'].width = 15  # Bultos
        ws.column_dimensions['C'].width = 20  # Kilos/Metros
        
        # Aplicar estilo a los datos
        for row in ws.iter_rows(min_row=16, max_row=ws.max_row-2):  # Ajusta según tu estructura
            for cell in row:
                cell.style = "data_style"
        
        # Aplicar estilo a los totales
        for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row):
            for cell in row:
                cell.style = "total_style"


    def _escribir_tabla_articulos(self, sheet):
        """Escribe la tabla de artículos en la hoja (blanco y negro)"""
        for i, item in enumerate(self.tree.get_children(), start=1):
            valores = self.tree.item(item)['values']
            
            # Escribir cantidad
            sheet.cell(row=7+i, column=1, value=int(valores[0])).style = 'TableRow'
            
            # Escribir descripción
            sheet.cell(row=7+i, column=2, value=valores[1]).style = 'TableRow'
            
            # Escribir peso total (extraer solo el valor numérico)
            peso_str = valores[2].replace(' kg', '')  # Eliminar "kg" del string
            try:
                peso = float(peso_str)
                peso_cell = sheet.cell(row=7+i, column=3, value=peso)
                peso_cell.style = 'TableRow'
            except ValueError:
                # Si no se puede convertir a float, escribir 0
                peso_cell = sheet.cell(row=7+i, column=3, value=0.0)
                peso_cell.style = 'TableRow'

    def _configurar_pagina_impresion(self, sheet, ultima_fila=None):
        """Configura los parámetros de impresión para el documento exportado"""
        # Configurar márgenes
        sheet.page_margins = PageMargins(
            left=0.5, right=0.5, top=0.75, bottom=0.75,
            header=0.3, footer=0.3
        )

        # Configurar orientación y ajuste
        sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0

        # Configurar área de impresión
        ultima_fila = 8 + len(self.tree.get_children()) + 3
        sheet.print_area = f"A1:F{ultima_fila}"

        # Configurar encabezado y pie de página (versión compatible)
        try:
            # Intenta usar header_footer si está disponible
            sheet.header_footer.center_header.text = "&\"Arial,Bold\"&14DESPACHO DE MATERIAL"
            sheet.header_footer.center_footer.text = "&\"Arial\"&10© 2024 - Sistema de Gestión de Despachos"
        except AttributeError:
            # Si header_footer no está disponible, usa alternativas
            sheet.oddHeader.center.text = "DESPACHO DE MATERIAL"
            sheet.oddHeader.center.size = 14
            sheet.oddHeader.center.font = "Arial,Bold"
            sheet.oddFooter.center.text = "© 2024 - Sistema de Gestión de Despachos"
            sheet.oddFooter.center.size = 10
            sheet.oddFooter.center.font = "Arial"

        # Repetir filas de encabezado
        sheet.print_title_rows = '7:7'
        # Configurar área de impresión si se proporciona ultima_fila
        if ultima_fila:
            sheet.print_area = f"A1:F{ultima_fila}"

    def _abrir_archivo(self, filepath):
        """Abre el archivo generado con el programa predeterminado"""
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":
                subprocess.run(["open", filepath])
            else:
                subprocess.run(["xdg-open", filepath])
        except Exception as e:
            self._mostrar_advertencia(
                "Aviso", 
                f"No se pudo abrir el archivo:\n{str(e)}\n"
                f"Puede encontrarlo en:\n{filepath}")

    def _mostrar_configuracion(self):
        """Muestra el diálogo de configuración de la aplicación"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Configuración del Sistema")
        dialog.geometry("700x550")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # Frame principal
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Pestañas
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        # Pestaña de Columnas
        columns_tab = ttk.Frame(notebook, padding=10)
        notebook.add(columns_tab, text="Columnas")

        # Configuración de columnas para clientes
        ttk.Label(
            columns_tab, 
            text="Nombres de columnas en hoja 'Clientes':",
            font=('Segoe UI', 10, 'bold')
        ).pack(anchor=tk.W, pady=(5, 2))

        self.columnas_clientes_var = tk.StringVar(
            value=", ".join(self.config.get("columnas_clientes", [])))
        
        columnas_clientes_frame = ttk.Frame(columns_tab)
        columnas_clientes_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(columnas_clientes_frame, text="Orden:").pack(side=tk.LEFT)
        self.columnas_clientes_entry = ttk.Entry(
            columnas_clientes_frame, 
            textvariable=self.columnas_clientes_var,
            width=50)
        self.columnas_clientes_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        ttk.Label(
            columns_tab, 
            text="Ejemplo: Nombre, RIF, Teléfono, Dirección",
            font=('Segoe UI', 8),
            foreground='#666666'
        ).pack(anchor=tk.W, pady=(0, 15))

        # Configuración de columnas para artículos
        ttk.Label(
            columns_tab, 
            text="Nombres de columnas en hoja 'Artículos':",
            font=('Segoe UI', 10, 'bold')
        ).pack(anchor=tk.W, pady=(5, 2))

        self.columnas_articulos_var = tk.StringVar(
            value=", ".join(self.config.get("columnas_articulos", [])))
        
        columnas_articulos_frame = ttk.Frame(columns_tab)
        columnas_articulos_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(columnas_articulos_frame, text="Orden:").pack(side=tk.LEFT)
        self.columnas_articulos_entry = ttk.Entry(
            columnas_articulos_frame, 
            textvariable=self.columnas_articulos_var,
            width=50)
        self.columnas_articulos_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        ttk.Label(
            columns_tab, 
            text="Ejemplo: Código Importado, Descripción Importado, Código Nacional, Descripción Nacional",
            font=('Segoe UI', 8),
            foreground='#666666'
        ).pack(anchor=tk.W, pady=(0, 15))

        # Pestaña de Preferencias
        prefs_tab = ttk.Frame(notebook, padding=10)
        notebook.add(prefs_tab, text="Preferencias")

        # Configuración de tipos de artículos
        ttk.Label(
            prefs_tab, 
            text="Tipos de Artículos:",
            font=('Segoe UI', 10, 'bold')
        ).pack(anchor=tk.W, pady=(5, 2))

        self.tipos_articulos_var = tk.StringVar(
            value=", ".join(self.config.get("tipo_articulos", [])))
        
        tipos_articulos_frame = ttk.Frame(prefs_tab)
        tipos_articulos_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(tipos_articulos_frame, text="Valores:").pack(side=tk.LEFT)
        self.tipos_articulos_entry = ttk.Entry(
            tipos_articulos_frame, 
            textvariable=self.tipos_articulos_var,
            width=50)
        self.tipos_articulos_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        ttk.Label(
            prefs_tab, 
            text="Ejemplo: Importado, Nacional, Especial",
            font=('Segoe UI', 8),
            foreground='#666666'
        ).pack(anchor=tk.W, pady=(0, 15))

        # Configuración de formato de fecha
        ttk.Label(
            prefs_tab, 
            text="Formato de Fecha:",
            font=('Segoe UI', 10, 'bold')
        ).pack(anchor=tk.W, pady=(5, 2))

        self.formato_fecha_var = tk.StringVar(
            value=self.config.get("formato_fecha", "%d/%m/%Y"))
        
        formato_fecha_frame = ttk.Frame(prefs_tab)
        formato_fecha_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(formato_fecha_frame, text="Formato:").pack(side=tk.LEFT)
        self.formato_fecha_combobox = ttk.Combobox(
            formato_fecha_frame,
            textvariable=self.formato_fecha_var,
            values=[
                "%d/%m/%Y", 
                "%m/%d/%Y", 
                "%Y-%m-%d", 
                "%d-%m-%Y", 
                "%d/%m/%Y %H:%M", 
                "%Y%m%d_%H%M%S"
            ],
            state="readonly",
            width=20)
        self.formato_fecha_combobox.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(
            prefs_tab, 
            text="Ejemplo actual: " + datetime.now().strftime(
                self.config.get("formato_fecha", "%d/%m/%Y")),
            font=('Segoe UI', 8),
            foreground='#666666'
        ).pack(anchor=tk.W, pady=(0, 15))

        # Configuración de valores por defecto
        ttk.Label(
            prefs_tab, 
            text="Valores por Defecto:",
            font=('Segoe UI', 10, 'bold')
        ).pack(anchor=tk.W, pady=(10, 2))

        # Nombre de tejido por defecto
        ttk.Label(
            prefs_tab, 
            text="Nombre de Tejido:",
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W, pady=(5, 2))

        self.tejido_default_var = tk.StringVar(
            value=self.config.get("tejido_default", "SABINA DRY FIT"))
        
        tejido_default_frame = ttk.Frame(prefs_tab)
        tejido_default_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Entry(
            tejido_default_frame,
            textvariable=self.tejido_default_var,
            width=40).pack(side=tk.LEFT, padx=5)

        # Código de tejido por defecto
        ttk.Label(
            prefs_tab, 
            text="Código de Tejido:",
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W, pady=(5, 2))

        self.codigo_tejido_default_var = tk.StringVar(
            value=self.config.get("codigo_tejido_default", "TN0218"))
        
        codigo_tejido_frame = ttk.Frame(prefs_tab)
        codigo_tejido_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Entry(
            codigo_tejido_frame,
            textvariable=self.codigo_tejido_default_var,
            width=20).pack(side=tk.LEFT, padx=5)

        # Opción para mostrar pesos individuales
        self.mostrar_pesos_var = tk.BooleanVar(
            value=self.config.get("mostrar_pesos_individuales", True))
        
        ttk.Checkbutton(
            prefs_tab,
            text="Mostrar pesos individuales de bultos",
            variable=self.mostrar_pesos_var
        ).pack(anchor=tk.W, pady=(15, 5))

        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(
            button_frame,
            text="Guardar Configuración",
            command=lambda: self._guardar_configuracion_dialogo(dialog),
            style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="Cancelar",
            command=dialog.destroy,
            style='Secondary.TButton'
        ).pack(side=tk.RIGHT, padx=5)

    def _guardar_configuracion_dialogo(self, dialog):
        """Guarda la configuración desde el diálogo"""
        try:
            # Validar y procesar columnas para clientes
            columnas_clientes = [
                col.strip() for col in self.columnas_clientes_var.get().split(",") 
                if col.strip()
            ]
            if len(columnas_clientes) < 2:
                raise ValueError("Debe especificar al menos 2 columnas para clientes")

            # Validar y procesar columnas para artículos
            columnas_articulos = [
                col.strip() for col in self.columnas_articulos_var.get().split(",") 
                if col.strip()
            ]
            if len(columnas_articulos) < 4:
                raise ValueError("Debe especificar al menos 4 columnas para artículos")

            # Validar y procesar tipos de artículos
            tipos_articulos = [
                tipo.strip() for tipo in self.tipos_articulos_var.get().split(",") 
                if tipo.strip()
            ]
            if len(tipos_articulos) < 2:
                raise ValueError("Debe especificar al menos 2 tipos de artículos")

            # Validar formato de fecha
            try:
                datetime.now().strftime(self.formato_fecha_var.get())
            except ValueError:
                raise ValueError("Formato de fecha inválido")

            # Actualizar configuración
            self.config.update({
                "columnas_clientes": columnas_clientes,
                "columnas_articulos": columnas_articulos,
                "tipo_articulos": tipos_articulos,
                "formato_fecha": self.formato_fecha_var.get(),
                "tejido_default": self.tejido_default_var.get(),
                "codigo_tejido_default": self.codigo_tejido_default_var.get(),
                "mostrar_pesos_individuales": self.mostrar_pesos_var.get()
            })

            # Guardar en archivo
            if self._guardar_configuracion():
                messagebox.showinfo(
                    "Configuración Guardada",
                    "Los cambios se guardaron correctamente.\n\n"
                    "Algunos cambios pueden requerir reiniciar la aplicación.",
                    parent=dialog
                )
                dialog.destroy()
                # Actualizar combobox de tipos de artículos
                self.tipo_combobox['values'] = tipos_articulos
                
        except ValueError as e:
            messagebox.showerror("Error en Configuración", str(e), parent=dialog)
        except Exception as e:
            messagebox.showerror(
                "Error", 
                f"No se pudo guardar la configuración:\n{str(e)}",
                parent=dialog)
            self._registrar_error(e)

    def _mostrar_documentacion(self):
        """Muestra un diálogo con documentación básica del sistema"""
        doc_text = """
        SISTEMA DE GESTIÓN DE DESPACHOS - MANUAL RÁPIDO
        
        1. SELECCIÓN DE CLIENTE
        - Busque clientes por nombre o RIF
        - Seleccione un cliente de la lista
        - Los datos se cargarán automáticamente
        
        2. AGREGAR ARTÍCULOS
        - Seleccione el tipo de artículo
        - Haga clic en 'Agregar Artículo'
        - Busque y seleccione el artículo deseado
        - Ingrese cantidad y peso (use 'Cálculo de Peso' para múltiples bultos)
        - Agregue observaciones si es necesario
        
        3. GESTIÓN DEL DESPACHO
        - Verifique el peso total
        - Elimine artículos si es necesario (Seleccione y click en 'Eliminar')
        - Use 'Nuevo Despacho' para limpiar el formulario
        
        4. GUARDAR Y EXPORTAR
        - 'Guardar Despacho': Guarda en el archivo Excel principal
        - 'Exportar a Excel': Crea un archivo independiente con formato profesional
        
        5. CÁLCULO DE PESO DE BULTOS
        - Para artículos con múltiples bultos:
          1. Ingrese la cantidad de bultos
          2. Click en 'Cálculo de Peso'
          3. Ingrese el peso individual de cada bulto
          4. El sistema calculará automáticamente el total
        
        CONFIGURACIÓN:
        - Personalice columnas, tipos de artículos y formatos en el menú Configuración
        """
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Documentación del Sistema")
        dialog.geometry("700x500")
        
        # Frame principal
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Área de texto con scroll
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        text_scroll = ttk.Scrollbar(text_frame)
        text_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        doc_text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            yscrollcommand=text_scroll.set,
            font=('Segoe UI', 10),
            padx=10,
            pady=10
        )
        doc_text_widget.pack(fill=tk.BOTH, expand=True)
        text_scroll.config(command=doc_text_widget.yview)
        
        # Insertar texto con formato
        doc_text_widget.insert(tk.END, doc_text)
        doc_text_widget.config(state=tk.DISABLED)
        
        # Botón de cierre
        ttk.Button(
            main_frame,
            text="Cerrar",
            command=dialog.destroy,
            style='Accent.TButton'
        ).pack(pady=10)

    def _mostrar_acerca_de(self):
        """Muestra el diálogo 'Acerca de' con información de la aplicación"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Acerca de Sistema de Despachos")
        dialog.geometry("500x400")
        dialog.resizable(False, False)
        
        # Frame principal
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Logo o icono
        ttk.Label(
            main_frame, 
            text="📦",  # Emoji de paquete
            font=('Arial', 48),
            justify=tk.CENTER
        ).pack(pady=10)
        
        # Información de la aplicación
        info_text = """
        SISTEMA DE GESTIÓN DE DESPACHOS v2.0
        
        Desarrollado por: [Tu Nombre]
        
        © 2024 Todos los derechos reservados
        
        Características principales:
        - Registro completo de despachos
        - Cálculo preciso de pesos
        - Exportación profesional a Excel
        - Configuración personalizable
        
        Contacto:
        [tu@email.com]
        [tu teléfono]
        """
        
        ttk.Label(
            main_frame, 
            text=info_text,
            justify=tk.CENTER,
            font=('Segoe UI', 10)
        ).pack(pady=10, fill=tk.X)
        
        # Frame para botones
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        # Botón de cierre
        ttk.Button(
            button_frame,
            text="Cerrar",
            command=dialog.destroy,
            style='Accent.TButton',
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        # Botón de documentación
        ttk.Button(
            button_frame,
            text="Documentación",
            command=self._mostrar_documentacion,
            style='Secondary.TButton',
            width=15
        ).pack(side=tk.RIGHT, padx=5)

    def _registrar_error(self, error: Exception):
        """Registra errores en un archivo de log"""
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'error_log.txt')
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f"\n[{timestamp}] ERROR:\n")
                f.write(f"Tipo: {type(error).__name__}\n")
                f.write(f"Mensaje: {str(error)}\n")
                
                # Registrar información del despacho actual si existe
                if hasattr(self, 'cliente_actual') and self.cliente_actual:
                    f.write(f"Cliente: {self.cliente_actual.get('Nombre', 'Desconocido')}\n")
                
                if hasattr(self, 'tree') and self.tree.get_children():
                    f.write(f"Artículos en despacho: {len(self.tree.get_children())}\n")
                
                f.write("-" * 50 + "\n")
        except Exception:
            pass  # Si no se puede escribir el log, no hacer nada

    def _mostrar_error(self, titulo: str, mensaje: str):
        """Muestra un mensaje de error"""
        messagebox.showerror(titulo, mensaje, parent=self.root)

    def _mostrar_advertencia(self, mensaje: str):
        """Muestra un mensaje de advertencia"""
        messagebox.showwarning("Advertencia", mensaje, parent=self.root)

    def _mostrar_info(self, titulo: str, mensaje: str):
        """Muestra un mensaje informativo"""
        messagebox.showinfo(titulo, mensaje, parent=self.root)

    def _actualizar_estado(self, mensaje: str):
        """Actualiza la barra de estado"""
        if hasattr(self, 'statusbar'):
            self.statusbar.config(text=mensaje)

def main():
    """Función principal para iniciar la aplicación"""
    try:
        # Configurar manejo de excepciones no capturadas
        sys.excepthook = lambda exc_type, exc_value, exc_traceback: (
            messagebox.showerror(
                "Error Crítico",
                f"Ocurrió un error inesperado:\n\n{''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))}\n\n"
                "La aplicación se cerrará. Revise el archivo error_log.txt para más detalles."),
            sys.exit(1)
        )
        
        # Crear ventana principal
        root = tk.Tk()
        
        # Configurar tema si está disponible
        try:
            import ttkthemes
            style = ttkthemes.ThemedStyle(root)
            style.set_theme("arc")  # Tema moderno
        except ImportError:
            pass
        
        # Crear y ejecutar aplicación
        app = AplicacionDespachos(root)
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror(
            "Error Inicial", 
            f"No se pudo iniciar la aplicación:\n{str(e)}\n\n"
            "Verifique que tenga todas las dependencias instaladas.")
        
        # Registrar error
        with open('error_log.txt', 'a', encoding='utf-8') as f:
            f.write(f"\n[{datetime.now()}] ERROR DE INICIO:\n")
            f.write(traceback.format_exc())
            f.write("\n" + "=" * 80 + "\n")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        
        # Configurar el tema moderno si está disponible
        try:
            from ttkthemes import ThemedStyle
            style = ThemedStyle(root)
            style.set_theme("arc")  # Tema moderno
        except ImportError:
            style = ttk.Style()
            style.theme_use('clam')
        
        # Crear instancia de la aplicación con manejo de errores
        try:
            app = AplicacionDespachos(root)
            root.mainloop()
        except Exception as e:
            error_msg = f"Error al iniciar la aplicación:\n{str(e)}\n\nTraceback:\n{traceback.format_exc()}"
            messagebox.showerror("Error Crítico", error_msg)
            with open("error_log.txt", "a") as f:
                f.write(f"\n{datetime.now()} - Error al iniciar:\n{error_msg}\n")
    
    except Exception as e:
        messagebox.showerror("Error Inicial", f"No se pudo iniciar la aplicación:\n{str(e)}")
        with open("error_log.txt", "a") as f:
            f.write(f"\n{datetime.now()} - Error fatal:\n{traceback.format_exc()}\n")           
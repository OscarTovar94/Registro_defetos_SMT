"""
------------------------------------------------------------
Proyecto : Registro de defectos SMT
Autor    : Oscar Tovar
Versión  : 2.0
------------------------------------------------------------

Historial de Revisiones

Rev 1.0 - 21/07/2026
- Creación inicial de la aplicación.
Rev 2.0 - 28/07/2026
- Cambio en el registro de defectos.
------------------------------------------------------------
"""
import csv
import os
from datetime import datetime, date
from tkinter import messagebox
import tkinter as tk
import sys
from tkcalendar import DateEntry
from PIL import Image
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import PercentFormatter
import customtkinter as ctk
import win32event
import win32api
import winerror
from matplotlib.figure import Figure
import pandas as pd

# ---- Control de instancia única ----
MUTEX_NAME = "DefectosSMT_UnicaInstancia"

mutex = win32event.CreateMutex(
    None,
    False,
    MUTEX_NAME
)

if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)


class RegistroDefectosSMT:
    """Aplicación para registrar defectos del área SMT."""

    def __init__(self, master):
        self.root = master
        self.ventana_analisis_defectos = None

        self.root.title("Registro de Defectos SMT")
        # self.root.state("zoomed")
        self.root.iconbitmap("C:/Registro_defetos_SMT/Image/elrad.ico")
        self.root.configure(fg_color="#21233C")

        # Archivos de configuración
        self.archivo_defectos = "C:/Registro_defetos_SMT/Settings/defects.ini"
        self.archivo_modelos = "C:/Registro_defetos_SMT/Settings/models.ini"
        self.archivo_log_pcb = (
            "C:/Registro_defetos_SMT/LogFile/LogFilePCB.csv"
        )

        os.makedirs(
            os.path.dirname(self.archivo_log_pcb),
            exist_ok=True
        )

        # Variables
        self.modelo_seleccionado = ctk.StringVar(value="")
        self.lista_defectos = []
        self.configuracion_modelos = {}
        self.cerrando_aplicacion = False
        self.after_reloj = None
        self.after_archivos = None
        self.after_dashboard = None
        self.canvas_pareto = None
        self.fig_pareto = None
        self.ax_pareto = None
        self.tarjetas_modelos = {}
        self.lbl_sin_modelos = None
        self.dashboard_actualizando = False
        self.ax_pareto_porcentaje = None
        self.modelo_actual = None
        self.numero_parte_actual = ""
        self.renglones_panel = 0
        self.columnas_panel = 0
        self.total_pcb_panel = 0
        self.posiciones_defectuosas = {}
        self.botones_pcb = {}
        self.frame_panel_pcb = None
        self.lbl_info_modelo = None
        self.btn_confirmar_panel = None
        self.panel_actual = None
        self.ventana_captura_ids = None
        self.entries_ids_pcb = {}
        self.labels_estado_ids = {}
        self.proceso_panel_activo = False
        self.ventana_registro_defectos = None
        self.posiciones_pendientes_defectos = []
        self.indice_pcb_defecto_actual = 0
        self.ventana_analisis_defectos = None
        self.canvas_analisis_defectos = None
        self.frame_dashboard_analisis = None

        self.defectos_pcb_actual = {}
        self.filas_defectos_pcb = {}
        self.guardando_panel = False
        self.opcion_seleccionar_panel = "Seleccionar modelo"
        self.opcion_otro = "Otro"
        self.descripcion_otro_pcb = ctk.StringVar(value="")

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.cerrar_aplicacion
        )

        # Fechas de modificación para detectar cambios
        self.fecha_modificacion_defectos = None
        self.fecha_modificacion_modelos = None

        self.crear_header()
        self.crear_panel_registro()

        # Cargar información inicial
        self.actualizar_archivos(forzar=True)

        self.restablecer_seleccion_panel()

        # Revisar periódicamente si cambiaron los archivos
        self.verificar_cambios_archivos()

        self.root.after(
            300,
            self.actualizar_dashboard_fecha
        )

        self.root.after(100, lambda: self.root.state("zoomed"))

    def crear_header(self):
        """Frames y widgets del encabezado principal de la aplicación."""
        self.frame_header = ctk.CTkFrame(
            self.root,
            height=50,
            corner_radius=0,
            fg_color="#2B2D42"
        )
        self.frame_header.pack(fill="x")
        self.frame_header.pack_propagate(False)

        self.frame_header.grid_columnconfigure(0, weight=1)
        self.frame_header.grid_columnconfigure(1, weight=1)
        self.frame_header.grid_columnconfigure(2, weight=1)
        self.frame_header.grid_rowconfigure(0, weight=0)

        self.logo = ctk.CTkImage(
            light_image=Image.open(
                "C:/Registro_defetos_SMT/Image/elrad_image.png"
            ),
            dark_image=Image.open(
                "C:/Registro_defetos_SMT/Image/elrad_image.png"
            ),
            size=(55, 55)
        )

        self.lbl_logo = ctk.CTkLabel(
            self.frame_header,
            image=self.logo,
            text=""
        )
        self.lbl_logo.grid(row=0, column=0, padx=20, pady=5, sticky="w")

        self.lbl_titulo = ctk.CTkLabel(
            self.frame_header,
            text="Registro de Defectos SMT",
            font=("Arial", 28, "bold"),
            text_color="white"
        )
        self.lbl_titulo.grid(
            row=0,
            column=1
        )

        self.lbl_fecha_hora = ctk.CTkLabel(
            self.frame_header,
            text="",
            font=("Arial", 16),
            text_color="#D9D9D9"
        )
        self.lbl_fecha_hora.grid(
            row=0,
            column=2,
            padx=25,
            sticky="e"
        )

        self.actualizar_fecha_hora()

    def crear_panel_registro(self):
        """Crea el flujo de inspección por panel y el dashboard por PCB."""

        # =========================================================
        # SELECCIÓN DEL MODELO
        # =========================================================
        self.frame_registro = ctk.CTkFrame(
            self.root,
            height=75,
            corner_radius=12,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70"
        )
        self.frame_registro.pack(fill="x", padx=15, pady=(0, 0))
        self.frame_registro.pack_propagate(False)
        self.frame_registro.grid_columnconfigure(0, weight=1)
        self.frame_registro.grid_columnconfigure(1, weight=2)

        self.lbl_titulo_modelo = ctk.CTkLabel(
            self.frame_registro,
            text="Modelo",
            font=("Arial", 14, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_titulo_modelo.grid(
            row=0, column=0, padx=15, pady=(5, 5), sticky="nsew"
        )

        self.combo_modelos = ctk.CTkComboBox(
            self.frame_registro,
            variable=self.modelo_seleccionado,
            values=["Sin modelos"],
            height=32,
            corner_radius=8,
            font=("Arial", 14),
            dropdown_font=("Arial", 14),
            state="readonly",
            command=self.seleccionar_modelo
        )
        self.combo_modelos.grid(
            row=1, column=0, padx=15, pady=(0, 8), sticky="ew"
        )

        self.lbl_titulo_proceso = ctk.CTkLabel(
            self.frame_registro,
            text="Proceso de inspección",
            font=("Arial", 14, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_titulo_proceso.grid(
            row=0, column=1, padx=15, pady=(5, 5), sticky="nsew"
        )

        self.lbl_estado_proceso = ctk.CTkLabel(
            self.frame_registro,
            text=(
                "Seleccione únicamente las posiciones con defecto. "
                "Las demás PCB se registrarán como PASS."
            ),
            height=32,
            corner_radius=8,
            fg_color="#1F2238",
            font=("Arial", 14),
            text_color="#DDE2FF"
        )
        self.lbl_estado_proceso.grid(
            row=1, column=1, padx=15, pady=(0, 8), sticky="ew"
        )

        # =========================================================
        # ÁREA DE SELECCIÓN DEL PANEL PCB
        # =========================================================
        self.frame_panel_contenedor = ctk.CTkFrame(
            self.root,
            corner_radius=12,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70"
        )
        self.frame_panel_contenedor.pack(fill="x", padx=15, pady=(5, 0))
        self.frame_panel_contenedor.grid_columnconfigure(0, weight=1)

        self.lbl_info_modelo = ctk.CTkLabel(
            self.frame_panel_contenedor,
            text="Seleccione un modelo",
            font=("Arial", 17, "bold"),
            text_color="#DDE2FF"
        )
        self.lbl_info_modelo.grid(row=0, column=0, padx=20, pady=(10, 5))

        self.frame_panel_pcb = ctk.CTkFrame(
            self.frame_panel_contenedor,
            fg_color="#1F2238",
            corner_radius=12
        )
        self.frame_panel_pcb.grid(row=1, column=0, padx=20, pady=5)

        self.btn_confirmar_panel = ctk.CTkButton(
            self.frame_panel_contenedor,
            text="Confirmar panel",
            height=38,
            font=("Arial", 15, "bold"),
            state="disabled",
            command=self.confirmar_panel
        )
        self.btn_confirmar_panel.grid(
            row=2, column=0, padx=20, pady=(5, 10), sticky="ew"
        )

        # =========================================================
        # DASHBOARD
        # =========================================================
        self.frame_graficas = ctk.CTkFrame(
            self.root,
            corner_radius=12,
            fg_color="#292C47",
            border_width=1,
            border_color="#454B70"
        )
        self.frame_graficas.pack(
            fill="both", expand=True, padx=15, pady=(5, 5)
        )
        self.frame_graficas.grid_columnconfigure(0, weight=0)
        self.frame_graficas.grid_columnconfigure(1, weight=1)
        self.frame_graficas.grid_rowconfigure(0, weight=0)
        self.frame_graficas.grid_rowconfigure(1, weight=0)
        self.frame_graficas.grid_rowconfigure(2, weight=1)

        self.frame_filtro_graficas = ctk.CTkFrame(
            self.frame_graficas,
            height=30,
            corner_radius=10,
            fg_color="#252842"
        )
        self.frame_filtro_graficas.grid(
            row=0, column=0, columnspan=2,
            padx=15, pady=(5, 5), sticky="ew"
        )
        self.frame_filtro_graficas.grid_columnconfigure(2, weight=1)

        self.lbl_filtro_fecha = ctk.CTkLabel(
            self.frame_filtro_graficas,
            text="Fecha:",
            font=("Arial", 12, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_filtro_fecha.grid(row=0, column=0, padx=(15, 8), pady=5)

        self.selector_fecha = DateEntry(
            self.frame_filtro_graficas,
            width=12,
            date_pattern="dd/mm/yyyy",
            font=("Arial", 12),
            background="#2878D0",
            foreground="white",
            borderwidth=0
        )
        self.selector_fecha.set_date(date.today())
        self.selector_fecha.grid(row=0, column=1, padx=5, pady=12, sticky="w")
        self.selector_fecha.bind(
            "<<DateEntrySelected>>",
            lambda evento: self.solicitar_actualizacion_dashboard()
        )

        self.lbl_frase = ctk.CTkLabel(
            self.frame_filtro_graficas,
            text=(
                "Tus manos definen la calidad del producto; "
                "tu atención asegura el orgullo de nuestro trabajo."
            ),
            font=("Arial", 14, "bold", "italic"),
            text_color="#FFD166",
            justify="left",
            anchor="w"
        )
        self.lbl_frase.grid(row=0, column=2, padx=20, sticky="w")

        self.btn_actualizar_fpy = ctk.CTkButton(
            self.frame_filtro_graficas,
            text="Actualizar",
            width=120,
            height=30,
            font=("Arial", 12, "bold"),
            command=self.solicitar_actualizacion_dashboard
        )
        self.btn_actualizar_fpy.grid(row=0, column=3, padx=15, pady=5)

        self.btn_otros_defectos = ctk.CTkButton(
            self.frame_filtro_graficas,
            text="Análisis de defectos",
            width=150,
            height=32,
            font=("Arial", 13, "bold"),
            fg_color="#2E8B57",
            hover_color="#246B45",
            command=self.abrir_ventana_analisis_defectos
        )

        self.btn_otros_defectos.grid(
            row=0,
            column=4,
            padx=(5, 5),
            pady=8
        )

        # FPY total
        self.frame_fpy_total = ctk.CTkFrame(
            self.frame_graficas,
            width=250,
            height=260,
            corner_radius=14,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70"
        )
        self.frame_fpy_total.grid(
            row=1, column=0, padx=(15, 8), pady=(5, 5), sticky="nsew"
        )
        self.frame_fpy_total.grid_propagate(False)
        self.frame_fpy_total.grid_columnconfigure(0, weight=1)

        self.lbl_titulo_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text="FPY TOTAL",
            font=("Arial", 20, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_titulo_fpy.grid(row=0, column=0, padx=15, pady=(8, 2))

        self.lbl_valor_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text="0.00 %",
            font=("Arial", 42, "bold"),
            text_color="#8F96B8"
        )
        self.lbl_valor_fpy.grid(row=1, column=0, padx=15, pady=2)

        self.barra_fpy_total = ctk.CTkProgressBar(
            self.frame_fpy_total,
            width=215,
            height=11,
            corner_radius=6,
            fg_color="#454B70"
        )
        self.barra_fpy_total.grid(
            row=2, column=0, padx=18, pady=(0, 6), sticky="ew"
        )
        self.barra_fpy_total.set(0)

        self.lbl_detalle_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text=(
                "Inspeccionadas: 0\n"
                "Buenas: 0\n"
                "Defectuosas: 0\n"
                "Defectos encontrados: 0"
            ),
            font=("Arial", 13),
            text_color="#AEB4D0",
            justify="left",
            anchor="w"
        )
        self.lbl_detalle_fpy.grid(
            row=3, column=0, padx=18, pady=3, sticky="w"
        )

        self.lbl_top_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text="Sin defectos registrados",
            font=("Arial", 11, "bold"),
            justify="left",
            anchor="w",
            text_color="#DDE2FF",
            wraplength=215
        )
        self.lbl_top_fpy.grid(
            row=4, column=0, padx=18, pady=(3, 8), sticky="w"
        )

        # Tarjetas por modelo
        self.frame_fpy_modelos = ctk.CTkScrollableFrame(
            self.frame_graficas,
            corner_radius=14,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70",
            orientation="horizontal",
            height=260
        )
        self.frame_fpy_modelos.grid(
            row=1, column=1, padx=(8, 15), pady=2, sticky="nsew"
        )

        # Pareto global
        self.frame_pareto_global = ctk.CTkFrame(
            self.frame_graficas,
            corner_radius=14,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70"
        )
        self.frame_pareto_global.grid(
            row=2, column=0, columnspan=2,
            padx=15, pady=(5, 5), sticky="nsew"
        )
        self.frame_pareto_global.grid_rowconfigure(0, weight=1)
        self.frame_pareto_global.grid_rowconfigure(1, weight=0)
        self.frame_pareto_global.grid_columnconfigure(0, weight=1)

        self.lbl_by = ctk.CTkLabel(
            self.frame_pareto_global,
            text="Rev: 2.0 (By: Oscar Tovar)",
            font=("Arial", 10),
            text_color="#AEB4D0"
        )
        self.lbl_by.grid(row=1, column=0, padx=15, pady=(1, 5), sticky="e")

        self.inicializar_pareto()

    @staticmethod
    def leer_lista_archivo(ruta):
        """
        Lee un archivo con un elemento por línea.

        Ignora:
        - Líneas vacías.
        - Líneas que comiencen con # o ;
        - Elementos repetidos.
        """
        if not os.path.exists(ruta):
            return []

        elementos = []

        try:
            with open(ruta, "r", encoding="utf-8-sig") as archivo:
                for linea in archivo:
                    elemento = linea.strip()

                    if not elemento:
                        continue

                    if elemento.startswith("#") or elemento.startswith(";"):
                        continue

                    if elemento not in elementos:
                        elementos.append(elemento)

        except UnicodeDecodeError:
            with open(ruta, "r", encoding="latin-1") as archivo:
                for linea in archivo:
                    elemento = linea.strip()

                    if (
                        elemento
                        and not elemento.startswith(("#", ";"))
                        and elemento not in elementos
                    ):
                        elementos.append(elemento)

        except OSError as error:
            messagebox.showerror(
                "Error de lectura",
                f"No fue posible leer el archivo:\n{ruta}\n\n{error}"
            )

        return elementos

    @staticmethod
    def leer_configuracion_modelos(ruta):
        """
        Lee models.ini con el formato:

        Modelo,Renglones,Columnas,NumeroParte

        Ejemplo:
        Lion Mite,4,5,635125

        Retorna:
        {
            "Lion Mite": {
                "renglones": 4,
                "columnas": 5,
                "numero_parte": "635125",
                "total_pcb": 20
            }
        }
        """

        modelos = {}

        if not os.path.exists(ruta):
            return modelos

        try:
            with open(
                ruta,
                mode="r",
                encoding="utf-8-sig"
            ) as archivo:

                for numero_linea, linea in enumerate(
                    archivo,
                    start=1
                ):
                    linea = linea.strip()

                    if not linea:
                        continue

                    if linea.startswith(("#", ";")):
                        continue

                    partes = linea.split(",")

                    if len(partes) != 4:
                        print(
                            f"Línea incorrecta en models.ini "
                            f"({numero_linea}): {linea}"
                        )
                        continue

                    modelo = partes[0].strip()
                    renglones_texto = partes[1].strip()
                    columnas_texto = partes[2].strip()
                    numero_parte = partes[3].strip()

                    if not modelo:
                        print(
                            f"Modelo vacío en línea {numero_linea}."
                        )
                        continue

                    try:
                        renglones = int(renglones_texto)
                        columnas = int(columnas_texto)

                    except ValueError:
                        print(
                            f"Renglones o columnas incorrectos "
                            f"para el modelo {modelo}."
                        )
                        continue

                    if renglones <= 0 or columnas <= 0:
                        print(
                            f"Los renglones y columnas de {modelo} "
                            "deben ser mayores que cero."
                        )
                        continue

                    if len(numero_parte) != 6:
                        print(
                            f"El número de parte de {modelo} debe "
                            "contener exactamente 6 dígitos."
                        )
                        continue

                    if not numero_parte.isdigit():
                        print(
                            f"El número de parte de {modelo} debe "
                            "contener solamente números."
                        )
                        continue

                    total_pcb = renglones * columnas

                    modelos[modelo] = {
                        "renglones": renglones,
                        "columnas": columnas,
                        "numero_parte": numero_parte,
                        "total_pcb": total_pcb
                    }

        except UnicodeDecodeError:
            messagebox.showerror(
                "Codificación incorrecta",
                (
                    "No fue posible leer models.ini.\n\n"
                    "Guarde el archivo con codificación UTF-8."
                )
            )

        except OSError as error:
            messagebox.showerror(
                "Error de lectura",
                (
                    "No fue posible leer models.ini.\n\n"
                    f"{error}"
                )
            )

        return modelos

    def cargar_defectos(self):
        """Carga la lista de defectos utilizada por el registro por PCB."""
        self.lista_defectos = self.leer_lista_archivo(
            self.archivo_defectos
        )

    def cargar_modelos(self):
        """
        Carga la configuración de modelos desde models.ini.
        """

        self.configuracion_modelos = (
            self.leer_configuracion_modelos(
                self.archivo_modelos
            )
        )

        modelos = [
            self.opcion_seleccionar_panel
        ] + list(
            self.configuracion_modelos.keys()
        )

        modelo_actual = self.modelo_seleccionado.get()

        if modelos:
            self.combo_modelos.configure(
                values=modelos,
                state="readonly"
            )

            if modelo_actual in modelos:
                self.modelo_seleccionado.set(
                    modelo_actual
                )
            else:
                self.modelo_seleccionado.set(
                    modelos[0]
                )

        else:
            self.combo_modelos.configure(
                values=["Sin modelos"],
                state="disabled"
            )

            self.modelo_seleccionado.set(
                "Sin modelos"
            )

    def verificar_cambios_archivos(self):
        """Verifica si los archivos defects.ini y 
        models.ini han cambiado y actualiza la información."""

        if self.cerrando_aplicacion:
            return

        try:
            if not self.root.winfo_exists():
                return

            # Aquí queda tu código actual para verificar
            # defects.ini y models.ini
            self.actualizar_archivos()

            self.after_archivos = self.root.after(
                2000,
                self.verificar_cambios_archivos
            )

        except tk.TclError:
            pass

    def actualizar_archivos(self, forzar=False):
        """Actualiza la lista de defectos y modelos si los archivos han cambiado."""
        modificacion_defectos = self.obtener_fecha_modificacion(
            self.archivo_defectos
        )
        modificacion_modelos = self.obtener_fecha_modificacion(
            self.archivo_modelos
        )

        if (
            forzar
            or modificacion_defectos != self.fecha_modificacion_defectos
        ):
            self.fecha_modificacion_defectos = modificacion_defectos
            self.cargar_defectos()

        if (
            forzar
            or modificacion_modelos != self.fecha_modificacion_modelos
        ):
            self.fecha_modificacion_modelos = modificacion_modelos
            self.cargar_modelos()

    @staticmethod
    def obtener_fecha_modificacion(ruta):
        """Obtiene la fecha de modificación de un archivo."""
        try:
            return os.path.getmtime(ruta)
        except OSError:
            return None

    def cargar_datos_dashboard(self):
        """Lee LogFilePCB.csv y filtra los registros por la fecha seleccionada."""
        if not os.path.exists(self.archivo_log_pcb):
            return [], [], "No existe LogFilePCB.csv"

        if os.path.getsize(self.archivo_log_pcb) == 0:
            return [], [], "LogFilePCB.csv está vacío"

        columnas_fijas = {
            "Sesion", "Modelo", "NumeroParte", "Posicion", "Renglon",
            "Columna", "ID_PCB", "Resultado", "Defectos", "DescripcionOtro", "Fecha/Hora"
        }

        try:
            with open(
                self.archivo_log_pcb,
                mode="r",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:
                lector = csv.DictReader(archivo)
                encabezados = lector.fieldnames or []

                columnas_requeridas = {
                    "Modelo", "NumeroParte", "Resultado",
                    "Defectos", "Fecha/Hora"
                }

                if not columnas_requeridas.issubset(encabezados):
                    return [], [], (
                        "LogFilePCB.csv no contiene las columnas requeridas"
                    )

                columnas_defectos = [
                    columna
                    for columna in encabezados
                    if columna not in columnas_fijas
                ]

                fecha_seleccionada = self.selector_fecha.get_date()
                registros = []

                for fila in lector:
                    fecha_texto = fila.get("Fecha/Hora", "").strip()

                    try:
                        fecha_registro = datetime.strptime(
                            fecha_texto,
                            "%d/%m/%Y %H:%M:%S"
                        ).date()
                    except ValueError:
                        continue

                    if fecha_registro != fecha_seleccionada:
                        continue

                    resultado = fila.get("Resultado", "").strip().upper()

                    if resultado not in {"PASS", "FAIL"}:
                        continue

                    registros.append(fila)

                return registros, columnas_defectos, ""

        except PermissionError:
            return [], [], (
                "LogFilePCB.csv está siendo utilizado por otro programa"
            )

        except OSError as error:
            return [], [], f"No fue posible leer LogFilePCB.csv: {error}"

    @staticmethod
    def convertir_entero(valor):
        """Convierte un valor del CSV a entero; un valor inválido equivale a cero."""
        try:
            return int(float(valor or 0))
        except (ValueError, TypeError):
            return 0

    def calcular_fpy_total(self, registros=None, columnas_defectos=None):
        """Calcula el FPY global mediante los resultados PASS y FAIL."""
        if registros is None or columnas_defectos is None:
            registros, columnas_defectos, error = self.cargar_datos_dashboard()

            if error:
                self.mostrar_fpy_sin_datos(error)
                return

        if not registros:
            self.mostrar_fpy_sin_datos(
                "Sin registros para la fecha seleccionada"
            )
            return

        inspeccionadas = len(registros)
        buenas = sum(
            1
            for fila in registros
            if fila.get("Resultado", "").strip().upper() == "PASS"
        )
        defectuosas = inspeccionadas - buenas
        defectos_encontrados = sum(
            self.convertir_entero(fila.get("Defectos", 0))
            for fila in registros
        )

        totales_defectos = {
            defecto: sum(
                self.convertir_entero(fila.get(defecto, 0))
                for fila in registros
            )
            for defecto in columnas_defectos
        }

        top_3 = sorted(
            (
                (defecto, cantidad)
                for defecto, cantidad in totales_defectos.items()
                if cantidad > 0
            ),
            key=lambda elemento: elemento[1],
            reverse=True
        )[:3]

        fpy = (buenas / inspeccionadas) * 100 if inspeccionadas else 0.0
        color = self.obtener_color_fpy(fpy)

        self.lbl_valor_fpy.configure(
            text=f"{fpy:.2f} %",
            text_color=color
        )
        self.barra_fpy_total.configure(progress_color=color)
        self.barra_fpy_total.set(max(0.0, min(fpy / 100, 1.0)))

        self.lbl_detalle_fpy.configure(
            text=(
                f"Inspeccionadas: {inspeccionadas}\n"
                f"Buenas: {buenas}\n"
                f"Defectuosas: {defectuosas}\n"
                f"Defectos encontrados: {defectos_encontrados}"
            )
        )
        self.lbl_top_fpy.configure(
            text=self.formatear_top_3_defectos(top_3)
        )

    def mostrar_fpy_sin_datos(self, mensaje):
        """Limpia la tarjeta del FPY total cuando no existen datos."""
        self.lbl_valor_fpy.configure(
            text="0.00 %",
            text_color="#8F96B8"
        )
        self.barra_fpy_total.configure(progress_color="#8F96B8")
        self.barra_fpy_total.set(0)
        self.lbl_detalle_fpy.configure(text=mensaje)
        self.lbl_top_fpy.configure(text="Sin defectos registrados")

    @staticmethod
    def obtener_color_fpy(fpy):
        """Retorna el color correspondiente al rango del FPY."""
        if fpy >= 98:
            return "#6FE3A1"
        if fpy >= 95:
            return "#FFD166"
        return "#FF6B6B"

    def aplicar_color_fpy(self, fpy):
        """Aplica un color al valor de FPY según el rango."""
        if fpy >= 98:
            color = "#6FE3A1"

        elif fpy >= 95:
            color = "#FFD166"

        else:
            color = "#FF6B6B"

        self.lbl_valor_fpy.configure(
            text_color=color
        )

    def actualizar_dashboard_fecha(self):
        """Actualiza una sola vez todos los elementos del dashboard."""
        if self.cerrando_aplicacion:
            return

        self.after_dashboard = None

        if self.dashboard_actualizando:
            return

        self.dashboard_actualizando = True

        try:
            registros, columnas_defectos, error = (
                self.cargar_datos_dashboard()
            )

            if error:
                self.mostrar_fpy_sin_datos(error)
                self.mostrar_sin_modelos(error)
                self.mostrar_mensaje_pareto(error)
                return

            self.calcular_fpy_total(
                registros,
                columnas_defectos
            )
            self.calcular_fpy_por_modelo(
                registros,
                columnas_defectos
            )
            self.actualizar_pareto_global(
                registros,
                columnas_defectos
            )

        finally:
            self.dashboard_actualizando = False

    def calcular_fpy_por_modelo(
        self,
        registros=None,
        columnas_defectos=None
    ):
        """Calcula y muestra las métricas de cada modelo por PCB."""
        for datos_tarjeta in self.tarjetas_modelos.values():
            datos_tarjeta["frame"].grid_remove()

        if registros is None or columnas_defectos is None:
            registros, columnas_defectos, error = (
                self.cargar_datos_dashboard()
            )

            if error:
                self.mostrar_sin_modelos(error)
                return

        if not registros:
            self.mostrar_sin_modelos(
                "Sin registros para la fecha seleccionada"
            )
            return

        datos_modelos = {}

        for fila in registros:
            modelo = fila.get("Modelo", "").strip()

            if not modelo:
                continue

            resultado = fila.get("Resultado", "").strip().upper()
            numero_parte = fila.get("NumeroParte", "").strip()

            if modelo not in datos_modelos:
                datos_modelos[modelo] = {
                    "numero_parte": numero_parte,
                    "inspeccionadas": 0,
                    "buenas": 0,
                    "defectuosas": 0,
                    "defectos_encontrados": 0,
                    "defectos_individuales": {
                        defecto: 0
                        for defecto in columnas_defectos
                    }
                }

            datos = datos_modelos[modelo]
            datos["inspeccionadas"] += 1

            if resultado == "PASS":
                datos["buenas"] += 1
            else:
                datos["defectuosas"] += 1

            datos["defectos_encontrados"] += self.convertir_entero(
                fila.get("Defectos", 0)
            )

            for defecto in columnas_defectos:
                datos["defectos_individuales"][defecto] += (
                    self.convertir_entero(fila.get(defecto, 0))
                )

        if not datos_modelos:
            self.mostrar_sin_modelos(
                "Sin registros para la fecha seleccionada"
            )
            return

        if self.lbl_sin_modelos is not None:
            self.lbl_sin_modelos.grid_remove()

        for columna, modelo in enumerate(sorted(datos_modelos)):
            datos = datos_modelos[modelo]
            inspeccionadas = datos["inspeccionadas"]
            buenas = datos["buenas"]
            fpy = (
                buenas / inspeccionadas * 100
                if inspeccionadas > 0
                else 0.0
            )

            top_3_defectos = sorted(
                (
                    (defecto, cantidad)
                    for defecto, cantidad
                    in datos["defectos_individuales"].items()
                    if cantidad > 0
                ),
                key=lambda elemento: elemento[1],
                reverse=True
            )[:3]

            self.actualizar_tarjeta_fpy_modelo(
                modelo=modelo,
                numero_parte=datos["numero_parte"],
                fpy=fpy,
                inspeccionadas=inspeccionadas,
                buenas=buenas,
                defectuosas=datos["defectuosas"],
                defectos_encontrados=datos["defectos_encontrados"],
                top_3_defectos=top_3_defectos,
                columna=columna
            )

    def formatear_top_3_defectos(self, top_3_defectos):
        """
        Forma el texto de los tres defectos principales.
        """

        if not top_3_defectos:
            return "Sin defectos registrados"

        posiciones = ["🥇", "🥈", "🥉"]
        lineas = []

        for indice, (defecto, cantidad) in enumerate(top_3_defectos):

            lineas.append(
                f"{posiciones[indice]}  {defecto}: {cantidad}"
            )

        return "\n".join(lineas)

    def mostrar_sin_modelos(self, mensaje):
        """
        Muestra un mensaje cuando no existen modelos
        para la fecha seleccionada.
        """

        for datos_tarjeta in self.tarjetas_modelos.values():
            datos_tarjeta["frame"].grid_remove()

        if self.lbl_sin_modelos is None:
            self.lbl_sin_modelos = ctk.CTkLabel(
                self.frame_fpy_modelos,
                text=mensaje,
                font=("Arial", 16),
                text_color="#8F96B8"
            )
        else:
            self.lbl_sin_modelos.configure(
                text=mensaje
            )

        self.lbl_sin_modelos.grid(
            row=0,
            column=0,
            padx=30,
            pady=2,
            sticky="nsew"
        )

    def actualizar_pareto_global(
        self,
        registros=None,
        columnas_defectos=None
    ):
        """Genera el Pareto global mediante ocurrencias de defectos."""
        if registros is None or columnas_defectos is None:
            registros, columnas_defectos, error = (
                self.cargar_datos_dashboard()
            )

            if error:
                self.mostrar_mensaje_pareto(error)
                return

        if not registros:
            self.mostrar_mensaje_pareto(
                "Sin registros para la fecha seleccionada"
            )
            return

        totales_defectos = {
            defecto: sum(
                self.convertir_entero(fila.get(defecto, 0))
                for fila in registros
            )
            for defecto in columnas_defectos
        }

        defectos_con_datos = sorted(
            (
                (defecto, cantidad)
                for defecto, cantidad in totales_defectos.items()
                if cantidad > 0
            ),
            key=lambda elemento: elemento[1],
            reverse=True
        )[:15]

        if not defectos_con_datos:
            self.mostrar_mensaje_pareto(
                "Sin defectos registrados para la fecha seleccionada"
            )
            return

        nombres = [elemento[0] for elemento in defectos_con_datos]
        cantidades = [elemento[1] for elemento in defectos_con_datos]
        total = sum(cantidades)

        acumulado = []
        suma_acumulada = 0

        for cantidad in cantidades:
            suma_acumulada += cantidad
            acumulado.append((suma_acumulada / total) * 100)

        self.crear_grafica_pareto(
            nombres=nombres,
            cantidades=cantidades,
            acumulado=acumulado
        )

    def crear_grafica_pareto(
        self,
        nombres,
        cantidades,
        acumulado
    ):
        """
        Actualiza el Pareto sin destruir el canvas.
        """

        if self.fig_pareto is None or self.canvas_pareto is None:
            return

        # Limpiar solamente el contenido de la figura
        self.fig_pareto.clear()

        self.ax_pareto = self.fig_pareto.add_subplot(111)
        self.ax_pareto_porcentaje = self.ax_pareto.twinx()

        self.fig_pareto.patch.set_facecolor("#252842")
        self.ax_pareto.set_facecolor("#252842")
        self.ax_pareto_porcentaje.set_facecolor("none")

        posiciones = list(range(len(nombres)))

        barras = self.ax_pareto.bar(
            posiciones,
            cantidades,
            color="#4DA3FF",
            edgecolor="#79C2FF",
            linewidth=0.8
        )

        self.ax_pareto.set_title(
            "PARETO GLOBAL DE DEFECTOS",
            fontsize=15,
            fontweight="bold",
            color="#DDE2FF",
            pad=15
        )

        self.ax_pareto.set_ylabel(
            "Cantidad de defectos",
            color="#AEB4D0",
            fontsize=11
        )

        self.ax_pareto.set_xticks(posiciones)

        self.ax_pareto.set_xticklabels(
            nombres,
            rotation=35,
            ha="right",
            color="#DDE2FF",
            fontsize=9
        )

        self.ax_pareto.tick_params(
            axis="y",
            colors="#DDE2FF"
        )

        self.ax_pareto.grid(
            axis="y",
            linestyle="--",
            alpha=0.20
        )

        self.ax_pareto.spines["top"].set_visible(False)
        self.ax_pareto.spines["right"].set_visible(False)
        self.ax_pareto.spines["bottom"].set_color("#454B70")
        self.ax_pareto.spines["left"].set_color("#454B70")

        for barra, cantidad in zip(barras, cantidades):
            self.ax_pareto.text(
                barra.get_x() + barra.get_width() / 2,
                barra.get_height(),
                str(cantidad),
                ha="center",
                va="bottom",
                color="#DDE2FF",
                fontsize=9,
                fontweight="bold"
            )

        # Línea acumulada
        self.ax_pareto_porcentaje.plot(
            posiciones,
            acumulado,
            color="#FFD166",
            marker="o",
            linewidth=2.2,
            markersize=5
        )

        self.ax_pareto_porcentaje.set_ylim(0, 110)

        self.ax_pareto_porcentaje.set_ylabel(
            "Porcentaje acumulado",
            color="#FFD166",
            fontsize=11
        )

        self.ax_pareto_porcentaje.tick_params(
            axis="y",
            colors="#FFD166"
        )

        self.ax_pareto_porcentaje.yaxis.set_major_formatter(
            PercentFormatter()
        )

        self.ax_pareto_porcentaje.spines["top"].set_visible(False)
        self.ax_pareto_porcentaje.spines["left"].set_visible(False)
        self.ax_pareto_porcentaje.spines["right"].set_color("#454B70")

        self.ax_pareto_porcentaje.axhline(
            y=80,
            color="#FF6B6B",
            linestyle="--",
            linewidth=1.2,
            alpha=0.85
        )

        self.ax_pareto_porcentaje.text(
            len(nombres) - 1,
            82,
            "80 %",
            color="#FF6B6B",
            ha="right",
            fontsize=9,
            fontweight="bold"
        )

        self.fig_pareto.tight_layout()

        # Actualización suave
        self.canvas_pareto.draw_idle()

    def mostrar_mensaje_pareto(self, mensaje):
        """
        Muestra un mensaje dentro del canvas sin destruirlo.
        """

        if self.fig_pareto is None or self.canvas_pareto is None:
            return

        self.fig_pareto.clear()

        self.ax_pareto = self.fig_pareto.add_subplot(111)

        self.fig_pareto.patch.set_facecolor("#252842")
        self.ax_pareto.set_facecolor("#252842")

        self.ax_pareto.text(
            0.5,
            0.5,
            mensaje,
            transform=self.ax_pareto.transAxes,
            ha="center",
            va="center",
            fontsize=15,
            color="#8F96B8"
        )

        self.ax_pareto.set_xticks([])
        self.ax_pareto.set_yticks([])

        for borde in self.ax_pareto.spines.values():
            borde.set_visible(False)

        self.canvas_pareto.draw_idle()

    def actualizar_fecha_hora(self):
        """Actualiza la etiqueta de fecha y hora cada segundo."""

        if self.cerrando_aplicacion:
            return

        try:
            if not self.root.winfo_exists():
                return

            ahora = datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S"
            )

            self.lbl_fecha_hora.configure(
                text=ahora
            )

            self.after_reloj = self.root.after(
                1000,
                self.actualizar_fecha_hora
            )

        except tk.TclError:
            pass

    def cerrar_aplicacion(self):
        """
        Cancela tareas pendientes y cierra correctamente
        la aplicación.
        """

        if self.cerrando_aplicacion:
            return

        self.cerrando_aplicacion = True

        tareas_after = [
            self.after_reloj,
            self.after_archivos,
            self.after_dashboard
        ]

        for tarea in tareas_after:
            if tarea is not None:
                try:
                    self.root.after_cancel(tarea)
                except (tk.TclError, ValueError):
                    pass

        # Cerrar figuras de Matplotlib
        try:
            if self.fig_pareto is not None:
                plt.close(self.fig_pareto)
        except (AttributeError, tk.TclError):
            pass

        # Destruir el canvas de Matplotlib
        try:
            if self.canvas_pareto is not None:
                widget_canvas = self.canvas_pareto.get_tk_widget()

                if widget_canvas.winfo_exists():
                    widget_canvas.destroy()
        except (AttributeError, tk.TclError):
            pass

        try:
            self.root.quit()
        except tk.TclError:
            pass

        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def inicializar_pareto(self):
        """
        Crea la figura y el canvas del Pareto una sola vez.
        """

        self.fig_pareto = plt.Figure(
            figsize=(12, 4.5),
            dpi=100,
            facecolor="#252842"
        )

        self.ax_pareto = self.fig_pareto.add_subplot(111)

        self.canvas_pareto = FigureCanvasTkAgg(
            self.fig_pareto,
            master=self.frame_pareto_global
        )

        self.canvas_pareto.get_tk_widget().grid(
            row=0,
            column=0,
            padx=10,
            pady=10,
            sticky="nsew"
        )

        self.mostrar_mensaje_pareto(
            "Sin información para mostrar"
        )

    def actualizar_tarjeta_fpy_modelo(
        self,
        modelo,
        numero_parte,
        fpy,
        inspeccionadas,
        buenas,
        defectuosas,
        defectos_encontrados,
        top_3_defectos,
        columna
    ):
        """Crea o actualiza una tarjeta de resultados por modelo."""
        color_fpy = self.obtener_color_fpy(fpy)
        texto_top = self.formatear_top_3_defectos(
            top_3_defectos
        )
        texto_totales = (
            f"Inspeccionadas: {inspeccionadas}\n"
            f"Buenas: {buenas}\n"
            f"Defectuosas: {defectuosas}\n"
            f"Defectos encontrados: {defectos_encontrados}"
        )

        if modelo in self.tarjetas_modelos:
            datos_tarjeta = self.tarjetas_modelos[modelo]

            datos_tarjeta["lbl_numero_parte"].configure(
                text=f"Número de parte: {numero_parte or 'N/D'}"
            )
            datos_tarjeta["lbl_fpy"].configure(
                text=f"{fpy:.2f} %",
                text_color=color_fpy
            )
            datos_tarjeta["barra_fpy"].configure(
                progress_color=color_fpy
            )
            datos_tarjeta["barra_fpy"].set(
                max(0.0, min(fpy / 100, 1.0))
            )
            datos_tarjeta["lbl_totales"].configure(
                text=texto_totales
            )
            datos_tarjeta["lbl_top"].configure(
                text=texto_top
            )
            datos_tarjeta["frame"].grid(
                row=0,
                column=columna,
                padx=8,
                pady=2,
                sticky="nsew"
            )
            return

        tarjeta = ctk.CTkFrame(
            self.frame_fpy_modelos,
            width=275,
            height=250,
            corner_radius=12,
            fg_color="#292C47",
            border_width=1,
            border_color="#454B70"
        )
        tarjeta.grid(
            row=0,
            column=columna,
            padx=8,
            pady=2,
            sticky="nsew"
        )
        tarjeta.grid_propagate(False)
        tarjeta.grid_columnconfigure(0, weight=1)

        lbl_modelo = ctk.CTkLabel(
            tarjeta,
            text=modelo,
            font=("Arial", 16, "bold"),
            text_color="#DDE2FF"
        )
        lbl_modelo.grid(row=0, column=0, padx=12, pady=(5, 0))

        lbl_numero_parte = ctk.CTkLabel(
            tarjeta,
            text=f"Número de parte: {numero_parte or 'N/D'}",
            font=("Arial", 11),
            text_color="#AEB4D0"
        )
        lbl_numero_parte.grid(row=1, column=0, padx=12, pady=0)

        lbl_fpy = ctk.CTkLabel(
            tarjeta,
            text=f"{fpy:.2f} %",
            font=("Arial", 30, "bold"),
            text_color=color_fpy
        )
        lbl_fpy.grid(row=2, column=0, padx=12, pady=1)

        barra_fpy = ctk.CTkProgressBar(
            tarjeta,
            width=230,
            height=10,
            corner_radius=5,
            progress_color=color_fpy,
            fg_color="#454B70"
        )
        barra_fpy.grid(
            row=3,
            column=0,
            padx=20,
            pady=(0, 4),
            sticky="ew"
        )
        barra_fpy.set(max(0.0, min(fpy / 100, 1.0)))

        lbl_totales = ctk.CTkLabel(
            tarjeta,
            text=texto_totales,
            font=("Arial", 11),
            text_color="#AEB4D0",
            justify="left",
            anchor="w"
        )
        lbl_totales.grid(
            row=4,
            column=0,
            padx=16,
            pady=1,
            sticky="w"
        )

        lbl_titulo_top = ctk.CTkLabel(
            tarjeta,
            text="TOP DEFECTOS",
            font=("Arial", 11, "bold"),
            text_color="#79C2FF"
        )
        lbl_titulo_top.grid(row=5, column=0, padx=12, pady=(2, 0))

        lbl_top = ctk.CTkLabel(
            tarjeta,
            text=texto_top,
            font=("Arial", 10, "bold"),
            text_color="#DDE2FF",
            justify="left",
            anchor="w",
            wraplength=240
        )
        lbl_top.grid(
            row=6,
            column=0,
            padx=15,
            pady=(0, 4),
            sticky="ew"
        )

        self.tarjetas_modelos[modelo] = {
            "frame": tarjeta,
            "lbl_modelo": lbl_modelo,
            "lbl_numero_parte": lbl_numero_parte,
            "lbl_fpy": lbl_fpy,
            "barra_fpy": barra_fpy,
            "lbl_totales": lbl_totales,
            "lbl_titulo_top": lbl_titulo_top,
            "lbl_top": lbl_top
        }

    def solicitar_actualizacion_dashboard(self):
        """
        Agrupa varias solicitudes de actualización en una sola.
        """
        if self.cerrando_aplicacion:
            return

        if self.after_dashboard is not None:
            try:
                self.root.after_cancel(
                    self.after_dashboard
                )
            except tk.TclError:
                pass

        self.after_dashboard = self.root.after(
            80,
            self.actualizar_dashboard_fecha
        )

    def seleccionar_modelo(self, modelo):
        """
        Carga en memoria la configuración del modelo seleccionado.
        """

        if self.proceso_panel_activo:
            messagebox.showwarning(
                "Inspección en proceso",
                (
                    "No puede cambiar el modelo mientras exista "
                    "una inspección en proceso."
                )
            )

            if self.modelo_actual:
                self.modelo_seleccionado.set(
                    self.modelo_actual
                )

            return

        modelo = modelo.strip()
        if modelo == self.opcion_seleccionar_panel:
            self.restablecer_seleccion_panel()
            return

        if modelo not in self.configuracion_modelos:
            self.modelo_actual = None
            self.numero_parte_actual = ""
            self.renglones_panel = 0
            self.columnas_panel = 0
            self.total_pcb_panel = 0
            return

        configuracion = self.configuracion_modelos[
            modelo
        ]

        self.modelo_actual = modelo

        self.numero_parte_actual = configuracion[
            "numero_parte"
        ]

        self.renglones_panel = configuracion[
            "renglones"
        ]

        self.columnas_panel = configuracion[
            "columnas"
        ]

        self.total_pcb_panel = configuracion[
            "total_pcb"
        ]

        # Reiniciar los datos del panel anterior
        self.posiciones_defectuosas.clear()
        self.botones_pcb.clear()

        print(
            f"Modelo: {self.modelo_actual} | "
            f"Número de parte: {self.numero_parte_actual} | "
            f"Configuración: "
            f"{self.renglones_panel} x "
            f"{self.columnas_panel} | "
            f"Total PCB: {self.total_pcb_panel}"
        )

        self.mostrar_panel_modelo()

    def mostrar_panel_modelo(self):
        """
        Dibuja la cuadrícula del modelo seleccionado.
        """
        self.frame_panel_pcb.grid()

        if self.frame_panel_pcb is None:
            return

        for widget in self.frame_panel_pcb.winfo_children():
            widget.destroy()

        self.botones_pcb.clear()
        self.posiciones_defectuosas.clear()

        if not self.modelo_actual:
            self.lbl_info_modelo.configure(
                text="Seleccione un modelo"
            )

            self.btn_confirmar_panel.configure(
                state="disabled"
            )
            return

        self.lbl_info_modelo.configure(
            text=(
                f"Modelo: {self.modelo_actual}     "
                f"Número de parte: {self.numero_parte_actual}     "
                f"Configuración: "
                f"{self.renglones_panel} × "
                f"{self.columnas_panel}     "
                f"Total PCB: {self.total_pcb_panel}"
            )
        )

        for columna in range(self.columnas_panel):
            self.frame_panel_pcb.grid_columnconfigure(
                columna,
                weight=1,
                uniform="pcb"
            )

        numero_pcb = 1

        for renglon in range(self.renglones_panel):
            for columna in range(self.columnas_panel):

                boton = ctk.CTkButton(
                    self.frame_panel_pcb,
                    text=f"PCB {numero_pcb}",
                    width=115,
                    height=58,
                    corner_radius=10,
                    font=("Arial", 15, "bold"),
                    fg_color="#454B70",
                    hover_color="#596083",
                    command=lambda posicion=numero_pcb: (
                        self.seleccionar_posicion_pcb(
                            posicion
                        )
                    )
                )

                boton.grid(
                    row=renglon,
                    column=columna,
                    padx=7,
                    pady=7,
                    sticky="nsew"
                )

                self.botones_pcb[numero_pcb] = boton
                numero_pcb += 1

        self.btn_confirmar_panel.configure(
            state="normal"
        )

    def seleccionar_posicion_pcb(self, posicion):
        """
        Marca o desmarca una posición como PCB defectuosa.
        """

        if posicion in self.posiciones_defectuosas:

            self.posiciones_defectuosas.pop(
                posicion,
                None
            )

            self.botones_pcb[posicion].configure(
                fg_color="#454B70",
                hover_color="#596083",
                text=f"PCB {posicion}"
            )

        else:
            self.posiciones_defectuosas[posicion] = {
                "posicion": posicion,
                "id_pcb": "",
                "defectos": [],
                "estado": "PENDIENTE_ID"
            }

            self.botones_pcb[posicion].configure(
                fg_color="#D97706",
                hover_color="#B45309",
                text=(
                    f"PCB {posicion}\n"
                    "DEFECTO"
                )
            )

    def confirmar_panel(self):
        """
        Confirma las posiciones defectuosas seleccionadas e inicia
        el proceso de captura de ID.
        """

        if not self.modelo_actual:
            messagebox.showwarning(
                "Modelo requerido",
                "Seleccione un modelo antes de confirmar el panel."
            )
            return

        total_defectuosas = len(
            self.posiciones_defectuosas
        )

        total_buenas = (
            self.total_pcb_panel
            - total_defectuosas
        )

        if total_defectuosas == 0:
            mensaje_adicional = (
                "\n\nNo se seleccionaron PCB defectuosas.\n"
                "Todas las posiciones se registrarán como PASS."
            )
        else:
            mensaje_adicional = (
                "\n\nDespués de confirmar deberá capturar "
                "el ID de cada PCB defectuosa."
            )

        # respuesta = messagebox.askyesno(
            # "Confirmar panel",
            # (
            # f"Modelo: {self.modelo_actual}\n"
            # f"Número de parte: {self.numero_parte_actual}\n\n"
            # f"PCB totales: {self.total_pcb_panel}\n"
            # f"PCB buenas: {total_buenas}\n"
            # f"PCB defectuosas: {total_defectuosas}"
            # f"{mensaje_adicional}\n\n"
            # "¿Desea confirmar la selección?"
            # )
        # )

        # if not respuesta:
            # return

        self.crear_panel_actual()
        self.bloquear_seleccion_panel()

        print(
            "Posiciones defectuosas:",
            sorted(
                self.posiciones_defectuosas.keys()
            )
        )

        if total_defectuosas == 0:
            self.panel_actual["estado"] = "RESUMEN"
            self.mostrar_resumen_provisional_panel()
            return

        self.abrir_captura_ids()

    def crear_panel_actual(self):
        """
        Construye en memoria la información del panel confirmado.
        """

        posiciones = sorted(
            self.posiciones_defectuosas.keys()
        )

        pcb_defectuosas = {}

        for posicion in posiciones:
            pcb_defectuosas[posicion] = {
                "posicion": posicion,
                "id_pcb": "",
                "resultado": "FAIL",
                "defectos": {},
                "estado": "PENDIENTE_ID"
            }

        sesion = datetime.now().strftime(
            "SES-%Y%m%d-%H%M%S"
        )

        sesion = (
            f"{sesion}-"
            f"{self.numero_parte_actual}"
        )

        self.panel_actual = {
            "sesion": sesion,
            "modelo": self.modelo_actual,
            "numero_parte": self.numero_parte_actual,
            "renglones": self.renglones_panel,
            "columnas": self.columnas_panel,
            "total_pcb": self.total_pcb_panel,
            "total_buenas": (
                self.total_pcb_panel
                - len(posiciones)
            ),
            "total_defectuosas": len(posiciones),
            "fecha_inicio": datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S"
            ),
            "pcb_defectuosas": pcb_defectuosas,
            "estado": "CAPTURA_IDS"
        }

    def bloquear_seleccion_panel(self):
        """
        Bloquea el modelo y las posiciones después de confirmar.
        """

        self.proceso_panel_activo = True

        self.combo_modelos.configure(
            state="disabled"
        )

        self.btn_confirmar_panel.configure(
            state="disabled"
        )

        for boton in self.botones_pcb.values():
            boton.configure(
                state="disabled"
            )

    def desbloquear_seleccion_panel(self):
        """
        Permite nuevamente seleccionar modelo y posiciones.
        """

        self.proceso_panel_activo = False

        self.combo_modelos.configure(
            state="readonly"
        )

        self.btn_confirmar_panel.configure(
            state="normal"
        )

        for boton in self.botones_pcb.values():
            boton.configure(
                state="normal"
            )

    def mostrar_resumen_provisional_panel(self):
        """
        Muestra temporalmente el resumen de un panel sin defectos.
        """

        if not self.panel_actual:
            return

        total = self.panel_actual["total_pcb"]
        buenas = self.panel_actual["total_buenas"]
        defectuosas = self.panel_actual[
            "total_defectuosas"
        ]

        if total > 0:
            fpy = buenas / total * 100
        else:
            fpy = 0.0

        # respuesta = messagebox.askyesno(
            # "Finalizar panel",
            # (
            # f"Modelo: {self.panel_actual['modelo']}\n"
            # f"Número de parte: "
            # f"{self.panel_actual['numero_parte']}\n\n"
            # f"PCB inspeccionadas: {total}\n"
            # f"PCB buenas: {buenas}\n"
            # f"PCB defectuosas: {defectuosas}\n"
            # f"FPY: {fpy:.2f} %\n\n"
            # "Todas las PCB serán registradas como PASS.\n\n"
            # "¿Desea finalizar y guardar este panel?"
            # ),
            # parent=self.root
        # )

        # if respuesta:
            # self.finalizar_y_guardar_panel()

        self.finalizar_y_guardar_panel()

    @staticmethod
    def validar_escritura_id(valor):
        """
        Permite únicamente números y un máximo de 16 caracteres.
        """

        if valor == "":
            return True

        return valor.isdigit() and len(valor) <= 16

    def abrir_captura_ids(self):
        """
        Abre una ventana modal para capturar los ID de todas las
        PCB defectuosas.
        """

        if not self.panel_actual:
            return

        if (
            self.ventana_captura_ids is not None
            and self.ventana_captura_ids.winfo_exists()
        ):
            self.ventana_captura_ids.lift()
            self.ventana_captura_ids.focus_force()
            return

        posiciones = sorted(
            self.panel_actual[
                "pcb_defectuosas"
            ].keys()
        )

        cantidad_pcb = len(posiciones)

        alto_ventana = min(
            720,
            max(390, 235 + cantidad_pcb * 62)
        )

        ventana = ctk.CTkToplevel(self.root)
        self.ventana_captura_ids = ventana

        ventana.title("Captura de ID de PCB defectuosas")
        ventana.geometry(
            f"720x{alto_ventana}"
        )
        ventana.minsize(650, 390)
        ventana.transient(self.root)
        ventana.grab_set()

        ventana.grid_columnconfigure(
            0,
            weight=1
        )

        ventana.grid_rowconfigure(
            2,
            weight=1
        )

        ventana.protocol(
            "WM_DELETE_WINDOW",
            self.cancelar_captura_ids
        )

        titulo = ctk.CTkLabel(
            ventana,
            text="CAPTURA DE ID DE PCB DEFECTUOSAS",
            font=("Arial", 22, "bold"),
            text_color="#FFFFFF"
        )

        titulo.grid(
            row=0,
            column=0,
            padx=25,
            pady=(20, 5)
        )

        texto_informacion = (
            f"Modelo: {self.panel_actual['modelo']}     "
            f"Número de parte: "
            f"{self.panel_actual['numero_parte']}\n"
            f"PCB defectuosas: {cantidad_pcb}"
        )

        lbl_informacion = ctk.CTkLabel(
            ventana,
            text=texto_informacion,
            font=("Arial", 15, "bold"),
            text_color="#BFC7EE",
            justify="center"
        )

        lbl_informacion.grid(
            row=1,
            column=0,
            padx=25,
            pady=(5, 10)
        )

        frame_lista = ctk.CTkScrollableFrame(
            ventana,
            corner_radius=12,
            fg_color="#252842",
            label_text=(
                "Escanee o capture los 16 dígitos "
                "correspondientes"
            ),
            label_font=("Arial", 14, "bold")
        )

        frame_lista.grid(
            row=2,
            column=0,
            padx=25,
            pady=5,
            sticky="nsew"
        )

        frame_lista.grid_columnconfigure(
            1,
            weight=1
        )

        self.entries_ids_pcb.clear()
        self.labels_estado_ids.clear()

        comando_validacion = (
            ventana.register(
                self.validar_escritura_id
            ),
            "%P"
        )

        for indice, posicion in enumerate(posiciones):

            lbl_posicion = ctk.CTkLabel(
                frame_lista,
                text=f"PCB {posicion}",
                width=90,
                font=("Arial", 16, "bold"),
                text_color="#FFFFFF"
            )

            lbl_posicion.grid(
                row=indice,
                column=0,
                padx=(12, 8),
                pady=8,
                sticky="w"
            )

            variable_id = ctk.StringVar()

            entry_id = ctk.CTkEntry(
                frame_lista,
                textvariable=variable_id,
                height=38,
                font=("Arial", 16),
                justify="center",
                placeholder_text="ID de 16 dígitos",
                validate="key",
                validatecommand=comando_validacion,
                border_width=2,
                border_color="#454B70"
            )

            entry_id.grid(
                row=indice,
                column=1,
                padx=8,
                pady=8,
                sticky="ew"
            )

            lbl_estado = ctk.CTkLabel(
                frame_lista,
                text="Pendiente",
                width=110,
                font=("Arial", 13, "bold"),
                text_color="#AEB4D0"
            )

            lbl_estado.grid(
                row=indice,
                column=2,
                padx=(8, 12),
                pady=8
            )

            self.entries_ids_pcb[posicion] = {
                "variable": variable_id,
                "entry": entry_id
            }

            self.labels_estado_ids[posicion] = (
                lbl_estado
            )

            variable_id.trace_add(
                "write",
                lambda *args, p=posicion: (
                    self.validar_id_en_tiempo_real(p)
                )
            )

            entry_id.bind(
                "<Return>",
                lambda event, p=posicion: (
                    self.enfocar_siguiente_id(p)
                )
            )

        frame_botones = ctk.CTkFrame(
            ventana,
            fg_color="transparent"
        )

        frame_botones.grid(
            row=3,
            column=0,
            padx=25,
            pady=(10, 20),
            sticky="ew"
        )

        frame_botones.grid_columnconfigure(
            0,
            weight=1
        )

        frame_botones.grid_columnconfigure(
            1,
            weight=1
        )

        btn_cancelar = ctk.CTkButton(
            frame_botones,
            text="Cancelar",
            height=42,
            font=("Arial", 15, "bold"),
            fg_color="#5B627E",
            hover_color="#484E66",
            command=self.cancelar_captura_ids
        )

        btn_cancelar.grid(
            row=0,
            column=0,
            padx=(0, 8),
            sticky="ew"
        )

        btn_continuar = ctk.CTkButton(
            frame_botones,
            text="Validar y continuar",
            height=42,
            font=("Arial", 15, "bold"),
            command=self.confirmar_ids_pcb
        )

        btn_continuar.grid(
            row=0,
            column=1,
            padx=(8, 0),
            sticky="ew"
        )

        ventana.after(
            100,
            self.enfocar_primer_id
        )

    def validar_id_en_tiempo_real(self, posicion):
        """
        Cambia el estado visual del ID conforme se captura.
        """

        datos_entry = self.entries_ids_pcb.get(
            posicion
        )

        if not datos_entry:
            return

        valor = datos_entry[
            "variable"
        ].get().strip()

        entry = datos_entry["entry"]
        lbl_estado = self.labels_estado_ids[
            posicion
        ]

        if not valor:
            entry.configure(
                border_color="#454B70"
            )

            lbl_estado.configure(
                text="Pendiente",
                text_color="#AEB4D0"
            )

            return

        if len(valor) < 16:
            entry.configure(
                border_color="#D97706"
            )

            lbl_estado.configure(
                text=f"{len(valor)}/16",
                text_color="#FBBF24"
            )

            return

        if not valor.startswith(
            self.numero_parte_actual
        ):
            entry.configure(
                border_color="#DC4C64"
            )

            lbl_estado.configure(
                text="Modelo incorrecto",
                text_color="#FF6B81"
            )

            return

        entry.configure(
            border_color="#2EB872"
        )

        lbl_estado.configure(
            text="Válido",
            text_color="#6FE3A1"
        )

    def enfocar_primer_id(self):
        """
        Coloca el cursor en el primer ID pendiente.
        """

        if not self.entries_ids_pcb:
            return

        primera_posicion = sorted(
            self.entries_ids_pcb.keys()
        )[0]

        entry = self.entries_ids_pcb[
            primera_posicion
        ]["entry"]

        entry.focus_set()

    def enfocar_siguiente_id(self, posicion_actual):
        """
        Al presionar Enter avanza al siguiente ID.
        """

        posiciones = sorted(
            self.entries_ids_pcb.keys()
        )

        try:
            indice_actual = posiciones.index(
                posicion_actual
            )
        except ValueError:
            return

        siguiente_indice = indice_actual + 1

        if siguiente_indice < len(posiciones):
            siguiente_posicion = posiciones[
                siguiente_indice
            ]

            siguiente_entry = self.entries_ids_pcb[
                siguiente_posicion
            ]["entry"]

            siguiente_entry.focus_set()
            siguiente_entry.select_range(
                0,
                "end"
            )

        else:
            self.confirmar_ids_pcb()

    def confirmar_ids_pcb(self):
        """
        Valida todos los ID y los guarda en panel_actual.
        """

        errores = []
        ids_capturados = {}
        posiciones_por_id = {}

        for posicion in sorted(
            self.entries_ids_pcb.keys()
        ):
            datos_entry = self.entries_ids_pcb[
                posicion
            ]

            valor = datos_entry[
                "variable"
            ].get().strip()

            entry = datos_entry["entry"]
            lbl_estado = self.labels_estado_ids[
                posicion
            ]

            if len(valor) != 16:
                errores.append(
                    f"PCB {posicion}: el ID debe "
                    "contener 16 dígitos."
                )

                entry.configure(
                    border_color="#DC4C64"
                )

                lbl_estado.configure(
                    text="Longitud incorrecta",
                    text_color="#FF6B81"
                )

                continue

            if not valor.isdigit():
                errores.append(
                    f"PCB {posicion}: el ID debe "
                    "contener solamente números."
                )

                entry.configure(
                    border_color="#DC4C64"
                )

                lbl_estado.configure(
                    text="ID incorrecto",
                    text_color="#FF6B81"
                )

                continue

            if not valor.startswith(
                self.numero_parte_actual
            ):
                errores.append(
                    f"PCB {posicion}: el ID no pertenece "
                    f"al número de parte "
                    f"{self.numero_parte_actual}."
                )

                entry.configure(
                    border_color="#DC4C64"
                )

                lbl_estado.configure(
                    text="Modelo incorrecto",
                    text_color="#FF6B81"
                )

                continue

            ids_capturados[posicion] = valor

            if valor not in posiciones_por_id:
                posiciones_por_id[valor] = []

            posiciones_por_id[valor].append(
                posicion
            )

        ids_repetidos = {
            identificador: posiciones
            for identificador, posiciones
            in posiciones_por_id.items()
            if len(posiciones) > 1
        }

        for identificador, posiciones in (
            ids_repetidos.items()
        ):
            texto_posiciones = ", ".join(
                f"PCB {posicion}"
                for posicion in posiciones
            )

            errores.append(
                f"ID repetido {identificador}: "
                f"{texto_posiciones}."
            )

            for posicion in posiciones:
                entry = self.entries_ids_pcb[
                    posicion
                ]["entry"]

                lbl_estado = self.labels_estado_ids[
                    posicion
                ]

                entry.configure(
                    border_color="#DC4C64"
                )

                lbl_estado.configure(
                    text="ID repetido",
                    text_color="#FF6B81"
                )

        if errores:
            texto_errores = "\n".join(
                f"• {error}"
                for error in errores[:10]
            )

            if len(errores) > 10:
                texto_errores += (
                    "\n• Existen más errores por corregir."
                )

            messagebox.showwarning(
                "ID incorrectos",
                (
                    "Corrija los siguientes datos:\n\n"
                    f"{texto_errores}"
                ),
                parent=self.ventana_captura_ids
            )

            self.enfocar_primer_id_invalido()
            return

        for posicion, identificador in (
            ids_capturados.items()
        ):
            datos_pcb = self.panel_actual[
                "pcb_defectuosas"
            ][posicion]

            datos_pcb["id_pcb"] = identificador
            datos_pcb["estado"] = (
                "PENDIENTE_DEFECTOS"
            )

        self.panel_actual["estado"] = (
            "REGISTRO_DEFECTOS"
        )

        print("\nID validados:")

        for posicion in sorted(
            ids_capturados.keys()
        ):
            print(
                f"PCB {posicion}: "
                f"{ids_capturados[posicion]}"
            )

        self.cerrar_ventana_captura_ids()

        self.iniciar_registro_defectos_panel()

    def enfocar_primer_id_invalido(self):
        """
        Coloca el cursor en el primer ID que no sea válido.
        """

        for posicion in sorted(
            self.entries_ids_pcb.keys()
        ):
            datos_entry = self.entries_ids_pcb[
                posicion
            ]

            valor = datos_entry[
                "variable"
            ].get().strip()

            if (
                len(valor) != 16
                or not valor.isdigit()
                or not valor.startswith(
                    self.numero_parte_actual
                )
            ):
                entry = datos_entry["entry"]

                entry.focus_set()
                entry.select_range(
                    0,
                    "end"
                )
                return

    def cancelar_captura_ids(self):
        """
        Cancela la captura y regresa a la selección del panel.
        """

        respuesta = messagebox.askyesno(
            "Cancelar captura",
            (
                "¿Desea cancelar la captura de ID?\n\n"
                "La selección de PCB defectuosas podrá "
                "modificarse nuevamente."
            ),
            parent=self.ventana_captura_ids
        )

        if not respuesta:
            return

        self.panel_actual = None

        self.cerrar_ventana_captura_ids()
        self.desbloquear_seleccion_panel()

    def cerrar_ventana_captura_ids(self):
        """
        Cierra y limpia las referencias de la ventana de ID.
        """

        if (
            self.ventana_captura_ids is not None
            and self.ventana_captura_ids.winfo_exists()
        ):
            try:
                self.ventana_captura_ids.grab_release()
            except tk.TclError:
                pass

            self.ventana_captura_ids.destroy()

        self.ventana_captura_ids = None
        self.entries_ids_pcb.clear()
        self.labels_estado_ids.clear()

    def iniciar_registro_defectos_panel(self):
        """
        Prepara el orden de las PCB defectuosas y abre la ventana
        de registro de defectos.
        """

        if not self.panel_actual:
            return

        self.posiciones_pendientes_defectos = sorted(
            self.panel_actual["pcb_defectuosas"].keys()
        )

        self.indice_pcb_defecto_actual = 0
        self.abrir_ventana_registro_defectos()

    def abrir_ventana_registro_defectos(self):
        """
        Abre una ventana modal para registrar los defectos
        de cada PCB.
        """

        if not self.panel_actual:
            return

        if not self.posiciones_pendientes_defectos:
            return

        if (
            self.ventana_registro_defectos is not None
            and self.ventana_registro_defectos.winfo_exists()
        ):
            self.ventana_registro_defectos.lift()
            self.ventana_registro_defectos.focus_force()
            return

        ventana = ctk.CTkToplevel(self.root)
        self.ventana_registro_defectos = ventana

        ventana.title("Registro de defectos por PCB")
        ventana.geometry("820x690")
        ventana.minsize(760, 620)
        ventana.transient(self.root)
        ventana.grab_set()

        ventana.grid_columnconfigure(0, weight=1)
        ventana.grid_rowconfigure(3, weight=1)

        ventana.protocol(
            "WM_DELETE_WINDOW",
            self.cancelar_registro_defectos_panel
        )

        self.lbl_titulo_registro_pcb = ctk.CTkLabel(
            ventana,
            text="REGISTRO DE DEFECTOS",
            font=("Arial", 23, "bold"),
            text_color="#FFFFFF"
        )

        self.lbl_titulo_registro_pcb.grid(
            row=0,
            column=0,
            padx=25,
            pady=(20, 5)
        )

        self.lbl_info_registro_pcb = ctk.CTkLabel(
            ventana,
            text="",
            font=("Arial", 16, "bold"),
            text_color="#BFC7EE",
            justify="center"
        )

        self.lbl_info_registro_pcb.grid(
            row=1,
            column=0,
            padx=25,
            pady=(5, 8)
        )

        self.barra_progreso_defectos = ctk.CTkProgressBar(
            ventana,
            height=14,
            corner_radius=7
        )

        self.barra_progreso_defectos.grid(
            row=2,
            column=0,
            padx=35,
            pady=(5, 10),
            sticky="ew"
        )

        frame_contenido = ctk.CTkFrame(
            ventana,
            corner_radius=12,
            fg_color="#252842"
        )

        frame_contenido.grid(
            row=3,
            column=0,
            padx=25,
            pady=5,
            sticky="nsew"
        )

        frame_contenido.grid_columnconfigure(0, weight=1)
        frame_contenido.grid_rowconfigure(2, weight=1)

        self.lbl_instruccion_defectos = ctk.CTkLabel(
            frame_contenido,
            text=(
                "Seleccione un defecto, indique la cantidad "
                "y agréguelo a la PCB."
            ),
            font=("Arial", 14),
            text_color="#DDE2FF"
        )

        self.lbl_instruccion_defectos.grid(
            row=0,
            column=0,
            padx=20,
            pady=(15, 8)
        )

        frame_captura = ctk.CTkFrame(
            frame_contenido,
            fg_color="transparent"
        )

        frame_captura.grid(
            row=1,
            column=0,
            padx=20,
            pady=5,
            sticky="ew"
        )

        frame_captura.grid_columnconfigure(0, weight=3)
        frame_captura.grid_columnconfigure(1, weight=1)

        self.defecto_pcb_seleccionado = ctk.StringVar(
            value="Seleccione un defecto"
        )

        self.combo_defectos_pcb = ctk.CTkComboBox(
            frame_captura,
            variable=self.defecto_pcb_seleccionado,
            values=self.obtener_lista_defectos_captura(),
            state="readonly",
            height=38,
            font=("Arial", 15),
            dropdown_font=("Arial", 14),
            command=self.cambiar_defecto_pcb
        )

        self.combo_defectos_pcb.grid(
            row=0,
            column=0,
            padx=(0, 8),
            sticky="ew"
        )

        self.cantidad_defecto_pcb = ctk.StringVar(value="1")

        self.entry_cantidad_defecto_pcb = ctk.CTkEntry(
            frame_captura,
            textvariable=self.cantidad_defecto_pcb,
            height=38,
            justify="center",
            font=("Arial", 15)
        )

        self.entry_cantidad_defecto_pcb.grid(
            row=0,
            column=1,
            padx=8,
            sticky="ew"
        )

        self.btn_agregar_defecto_pcb = ctk.CTkButton(
            frame_captura,
            text="Agregar defecto",
            height=38,
            font=("Arial", 14, "bold"),
            command=self.agregar_defecto_pcb_actual
        )

        self.btn_agregar_defecto_pcb.grid(
            row=0,
            column=2,
            padx=(8, 0)
        )

        self.frame_descripcion_otro = ctk.CTkFrame(
            frame_captura,
            fg_color="transparent"
        )

        self.frame_descripcion_otro.grid(
            row=1,
            column=0,
            columnspan=3,
            padx=0,
            pady=(10, 0),
            sticky="ew"
        )

        self.frame_descripcion_otro.grid_columnconfigure(
            1,
            weight=1
        )

        self.lbl_descripcion_otro = ctk.CTkLabel(
            self.frame_descripcion_otro,
            text="Descripción de otro defecto:",
            font=("Arial", 14, "bold"),
            text_color="#DDE2FF"
        )

        self.lbl_descripcion_otro.grid(
            row=0,
            column=0,
            padx=(0, 10),
            sticky="w"
        )

        self.entry_descripcion_otro = ctk.CTkEntry(
            self.frame_descripcion_otro,
            textvariable=self.descripcion_otro_pcb,
            height=38,
            font=("Arial", 15),
            placeholder_text="Escriba una descripción clara del defecto"
        )

        self.entry_descripcion_otro.grid(
            row=0,
            column=1,
            sticky="ew"
        )

        # Oculto hasta seleccionar Otro
        self.frame_descripcion_otro.grid_remove()

        self.frame_lista_defectos_pcb = ctk.CTkScrollableFrame(
            frame_contenido,
            corner_radius=10,
            fg_color="#1F2238",
            label_text="Defectos registrados en esta PCB",
            label_font=("Arial", 14, "bold")
        )

        self.frame_lista_defectos_pcb.grid(
            row=2,
            column=0,
            padx=20,
            pady=10,
            sticky="nsew"
        )

        self.frame_lista_defectos_pcb.grid_columnconfigure(
            0,
            weight=1
        )

        frame_botones = ctk.CTkFrame(
            ventana,
            fg_color="transparent"
        )

        frame_botones.grid(
            row=4,
            column=0,
            padx=25,
            pady=(10, 20),
            sticky="ew"
        )

        frame_botones.grid_columnconfigure(0, weight=1)
        frame_botones.grid_columnconfigure(1, weight=1)

        self.btn_cancelar_registro_defectos = ctk.CTkButton(
            frame_botones,
            text="Cancelar inspección",
            height=42,
            font=("Arial", 15, "bold"),
            fg_color="#5B627E",
            hover_color="#484E66",
            command=self.cancelar_registro_defectos_panel
        )

        self.btn_cancelar_registro_defectos.grid(
            row=0,
            column=0,
            padx=(0, 8),
            sticky="ew"
        )

        self.btn_guardar_siguiente_pcb = ctk.CTkButton(
            frame_botones,
            text="Guardar y siguiente",
            height=42,
            font=("Arial", 15, "bold"),
            command=self.guardar_defectos_pcb_actual
        )

        self.btn_guardar_siguiente_pcb.grid(
            row=0,
            column=1,
            padx=(8, 0),
            sticky="ew"
        )

        self.cargar_pcb_actual_en_ventana()

    def cargar_pcb_actual_en_ventana(self):
        """
        Muestra la PCB correspondiente al índice actual.
        """

        total = len(self.posiciones_pendientes_defectos)

        if total == 0:
            return

        if self.indice_pcb_defecto_actual >= total:
            self.finalizar_captura_defectos_panel()
            return

        posicion = self.posiciones_pendientes_defectos[
            self.indice_pcb_defecto_actual
        ]

        datos_pcb = self.panel_actual[
            "pcb_defectuosas"
        ][posicion]

        numero_actual = self.indice_pcb_defecto_actual + 1

        self.lbl_info_registro_pcb.configure(
            text=(
                f"PCB {numero_actual} de {total}\n"
                f"Posición: PCB {posicion}     "
                f"ID: {datos_pcb['id_pcb']}"
            )
        )

        progreso = numero_actual / total
        self.barra_progreso_defectos.set(progreso)

        self.defectos_pcb_actual = dict(
            datos_pcb.get("defectos", {})
        )

        self.defecto_pcb_seleccionado.set(
            "Seleccione un defecto"
        )

        self.cantidad_defecto_pcb.set("1")
        self.descripcion_otro_pcb.set("")
        self.frame_descripcion_otro.grid_remove()

        self.actualizar_lista_defectos_pcb()

        if numero_actual == total:
            self.btn_guardar_siguiente_pcb.configure(
                text="Guardar y finalizar"
            )
        else:
            self.btn_guardar_siguiente_pcb.configure(
                text="Guardar y siguiente"
            )

    def agregar_defecto_pcb_actual(self):
        """
        Agrega o acumula un defecto en la PCB actual.

        Cuando se selecciona Otro, la descripción forma parte de la
        identificación interna del defecto.
        """

        defecto = self.defecto_pcb_seleccionado.get().strip()
        cantidad_texto = self.cantidad_defecto_pcb.get().strip()

        if (
            not defecto
            or defecto == "Seleccione un defecto"
        ):
            messagebox.showwarning(
                "Defecto requerido",
                "Seleccione un defecto.",
                parent=self.ventana_registro_defectos
            )
            return

        try:
            cantidad = int(cantidad_texto)
        except ValueError:
            messagebox.showwarning(
                "Cantidad incorrecta",
                "La cantidad debe ser un número entero.",
                parent=self.ventana_registro_defectos
            )
            return

        if cantidad <= 0:
            messagebox.showwarning(
                "Cantidad incorrecta",
                "La cantidad debe ser mayor que cero.",
                parent=self.ventana_registro_defectos
            )
            return

        # -----------------------------------------------------
        # DEFECTO PERSONALIZADO
        # -----------------------------------------------------
        if defecto == self.opcion_otro:
            descripcion = (
                self.descripcion_otro_pcb.get()
                .strip()
            )

            if not descripcion:
                messagebox.showwarning(
                    "Descripción requerida",
                    (
                        "Escriba una descripción para el defecto "
                        'seleccionado como "Otro".'
                    ),
                    parent=self.ventana_registro_defectos
                )

                self.entry_descripcion_otro.focus_set()
                return

            if len(descripcion) < 3:
                messagebox.showwarning(
                    "Descripción incorrecta",
                    (
                        "La descripción debe contener al menos "
                        "3 caracteres."
                    ),
                    parent=self.ventana_registro_defectos
                )
                return

            if len(descripcion) > 100:
                messagebox.showwarning(
                    "Descripción demasiado larga",
                    (
                        "La descripción no debe superar "
                        "los 100 caracteres."
                    ),
                    parent=self.ventana_registro_defectos
                )
                return

            # La descripción queda diferenciada en memoria
            clave_defecto = f"Otro: {descripcion}"

        else:
            clave_defecto = defecto

        cantidad_actual = self.defectos_pcb_actual.get(
            clave_defecto,
            0
        )

        self.defectos_pcb_actual[clave_defecto] = (
            cantidad_actual + cantidad
        )

        self.defecto_pcb_seleccionado.set(
            "Seleccione un defecto"
        )

        self.cantidad_defecto_pcb.set("1")
        self.descripcion_otro_pcb.set("")
        self.frame_descripcion_otro.grid_remove()

        self.actualizar_lista_defectos_pcb()

    def actualizar_lista_defectos_pcb(self):
        """
        Actualiza visualmente los defectos agregados a la PCB.
        """

        for widget in (
            self.frame_lista_defectos_pcb.winfo_children()
        ):
            widget.destroy()

        self.filas_defectos_pcb.clear()

        if not self.defectos_pcb_actual:
            lbl_vacio = ctk.CTkLabel(
                self.frame_lista_defectos_pcb,
                text="Aún no se han agregado defectos.",
                font=("Arial", 14),
                text_color="#AEB4D0"
            )

            lbl_vacio.grid(
                row=0,
                column=0,
                padx=15,
                pady=20
            )
            return

        for fila, (defecto, cantidad) in enumerate(
            self.defectos_pcb_actual.items()
        ):
            frame_fila = ctk.CTkFrame(
                self.frame_lista_defectos_pcb,
                fg_color="#30344F",
                corner_radius=8
            )

            frame_fila.grid(
                row=fila,
                column=0,
                padx=8,
                pady=5,
                sticky="ew"
            )

            frame_fila.grid_columnconfigure(0, weight=1)

            lbl_defecto = ctk.CTkLabel(
                frame_fila,
                text=defecto,
                font=("Arial", 14, "bold"),
                text_color="#FFFFFF",
                anchor="w"
            )

            lbl_defecto.grid(
                row=0,
                column=0,
                padx=12,
                pady=10,
                sticky="ew"
            )

            lbl_cantidad = ctk.CTkLabel(
                frame_fila,
                text=f"Cantidad: {cantidad}",
                width=115,
                font=("Arial", 14, "bold"),
                text_color="#6FE3A1"
            )

            lbl_cantidad.grid(
                row=0,
                column=1,
                padx=8,
                pady=10
            )

            btn_restar = ctk.CTkButton(
                frame_fila,
                text="−",
                width=38,
                height=30,
                font=("Arial", 18, "bold"),
                fg_color="#D97706",
                hover_color="#B45309",
                command=lambda d=defecto: (
                    self.restar_defecto_pcb(d)
                )
            )

            btn_restar.grid(
                row=0,
                column=2,
                padx=4
            )

            btn_eliminar = ctk.CTkButton(
                frame_fila,
                text="Eliminar",
                width=80,
                height=30,
                font=("Arial", 13, "bold"),
                fg_color="#C24155",
                hover_color="#9F3345",
                command=lambda d=defecto: (
                    self.eliminar_defecto_pcb(d)
                )
            )

            btn_eliminar.grid(
                row=0,
                column=3,
                padx=(4, 10)
            )

            self.filas_defectos_pcb[defecto] = (
                frame_fila
            )

    def restar_defecto_pcb(self, defecto):
        """
        Resta una unidad al defecto indicado.
        """

        if defecto not in self.defectos_pcb_actual:
            return

        nueva_cantidad = (
            self.defectos_pcb_actual[defecto] - 1
        )

        if nueva_cantidad <= 0:
            self.defectos_pcb_actual.pop(
                defecto,
                None
            )
        else:
            self.defectos_pcb_actual[defecto] = (
                nueva_cantidad
            )

        self.actualizar_lista_defectos_pcb()

    def eliminar_defecto_pcb(self, defecto):
        """
        Elimina completamente un defecto de la PCB actual.
        """

        self.defectos_pcb_actual.pop(
            defecto,
            None
        )

        self.actualizar_lista_defectos_pcb()

    def guardar_defectos_pcb_actual(self):
        """
        Guarda en memoria los defectos de la PCB actual.
        """

        if not self.defectos_pcb_actual:
            messagebox.showwarning(
                "Defecto requerido",
                (
                    "La PCB fue seleccionada como defectuosa.\n\n"
                    "Debe registrar al menos un defecto."
                ),
                parent=self.ventana_registro_defectos
            )
            return

        posicion = self.posiciones_pendientes_defectos[
            self.indice_pcb_defecto_actual
        ]

        datos_pcb = self.panel_actual[
            "pcb_defectuosas"
        ][posicion]

        datos_pcb["defectos"] = dict(
            self.defectos_pcb_actual
        )

        datos_pcb["cantidad_defectos"] = sum(
            self.defectos_pcb_actual.values()
        )

        datos_pcb["estado"] = "COMPLETADA"

        if posicion in self.botones_pcb:
            self.botones_pcb[posicion].configure(
                text=(
                    f"PCB {posicion}\n"
                    f"{datos_pcb['cantidad_defectos']} DEF."
                ),
                fg_color="#C24155",
                hover_color="#9F3345"
            )

        self.indice_pcb_defecto_actual += 1

        if (
            self.indice_pcb_defecto_actual
            >= len(self.posiciones_pendientes_defectos)
        ):
            self.finalizar_captura_defectos_panel()
            return

        self.cargar_pcb_actual_en_ventana()

    def finalizar_captura_defectos_panel(self):
        """
        Finaliza la captura de defectos y muestra un resumen
        provisional.
        """

        self.panel_actual["estado"] = "RESUMEN"

        total_defectos = 0

        for datos_pcb in self.panel_actual[
            "pcb_defectuosas"
        ].values():
            total_defectos += datos_pcb.get(
                "cantidad_defectos",
                0
            )

        self.panel_actual[
            "total_defectos"
        ] = total_defectos

        self.cerrar_ventana_registro_defectos()

        self.mostrar_resumen_panel_con_defectos()

    def mostrar_resumen_panel_con_defectos(self):
        """
        Muestra el resumen y solicita confirmar el guardado.
        """

        if not self.panel_actual:
            return

        total = self.panel_actual["total_pcb"]
        buenas = self.panel_actual["total_buenas"]

        defectuosas = self.panel_actual[
            "total_defectuosas"
        ]

        total_defectos = self.panel_actual.get(
            "total_defectos",
            0
        )

        fpy = (
            buenas / total * 100
            if total > 0
            else 0.0
        )

        detalle_pcb = []

        for posicion, datos_pcb in sorted(
            self.panel_actual[
                "pcb_defectuosas"
            ].items()
        ):
            detalle_pcb.append(
                (
                    f"PCB {posicion} | "
                    f"{datos_pcb['id_pcb']} | "
                    f"{datos_pcb.get('cantidad_defectos', 0)} "
                    "defecto(s)"
                )
            )

        texto_detalle = "\n".join(
            detalle_pcb
        )

        # respuesta = messagebox.askyesno(
        # "Finalizar panel",
        # (
        # f"Modelo: {self.panel_actual['modelo']}\n"
        # f"Número de parte: "
        # f"{self.panel_actual['numero_parte']}\n\n"
        # f"PCB inspeccionadas: {total}\n"
        # f"PCB buenas: {buenas}\n"
        # f"PCB defectuosas: {defectuosas}\n"
        # f"Defectos encontrados: {total_defectos}\n"
        # f"FPY: {fpy:.2f} %\n\n"
        # f"{texto_detalle}\n\n"
        # "¿Desea finalizar y guardar este panel?"
        # ),
        # parent=self.root
        # )

        # if not respuesta:
        # return

        self.finalizar_y_guardar_panel()

    def cerrar_ventana_registro_defectos(self):
        """
        Cierra la ventana de registro de defectos.
        """

        if (
            self.ventana_registro_defectos is not None
            and self.ventana_registro_defectos.winfo_exists()
        ):
            try:
                self.ventana_registro_defectos.grab_release()
            except tk.TclError:
                pass

            self.ventana_registro_defectos.destroy()

        self.ventana_registro_defectos = None
        self.defectos_pcb_actual.clear()
        self.filas_defectos_pcb.clear()

    def cancelar_registro_defectos_panel(self):
        """
        Cancela completamente la inspección actual.
        """

        respuesta = messagebox.askyesno(
            "Cancelar inspección",
            (
                "¿Desea cancelar toda la inspección actual?\n\n"
                "Los ID y defectos capturados se perderán."
            ),
            parent=self.ventana_registro_defectos
        )

        if not respuesta:
            return

        self.cerrar_ventana_registro_defectos()

        self.panel_actual = None
        self.posiciones_pendientes_defectos.clear()
        self.indice_pcb_defecto_actual = 0

        self.desbloquear_seleccion_panel()

    def obtener_encabezados_log_pcb(self):
        """
        Retorna los encabezados del archivo LogFilePCB.csv.
        """

        columnas_fijas = [
            "Sesion",
            "Modelo",
            "NumeroParte",
            "Posicion",
            "Renglon",
            "Columna",
            "ID_PCB",
            "Resultado",
            "Defectos",
            "DescripcionOtro",
            "Fecha/Hora"
        ]

        columnas_defectos = list(self.lista_defectos)

        if self.opcion_otro not in columnas_defectos:
            columnas_defectos.append(self.opcion_otro)

        return columnas_fijas + columnas_defectos

    def obtener_coordenadas_posicion(self, posicion):
        """
        Convierte el número de posición en renglón y columna.

        La numeración comienza en 1.
        """

        if self.columnas_panel <= 0:
            return 0, 0

        renglon = (
            (posicion - 1)
            // self.columnas_panel
        ) + 1

        columna = (
            (posicion - 1)
            % self.columnas_panel
        ) + 1

        return renglon, columna

    def asegurar_encabezados_log_pcb(self):
        """
        Crea LogFilePCB.csv o agrega nuevas columnas de defectos
        conservando los registros existentes.
        """

        encabezados_requeridos = (
            self.obtener_encabezados_log_pcb()
        )

        if not os.path.exists(self.archivo_log_pcb):
            return encabezados_requeridos

        if os.path.getsize(self.archivo_log_pcb) == 0:
            return encabezados_requeridos

        with open(
            self.archivo_log_pcb,
            mode="r",
            newline="",
            encoding="utf-8-sig"
        ) as archivo:

            lector = csv.DictReader(archivo)
            encabezados_actuales = lector.fieldnames or []
            registros = list(lector)

        columnas_fijas = [
            "Sesion",
            "Modelo",
            "NumeroParte",
            "Posicion",
            "Renglon",
            "Columna",
            "ID_PCB",
            "Resultado",
            "Defectos",
            "DescripcionOtro",
            "Fecha/Hora"
        ]

        defectos_actuales = [
            encabezado
            for encabezado in encabezados_actuales
            if encabezado not in columnas_fijas
        ]

        defectos_finales = defectos_actuales.copy()

        for defecto in self.lista_defectos:
            if defecto not in defectos_finales:
                defectos_finales.append(defecto)

        encabezados_finales = (
            columnas_fijas
            + defectos_finales
        )

        if encabezados_actuales == encabezados_finales:
            return encabezados_finales

        archivo_temporal = (
            self.archivo_log_pcb
            + ".tmp"
        )

        with open(
            archivo_temporal,
            mode="w",
            newline="",
            encoding="utf-8-sig"
        ) as archivo:

            escritor = csv.DictWriter(
                archivo,
                fieldnames=encabezados_finales
            )

            escritor.writeheader()

            for registro in registros:
                fila = {}

                for encabezado in encabezados_finales:
                    valor = registro.get(
                        encabezado,
                        ""
                    )

                    if (
                        encabezado not in columnas_fijas
                        and valor == ""
                    ):
                        valor = 0

                    fila[encabezado] = valor

                escritor.writerow(fila)

        os.replace(
            archivo_temporal,
            self.archivo_log_pcb
        )

        return encabezados_finales

    def construir_filas_panel_actual(self):
        """
        Crea una fila por cada posición del panel.
        """

        if not self.panel_actual:
            return []

        filas = []

        fecha_hora = datetime.now().strftime(
            "%d/%m/%Y %H:%M:%S"
        )

        pcb_defectuosas = self.panel_actual[
            "pcb_defectuosas"
        ]

        for posicion in range(
            1,
            self.panel_actual["total_pcb"] + 1
        ):
            renglon, columna = (
                self.obtener_coordenadas_posicion(
                    posicion
                )
            )

            columnas_defectos = list(
                self.lista_defectos
            )

            if self.opcion_otro not in columnas_defectos:
                columnas_defectos.append(
                    self.opcion_otro
                )

            cantidades_defectos = {
                defecto: 0
                for defecto in columnas_defectos
            }

            descripciones_otro = []

            if posicion in pcb_defectuosas:
                datos_pcb = pcb_defectuosas[
                    posicion
                ]

                resultado = "FAIL"
                id_pcb = datos_pcb["id_pcb"]

                defectos_pcb = datos_pcb.get(
                    "defectos",
                    {}
                )

                for defecto, cantidad in defectos_pcb.items():

                    if defecto.startswith("Otro: "):
                        descripcion = defecto[
                            len("Otro: "):
                        ].strip()

                        cantidades_defectos[
                            self.opcion_otro
                        ] += cantidad

                        descripciones_otro.append(
                            f"{descripcion} ({cantidad})"
                        )

                    else:
                        cantidades_defectos[
                            defecto
                        ] = cantidad

                total_defectos = sum(
                    defectos_pcb.values()
                )

            else:
                resultado = "PASS"
                id_pcb = ""
                total_defectos = 0
                descripciones_otro = []

            fila = {
                "Sesion": self.panel_actual[
                    "sesion"
                ],
                "Modelo": self.panel_actual[
                    "modelo"
                ],
                "NumeroParte": self.panel_actual[
                    "numero_parte"
                ],
                "Posicion": f"PCB {posicion}",
                "Renglon": renglon,
                "Columna": columna,
                "ID_PCB": id_pcb,
                "Resultado": resultado,
                "Defectos": total_defectos,
                "DescripcionOtro": " | ".join(
                    descripciones_otro
                ),
                "Fecha/Hora": fecha_hora
            }

            fila.update(
                cantidades_defectos
            )

            filas.append(fila)

        return filas

    def guardar_panel_actual_csv(self):
        """
        Guarda todas las posiciones del panel en una sola operación.
        """
        if not self.panel_actual:
            return False

        sesion = self.panel_actual["sesion"]

        if self.sesion_ya_registrada(sesion):
            messagebox.showerror(
                "Panel duplicado",
                (
                    "Esta sesión ya fue registrada.\n\n"
                    f"Sesión: {sesion}"
                )
            )
            return False

        if self.guardando_panel:
            return False

        self.guardando_panel = True

        try:
            encabezados = (
                self.asegurar_encabezados_log_pcb()
            )
            print("Encabezados para guardar:")
            print(encabezados)

            filas = (
                self.construir_filas_panel_actual()
            )

            if not filas:
                messagebox.showerror(
                    "Error de guardado",
                    "No existen registros para guardar."
                )
                return False

            archivo_existe = (
                os.path.exists(self.archivo_log_pcb)
                and os.path.getsize(
                    self.archivo_log_pcb
                ) > 0
            )

            with open(
                self.archivo_log_pcb,
                mode="a",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:

                escritor = csv.DictWriter(
                    archivo,
                    fieldnames=encabezados
                )

                if not archivo_existe:
                    escritor.writeheader()

                escritor.writerows(filas)

            return True

        except PermissionError:
            messagebox.showerror(
                "Archivo en uso",
                (
                    "No fue posible guardar el panel.\n\n"
                    "Cierre LogFilePCB.csv si está abierto "
                    "en Excel u otro programa."
                )
            )

            return False

        except OSError as error:
            messagebox.showerror(
                "Error de guardado",
                (
                    "No fue posible guardar el panel.\n\n"
                    f"{error}"
                )
            )

            return False

        finally:
            self.guardando_panel = False

    def sesion_ya_registrada(self, sesion):
        """
        Verifica si una sesión ya existe en LogFilePCB.csv.
        """

        if not os.path.exists(self.archivo_log_pcb):
            return False

        if os.path.getsize(self.archivo_log_pcb) == 0:
            return False

        try:
            with open(
                self.archivo_log_pcb,
                mode="r",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:

                lector = csv.DictReader(archivo)

                for fila in lector:
                    if fila.get("Sesion", "") == sesion:
                        return True

        except OSError:
            return False

        return False

    def finalizar_y_guardar_panel(self):
        """
        Guarda el panel, actualiza la interfaz y prepara
        una nueva inspección.
        """

        guardado = self.guardar_panel_actual_csv()

        if not guardado:
            return

        sesion = self.panel_actual["sesion"]
        total = self.panel_actual["total_pcb"]
        defectuosas = self.panel_actual[
            "total_defectuosas"
        ]

        # messagebox.showinfo(
        # "Panel guardado",
        # (
        # "El panel se guardó correctamente.\n\n"
        # f"Sesión: {sesion}\n"
        # f"PCB registradas: {total}\n"
        # f"PCB defectuosas: {defectuosas}"
        # ),
        # parent=self.root
        # )

        self.reiniciar_inspeccion_panel()

    def reiniciar_inspeccion_panel(self):
        """
        Limpia la inspección terminada y deja listo el mismo modelo
        para iniciar otro panel.
        """

        self.panel_actual = None
        self.proceso_panel_activo = False

        self.posiciones_defectuosas.clear()
        self.posiciones_pendientes_defectos.clear()
        self.defectos_pcb_actual.clear()

        self.indice_pcb_defecto_actual = 0

        self.combo_modelos.configure(
            state="readonly"
        )

        self.restablecer_seleccion_panel()

        self.solicitar_actualizacion_dashboard()

    def restablecer_seleccion_panel(self):
        """
        Regresa la interfaz al estado inicial sin un panel seleccionado.
        """
        self.frame_panel_pcb.grid_remove()

        self.modelo_actual = None
        self.numero_parte_actual = ""
        self.renglones_panel = 0
        self.columnas_panel = 0
        self.total_pcb_panel = 0

        self.modelo_seleccionado.set(
            self.opcion_seleccionar_panel
        )

        self.posiciones_defectuosas.clear()
        self.botones_pcb.clear()

        if self.frame_panel_pcb is not None:
            for widget in self.frame_panel_pcb.winfo_children():
                widget.destroy()

        if self.lbl_info_modelo is not None:
            self.lbl_info_modelo.configure(
                text=(
                    "Seleccione un panel para iniciar la inspección"
                )
            )

        if self.btn_confirmar_panel is not None:
            self.btn_confirmar_panel.configure(
                state="disabled"
            )

    def obtener_lista_defectos_captura(self):
        """
        Retorna los defectos configurados y agrega la opción especial Otro.
        """

        defectos = list(self.lista_defectos)

        if self.opcion_otro not in defectos:
            defectos.append(self.opcion_otro)

        return defectos

    def cambiar_defecto_pcb(self, defecto):
        """
        Muestra el campo de descripción cuando se selecciona Otro.
        """

        defecto = defecto.strip()

        if defecto == self.opcion_otro:
            self.frame_descripcion_otro.grid()

            self.ventana_registro_defectos.after(
                80,
                self.entry_descripcion_otro.focus_set
            )
        else:
            self.frame_descripcion_otro.grid_remove()
            self.descripcion_otro_pcb.set("")

    def abrir_ventana_analisis_defectos(self):
        """
        Abre el dashboard independiente para analizar todos los defectos.
        """

        if (
            self.ventana_analisis_defectos is not None
            and self.ventana_analisis_defectos.winfo_exists()
        ):
            self.ventana_analisis_defectos.lift()
            self.ventana_analisis_defectos.focus_force()
            return

        self.ventana_analisis_defectos = ctk.CTkToplevel(
            self.root
        )
        self.ventana_analisis_defectos.protocol(
            "WM_DELETE_WINDOW",
            self.cerrar_ventana_analisis_defectos
        )

        self.ventana_analisis_defectos.title(
            "Análisis de defectos SMT"
        )

        self.ventana_analisis_defectos.minsize(
            1000,
            650
        )
        self.ventana_analisis_defectos.after(
            100,
            lambda: self.ventana_analisis_defectos.state("zoomed")
        )

        self.ventana_analisis_defectos.transient(
            self.root
        )

        self.ventana_analisis_defectos.grid_columnconfigure(
            0,
            weight=1
        )

        self.ventana_analisis_defectos.grid_rowconfigure(
            1,
            weight=1
        )

        # =====================================================
        # FILTROS
        # =====================================================

        self.frame_filtros_analisis = ctk.CTkFrame(
            self.ventana_analisis_defectos,
            corner_radius=10
        )

        self.frame_filtros_analisis.grid(
            row=0,
            column=0,
            padx=15,
            pady=(15, 8),
            sticky="ew"
        )

        self.frame_filtros_analisis.grid_columnconfigure(
            8,
            weight=1
        )

        titulo = ctk.CTkLabel(
            self.frame_filtros_analisis,
            text="ANÁLISIS DE DEFECTOS",
            font=("Arial", 21, "bold")
        )

        titulo.grid(
            row=0,
            column=0,
            padx=(15, 25),
            pady=15
        )

        # -----------------------------------------------------
        # FECHA INDEPENDIENTE
        # -----------------------------------------------------

        ctk.CTkLabel(
            self.frame_filtros_analisis,
            text="Fecha:",
            font=("Arial", 13, "bold")
        ).grid(
            row=0,
            column=1,
            padx=(5, 5)
        )

        self.selector_fecha_analisis = DateEntry(
            self.frame_filtros_analisis,
            width=12,
            date_pattern="dd/mm/yyyy",
            font=("Arial", 12),
            background="#2878D0",
            foreground="white",
            borderwidth=0
        )

        # Inicia con la misma fecha del dashboard principal,
        # pero después funciona de manera independiente.
        try:
            self.selector_fecha_analisis.set_date(
                self.selector_fecha.get_date()
            )
        except Exception:
            self.selector_fecha_analisis.set_date(
                date.today()
            )

        self.selector_fecha_analisis.grid(
            row=0,
            column=2,
            padx=(0, 15),
            pady=12
        )

        # -----------------------------------------------------
        # HORA DE INICIO
        # -----------------------------------------------------

        ctk.CTkLabel(
            self.frame_filtros_analisis,
            text="Hora inicial:",
            font=("Arial", 13, "bold")
        ).grid(
            row=0,
            column=3,
            padx=(5, 5)
        )

        horarios = self.generar_opciones_hora(
            intervalo_minutos=30
        )

        self.hora_inicio_analisis = ctk.StringVar(
            value="00:00"
        )

        self.combo_hora_inicio_analisis = ctk.CTkComboBox(
            self.frame_filtros_analisis,
            variable=self.hora_inicio_analisis,
            values=horarios,
            width=95,
            height=32,
            state="readonly",
            font=("Arial", 13),
            dropdown_font=("Arial", 12)
        )

        self.combo_hora_inicio_analisis.grid(
            row=0,
            column=4,
            padx=(0, 15)
        )

        # -----------------------------------------------------
        # HORA FINAL
        # -----------------------------------------------------

        ctk.CTkLabel(
            self.frame_filtros_analisis,
            text="Hora final:",
            font=("Arial", 13, "bold")
        ).grid(
            row=0,
            column=5,
            padx=(5, 5)
        )

        horarios_finales = list(horarios)

        if "23:59" not in horarios_finales:
            horarios_finales.append("23:59")

        self.hora_final_analisis = ctk.StringVar(
            value="23:59"
        )

        self.combo_hora_final_analisis = ctk.CTkComboBox(
            self.frame_filtros_analisis,
            variable=self.hora_final_analisis,
            values=horarios_finales,
            width=95,
            height=32,
            state="readonly",
            font=("Arial", 13),
            dropdown_font=("Arial", 12)
        )

        self.combo_hora_final_analisis.grid(
            row=0,
            column=6,
            padx=(0, 15)
        )

        # -----------------------------------------------------
        # MODELO
        # -----------------------------------------------------

        ctk.CTkLabel(
            self.frame_filtros_analisis,
            text="Modelo:",
            font=("Arial", 13, "bold")
        ).grid(
            row=0,
            column=7,
            padx=(5, 5)
        )

        modelos = ["Todos los modelos"]

        modelos.extend(
            sorted(
                self.configuracion_modelos.keys()
            )
        )

        self.modelo_analisis_defectos = ctk.StringVar(
            value="Todos los modelos"
        )

        self.combo_modelo_analisis = ctk.CTkComboBox(
            self.frame_filtros_analisis,
            variable=self.modelo_analisis_defectos,
            values=modelos,
            width=190,
            height=32,
            state="readonly",
            font=("Arial", 13),
            dropdown_font=("Arial", 12)
        )

        self.combo_modelo_analisis.grid(
            row=0,
            column=8,
            padx=(0, 15)
        )

        # -----------------------------------------------------
        # ACTUALIZAR
        # -----------------------------------------------------

        self.btn_actualizar_analisis = ctk.CTkButton(
            self.frame_filtros_analisis,
            text="Actualizar",
            width=110,
            height=32,
            font=("Arial", 13, "bold"),
            command=self.actualizar_dashboard_analisis_defectos
        )

        self.btn_actualizar_analisis.grid(
            row=0,
            column=9,
            padx=(5, 15),
            pady=12
        )

        # Actualizar automáticamente al seleccionar fecha
        self.selector_fecha_analisis.bind(
            "<<DateEntrySelected>>",
            lambda evento: (
                self.actualizar_dashboard_analisis_defectos()
            )
        )

        # =====================================================
        # CONTENEDOR DEL DASHBOARD
        # =====================================================

        self.frame_dashboard_analisis = ctk.CTkFrame(
            self.ventana_analisis_defectos,
            corner_radius=10
        )

        self.frame_dashboard_analisis.grid(
            row=1,
            column=0,
            padx=15,
            pady=(8, 15),
            sticky="nsew"
        )

        self.frame_dashboard_analisis.grid_columnconfigure(
            0,
            weight=1
        )

        self.frame_dashboard_analisis.grid_rowconfigure(
            1,
            weight=1
        )

        self.actualizar_dashboard_analisis_defectos()

    def separar_descripciones_otro(self, texto):
        """
        Convierte el contenido de DescripcionOtro en una lista
        de tuplas: [(descripcion, cantidad), ...].
        """

        import re

        resultados = []

        if texto is None:
            return resultados

        texto = str(texto).strip()

        if not texto or texto.lower() == "nan":
            return resultados

        partes = texto.split(" | ")

        for parte in partes:
            parte = parte.strip()

            if not parte:
                continue

            coincidencia = re.match(
                r"^(.*?)\s*\((\d+)\)\s*$",
                parte
            )

            if coincidencia:
                descripcion = (
                    coincidencia.group(1).strip()
                )

                cantidad = int(
                    coincidencia.group(2)
                )
            else:
                # Compatibilidad con registros anteriores
                descripcion = parte
                cantidad = 1

            if descripcion:
                resultados.append(
                    (descripcion, cantidad)
                )

        return resultados

    def obtener_fecha_dashboard_seleccionada(self):
        """
        Obtiene la fecha seleccionada en el dashboard.
        """

        return self.selector_fecha.get().strip()

    def generar_opciones_hora(self, intervalo_minutos=30):
        """
        Genera horarios desde 00:00 hasta 23:30.

        Ejemplo con intervalo de 30 minutos:
        00:00, 00:30, 01:00, 01:30...
        """

        horarios = []

        for hora in range(24):
            for minuto in range(0, 60, intervalo_minutos):
                horarios.append(
                    f"{hora:02d}:{minuto:02d}"
                )

        return horarios

    def obtener_datos_analisis_defectos(
        self,
        fecha_seleccionada,
        hora_inicio,
        hora_final,
        modelo_seleccionado
    ):
        """
        Lee LogFilePCB.csv y filtra los registros por:

        - Fecha
        - Hora inicial
        - Hora final
        - Modelo
        """

        import os
        import pandas as pd

        if not os.path.exists(self.archivo_log_pcb):
            return pd.DataFrame()

        try:
            df = pd.read_csv(
                self.archivo_log_pcb,
                encoding="utf-8-sig"
            )

        except Exception as error:
            messagebox.showerror(
                "Error",
                (
                    "No fue posible leer LogFilePCB.csv.\n\n"
                    f"{error}"
                ),
                parent=self.ventana_analisis_defectos
            )

            return pd.DataFrame()

        if "Fecha/Hora" not in df.columns:
            return pd.DataFrame()

        # Convierte tanto fecha como hora.
        df["FechaHoraConvertida"] = pd.to_datetime(
            df["Fecha/Hora"],
            dayfirst=True,
            errors="coerce"
        )

        df = df.dropna(
            subset=["FechaHoraConvertida"]
        ).copy()

        try:
            fecha_inicio = pd.to_datetime(
                f"{fecha_seleccionada} {hora_inicio}",
                format="%d/%m/%Y %H:%M"
            )

            fecha_final = pd.to_datetime(
                f"{fecha_seleccionada} {hora_final}",
                format="%d/%m/%Y %H:%M"
            )

        except ValueError:
            messagebox.showwarning(
                "Filtro incorrecto",
                "La fecha o el horario seleccionado no es válido.",
                parent=self.ventana_analisis_defectos
            )

            return pd.DataFrame()

        if fecha_inicio > fecha_final:
            messagebox.showwarning(
                "Horario incorrecto",
                (
                    "La hora inicial no puede ser mayor "
                    "que la hora final."
                ),
                parent=self.ventana_analisis_defectos
            )

            return pd.DataFrame()

        df = df[
            (
                df["FechaHoraConvertida"]
                >= fecha_inicio
            )
            & (
                df["FechaHoraConvertida"]
                <= fecha_final
            )
        ].copy()

        if modelo_seleccionado != "Todos los modelos":
            df = df[
                df["Modelo"].astype(str).str.strip()
                == modelo_seleccionado
            ].copy()

        return df

    def obtener_columnas_defectos_analisis(self, df):
        """
        Retorna todos los defectos configurados en defects.ini,
        más la categoría especial Otro.
        """

        defectos = list(self.lista_defectos)

        if self.opcion_otro not in defectos:
            defectos.append(
                self.opcion_otro
            )

        # Solo usar columnas que realmente existan en el CSV.
        return [
            defecto
            for defecto in defectos
            if defecto in df.columns
        ]

    def calcular_totales_defectos_analisis(self, df):
        """
        Calcula la cantidad de cada defecto configurado.
        Incluye defectos con cero ocurrencias.
        """

        import pandas as pd

        defectos_configurados = list(
            self.lista_defectos
        )

        if self.opcion_otro not in defectos_configurados:
            defectos_configurados.append(
                self.opcion_otro
            )

        resultados = []

        for defecto in defectos_configurados:

            if defecto in df.columns:
                cantidades = pd.to_numeric(
                    df[defecto],
                    errors="coerce"
                ).fillna(0)

                total = int(cantidades.sum())
            else:
                total = 0

            resultados.append(
                {
                    "Defecto": defecto,
                    "Cantidad": total
                }
            )

        return pd.DataFrame(resultados)

    def actualizar_dashboard_analisis_defectos(self):
        """
        Actualiza el análisis completo de defectos.
        """

        import pandas as pd
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_tkagg import (
            FigureCanvasTkAgg
        )

        fecha = (
            self.selector_fecha_analisis
            .get()
            .strip()
        )

        hora_inicio = (
            self.hora_inicio_analisis
            .get()
            .strip()
        )

        hora_final = (
            self.hora_final_analisis
            .get()
            .strip()
        )

        modelo = (
            self.modelo_analisis_defectos
            .get()
            .strip()
        )

        # Limpiar dashboard anterior
        for widget in (
            self.frame_dashboard_analisis
            .winfo_children()
        ):
            widget.destroy()

        df = self.obtener_datos_analisis_defectos(
            fecha,
            hora_inicio,
            hora_final,
            modelo
        )

        if df.empty:
            mensaje = ctk.CTkLabel(
                self.frame_dashboard_analisis,
                text=(
                    "No existen registros para los filtros "
                    "seleccionados."
                ),
                font=("Arial", 18, "bold"),
                text_color="#AEB4D0"
            )

            mensaje.grid(
                row=0,
                column=0,
                padx=20,
                pady=40
            )

            return

        datos_defectos = (
            self.calcular_totales_defectos_analisis(
                df
            )
        )

        datos_otros = self.calcular_detalle_otros_defectos(
            df
        )

        # Eliminar la barra general "Otro".
        datos_defectos = datos_defectos[
            datos_defectos["Defecto"].astype(str).str.strip()
            != self.opcion_otro
        ].copy()

        # Agregar el detalle de Otros.
        if not datos_otros.empty:
            datos_defectos = pd.concat(
                [
                    datos_defectos,
                    datos_otros
                ],
                ignore_index=True
            )

        total_pcb = len(df)

        if "Resultado" in df.columns:
            total_fail = int(
                (
                    df["Resultado"]
                    .astype(str)
                    .str.upper()
                    .str.strip()
                    == "FAIL"
                ).sum()
            )
        else:
            total_fail = 0

        total_ocurrencias = int(
            datos_defectos["Cantidad"].sum()
        )

        defectos_con_registro = int(
            (
                datos_defectos["Cantidad"] > 0
            ).sum()
        )

        # =====================================================
        # TARJETAS DE RESUMEN
        # =====================================================

        frame_resumen = ctk.CTkFrame(
            self.frame_dashboard_analisis,
            fg_color="transparent"
        )

        frame_resumen.grid(
            row=0,
            column=0,
            padx=12,
            pady=(12, 5),
            sticky="ew"
        )

        for columna in range(4):
            frame_resumen.grid_columnconfigure(
                columna,
                weight=1
            )

        datos_tarjetas = [
            (
                "PCB inspeccionadas",
                total_pcb
            ),
            (
                "PCB defectuosas",
                total_fail
            ),
            (
                "Ocurrencias",
                total_ocurrencias
            ),
            (
                "Tipos encontrados",
                defectos_con_registro
            )
        ]

        for indice, (titulo, valor) in enumerate(
            datos_tarjetas
        ):
            tarjeta = ctk.CTkFrame(
                frame_resumen,
                corner_radius=10,
                border_width=1,
                border_color="#454B70"
            )

            tarjeta.grid(
                row=0,
                column=indice,
                padx=6,
                pady=5,
                sticky="ew"
            )

            ctk.CTkLabel(
                tarjeta,
                text=titulo,
                font=("Arial", 13, "bold"),
                text_color="#AEB4D0"
            ).pack(
                pady=(10, 2)
            )

            ctk.CTkLabel(
                tarjeta,
                text=str(valor),
                font=("Arial", 25, "bold")
            ).pack(
                pady=(0, 10)
            )

        # =====================================================
        # PREPARAR GRÁFICA
        # =====================================================

        # Mantiene todos los defectos de defects.ini.
        grafica = datos_defectos.sort_values(
            by="Cantidad",
            ascending=True
        ).copy()

        figura = Figure(
            figsize=(11, 6),
            dpi=100
        )

        figura.patch.set_facecolor("#252842")

        eje = figura.add_subplot(111)
        eje.set_facecolor("#252842")

        barras = eje.barh(
            grafica["Defecto"],
            grafica["Cantidad"]
        )

        titulo_modelo = (
            modelo
            if modelo != "Todos los modelos"
            else "Todos los modelos"
        )

        eje.set_title(
            (
                "DEFECTOS REGISTRADOS\n"
                f"{titulo_modelo} | "
                f"{fecha} | "
                f"{hora_inicio} - {hora_final}"
            ),
            fontsize=15,
            fontweight="bold",
            color="#E6E8FF",
            pad=15
        )

        eje.set_xlabel(
            "Cantidad de ocurrencias",
            color="#D8DCF5"
        )

        eje.tick_params(
            axis="x",
            colors="#D8DCF5"
        )

        eje.tick_params(
            axis="y",
            colors="#D8DCF5",
            labelsize=9
        )

        eje.grid(
            axis="x",
            linestyle="--",
            alpha=0.25
        )

        cantidad_maxima = max(
            int(grafica["Cantidad"].max()),
            1
        )

        eje.set_xlim(
            0,
            cantidad_maxima * 1.15
        )

        for barra, cantidad in zip(
            barras,
            grafica["Cantidad"]
        ):
            eje.text(
                barra.get_width() + (
                    cantidad_maxima * 0.01
                ),
                barra.get_y()
                + barra.get_height() / 2,
                str(int(cantidad)),
                va="center",
                fontsize=9,
                fontweight="bold",
                color="#FFFFFF"
            )

        eje.spines["top"].set_visible(False)
        eje.spines["right"].set_visible(False)

        eje.spines["bottom"].set_color(
            "#6F7390"
        )

        eje.spines["left"].set_color(
            "#6F7390"
        )

        figura.tight_layout()

        canvas = FigureCanvasTkAgg(
            figura,
            master=self.frame_dashboard_analisis
        )

        canvas.draw()

        canvas.get_tk_widget().grid(
            row=1,
            column=0,
            padx=12,
            pady=(5, 12),
            sticky="nsew"
        )

        self.canvas_analisis_defectos = canvas

    def cerrar_ventana_analisis_defectos(self):
        """Cerrar vetana de anlisis"""
        if (
            self.ventana_analisis_defectos is not None
            and self.ventana_analisis_defectos.winfo_exists()
        ):
            self.ventana_analisis_defectos.destroy()

        self.ventana_analisis_defectos = None

    def calcular_detalle_otros_defectos(self, df):
        """
        Obtiene las descripciones registradas en la columna
        DescripcionOtro y suma sus cantidades.

        Ejemplo:
        Cap roto (2) | PCB rayada (1)

        Resultado:
        Otro: Cap roto = 2
        Otro: PCB rayada = 1
        """

        import re
        import pandas as pd

        resultados = {}

        if (
            df.empty
            or "DescripcionOtro" not in df.columns
        ):
            return pd.DataFrame(
                columns=["Defecto", "Cantidad"]
            )

        for valor in df["DescripcionOtro"].fillna(""):

            texto = str(valor).strip()

            if not texto or texto.lower() == "nan":
                continue

            partes = texto.split("|")

            for parte in partes:
                parte = parte.strip()

                if not parte:
                    continue

                coincidencia = re.match(
                    r"^(.*?)\s*\((\d+)\)\s*$",
                    parte
                )

                if coincidencia:
                    descripcion = (
                        coincidencia.group(1).strip()
                    )

                    cantidad = int(
                        coincidencia.group(2)
                    )
                else:
                    descripcion = parte
                    cantidad = 1

                if not descripcion:
                    continue

                etiqueta = f"Otro: {descripcion}"

                resultados[etiqueta] = (
                    resultados.get(etiqueta, 0)
                    + cantidad
                )

        if not resultados:
            return pd.DataFrame(
                columns=["Defecto", "Cantidad"]
            )

        return pd.DataFrame(
            [
                {
                    "Defecto": defecto,
                    "Cantidad": cantidad
                }
                for defecto, cantidad in resultados.items()
            ]
        )


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    app = RegistroDefectosSMT(root)
    root.mainloop()

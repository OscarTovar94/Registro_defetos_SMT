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

        self.root.title("Registro de Defectos SMT")
        # self.root.state("zoomed")
        self.root.iconbitmap("C:/Registro_defetos_SMT/Image/elrad.ico")
        self.root.configure(fg_color="#21233C")

        # Archivos de configuración
        self.archivo_defectos = "C:/Registro_defetos_SMT/Settings/defects.ini"
        self.archivo_modelos = "C:/Registro_defetos_SMT/Settings/models.ini"
        self.archivo_log = "C:/Registro_defetos_SMT/LogFile/LogFile.csv"

        os.makedirs(
            os.path.dirname(self.archivo_log),
            exist_ok=True
        )

        # Variables
        self.defecto_seleccionado = ctk.StringVar(value="")
        self.modelo_seleccionado = ctk.StringVar(value="")
        self.cantidad = ctk.StringVar(value="")
        self.opcion_sin_defecto = "Ningún defecto"
        self.lista_defectos = []
        self.defectos_seleccionados = []
        self.entries_cantidades = {}
        self.estandares_modelos = {}
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
        """
        Crea el panel para seleccionar el defecto, ingresar la cantidad,
        seleccionar el modelo y registrar la información.
        """

        # =========================================================
        # FRAME DE REGISTRO
        # =========================================================
        self.frame_registro = ctk.CTkFrame(
            self.root,
            height=75,
            corner_radius=12,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70"
        )
        self.frame_registro.pack(
            fill="x",
            padx=15,
            pady=(0, 0)
        )
        self.frame_registro.pack_propagate(False)

        # Configuración de columnas
        self.frame_registro.grid_columnconfigure(
            0,
            weight=1,
            uniform="registro"
        )
        self.frame_registro.grid_columnconfigure(
            1,
            weight=1,
            uniform="registro"
        )
        self.frame_registro.grid_columnconfigure(
            2,
            weight=1,
            uniform="registro"
        )
        self.frame_registro.grid_columnconfigure(
            3,
            weight=1,
            uniform="registro"
        )

        # =========================================================
        # DEFECTO
        # =========================================================
        self.lbl_titulo_defecto = ctk.CTkLabel(
            self.frame_registro,
            text="Defecto",
            font=("Arial", 14, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_titulo_defecto.grid(
            row=0,
            column=0,
            padx=15,
            pady=(5, 5),
            sticky="nsew"
        )

        self.btn_seleccionar_defectos = ctk.CTkButton(
            self.frame_registro,
            text="Seleccionar defectos",
            height=30,
            corner_radius=14,
            font=("Arial", 14),
            fg_color="#343853",
            hover_color="#414662",
            anchor="w",
            command=self.abrir_selector_defectos
        )

        self.btn_seleccionar_defectos.grid(
            row=1,
            column=0,
            padx=15,
            sticky="ew"
        )

        # =========================================================
        # CANTIDAD
        # =========================================================
        self.lbl_titulo_cantidad = ctk.CTkLabel(
            self.frame_registro,
            text="Cantidad",
            font=("Arial", 14, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_titulo_cantidad.grid(
            row=0,
            column=1,
            padx=15,
            pady=(5, 5),
            sticky="nsew"
        )

        self.frame_cantidades = ctk.CTkFrame(
            self.frame_registro,
            fg_color="#1F2238",
            corner_radius=14,
            height=30
        )

        self.frame_cantidades.grid(
            row=1,
            column=1,
            padx=15,
            pady=(5, 5),
            sticky="ew"
        )

        self.frame_cantidades.grid_propagate(False)

        # =========================================================
        # MODELO
        # =========================================================
        self.lbl_titulo_modelo = ctk.CTkLabel(
            self.frame_registro,
            text="Modelo",
            font=("Arial", 14, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_titulo_modelo.grid(
            row=0,
            column=2,
            padx=15,
            pady=(5, 5),
            sticky="nsew"
        )

        self.combo_modelos = ctk.CTkComboBox(
            self.frame_registro,
            variable=self.modelo_seleccionado,
            values=["Sin modelos"],
            height=30,
            corner_radius=8,
            font=("Arial", 14),
            dropdown_font=("Arial", 14),
            state="readonly"
        )
        self.combo_modelos.grid(
            row=1,
            column=2,
            padx=15,
            sticky="ew"
        )

        # =========================================================
        # BOTÓN REGISTRAR
        # =========================================================
        self.lbl_titulo_accion = ctk.CTkLabel(
            self.frame_registro,
            text="Acción",
            font=("Arial", 14, "bold"),
            text_color="#AEB4D0"
        )
        self.lbl_titulo_accion.grid(
            row=0,
            column=3,
            padx=15,
            pady=(5, 5),
            sticky="nsew"
        )

        self.btn_registrar = ctk.CTkButton(
            self.frame_registro,
            text="Registrar defecto",
            height=30,
            corner_radius=14,
            font=("Arial", 14, "bold"),
            fg_color="#2878D0",
            hover_color="#1E609F",
            command=self.registrar_defecto
        )
        self.btn_registrar.grid(
            row=1,
            column=3,
            padx=15,
            sticky="ew"
        )

        # =========================================================
        # FRAME PARA GRÁFICAS
        # =========================================================
        self.frame_graficas = ctk.CTkFrame(
            self.root,
            corner_radius=12,
            fg_color="#292C47",
            border_width=1,
            border_color="#454B70"
        )
        self.frame_graficas.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=(5, 5)
        )
        # =========================================================
        # CONFIGURACIÓN DEL ÁREA DE GRÁFICAS
        # =========================================================
        self.frame_graficas.grid_columnconfigure(0, weight=0)
        self.frame_graficas.grid_columnconfigure(1, weight=0)
        self.frame_graficas.grid_columnconfigure(2, weight=1)
        self.frame_graficas.grid_columnconfigure(3, weight=0)
        self.frame_graficas.grid_rowconfigure(0, weight=0)
        self.frame_graficas.grid_rowconfigure(1, weight=1)
        self.frame_graficas.grid_rowconfigure(2, weight=1)
        self.frame_graficas.grid_rowconfigure(3, weight=0)

        # =========================================================
        # ENCABEZADO Y FILTRO POR FECHA
        # =========================================================
        self.frame_filtro_graficas = ctk.CTkFrame(
            self.frame_graficas,
            height=30,
            corner_radius=10,
            fg_color="#252842"
        )

        self.frame_filtro_graficas.grid(
            row=0,
            column=0,
            columnspan=4,
            padx=15,
            pady=(5, 5),
            sticky="w"
        )

        self.frame_filtro_graficas.grid_columnconfigure(1, weight=1)

        self.lbl_filtro_fecha = ctk.CTkLabel(
            self.frame_filtro_graficas,
            text="Fecha:",
            font=("Arial", 12, "bold"),
            text_color="#AEB4D0"
        )

        self.lbl_filtro_fecha.grid(
            row=0,
            column=0,
            padx=(15, 8),
            pady=5
        )

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

        self.selector_fecha.grid(
            row=0,
            column=1,
            padx=5,
            pady=12,
            sticky="w"
        )

        self.selector_fecha.bind(
            "<<DateEntrySelected>>",
            lambda evento: self.solicitar_actualizacion_dashboard()
        )

        self.lbl_frase = ctk.CTkLabel(
            self.frame_filtro_graficas,
            text="Tus manos definen la calidad del producto; tu atención asegura el orgullo de nuestro trabajo.",
            font=("Arial", 14, "bold", "italic"),
            text_color="#FFD166",
            justify="left",
            anchor="w"
        )

        self.lbl_frase.grid(
            row=0,
            column=2,
            padx=20,
            pady=0,
            sticky="w"
        )

        self.btn_actualizar_fpy = ctk.CTkButton(
            self.frame_filtro_graficas,
            text="Actualizar",
            width=120,
            height=30,
            font=("Arial", 12, "bold"),
            command=self.solicitar_actualizacion_dashboard
        )

        self.btn_actualizar_fpy.grid(
            row=0,
            column=3,
            padx=15,
            pady=5
        )

        # =========================================================
        # TARJETA DEL FPY TOTAL
        # =========================================================
        self.frame_fpy_total = ctk.CTkFrame(
            self.frame_graficas,
            width=200,
            height=150,
            corner_radius=14,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70"
        )

        self.frame_fpy_total.grid(
            row=1,
            column=0,
            padx=(15, 8),
            pady=(5, 5),
            sticky="nsew"
        )

        self.frame_fpy_total.grid_columnconfigure(0, weight=1)

        self.lbl_titulo_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text="FPY TOTAL",
            font=("Arial", 20, "bold"),
            text_color="#AEB4D0"
        )

        self.lbl_titulo_fpy.grid(
            row=0,
            column=0,
            padx=15,
            pady=(5, 5)
        )

        self.lbl_valor_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text="0.00 %",
            font=("Arial", 50, "bold"),
            text_color="#8F96B8"
        )

        self.lbl_valor_fpy.grid(
            row=1,
            column=0,
            padx=15,
            pady=(5, 5)
        )

        self.lbl_detalle_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text=(
                "Producción total: 0\n"
                "Defectos totales: 0"
            ),
            font=("Arial", 15),
            text_color="#AEB4D0",
            justify="center"
        )

        self.lbl_detalle_fpy.grid(
            row=2,
            column=0,
            padx=15,
            pady=(5, 5)
        )

        self.lbl_top_fpy = ctk.CTkLabel(
            self.frame_fpy_total,
            text="",
            font=("Arial", 14, "bold"),
            justify="left",
            anchor="w",
            text_color="#DDE2FF"
        )

        self.lbl_top_fpy.grid(
            row=4,
            column=0,
            padx=18,
            pady=(5, 5),
            sticky="w"
        )

        # =========================================================
        # ÁREA DE FPY POR MODELO
        # =========================================================
        self.frame_fpy_modelos = ctk.CTkScrollableFrame(
            self.frame_graficas,
            # label_text="FPY POR MODELO",
            label_font=("Arial", 10, "bold"),
            corner_radius=14,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70",
            orientation="horizontal"
        )

        self.frame_fpy_modelos.grid(
            row=1,
            column=1,
            columnspan=3,
            padx=(8, 15),
            pady=(5, 5),
            sticky="nsew"
        )
        # =========================================================
        # PARETO GLOBAL
        # =========================================================
        self.frame_pareto_global = ctk.CTkFrame(
            self.frame_graficas,
            corner_radius=14,
            fg_color="#252842",
            border_width=1,
            border_color="#454B70"
        )

        self.frame_pareto_global.grid(
            row=2,
            column=0,
            columnspan=4,
            padx=15,
            pady=(5, 5),
            sticky="nsew"
        )

        self.lbl_by = ctk.CTkLabel(
            self.frame_pareto_global,
            text="Rev: 2.0 (By: Oscar Tovar)",
            font=("Arial", 10),
            text_color="#AEB4D0"
        )

        self.lbl_by.grid(
            row=1,
            column=0,
            padx=15,
            pady=(1, 5),
            sticky="e"
        )

        self.frame_pareto_global.grid_rowconfigure(0, weight=1)
        self.frame_pareto_global.grid_rowconfigure(1, weight=0)
        self.frame_pareto_global.grid_columnconfigure(0, weight=1)

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
    def leer_modelos_estandar(ruta):
        """
        Lee models.ini con el formato:

        Modelo,Estandar

        Ejemplo:
        ROUTER,10
        PR20,15

        Retorna:
        {
            "ROUTER": 10,
            "PR20": 15
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

                for numero_linea, linea in enumerate(archivo, start=1):
                    linea = linea.strip()

                    if not linea:
                        continue

                    if linea.startswith(("#", ";")):
                        continue

                    partes = linea.split(",")

                    if len(partes) != 2:
                        print(
                            f"Línea incorrecta en models.ini "
                            f"({numero_linea}): {linea}"
                        )
                        continue

                    modelo = partes[0].strip()
                    estandar_texto = partes[1].strip()

                    if not modelo:
                        continue

                    try:
                        estandar = int(estandar_texto)
                    except ValueError:
                        print(
                            f"Estándar incorrecto para {modelo}: "
                            f"{estandar_texto}"
                        )
                        continue

                    if estandar <= 0:
                        print(
                            f"El estándar de {modelo} debe ser "
                            "mayor que cero."
                        )
                        continue

                    modelos[modelo] = estandar

        except UnicodeDecodeError:
            try:
                with open(
                    ruta,
                    mode="r",
                    encoding="latin-1"
                ) as archivo:

                    for linea in archivo:
                        linea = linea.strip()

                        if not linea or linea.startswith(("#", ";")):
                            continue

                        partes = linea.split(",")

                        if len(partes) != 2:
                            continue

                        modelo = partes[0].strip()

                        try:
                            estandar = int(partes[1].strip())
                        except ValueError:
                            continue

                        if modelo and estandar > 0:
                            modelos[modelo] = estandar

            except OSError as error:
                messagebox.showerror(
                    "Error de lectura",
                    (
                        "No fue posible leer models.ini.\n\n"
                        f"{error}"
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
        """
        Carga los defectos desde defects.ini.

        Ya no actualiza un ComboBox, porque los defectos se seleccionan
        desde la ventana con CheckBox.
        """
        defectos = self.leer_lista_archivo(
            self.archivo_defectos
        )

        self.lista_defectos = defectos

        # Eliminar selecciones que ya no existan en defects.ini
        self.defectos_seleccionados = [
            defecto
            for defecto in self.defectos_seleccionados
            if defecto in self.lista_defectos
        ]

        # Actualizar los campos de cantidades si la interfaz ya existe
        if hasattr(self, "frame_cantidades"):
            self.actualizar_campos_cantidades()

    def cargar_modelos(self):
        """
        Carga los modelos y sus estándares desde models.ini.
        """
        self.estandares_modelos = self.leer_modelos_estandar(
            self.archivo_modelos
        )

        modelos = list(self.estandares_modelos.keys())
        modelo_actual = self.modelo_seleccionado.get()

        if modelos:
            self.combo_modelos.configure(
                values=modelos,
                state="readonly"
            )

            if modelo_actual in modelos:
                self.modelo_seleccionado.set(modelo_actual)
            else:
                self.modelo_seleccionado.set(modelos[0])

        else:
            self.combo_modelos.configure(
                values=["Sin modelos"],
                state="disabled"
            )
            self.modelo_seleccionado.set("Sin modelos")

    def verificar_cambios_archivos(self):
        """Verifica si los archivos defects.ini y models.ini han cambiado y actualiza la información."""

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

    @staticmethod
    def validar_entrada_cantidad(valor):
        """Permite campo vacío o únicamente números enteros."""

        return valor == "" or valor.isdigit()

    def registrar_defecto(self):
        """Registra un panel completo en una sola fila del CSV."""
        modelo = self.modelo_seleccionado.get().strip()

        # =====================================================
        # VALIDAR MODELO
        # =====================================================
        if not modelo or modelo == "Sin modelos":
            messagebox.showwarning(
                "Modelo requerido",
                "Seleccione un modelo válido."
            )
            return

        if modelo not in self.estandares_modelos:
            messagebox.showwarning(
                "Estándar no encontrado",
                (
                    f'No se encontró el estándar del modelo "{modelo}" '
                    "en models.ini."
                )
            )
            return

        estandar = self.estandares_modelos[modelo]

        # Todos los defectos empiezan con cantidad 0
        cantidades_defectos = {
            defecto: 0
            for defecto in self.lista_defectos
        }

        total_defectos = 0

        # =====================================================
        # LEER DEFECTOS SELECCIONADOS
        # =====================================================
        for defecto in self.defectos_seleccionados:

            if defecto not in self.entries_cantidades:
                continue

            cantidad_texto = (
                self.entries_cantidades[defecto]
                .get()
                .strip()
            )

            if not cantidad_texto:
                messagebox.showwarning(
                    "Cantidad requerida",
                    f'Ingrese la cantidad para "{defecto}".'
                )
                return

            try:
                cantidad = int(cantidad_texto)
            except ValueError:
                messagebox.showwarning(
                    "Cantidad incorrecta",
                    (
                        f'La cantidad ingresada para "{defecto}" '
                        "no es válida."
                    )
                )
                return

            if cantidad <= 0:
                messagebox.showwarning(
                    "Cantidad incorrecta",
                    (
                        f'La cantidad de "{defecto}" debe ser '
                        "mayor que cero."
                    )
                )
                return

            cantidades_defectos[defecto] = cantidad
            total_defectos += cantidad

        fecha_hora = datetime.now().strftime(
            "%d/%m/%Y %H:%M:%S"
        )

        guardado = self.guardar_registro_csv(
            modelo=modelo,
            total_defectos=total_defectos,
            estandar=estandar,
            fecha_hora=fecha_hora,
            cantidades_defectos=cantidades_defectos
        )

        if not guardado:
            return

        self.solicitar_actualizacion_dashboard()
        self.defectos_seleccionados.clear()
        self.actualizar_campos_cantidades()

    def calcular_fpy_total(self):
        """Calcula el FPY total de la fecha seleccionada. FPY = ((Estandar total - Defectos totales)/ Estandar total) * 100"""

        if not os.path.exists(self.archivo_log):
            self.mostrar_fpy_sin_datos(
                "No existe LogFile.csv"
            )
            return

        try:
            fecha_seleccionada = self.selector_fecha.get_date()

            estandar_total = 0
            defectos_totales = 0
            registros_encontrados = 0

            with open(
                self.archivo_log,
                mode="r",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:

                lector = csv.DictReader(archivo)

                encabezados = lector.fieldnames or []

                columnas_requeridas = {
                    "Defectos",
                    "Estandar",
                    "Fecha/Hora"
                }

                if not columnas_requeridas.issubset(encabezados):
                    messagebox.showerror(
                        "Formato incorrecto",
                        (
                            "LogFile.csv no contiene las columnas "
                            "requeridas:\n\n"
                            "Defectos, Estandar y Fecha/Hora"
                        )
                    )
                    return

                columnas_fijas = {
                    "Modelo",
                    "Defectos",
                    "Estandar",
                    "Fecha/Hora"
                }

                columnas_defectos = [
                    columna
                    for columna in encabezados
                    if columna not in columnas_fijas
                ]

                totales_defectos = {
                    defecto: 0
                    for defecto in columnas_defectos
                }

                for fila in lector:
                    fecha_texto = fila.get(
                        "Fecha/Hora",
                        ""
                    ).strip()

                    try:
                        fecha_registro = datetime.strptime(
                            fecha_texto,
                            "%d/%m/%Y %H:%M:%S"
                        ).date()
                    except ValueError:
                        continue

                    if fecha_registro != fecha_seleccionada:
                        continue

                    try:
                        defectos = int(
                            float(fila.get("Defectos", 0))
                        )

                        estandar = int(
                            float(fila.get("Estandar", 0))
                        )

                    except (ValueError, TypeError):
                        continue

                    defectos_totales += defectos
                    estandar_total += estandar
                    registros_encontrados += 1

                    for defecto in columnas_defectos:

                        try:
                            cantidad = int(
                                float(
                                    fila.get(defecto, 0) or 0
                                )
                            )

                        except (ValueError, TypeError):
                            cantidad = 0

                        totales_defectos[defecto] += cantidad

                top_3 = sorted(
                    (
                        (defecto, cantidad)
                        for defecto, cantidad
                        in totales_defectos.items()
                        if cantidad > 0
                    ),
                    key=lambda x: x[1],
                    reverse=True
                )[:3]

            if registros_encontrados == 0 or estandar_total <= 0:
                self.mostrar_fpy_sin_datos(
                    "Sin registros para la fecha seleccionada"
                )
                return

            piezas_buenas = estandar_total - defectos_totales

            # Evitar resultados menores de cero
            piezas_buenas = max(piezas_buenas, 0)

            fpy = (
                piezas_buenas
                / estandar_total
            ) * 100

            # Evitar valores fuera del rango
            fpy = max(0.0, min(fpy, 100.0))

            self.lbl_valor_fpy.configure(
                text=f"{fpy:.2f} %"
            )

            self.lbl_detalle_fpy.configure(
                text=(
                    f"Producción total: {estandar_total}\n"
                    f"Defectos totales: {defectos_totales}\n"
                    f"Registros: {registros_encontrados}"
                )
            )

            # -------------------------------
            # Actualizar Top 3 de defectos
            # -------------------------------
            self.lbl_top_fpy.configure(
                text=self.formatear_top_3_defectos(top_3)
            )

            self.aplicar_color_fpy(fpy)

        except PermissionError:
            messagebox.showerror(
                "Archivo en uso",
                (
                    "No fue posible leer LogFile.csv.\n\n"
                    "Verifique que el archivo no esté bloqueado "
                    "por otro programa."
                )
            )

        except OSError as error:
            messagebox.showerror(
                "Error de lectura",
                (
                    "No fue posible calcular el FPY.\n\n"
                    f"{error}"
                )
            )

    def mostrar_fpy_sin_datos(self, mensaje):
        """Muestra un mensaje cuando no hay datos para calcular el FPY."""

        self.lbl_valor_fpy.configure(
            text="0.00 %",
            text_color="#8F96B8"
        )

        self.lbl_detalle_fpy.configure(
            text=mensaje
        )

        self.lbl_top_fpy.configure(
            text="Sin defectos registrados"
        )

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
        """
        Actualiza una sola vez todo el dashboard.
        """
        if self.cerrando_aplicacion:
            return

        self.after_dashboard = None

        if self.dashboard_actualizando:
            return

        self.dashboard_actualizando = True

        try:
            self.calcular_fpy_total()
            self.calcular_fpy_por_modelo()
            self.actualizar_pareto_global()

        finally:
            self.dashboard_actualizando = False

    def calcular_fpy_por_modelo(self):
        """
        Calcula el FPY de cada modelo para la fecha seleccionada.

        Solo muestra modelos que tengan registros.

        También calcula el defecto con mayor cantidad
        para cada modelo.
        """

        for datos_tarjeta in self.tarjetas_modelos.values():
            datos_tarjeta["frame"].grid_remove()

        if not os.path.exists(self.archivo_log):
            self.mostrar_sin_modelos(
                "No existe LogFile.csv"
            )
            return

        fecha_seleccionada = self.selector_fecha.get_date()

        columnas_fijas = {
            "Modelo",
            "Defectos",
            "Estandar",
            "Fecha/Hora"
        }

        datos_modelos = {}

        try:
            with open(
                self.archivo_log,
                mode="r",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:

                lector = csv.DictReader(archivo)
                encabezados = lector.fieldnames or []

                columnas_requeridas = {
                    "Modelo",
                    "Defectos",
                    "Estandar",
                    "Fecha/Hora"
                }

                if not columnas_requeridas.issubset(encabezados):
                    messagebox.showerror(
                        "Formato incorrecto",
                        (
                            "LogFile.csv no contiene las columnas "
                            "requeridas:\n\n"
                            "Modelo, Defectos, Estandar y Fecha/Hora"
                        )
                    )
                    return

                # Todas las columnas que no sean fijas son defectos
                columnas_defectos = [
                    columna
                    for columna in encabezados
                    if columna not in columnas_fijas
                ]

                for fila in lector:
                    fecha_texto = fila.get(
                        "Fecha/Hora",
                        ""
                    ).strip()

                    try:
                        fecha_registro = datetime.strptime(
                            fecha_texto,
                            "%d/%m/%Y %H:%M:%S"
                        ).date()
                    except ValueError:
                        continue

                    if fecha_registro != fecha_seleccionada:
                        continue

                    modelo = fila.get(
                        "Modelo",
                        ""
                    ).strip()

                    if not modelo:
                        continue

                    try:
                        defectos = int(
                            float(fila.get("Defectos", 0) or 0)
                        )

                        estandar = int(
                            float(fila.get("Estandar", 0) or 0)
                        )

                    except (ValueError, TypeError):
                        continue

                    if modelo not in datos_modelos:
                        datos_modelos[modelo] = {
                            "estandar": 0,
                            "defectos": 0,
                            "defectos_individuales": {
                                defecto: 0
                                for defecto in columnas_defectos
                            }
                        }

                    datos_modelos[modelo]["estandar"] += estandar
                    datos_modelos[modelo]["defectos"] += defectos

                    # Sumar cada defecto del modelo
                    for defecto in columnas_defectos:
                        try:
                            cantidad = int(
                                float(fila.get(defecto, 0) or 0)
                            )
                        except (ValueError, TypeError):
                            cantidad = 0

                        datos_modelos[modelo][
                            "defectos_individuales"
                        ][defecto] += cantidad

            if not datos_modelos:
                self.mostrar_sin_modelos(
                    "Sin registros para la fecha seleccionada"
                )
                return

            if self.lbl_sin_modelos is not None:
                self.lbl_sin_modelos.grid_remove()

            # Crear una tarjeta para cada modelo con registros
            for columna, modelo in enumerate(datos_modelos):
                datos = datos_modelos[modelo]

                estandar_total = datos["estandar"]
                defectos_totales = datos["defectos"]

                piezas_buenas = max(
                    estandar_total - defectos_totales,
                    0
                )

                if estandar_total > 0:
                    fpy = (
                        piezas_buenas
                        / estandar_total
                    ) * 100
                else:
                    fpy = 0.0

                fpy = max(
                    0.0,
                    min(fpy, 100.0)
                )

                defectos_individuales = datos[
                    "defectos_individuales"
                ]

                defectos_con_cantidad = {
                    defecto: cantidad
                    for defecto, cantidad
                    in defectos_individuales.items()
                    if cantidad > 0
                }

                if defectos_con_cantidad:
                    top_3_defectos = sorted(
                        defectos_con_cantidad.items(),
                        key=lambda elemento: elemento[1],
                        reverse=True
                    )[:3]

                else:
                    top_3_defectos = []

                self.actualizar_tarjeta_fpy_modelo(
                    modelo=modelo,
                    fpy=fpy,
                    estandar=estandar_total,
                    defectos=defectos_totales,
                    top_3_defectos=top_3_defectos,
                    columna=columna
                )
        except PermissionError:
            messagebox.showerror(
                "Archivo en uso",
                (
                    "No fue posible leer LogFile.csv.\n\n"
                    "Verifique que el archivo no esté abierto "
                    "o bloqueado por otro programa."
                )
            )

        except OSError as error:
            messagebox.showerror(
                "Error de lectura",
                (
                    "No fue posible calcular el FPY por modelo.\n\n"
                    f"{error}"
                )
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
            pady=50,
            sticky="nsew"
        )

    def formatear_top_3(self, top_3):
        """Formatea top"""

        if not top_3:
            return "Sin defectos"

        medallas = ["🥇", "🥈", "🥉"]

        texto = "TOP DEFECTOS\n\n"

        for indice, (defecto, cantidad) in enumerate(top_3):

            texto += (
                f"{medallas[indice]} "
                f"{defecto}: {cantidad}\n"
            )

        return texto

    def actualizar_pareto_global(self):
        """
        Genera el Pareto global de defectos correspondiente
        a la fecha seleccionada.
        """

        if not os.path.exists(self.archivo_log):
            self.mostrar_mensaje_pareto(
                "No existe LogFile.csv"
            )
            return

        fecha_seleccionada = self.selector_fecha.get_date()

        columnas_fijas = {
            "Modelo",
            "Defectos",
            "Estandar",
            "Fecha/Hora"
        }

        try:
            with open(
                self.archivo_log,
                mode="r",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:

                lector = csv.DictReader(archivo)
                encabezados = lector.fieldnames or []

                columnas_requeridas = {
                    "Modelo",
                    "Defectos",
                    "Estandar",
                    "Fecha/Hora"
                }

                if not columnas_requeridas.issubset(encabezados):
                    self.mostrar_mensaje_pareto(
                        "LogFile.csv no tiene el formato correcto"
                    )
                    return

                columnas_defectos = [
                    columna
                    for columna in encabezados
                    if columna not in columnas_fijas
                ]

                totales_defectos = {
                    defecto: 0
                    for defecto in columnas_defectos
                }

                for fila in lector:
                    fecha_texto = fila.get(
                        "Fecha/Hora",
                        ""
                    ).strip()

                    try:
                        fecha_registro = datetime.strptime(
                            fecha_texto,
                            "%d/%m/%Y %H:%M:%S"
                        ).date()

                    except ValueError:
                        continue

                    if fecha_registro != fecha_seleccionada:
                        continue

                    for defecto in columnas_defectos:
                        try:
                            cantidad = int(
                                float(
                                    fila.get(defecto, 0) or 0
                                )
                            )

                        except (ValueError, TypeError):
                            cantidad = 0

                        totales_defectos[defecto] += cantidad

            # Quitar defectos con cantidad cero
            defectos_con_datos = [
                (defecto, cantidad)
                for defecto, cantidad
                in totales_defectos.items()
                if cantidad > 0
            ]

            if not defectos_con_datos:
                self.mostrar_mensaje_pareto(
                    "Sin defectos registrados para la fecha seleccionada"
                )
                return

            # Ordenar de mayor a menor
            defectos_con_datos.sort(
                key=lambda elemento: elemento[1],
                reverse=True
            )
            defectos_con_datos = defectos_con_datos[:10]

            nombres = [
                elemento[0]
                for elemento in defectos_con_datos
            ]

            cantidades = [
                elemento[1]
                for elemento in defectos_con_datos
            ]

            total = sum(cantidades)

            acumulado = []
            suma_acumulada = 0

            for cantidad in cantidades:
                suma_acumulada += cantidad

                porcentaje = (
                    suma_acumulada
                    / total
                ) * 100

                acumulado.append(porcentaje)

            self.crear_grafica_pareto(
                nombres=nombres,
                cantidades=cantidades,
                acumulado=acumulado
            )

        except PermissionError:
            self.mostrar_mensaje_pareto(
                "LogFile.csv está siendo utilizado por otro programa"
            )

        except OSError as error:
            self.mostrar_mensaje_pareto(
                f"No fue posible generar el Pareto:\n{error}"
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

    def actualizar_campos_cantidades(self):
        """Actualiza los campos de cantidad según los defectos seleccionados."""
        cantidades_anteriores = {
            defecto: variable.get()
            for defecto, variable in self.entries_cantidades.items()
        }

        for widget in self.frame_cantidades.winfo_children():
            widget.destroy()

        self.entries_cantidades.clear()

        if not self.defectos_seleccionados:
            self.btn_seleccionar_defectos.configure(
                text="Ningún defecto"
            )

            etiqueta = ctk.CTkLabel(
                self.frame_cantidades,
                text="Cantidad automática: 0",
                text_color="#6FE3A1"
            )
            etiqueta.pack(
                padx=10,
                pady=5
            )

            return

        texto_boton = ", ".join(
            self.defectos_seleccionados
        )

        if len(texto_boton) > 35:
            texto_boton = texto_boton[:32] + "..."

        self.btn_seleccionar_defectos.configure(
            text=texto_boton
        )

        for defecto in self.defectos_seleccionados:
            fila = ctk.CTkFrame(
                self.frame_cantidades,
                fg_color="transparent"
            )
            fila.pack(
                fill="x",
                padx=5,
                pady=5
            )

            fila.grid_columnconfigure(0, weight=1)

            etiqueta = ctk.CTkLabel(
                fila,
                text=defecto,
                font=("Arial", 18),
                anchor="w"
            )
            etiqueta.grid(
                row=0,
                column=0,
                padx=(5, 8),
                sticky="ew"
            )

            variable = ctk.StringVar(
                value=cantidades_anteriores.get(defecto, "")
            )

            entry = ctk.CTkEntry(
                fila,
                width=65,
                height=32,
                textvariable=variable,
                justify="center"
            )
            entry.grid(
                row=0,
                column=1,
                padx=5
            )

            self.entries_cantidades[defecto] = variable

    def abrir_selector_defectos(self):
        """Abre una ventana para seleccionar uno o varios defectos."""
        defectos = self.leer_lista_archivo(
            self.archivo_defectos
        )

        ventana = ctk.CTkToplevel(self.root)
        ventana.title("Seleccionar defectos")
        ventana.geometry("480x550")
        ventana.transient(self.root)
        ventana.grab_set()

        ventana.grid_columnconfigure(0, weight=1)
        ventana.grid_rowconfigure(1, weight=1)

        titulo = ctk.CTkLabel(
            ventana,
            text="Seleccione uno o varios defectos",
            font=("Arial", 20, "bold")
        )
        titulo.grid(
            row=0,
            column=0,
            padx=20,
            pady=(5, 5)
        )

        frame_lista = ctk.CTkScrollableFrame(
            ventana
        )
        frame_lista.grid(
            row=1,
            column=0,
            padx=20,
            pady=5,
            sticky="nsew"
        )

        variables = {}

        for defecto in defectos:
            variable = ctk.BooleanVar(
                value=defecto in self.defectos_seleccionados
            )

            variables[defecto] = variable

            checkbox = ctk.CTkCheckBox(
                frame_lista,
                text=defecto,
                variable=variable,
                font=("Arial", 15)
            )
            checkbox.pack(
                anchor="w",
                padx=10,
                pady=5
            )

        def confirmar():
            seleccionados = [
                defecto
                for defecto, variable in variables.items()
                if variable.get()
            ]

            self.defectos_seleccionados = seleccionados
            self.actualizar_campos_cantidades()
            ventana.destroy()

        btn_confirmar = ctk.CTkButton(
            ventana,
            text="Confirmar selección",
            height=42,
            font=("Arial", 16, "bold"),
            command=confirmar
        )
        btn_confirmar.grid(
            row=2,
            column=0,
            padx=20,
            pady=5,
            sticky="ew"
        )

    def preparar_encabezados_csv(self, encabezados_nuevos):
        """
        Verifica que LogFile.csv tenga las columnas actuales.

        Si defects.ini contiene nuevos defectos, reescribe el CSV
        agregando las nuevas columnas y conservando los datos
        anteriores.
        """
        if not os.path.exists(self.archivo_log):
            return

        if os.path.getsize(self.archivo_log) == 0:
            return

        try:
            with open(
                self.archivo_log,
                mode="r",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:

                lector = csv.DictReader(archivo)

                encabezados_actuales = lector.fieldnames or []
                registros_anteriores = list(lector)

            # Si ya tiene exactamente los encabezados necesarios,
            # no se modifica el archivo.
            if encabezados_actuales == encabezados_nuevos:
                return

            # Conservar columnas antiguas de defectos aunque se hayan
            # retirado de defects.ini.
            columnas_fijas = [
                "Modelo",
                "Defectos",
                "Estandar",
                "Fecha/Hora"
            ]

            defectos_anteriores = [
                columna
                for columna in encabezados_actuales
                if columna not in columnas_fijas
            ]

            defectos_nuevos = [
                columna
                for columna in encabezados_nuevos
                if columna not in columnas_fijas
            ]

            todos_los_defectos = defectos_anteriores.copy()

            for defecto in defectos_nuevos:
                if defecto not in todos_los_defectos:
                    todos_los_defectos.append(defecto)

            encabezados_finales = (
                columnas_fijas
                + todos_los_defectos
            )

            archivo_temporal = (
                self.archivo_log
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

                for registro in registros_anteriores:
                    fila_actualizada = {}

                    for encabezado in encabezados_finales:
                        valor = registro.get(encabezado, "")

                        # Los nuevos defectos deben iniciar en cero
                        if (
                            encabezado not in columnas_fijas
                            and valor == ""
                        ):
                            valor = 0

                        fila_actualizada[encabezado] = valor

                    escritor.writerow(fila_actualizada)

            os.replace(
                archivo_temporal,
                self.archivo_log
            )

            # Actualizar también la lista utilizada para escribir
            self.lista_defectos = todos_los_defectos

        except PermissionError:
            raise

        except OSError as error:
            raise OSError(
                f"No se pudieron actualizar los encabezados: {error}") from error

    def guardar_registro_csv(
        self,
        modelo,
        total_defectos,
        estandar,
        fecha_hora,
        cantidades_defectos
    ):
        """guarda un registro en LogFile.csv con todos los defectos y sus cantidades."""
        encabezados = [
            "Modelo",
            "Defectos",
            "Estandar",
            "Fecha/Hora"
        ] + self.lista_defectos

        try:
            self.preparar_encabezados_csv(encabezados)

            encabezados = [
                "Modelo",
                "Defectos",
                "Estandar",
                "Fecha/Hora"
            ] + self.lista_defectos

            registro = {
                "Modelo": modelo,
                "Defectos": total_defectos,
                "Estandar": estandar,
                "Fecha/Hora": fecha_hora
            }

            # Agregar la cantidad de cada defecto
            for defecto in self.lista_defectos:
                registro[defecto] = cantidades_defectos.get(
                    defecto,
                    0
                )

            with open(
                self.archivo_log,
                mode="a",
                newline="",
                encoding="utf-8-sig"
            ) as archivo:

                escritor = csv.DictWriter(
                    archivo,
                    fieldnames=encabezados
                )

                if os.path.getsize(self.archivo_log) == 0:
                    escritor.writeheader()

                escritor.writerow(registro)

            return True

        except PermissionError:
            messagebox.showerror(
                "Archivo en uso",
                (
                    "No fue posible guardar el registro.\n\n"
                    "LogFile.csv puede estar abierto en Excel "
                    "u otro programa."
                )
            )
            return False

        except OSError as error:
            messagebox.showerror(
                "Error al guardar",
                (
                    "No fue posible guardar el registro en "
                    f"LogFile.csv.\n\n{error}"
                )
            )
            return False

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
        fpy,
        estandar,
        defectos,
        top_3_defectos,
        columna
    ):
        """
        Actualiza una tarjeta existente.
        Si todavía no existe, la crea una sola vez.
        """

        if fpy >= 98:
            color_fpy = "#6FE3A1"

        elif fpy >= 95:
            color_fpy = "#FFD166"

        else:
            color_fpy = "#FF6B6B"

        texto_top = self.formatear_top_3_defectos(
            top_3_defectos
        )

        # =====================================================
        # SI YA EXISTE, SOLO ACTUALIZAR
        # =====================================================
        if modelo in self.tarjetas_modelos:

            datos_tarjeta = self.tarjetas_modelos[modelo]

            datos_tarjeta["lbl_fpy"].configure(
                text=f"{fpy:.2f} %",
                text_color=color_fpy
            )

            datos_tarjeta["lbl_totales"].configure(
                text=(
                    f"Producción: {estandar}\n"
                    f"Defectos: {defectos}"
                )
            )

            datos_tarjeta["lbl_top"].configure(
                text=texto_top
            )

            datos_tarjeta["barra_fpy"].configure(
                progress_color=color_fpy
            )

            datos_tarjeta["barra_fpy"].set(
                max(0.0, min(fpy / 100, 1.0))
            )

            datos_tarjeta["frame"].grid(
                row=0,
                column=columna,
                padx=8,
                pady=8,
                sticky="nsew"
            )

            return

        # =====================================================
        # SI NO EXISTE, CREARLA
        # =====================================================
        tarjeta = ctk.CTkFrame(
            self.frame_fpy_modelos,
            width=250,
            height=230,
            corner_radius=12,
            fg_color="#292C47",
            border_width=1,
            border_color="#454B70"
        )

        tarjeta.grid(
            row=0,
            column=columna,
            padx=8,
            pady=8,
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

        lbl_modelo.grid(
            row=0,
            column=0,
            padx=12,
            pady=(10, 2)
        )

        lbl_fpy = ctk.CTkLabel(
            tarjeta,
            text=f"{fpy:.2f} %",
            font=("Arial", 30, "bold"),
            text_color=color_fpy
        )

        lbl_fpy.grid(
            row=1,
            column=0,
            padx=12,
            pady=(2, 2)
        )

        lbl_totales = ctk.CTkLabel(
            tarjeta,
            text=(
                f"Producción: {estandar}\n"
                f"Defectos: {defectos}"
            ),
            font=("Arial", 12),
            text_color="#AEB4D0",
            justify="center"
        )

        lbl_totales.grid(
            row=2,
            column=0,
            padx=12,
            pady=(0, 5)
        )

        barra_fpy = ctk.CTkProgressBar(
            tarjeta,
            width=210,
            height=10,
            corner_radius=5,
            progress_color=color_fpy,
            fg_color="#454B70"
        )

        barra_fpy.grid(
            row=3,
            column=0,
            padx=20,
            pady=(3, 8),
            sticky="ew"
        )

        barra_fpy.set(
            max(0.0, min(fpy / 100, 1.0))
        )

        lbl_titulo_top = ctk.CTkLabel(
            tarjeta,
            text="TOP DEFECTOS",
            font=("Arial", 12, "bold"),
            text_color="#79C2FF"
        )

        lbl_titulo_top.grid(
            row=4,
            column=0,
            padx=12,
            pady=(2, 3)
        )

        lbl_top = ctk.CTkLabel(
            tarjeta,
            text=texto_top,
            font=("Arial", 11, "bold"),
            text_color="#DDE2FF",
            justify="left",
            anchor="w",
            wraplength=215
        )

        lbl_top.grid(
            row=5,
            column=0,
            padx=15,
            pady=(0, 10),
            sticky="ew"
        )

        # Guardar referencias para próximas actualizaciones
        self.tarjetas_modelos[modelo] = {
            "frame": tarjeta,
            "lbl_modelo": lbl_modelo,
            "lbl_fpy": lbl_fpy,
            "lbl_totales": lbl_totales,
            "barra_fpy": barra_fpy,
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


if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    app = RegistroDefectosSMT(root)
    root.mainloop()

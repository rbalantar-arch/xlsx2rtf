#!/usr/bin/env python3
"""
Interfaz gráfica de xlsx2rtf: convierte Excel (.xlsx / .xlsm) a RTF.

Usa el mismo motor que xlsx2rtf.py. Se empaqueta como Xlsx2RTF.exe para
Windows con build_exe.bat (o con el flujo de GitHub Actions).
"""

import ctypes
import logging
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import xlsx2rtf as core

NOMBRE_APP = "Excel a RTF"


class ColaLog(logging.Handler):
    """Envía los mensajes del logger 'xlsx2rtf' a la cola de la interfaz."""

    def __init__(self, cola):
        super().__init__()
        self.cola = cola
        self.setFormatter(logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S"))

    def emit(self, registro):
        hora = self.formatter.formatTime(registro, self.formatter.datefmt)
        prefijo = "ERROR: " if registro.levelno >= logging.ERROR else ""
        self.cola.put(("log", f"{hora}  {prefijo}{registro.getMessage()}"))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(NOMBRE_APP)
        self.minsize(640, 520)

        self.cola = queue.Queue()
        self.parar = threading.Event()
        self.ocupado = False
        self.vigilando = False

        self.var_origen = tk.StringVar()
        self.var_salida = tk.StringVar()
        self.var_horizontal = tk.BooleanVar(value=False)
        self.var_intervalo = tk.StringVar(value="5")
        self.var_estado = tk.StringVar(value="Listo")

        logger = logging.getLogger("xlsx2rtf")
        logger.setLevel(logging.INFO)
        logger.addHandler(ColaLog(self.cola))

        self._construir()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        self.after(100, self._vaciar_cola)

    # -- construcción de la ventana ---------------------------------------

    def _construir(self):
        marco = ttk.Frame(self, padding=12)
        marco.pack(fill="both", expand=True)
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(8, weight=1)

        ttk.Label(marco, text="Archivo de Excel o carpeta de origen:").grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Entry(marco, textvariable=self.var_origen).grid(
            row=1, column=0, sticky="ew", pady=(2, 8))
        ttk.Button(marco, text="Archivo...", command=self._elegir_archivo).grid(
            row=1, column=1, padx=(6, 0))
        ttk.Button(marco, text="Carpeta...", command=self._elegir_origen_carpeta).grid(
            row=1, column=2, padx=(6, 0))

        ttk.Label(marco, text="Carpeta de salida (vacío = misma carpeta del origen):").grid(
            row=2, column=0, columnspan=3, sticky="w")
        ttk.Entry(marco, textvariable=self.var_salida).grid(
            row=3, column=0, sticky="ew", pady=(2, 8))
        ttk.Button(marco, text="Examinar...", command=self._elegir_salida).grid(
            row=3, column=1, columnspan=2, sticky="ew", padx=(6, 0))

        opciones = ttk.Frame(marco)
        opciones.grid(row=4, column=0, columnspan=3, sticky="w", pady=4)
        ttk.Checkbutton(opciones, text="Orientación horizontal",
                        variable=self.var_horizontal).pack(side="left", padx=(0, 20))
        ttk.Label(opciones, text="Revisar la carpeta cada").pack(side="left")
        ttk.Spinbox(opciones, from_=1, to=3600, width=5,
                    textvariable=self.var_intervalo).pack(side="left", padx=4)
        ttk.Label(opciones, text="segundos").pack(side="left")

        botones = ttk.Frame(marco)
        botones.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 6))
        self.btn_convertir = ttk.Button(botones, text="Convertir ahora",
                                        command=self._convertir_ahora)
        self.btn_convertir.pack(side="left")
        self.btn_vigilar = ttk.Button(botones, text="Iniciar vigilancia",
                                      command=self._alternar_vigilancia)
        self.btn_vigilar.pack(side="left", padx=8)

        ttk.Label(marco, textvariable=self.var_estado, foreground="#555").grid(
            row=6, column=0, columnspan=3, sticky="w")
        ttk.Label(marco, text="Registro:").grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.registro = ScrolledText(marco, height=14, state="disabled",
                                     wrap="word", font=("Consolas", 10))
        self.registro.grid(row=8, column=0, columnspan=3, sticky="nsew")

    # -- selección de rutas -----------------------------------------------

    def _elegir_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Elige un archivo de Excel",
            filetypes=[("Libros de Excel", "*.xlsx *.xlsm")])
        if ruta:
            self.var_origen.set(ruta)

    def _elegir_origen_carpeta(self):
        ruta = filedialog.askdirectory(title="Elige la carpeta con archivos de Excel")
        if ruta:
            self.var_origen.set(ruta)

    def _elegir_salida(self):
        ruta = filedialog.askdirectory(title="Elige la carpeta de salida")
        if ruta:
            self.var_salida.set(ruta)

    def _rutas(self):
        """Devuelve (origen, salida) validadas, o None si hay un problema."""
        origen_txt = self.var_origen.get().strip().strip('"')
        if not origen_txt:
            messagebox.showwarning(NOMBRE_APP, "Indica un archivo o una carpeta de origen.")
            return None
        origen = Path(origen_txt).expanduser()
        if not origen.exists():
            messagebox.showerror(NOMBRE_APP, f"No existe:\n{origen}")
            return None
        if origen.is_file() and origen.suffix.lower() not in core.EXTENSIONES:
            messagebox.showerror(NOMBRE_APP, "El archivo debe ser .xlsx o .xlsm.")
            return None

        salida_txt = self.var_salida.get().strip().strip('"')
        if salida_txt:
            salida = Path(salida_txt).expanduser()
        else:
            salida = origen.parent if origen.is_file() else origen
        return origen.resolve(), salida.resolve()

    # -- conversión única -------------------------------------------------

    def _convertir_ahora(self):
        if self.ocupado:
            return
        rutas = self._rutas()
        if not rutas:
            return
        origen, salida = rutas
        self._ocupar(True)
        self.var_estado.set("Convirtiendo...")
        threading.Thread(target=self._trabajo,
                         args=(origen, salida, self.var_horizontal.get()),
                         daemon=True).start()

    def _trabajo(self, origen, salida, horizontal):
        """Se ejecuta en un hilo aparte para no congelar la ventana."""
        log = logging.getLogger("xlsx2rtf")
        try:
            salida.mkdir(parents=True, exist_ok=True)
            if origen.is_file():
                destino = core.convertir(origen, salida, horizontal)
                log.info("%s -> %s", origen.name, destino)
            else:
                n = core.procesar_carpeta(origen, salida, horizontal, {})
                log.info("Archivos convertidos: %d", n)
        except Exception as exc:
            log.error("%s", exc)
        finally:
            self.cola.put(("fin", None))

    # -- vigilancia continua ----------------------------------------------

    def _alternar_vigilancia(self):
        if self.vigilando:
            self.parar.set()
            self.btn_vigilar.configure(state="disabled")
            self.var_estado.set("Deteniendo vigilancia...")
            return
        if self.ocupado:
            return
        rutas = self._rutas()
        if not rutas:
            return
        origen, salida = rutas
        if not origen.is_dir():
            messagebox.showwarning(NOMBRE_APP,
                                   "Para vigilar, elige una carpeta (no un archivo).")
            return
        try:
            intervalo = float(self.var_intervalo.get().replace(",", "."))
            if intervalo < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning(NOMBRE_APP,
                                   "El intervalo debe ser un número de al menos 1 segundo.")
            return

        self.parar.clear()
        self.vigilando = True
        self._ocupar(True)
        self.btn_vigilar.configure(text="Detener vigilancia", state="normal")
        self.var_estado.set("Vigilando carpeta...")
        threading.Thread(target=self._bucle_vigilancia,
                         args=(origen, salida, self.var_horizontal.get(), intervalo),
                         daemon=True).start()

    def _bucle_vigilancia(self, carpeta, salida, horizontal, intervalo):
        log = logging.getLogger("xlsx2rtf")
        estado = {}  # marcas de cambio de cada archivo ya procesado
        try:
            salida.mkdir(parents=True, exist_ok=True)
            log.info("Vigilando %s -> %s", carpeta, salida)
            while not self.parar.is_set():
                try:
                    core.procesar_carpeta(carpeta, salida, horizontal, estado)
                except Exception as exc:
                    log.error("%s", exc)
                self.parar.wait(intervalo)
            log.info("Vigilancia detenida.")
        except Exception as exc:
            log.error("%s", exc)
        finally:
            self.cola.put(("fin", None))

    # -- estado de los controles y comunicación con el hilo ---------------

    def _ocupar(self, ocupado: bool):
        self.ocupado = ocupado
        self.btn_convertir.configure(state="disabled" if ocupado else "normal")
        self.btn_vigilar.configure(state="disabled" if ocupado and not self.vigilando else "normal")

    def _fin(self):
        self.vigilando = False
        self.ocupado = False
        self.btn_convertir.configure(state="normal")
        self.btn_vigilar.configure(text="Iniciar vigilancia", state="normal")
        self.var_estado.set("Listo")

    def _vaciar_cola(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self.registro.configure(state="normal")
                    self.registro.insert("end", dato + "\n")
                    self.registro.see("end")
                    self.registro.configure(state="disabled")
                elif tipo == "fin":
                    self._fin()
        except queue.Empty:
            pass
        self.after(100, self._vaciar_cola)

    def _cerrar(self):
        if self.vigilando and not messagebox.askokcancel(
                NOMBRE_APP, "La vigilancia está activa. ¿Cerrar la aplicación?"):
            return
        self.parar.set()
        self.destroy()


def main():
    if sys.platform == "win32":
        try:  # texto nítido en pantallas con escalado (125 %, 150 %...)
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()

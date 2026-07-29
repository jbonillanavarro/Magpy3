"""
gui_magpy3.py
─────────────
Interfaz gráfica (customtkinter) para automatizacion_magpy3_V14.py.

DEPENDENCIA NUEVA: requiere 'customtkinter' (pip install customtkinter).
La ventana ya no usa overrideredirect ni barra de título dibujada a
mano: se apoya en la gestión nativa de ventana de Tk/CTk (aparece en
la barra de tareas, responde a alt-tab, minimizar/maximizar nativos).

No modifica el script original: lo importa como módulo y sustituye
(monkeypatch) sus primitivas de interacción con el operador:

    msg, enter, cuadro, input, pedir_* , barra_progreso

por versiones que hablan con la GUI a través de dos colas:

    eventos_q   : hilo de trabajo -> GUI   (qué mostrar)
    respuesta_q : GUI -> hilo de trabajo   (qué contestó el operador)

La consola de sistema se deja funcionando en paralelo: msg()/print()
del script original siguen escribiendo ahí (log técnico), tal como
antes. La GUI es una capa adicional, no un reemplazo del log.

El core (adquisición, promediado, escritura Excel) se ejecuta sin
cambios en un hilo de trabajo (daemon) para no bloquear la ventana.
"""

import os
import sys
import glob
import queue
import threading
import importlib.util
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# ═══════════════════════════════════════════════════════════════
# CARGA DEL MÓDULO DE MEDIDA (core)
# ═══════════════════════════════════════════════════════════════
#
# El nombre del core se fija a propósito (no se "adivina" el primer
# .py que haya en la carpeta): si esta GUI cogiera cualquier script
# python al azar, un fichero de prueba o una copia antigua dejada en
# la carpeta podría acabar escribiendo mal el Excel de compliance sin
# que el operador se entere. Preferimos fallar alto y claro.
NOMBRE_CORE_ESPERADO = "automatizacion_magpy3_V14"
RUTA_CORE_ESPERADA   = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    NOMBRE_CORE_ESPERADO + ".py")


def _candidatos_automatizacion():
    """Otros scripts 'automatizacion_magpy3_*.py' presentes en la carpeta,
    para poder avisar con un mensaje útil si falta justo el esperado."""
    carpeta = os.path.dirname(os.path.abspath(__file__))
    patron = os.path.join(carpeta, "automatizacion_magpy3_*.py")
    return sorted(os.path.basename(p) for p in glob.glob(patron))


def cargar_core():
    """
    Carga automatizacion_magpy3_V14.py desde la misma carpeta que gui_magpy3.py
    (independiente del directorio de trabajo desde el que se lance 'py gui_magpy3.py').
    Si no está, lanza un error legible en vez de un ModuleNotFoundError críptico.
    """
    if not os.path.isfile(RUTA_CORE_ESPERADA):
        otros = [c for c in _candidatos_automatizacion()
                 if c != NOMBRE_CORE_ESPERADO + ".py"]
        mensaje = (
            f"No se encuentra '{NOMBRE_CORE_ESPERADO}.py' en esta carpeta:\n\n"
            f"    {os.path.dirname(RUTA_CORE_ESPERADA)}\n\n"
            f"Copia ahí el archivo '{NOMBRE_CORE_ESPERADO}.py' (con ese nombre exacto)."
        )
        if otros:
            mensaje += (
                f"\n\nSe ha encontrado en su lugar: {', '.join(otros)}.\n"
                f"Esa versión puede no tener las últimas correcciones "
                f"(p. ej. la clasificación RMS/Peak en la hoja ISED). "
                f"No se carga automáticamente para no arriesgar una medida "
                f"de compliance con el script equivocado."
            )
        raise FileNotFoundError(mensaje)

    spec = importlib.util.spec_from_file_location(NOMBRE_CORE_ESPERADO, RUTA_CORE_ESPERADA)
    modulo = importlib.util.module_from_spec(spec)
    sys.modules[NOMBRE_CORE_ESPERADO] = modulo
    spec.loader.exec_module(modulo)
    return modulo


try:
    core = cargar_core()
except FileNotFoundError as e:
    # Tkinter aún no tiene ventana principal creada; usamos un diálogo
    # mínimo para mostrar el error igualmente, con consola como respaldo.
    print(f"\n[ERROR] {e}\n")
    try:
        _root_error = tk.Tk()
        _root_error.withdraw()
        messagebox.showerror("No se encuentra el script de medida", str(e))
        _root_error.destroy()
    except Exception:
        pass
    sys.exit(1)

import config_store


# ═══════════════════════════════════════════════════════════════
# COMUNICACIÓN HILO DE TRABAJO <-> GUI
# ═══════════════════════════════════════════════════════════════

eventos_q   = queue.Queue()   # (tipo, payload) generados por el hilo de trabajo
respuesta_q = queue.Queue()   # respuestas que la GUI deposita para el hilo


def _pedir_respuesta():
    """Bloquea el hilo de trabajo hasta que la GUI deposite una respuesta."""
    return respuesta_q.get()


# ── Sustitutos de las primitivas de consola ───────────────────────────────

def gui_msg(texto, tipo="INFO"):
    # Mantenemos el log real en consola (comportamiento original)...
    core._msg_original(texto, tipo)
    # ...y además lo reflejamos en el panel de log de la GUI.
    eventos_q.put(("LOG", (texto, tipo)))


def gui_enter(texto="Presiona Continuar para seguir..."):
    """Sustituye a enter(): panel fijo con instrucción + botón Continuar."""
    eventos_q.put(("PASO_MANUAL", {"titulo": "Acción requerida", "lineas": [texto],
                                   "opciones": ["Continuar"]}))
    _pedir_respuesta()


def gui_cuadro(titulo, lineas):
    """Sustituye a cuadro(): solo pinta, no espera (igual que el original)."""
    eventos_q.put(("CUADRO", {"titulo": titulo, "lineas": lineas}))


def gui_input(prompt=""):
    """
    Sustituye al input() builtin usado para los menús B/R/Q y A/ENTER.
    Detecta el tipo de menú por el contenido del prompt reciente (lo
    resuelven las funciones pedir_* de más abajo, que llaman a esto
    indirectamente vía input_config y los puntos sueltos de 'accion = input(...)').
    """
    eventos_q.put(("PASO_MANUAL", {"titulo": "Confirmación", "lineas": [prompt.strip()],
                                   "opciones": ["Continuar"]}))
    _pedir_respuesta()
    return ""  # equivalente a pulsar ENTER


def gui_barra_progreso(segundos_total, etiqueta="Midiendo"):
    """Sustituye a barra_progreso(): progreso real por tiempo, sin bloquear la GUI."""
    import time
    segundos_total = max(1, int(segundos_total))
    inicio = time.monotonic()
    eventos_q.put(("PROGRESO_INICIO", {"etiqueta": etiqueta, "segundos_total": segundos_total}))
    while True:
        transcurrido = time.monotonic() - inicio
        pct = min(100, int(transcurrido / segundos_total * 100))
        restante = max(0, int(round(segundos_total - transcurrido)))
        eventos_q.put(("PROGRESO", {"pct": pct, "restante_s": restante}))
        if transcurrido >= segundos_total:
            break
        time.sleep(0.2)
    eventos_q.put(("PROGRESO_FIN", {}))


def gui_beep(n=1):
    # Sin hardware de sonido garantizado en todos los puestos; lo hacemos
    # best-effort y no rompemos la sesión si falla.
    try:
        core._beep_original(n)
    except Exception:
        pass


def gui_getpass_excel(prompt=""):
    """
    Sustituye a getpass.getpass() dentro de _abrir_excel_protegido().
    Pide la contraseña del Excel con un popup (campo enmascarado) en vez
    de por consola. Bloquea el hilo de trabajo hasta que el operador
    confirme o cancele.
    """
    eventos_q.put(("PEDIR_PASSWORD_EXCEL", {}))
    resultado = _pedir_respuesta()   # dict {"password": str, "recordar": bool} o None si canceló
    if resultado is None:
        # El operador canceló el diálogo: devolvemos cadena vacía, que
        # msoffcrypto rechazará como contraseña incorrecta de forma
        # controlada (el core ya sabe reintentar/avisar en ese caso).
        return ""
    if resultado.get("recordar"):
        config_store.guardar_password_excel(resultado["password"])
    return resultado["password"]


def gui_pedir_numero(texto, mn=0.0, mx=999.0, actual=None):
    """
    Sustituye a pedir_numero() durante el flujo de medida (se usa para
    pedir la distancia PEAK en ISED). La versión original llama a
    input_config()/input(), que en la GUI queda interceptado por
    gui_input() -- pensado para menús ENTER/A, no para valores numéricos
    -- y provocaba que este paso se saltara en silencio. Aquí se pide con
    un panel/diálogo real y no se continúa hasta tener un valor válido.
    """
    eventos_q.put(("PEDIR_NUMERO", {"texto": texto, "mn": mn, "mx": mx, "actual": actual}))
    return _pedir_respuesta()   # float ya validado en rango por la GUI


# ═══════════════════════════════════════════════════════════════
# APLICACIÓN DE LOS MONKEYPATCHES SOBRE EL MÓDULO core
# ═══════════════════════════════════════════════════════════════

def instalar_monkeypatches():
    # Guardamos los originales por si se quieren restaurar o para el beep.
    core._msg_original  = core.msg
    core._beep_original = core.beep

    core.msg            = gui_msg
    core.enter           = gui_enter
    core.cuadro           = gui_cuadro
    core.barra_progreso  = gui_barra_progreso
    core.beep            = gui_beep

    # El builtin input() se usa en tres puntos del flujo de medida
    # (main, iniciar_medida_con_reintentos) para menús ENTER/A.
    # Se sustituye a nivel de módulo 'core' para no tocar otros usos de input().
    core.input           = gui_input

    # getpass.getpass() se usa solo dentro de _abrir_excel_protegido().
    # Sustituimos el método sobre el objeto 'getpass' que el core importó,
    # así no afecta a ningún otro uso de getpass en el proceso.
    core.getpass.getpass = gui_getpass_excel

    # pedir_numero() se usa durante procesar_posicion() para la distancia
    # PEAK (ISED). Sin este patch caía en input_config()/input() y el
    # paso se saltaba en silencio (ver gui_pedir_numero para el detalle).
    core.pedir_numero = gui_pedir_numero

    # Si ya hay una contraseña guardada (DPAPI) de una sesión anterior,
    # la precargamos: _abrir_excel_protegido() solo pide contraseña
    # cuando _EXCEL_PASSWORD es falsy, así que si esto tiene éxito no
    # se mostrará ningún popup al abrir el Excel.
    password_guardada = config_store.cargar_password_excel()
    if password_guardada:
        core._EXCEL_PASSWORD = password_guardada


# ═══════════════════════════════════════════════════════════════
# PASOS DE CONFIGURACIÓN "EN UNA PANTALLA"
# En vez de llamar a pedir_configuracion_sesion() (paso a paso por CLI),
# la GUI construye directamente el dict de configuración y valida con
# las mismas reglas que las funciones pedir_* originales.
# ═══════════════════════════════════════════════════════════════

class ErrorValidacion(Exception):
    pass


def validar_config_formulario(datos):
    """
    Reaplica las validaciones que hacían pedir_entero/pedir_numero/pedir_frecuencia
    en la CLI, para que la GUI no permita continuar con datos fuera de rango.
    datos: dict con claves crudas del formulario (strings).
    Devuelve el dict de configuración ya tipado, o lanza ErrorValidacion.
    """
    errores = []

    regiones = list(datos.get("regiones", []))
    if not regiones:
        errores.append("Selecciona al menos una región (CE, FCC, ISED).")

    carpeta_raiz = datos.get("carpeta_raiz", "").strip()
    if not carpeta_raiz:
        errores.append("Indica la carpeta raíz de resultados.")

    ruta_excel = datos.get("ruta_excel", "").strip() or None
    if ruta_excel and not os.path.isfile(ruta_excel):
        errores.append(f"El archivo Excel no existe: {ruta_excel}")

    def _to_int(valor, campo, mn, mx):
        try:
            v = int(valor)
        except (TypeError, ValueError):
            errores.append(f"{campo}: introduce un número entero.")
            return None
        if not (mn <= v <= mx):
            errores.append(f"{campo}: debe estar entre {mn} y {mx}.")
            return None
        return v

    def _to_float(valor, campo, mn, mx):
        try:
            v = float(str(valor).replace(",", "."))
        except (TypeError, ValueError):
            errores.append(f"{campo}: introduce un número.")
            return None
        if not (mn <= v <= mx):
            errores.append(f"{campo}: debe estar entre {mn} y {mx}.")
            return None
        return v

    n_posiciones = _to_int(datos.get("n_posiciones"), "Número de posiciones", 1, 50)
    distancia_cm = _to_float(datos.get("distancia_cm"), "Distancia (cm)", 0.0, 200.0)
    frecuencia   = _to_float(datos.get("frecuencia"), "Frecuencia (kHz)", 3, 10000)

    technology = datos.get("technology", "").strip() or "N/A"
    wpt_client = datos.get("wpt_client", "").strip() or "N/A"
    battery    = datos.get("battery", "").strip() or "N/A"

    if errores:
        raise ErrorValidacion("\n".join(f"• {e}" for e in errores))

    if carpeta_raiz:
        try:
            os.makedirs(carpeta_raiz, exist_ok=True)
        except Exception as e:
            raise ErrorValidacion(f"No se pudo crear/acceder a la carpeta raíz: {e}")

    ts_sesion        = datetime.now().strftime("%Y%m%d_%H%M%S")
    archivo_csv      = os.path.join(carpeta_raiz, f"promedios_{ts_sesion}.csv")
    archivo_csv_peak = os.path.join(carpeta_raiz, f"peak_{ts_sesion}.csv")
    plan             = core.calcular_plan_medida(regiones)
    nombres_regiones = ", ".join(core.CONFIG_REGION[r]["nombre"] for r in regiones)

    return {
        "regiones": regiones,
        "carpeta_raiz": carpeta_raiz,
        "ruta_excel": ruta_excel,
        "n_posiciones": n_posiciones,
        "distancia_cm": distancia_cm,
        "frecuencia": frecuencia,
        "technology": technology,
        "wpt_client": wpt_client,
        "battery": battery,
        "archivo_csv": archivo_csv,
        "archivo_csv_peak": archivo_csv_peak,
        "plan": plan,
        "nombres_regiones": nombres_regiones,
    }


# ═══════════════════════════════════════════════════════════════
# HILO DE TRABAJO: replica main() del script original, pero sin
# pedir_configuracion_sesion() (esa parte ya la resolvió la GUI).
# ═══════════════════════════════════════════════════════════════

def hilo_sesion_medida(config):
    try:
        eventos_q.put(("FASE", "Comprobando conexión con MAGpy3..."))
        if not core.verificar_conexion():
            eventos_q.put(("ERROR_FATAL", "No hay conexión con MAGpy3. Ábrelo y verifica que el backend esté activo."))
            return

        regiones         = config["regiones"]
        carpeta_raiz     = config["carpeta_raiz"]
        ruta_excel       = config["ruta_excel"]
        n_posiciones     = config["n_posiciones"]
        distancia_cm     = config["distancia_cm"]
        frecuencia       = config["frecuencia"]
        technology       = config["technology"]
        wpt_client       = config["wpt_client"]
        battery          = config["battery"]
        archivo_csv      = config["archivo_csv"]
        archivo_csv_peak = config["archivo_csv_peak"]
        nombres_regiones = config["nombres_regiones"]

        if not core.validar_configuracion_magpy(regiones[0]):
            eventos_q.put(("ERROR_FATAL", "Corrige la configuración en MAGpy3 (Standard/Environment/Health effect/Peak-RMS) y vuelve a intentarlo."))
            return

        eventos_q.put(("FASE", "Configurando adquisición vía API..."))
        core.configurar_adquisicion(frecuencia)
        core.time.sleep(1)

        eventos_q.put(("FASE", "Medición en curso"))
        t_inicio = datetime.now()
        resultados, resultados_peak = [], []

        for i in range(1, n_posiciones + 1):
            eventos_q.put(("PEDIR_NOMBRE_POSICION", {"num": i, "total": n_posiciones}))
            nombre_pos = respuesta_q.get()
            if nombre_pos is None:   # abortado desde la GUI
                eventos_q.put(("SESION_ABORTADA", None))
                core.guardar_csv_resumen(resultados, archivo_csv)
                if resultados_peak:
                    core.guardar_csv_resumen(resultados_peak, archivo_csv_peak)
                return

            while True:
                carpeta_pos, promedios, promedios_peak, distancia_peak_usada = core.procesar_posicion(
                    nombre_pos, carpeta_raiz, regiones,
                    distancia_cm, frecuencia, ruta_excel,
                    technology, wpt_client, battery
                )
                if promedios:
                    resultados.append({
                        "posicion": nombre_pos, "regiones": "+".join(regiones),
                        "distancia_cm": distancia_cm, "frecuencia_khz": frecuencia,
                        "freq_mhz": promedios.get("freq_mhz"),
                        "h_prom_A_m": promedios.get("h_prom"),
                        "e_prom_V_m": promedios.get("e_prom"),
                    })
                    if promedios_peak:
                        resultados_peak.append({
                            "posicion": nombre_pos, "regiones": "+".join(regiones),
                            "distancia_cm": distancia_peak_usada, "frecuencia_khz": frecuencia,
                            "freq_mhz": promedios_peak.get("freq_mhz"),
                            "h_prom_A_m": promedios_peak.get("h_prom"),
                            "e_prom_V_m": promedios_peak.get("e_prom"),
                        })
                    break

                eventos_q.put(("POSICION_FALLIDA", {"nombre": nombre_pos}))
                accion = respuesta_q.get()  # "REPETIR" o "ABORTAR"
                if accion == "ABORTAR":
                    core.msg("Sesión abortada. No se continuará con la siguiente posición.", "ERROR")
                    core.guardar_csv_resumen(resultados, archivo_csv)
                    if resultados_peak:
                        core.guardar_csv_resumen(resultados_peak, archivo_csv_peak)
                    eventos_q.put(("SESION_ABORTADA", None))
                    return

        core.guardar_csv_resumen(resultados, archivo_csv)
        if resultados_peak:
            core.guardar_csv_resumen(resultados_peak, archivo_csv_peak)

        t_fin = datetime.now()
        dur = t_fin - t_inicio
        h, m = int(dur.total_seconds() // 3600), int((dur.total_seconds() % 3600) // 60)

        eventos_q.put(("SESION_COMPLETADA", {
            "n_posiciones": n_posiciones, "nombres_regiones": nombres_regiones,
            "frecuencia": frecuencia, "t_inicio": t_inicio, "t_fin": t_fin,
            "h": h, "m": m, "carpeta_raiz": carpeta_raiz,
            "archivo_csv": archivo_csv,
            "archivo_csv_peak": archivo_csv_peak if resultados_peak else None,
            "ruta_excel": ruta_excel,
        }))

    except Exception as e:
        eventos_q.put(("ERROR_FATAL", f"Error inesperado: {type(e).__name__}: {e}"))

import customtkinter as ctk

ctk.set_appearance_mode("dark")
# ctk.deactivate_automatic_dpi_awareness()

# ═══════════════════════════════════════════════════════════════
# THEME TOKENS — paleta exacta pedida (no aproximaciones)
# ═══════════════════════════════════════════════════════════════

FUENTE_FAMILIA = "Segoe UI"

FUENTE_TITULO_MODAL = (FUENTE_FAMILIA, 24)
FUENTE_SECCION       = (FUENTE_FAMILIA, 19)     # sin bold, según referencia
FUENTE_SECCION_TITULO = (FUENTE_FAMILIA, 19)    # títulos de sección: mismo tamaño, sin bold
FUENTE_LABEL         = (FUENTE_FAMILIA, 16)
FUENTE_VALOR         = (FUENTE_FAMILIA, 16)     # sin bold, según referencia
FUENTE_LOG           = ("Consolas", 13)

COLORS = {
    "bg_modal":       "#131313",   # fondo general de la ventana
    "bg_header":      "#383838",   # cabecera/título de cada tarjeta
    "bg_input":       "#303030",   # cuerpo de la tarjeta / campos de valor
    "border":         "#3d3d3d",   # borde fino que envuelve tarjetas y celdas
    "text_primary":   "#ffffff",   # texto en blanco, sin bold
    "text_secondary": "#8a8d93",
    "accent_blue":    "#6ec3df",   # ya no se usa en títulos de sección; queda disponible por si hace falta en otro sitio
    "accent_yellow":  "#fdef5d",   # amarillo de acento: texto del botón principal y hover de checkboxes
    "separator":      "#3d3d3d",
}

BORDE_GROSOR = 2   # grosor del borde de tarjetas/celdas (antes 1px)

TOKEN_BG_MODAL       = COLORS["bg_modal"]
TOKEN_BG_HEADER      = COLORS["bg_header"]
TOKEN_BG_INPUT       = COLORS["bg_input"]
TOKEN_BG_INPUT_HOVER = "#33363c"
TOKEN_BORDER_SUBTLE  = COLORS["border"]
TOKEN_TEXT_PRIMARY   = COLORS["text_primary"]
TOKEN_TEXT_SECONDARY = COLORS["text_secondary"]
TOKEN_ACCENT_BLUE    = COLORS["accent_blue"]
TOKEN_ACCENT_YELLOW  = COLORS["accent_yellow"]
TOKEN_SEPARATOR      = COLORS["separator"]
TOKEN_DOT_STATUS     = "#c96b6b"

# Botón principal ("Iniciar sesión de medida"): fondo gris oscuro,
# texto amarillo — ya no es el pill azul de antes.
TOKEN_ACCION_FONDO       = "#323232"
TOKEN_ACCION_FONDO_HOVER = "#3a3a3a"
TOKEN_ACCION_TEXTO       = TOKEN_ACCENT_YELLOW

TOKEN_ACCENT_ACTION       = TOKEN_ACCENT_BLUE
TOKEN_ACCENT_ACTION_HOVER = "#4db0e6"
TOKEN_ON_ACCENT            = "#0e1013"

COLOR_OK    = "#8fbf8f"
COLOR_AVISO = "#d9b25a"
COLOR_ERROR = TOKEN_DOT_STATUS
COLOR_INFO  = TOKEN_TEXT_PRIMARY

MAGPY_NEGRO           = TOKEN_BG_MODAL
MAGPY_GRIS_TARJETA     = TOKEN_BG_HEADER
MAGPY_GRIS_CAMPO       = TOKEN_BG_INPUT
MAGPY_GRIS_BORDE       = TOKEN_BORDER_SUBTLE
MAGPY_BLANCO           = TOKEN_TEXT_PRIMARY
MAGPY_GRIS_TEXTO       = TOKEN_TEXT_SECONDARY
MAGPY_AMARILLO         = TOKEN_ACCENT_ACTION
FUENTE_TITULO = FUENTE_SECCION
FUENTE_NORMAL = FUENTE_LABEL
FUENTE_BOTON  = FUENTE_VALOR

RADIO_PILL = 8

ctk.ThemeManager.theme["CTkFrame"]["fg_color"] = [TOKEN_BG_MODAL, TOKEN_BG_MODAL]


class _BotonPrimario(ctk.CTkFrame):
    """
    Botón custom para los casos 'primario=True'. CTkButton en
    customtkinter 6.0.0 tiene un bug de renderizado que ignora el
    fg_color pasado (tanto en el constructor como en .configure()
    posterior) y siempre pinta el azul del tema por defecto. CTkFrame +
    CTkLabel no pasan por ese mismo pipeline de color y sí respetan el
    fg_color exacto, así que replicamos aquí el comportamiento de un
    botón (click, hover, estado disabled) a mano.
    """
    def __init__(self, parent, texto, command=None, alto=44, ancho=None,
                 font=None, fg_color=None):
        fg = fg_color or TOKEN_ACCION_FONDO
        hover = TOKEN_ACCION_FONDO_HOVER
        super().__init__(
            parent, height=alto, width=ancho or 0, fg_color=fg,
            corner_radius=RADIO_PILL, border_width=BORDE_GROSOR, border_color=TOKEN_BORDER_SUBTLE,
        )
        # Si se pide un ancho explícito, se respeta (no se deja encoger);
        # si no, el frame se ajusta al contenido del CTkLabel interno,
        # como un botón normal con padding horizontal.
        if ancho:
            self.pack_propagate(False)
        self._fg = fg
        self._hover = hover
        self._command = command
        self._enabled = True

        self.lbl = ctk.CTkLabel(
            self, text=texto, text_color=TOKEN_ACCION_TEXTO,
            font=font or FUENTE_VALOR, fg_color="transparent",
        )
        self.lbl.pack(expand=True, fill="both", padx=20, pady=4)

        # Bug de customtkinter 6.0.0: en algunos casos el ThemeManager
        # reaplica el color por defecto DESPUÉS de que el widget ya se
        # ha construido con fg_color explícito (al montarse en pantalla).
        # Forzamos el color de nuevo una vez el bucle de eventos ha
        # terminado el ciclo de creación, para que gane el nuestro.
        self.after_idle(self._forzar_color_inicial)

        for widget in (self, self.lbl):
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
        self.configure(cursor="hand2")
        self.lbl.configure(cursor="hand2")

    def _forzar_color_inicial(self):
        if self._enabled:
            super().configure(fg_color=self._fg)

    def _on_click(self, event=None):
        if self._enabled and self._command:
            self._command()

    def _on_enter(self, event=None):
        if self._enabled:
            super().configure(fg_color=self._hover)

    def _on_leave(self, event=None):
        if self._enabled:
            super().configure(fg_color=self._fg)

    def configure(self, **kwargs):
        # Soporta el mismo patrón que ya usa el resto del código con
        # CTkButton: btn.configure(state="disabled"/"normal", text=..., command=...)
        if "state" in kwargs:
            estado = kwargs.pop("state")
            self._enabled = (estado == "normal")
            color = self._fg if self._enabled else TOKEN_BG_INPUT
            super().configure(fg_color=color,
                              cursor="hand2" if self._enabled else "arrow")
            self.lbl.configure(text_color=TOKEN_ACCION_TEXTO if self._enabled else TOKEN_TEXT_SECONDARY)
        if "text" in kwargs:
            self.lbl.configure(text=kwargs.pop("text"))
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if kwargs:
            super().configure(**kwargs)


def _boton(parent, texto, command=None, primario=False, alto=44, ancho=None,
          font=None, fg_color=None):
    if primario:
        return _BotonPrimario(parent, texto, command=command, alto=alto,
                              ancho=ancho, font=font, fg_color=fg_color)
    fg = fg_color or TOKEN_BG_INPUT
    hover = TOKEN_BG_INPUT_HOVER
    txt = TOKEN_TEXT_PRIMARY
    borde = TOKEN_BORDER_SUBTLE
    btn = ctk.CTkButton(
        parent, text=texto, command=command, height=alto, width=ancho or 0,
        corner_radius=RADIO_PILL, fg_color=fg, hover_color=hover,
        text_color=txt, border_width=BORDE_GROSOR, border_color=borde,
        font=font or FUENTE_VALOR,
    )
    btn.configure(fg_color=fg, hover_color=hover, text_color=txt)
    return btn

def _entrada(parent, textvariable, width=220, show=None, justify="left"):
    return ctk.CTkEntry(
        parent, textvariable=textvariable, width=width, height=38,
        corner_radius=RADIO_PILL, fg_color=TOKEN_BG_INPUT,
        border_width=BORDE_GROSOR, border_color=TOKEN_BORDER_SUBTLE,
        text_color=TOKEN_TEXT_PRIMARY, font=FUENTE_VALOR,
        show=show or "", justify=justify,
    )


def _label(parent, texto, secundario=False, seccion=False, wraplength=None, font=None, **kw):
    color = TOKEN_TEXT_PRIMARY if seccion else (TOKEN_TEXT_SECONDARY if secundario else TOKEN_TEXT_PRIMARY)
    fuente = font or (FUENTE_SECCION_TITULO if seccion else FUENTE_LABEL)
    return ctk.CTkLabel(parent, text=texto, text_color=color, font=fuente,
                        wraplength=wraplength or 0, justify="left", **kw)


class CheckboxRegion(ctk.CTkCheckBox):
    def __init__(self, parent, variable, texto):
        super().__init__(
            parent, text=texto, variable=variable, font=FUENTE_LABEL,
            text_color=TOKEN_TEXT_PRIMARY, fg_color=TOKEN_ACCENT_YELLOW,
            hover_color=TOKEN_ACCENT_YELLOW, checkmark_color=TOKEN_BG_MODAL,
            border_color=TOKEN_BORDER_SUBTLE, border_width=BORDE_GROSOR,
            corner_radius=6, checkbox_width=28, checkbox_height=28,
        )


class SeccionCard(ctk.CTkFrame):
    """
    Tarjeta de sección con cabecera diferenciada (estilo 'Peak Search'):
    fila de título con su propio fondo (bg_header), separada por una
    línea del cuerpo (bg_input) donde van las filas de valores.

    El frame exterior lleva corner_radius=10 + border_width nativo.
    header/cuerpo (rectos, corner_radius=0) se dejan con un margen
    igual al grosor del borde en vez de llenar el ancho a tope: así el
    radio del exterior queda siempre visible en las 4 esquinas (si los
    sub-frames llenan el ancho completo, sus esquinas rectas tapan el
    redondeo del padre y la tarjeta se ve a pico).
    """
    def __init__(self, parent, titulo):
        super().__init__(parent, fg_color=TOKEN_BG_INPUT, corner_radius=10,
                         border_width=BORDE_GROSOR, border_color=TOKEN_BORDER_SUBTLE)

        m = BORDE_GROSOR  # margen = grosor del borde, para dejar ver el redondeo exterior

        self.header = ctk.CTkFrame(self, fg_color=TOKEN_BG_HEADER, corner_radius=0)
        self.header.pack(fill="x", padx=m, pady=(m, 0))
        self.header.columnconfigure(1, weight=1)
        _label(self.header, titulo, seccion=True).grid(
            row=0, column=0, sticky="w", padx=24, pady=12)

        self.separador = ctk.CTkFrame(self, fg_color=TOKEN_BORDER_SUBTLE, height=BORDE_GROSOR, corner_radius=0)
        self.separador.pack(fill="x", padx=m)

        self.cuerpo = ctk.CTkFrame(self, fg_color=TOKEN_BG_INPUT, corner_radius=0)
        self.cuerpo.pack(fill="both", expand=True, padx=m, pady=(0, m))
        self.cuerpo.columnconfigure(1, weight=1)
        self._fila = 0

    def header_extra(self, widget_factory):
        """
        Añade contenido a la derecha del título, dentro de la misma
        cabecera (p.ej. 'Dejar en blanco = N/A'). Usa grid dentro de
        self.header, que es un contenedor propio y no comparte gestor
        de geometría con el resto del SeccionCard (que usa pack).
        """
        widget = widget_factory(self.header)
        widget.grid(row=0, column=1, sticky="e", padx=24, pady=14)
        return widget

    def fila(self, texto_label, widget_factory, ayuda=None):
        _label(self.cuerpo, texto_label).grid(
            row=self._fila, column=0, sticky="w", padx=(24, 16), pady=6)
        celda = ctk.CTkFrame(self.cuerpo, fg_color="transparent")
        celda.grid(row=self._fila, column=1, sticky="we", padx=(0, 24), pady=6)
        celda.columnconfigure(0, weight=1)
        widget = widget_factory(celda)
        widget.grid(row=0, column=0, sticky="w")
        if ayuda:
            _label(celda, ayuda, secundario=True).grid(row=0, column=1, sticky="w", padx=(12, 0))
        self._fila += 1
        return widget

    def fila_libre(self, widget_factory):
        celda = ctk.CTkFrame(self.cuerpo, fg_color="transparent")
        celda.grid(row=self._fila, column=0, columnspan=2, sticky="we", padx=24, pady=6)
        widget = widget_factory(celda)
        widget.pack(fill="x", anchor="w")
        self._fila += 1
        return widget


class ScrollArea(ctk.CTkScrollableFrame):
    """
    Contenedor con scrollbar vertical fina, tema oscuro, en línea con
    el resto de la paleta (sustituye al patrón manual Canvas+Frame+
    Scrollbar; CTkScrollableFrame ya lo resuelve internamente).
    """
    def __init__(self, parent):
        super().__init__(parent, fg_color=TOKEN_BG_MODAL,
                         scrollbar_button_color=TOKEN_BG_INPUT,
                         scrollbar_button_hover_color=TOKEN_BG_INPUT_HOVER)


class DialogoModal(ctk.CTkToplevel):
    """
    Toplevel modal estilo 'Settings': barra de título propia en
    TOKEN_BG_HEADER, título a la izquierda. NO usa overrideredirect: se
    deja la gestión de ventana nativa (foco, cierre y superposición
    correctos) y solo se restylea el contenido.
    """
    def __init__(self, master, titulo, ancho=460):
        super().__init__(master)
        self.configure(fg_color=TOKEN_BG_MODAL)
        self.title(titulo)
        self.resizable(False, False)
        self.transient(master)
        self.attributes("-topmost", True)
        self._ancho_fijo = ancho   # guardado aparte: winfo_width() puede reportar
                                    # un valor ya encogido por Tk antes de ajustar_alto()

        header = ctk.CTkFrame(self, fg_color=TOKEN_BG_HEADER, corner_radius=0, height=56)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        _label(header, titulo).pack(side="left", padx=16)

        self.contenido = ctk.CTkFrame(self, fg_color=TOKEN_BG_MODAL, corner_radius=0,
                                      width=ancho)
        self.contenido.pack(fill="both", expand=True, padx=20, pady=18)
        self.contenido.pack_propagate(False)  # el ancho no lo decide el contenido interno

        self.update_idletasks()
        x = master.winfo_x() + (master.winfo_width() - ancho) // 2
        y = master.winfo_y() + 80
        self.geometry(f"{ancho}x1+{max(0,x)}+{max(0,y)}")

    def ajustar_alto(self):
        self.contenido.pack_propagate(True)  # ya con todo el contenido creado, medir alto real
        self.update_idletasks()
        alto = self.winfo_reqheight()
        self.geometry(f"{self._ancho_fijo}x{alto}")


# ═══════════════════════════════════════════════════════════════
# VENTANA PRINCIPAL
# ═══════════════════════════════════════════════════════════════

class AppMagpy3(ctk.CTk):
    def __init__(self):
        super().__init__()
        # CTk gestiona la ventana con el gestor nativo del SO: aparece
        # en la barra de tareas, responde a alt-tab, minimizar y
        # maximizar funcionan de forma estándar. No se usa
        # overrideredirect en ningún punto de la app.
        self.title("MAGpy3 — Automatización de medidas EM (DEKRA)")
        self.geometry("1280x1080")
        self.minsize(1000, 700)
        self.configure(fg_color=TOKEN_BG_MODAL)

        self.hilo = None
        self.esperando_nombre_posicion = False

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._construir_pantalla_formulario()
        self._construir_pantalla_sesion()

        self._precargar_perfil()
        self.frame_formulario.grid(row=0, column=0, sticky="nsew")

        self.after(150, self._procesar_eventos)

    # ── Pantalla 1: formulario de configuración ─────────────────────────

    def _construir_pantalla_formulario(self):
        scroll = ScrollArea(self)
        self.frame_formulario = scroll
        scroll.grid_columnconfigure(0, weight=1)

        cont = ctk.CTkFrame(scroll, fg_color="transparent")
        cont.pack(fill="both", expand=True, padx=24, pady=16)
        cont.columnconfigure(0, weight=1)

        _label(cont, "Configuración de la sesión de medida", seccion=True,
              font=FUENTE_TITULO_MODAL).grid(row=0, column=0, sticky="w", pady=(0, 14))

        fila = 1

        card_regiones = SeccionCard(cont, "Regiones")
        card_regiones.grid(row=fila, column=0, sticky="we", pady=(0, 16))
        self.vars_regiones = {}

        def _crear_checks(celda):
            wrapper = ctk.CTkFrame(celda, fg_color="transparent")
            for k, v in core.CONFIG_REGION.items():
                fases_str = " + ".join(fs["etiqueta"] for fs in v["fases"])
                var = ctk.BooleanVar(value=False)
                self.vars_regiones[k] = var
                CheckboxRegion(wrapper, var, f"{k}  —  {v['nombre']}  ({fases_str})")\
                    .pack(anchor="w", pady=9)
            return wrapper

        card_regiones.fila_libre(_crear_checks)
        fila += 1

        card_rutas = SeccionCard(cont, "Configuración")
        card_rutas.grid(row=fila, column=0, sticky="we", pady=(0, 16))

        self.var_carpeta = ctk.StringVar()
        def _fila_carpeta(celda):
            wrap = ctk.CTkFrame(celda, fg_color="transparent")
            wrap.columnconfigure(0, weight=1)
            e = _entrada(wrap, self.var_carpeta, width=380)
            e.grid(row=0, column=0, sticky="we")
            _boton(wrap, "Examinar…", command=self._elegir_carpeta, alto=48)\
                .grid(row=0, column=1, padx=(8, 0))
            return wrap
        card_rutas.fila("Carpeta raíz de resultados", _fila_carpeta)

        self.var_excel = ctk.StringVar()
        def _fila_excel(celda):
            wrap = ctk.CTkFrame(celda, fg_color="transparent")
            wrap.columnconfigure(0, weight=1)
            e = _entrada(wrap, self.var_excel, width=380)
            e.grid(row=0, column=0, sticky="we")
            _boton(wrap, "Examinar…", command=self._elegir_excel, alto=48)\
                .grid(row=0, column=1, padx=(8, 0))
            return wrap
        card_rutas.fila("Archivo Excel (.xlsm/.xlsx)", _fila_excel)

        self.var_n_posiciones = ctk.StringVar(value="1")
        card_rutas.fila("Número de posiciones",
                       lambda celda: _entrada(celda, self.var_n_posiciones, width=120, justify="center"))

        self.var_distancia = ctk.StringVar(value="")
        card_rutas.fila("Distancia de medida (cm)",
                       lambda celda: _entrada(celda, self.var_distancia, width=120, justify="center"))

        self.var_frecuencia = ctk.StringVar(value="")
        card_rutas.fila("Frecuencia de búsqueda de pico (kHz)",
                       lambda celda: _entrada(celda, self.var_frecuencia, width=140, justify="center"),
                       ayuda="Rango 3–10000 kHz")

        fila += 1

        card_excel = SeccionCard(cont, "Datos para el Excel")
        card_excel.grid(row=fila, column=0, sticky="we", pady=(0, 16))
        card_excel.header_extra(
            lambda h: _label(h, "Dejar en blanco = N/A", secundario=True))

        self.var_technology = ctk.StringVar()
        card_excel.fila("Technology", lambda celda: _entrada(celda, self.var_technology, width=260))

        self.var_wpt = ctk.StringVar()
        card_excel.fila("WPT Client", lambda celda: _entrada(celda, self.var_wpt, width=260))

        self.var_battery = ctk.StringVar()
        card_excel.fila("Battery level", lambda celda: _entrada(celda, self.var_battery, width=260))

        fila += 1

        self.lbl_debug_fast = _label(cont, "", secundario=True, wraplength=900)
        self.lbl_debug_fast.configure(text_color=COLOR_AVISO)
        self.lbl_debug_fast.grid(row=fila, column=0, sticky="w", pady=(4, 0))
        if core.DEBUG_FAST:
            self.lbl_debug_fast.configure(
                text="⚠ MODO DEPURACIÓN RÁPIDA ACTIVO (MAGPY_DEBUG_FAST=1). "
                     "Los valores resultantes NO son válidos para compliance."
            )
        fila += 1

        self.lbl_error_form = _label(cont, "", wraplength=900)
        self.lbl_error_form.configure(text_color=COLOR_ERROR)
        self.lbl_error_form.grid(row=fila, column=0, sticky="w", pady=(8, 0))
        fila += 1

        frame_accion = ctk.CTkFrame(cont, fg_color="transparent")
        frame_accion.grid(row=fila, column=0, sticky="e", pady=(20, 0))
        _boton(frame_accion, "Iniciar sesión de medida →", command=self._on_iniciar_sesion,
              primario=True, alto=48, font=FUENTE_SECCION).pack()

    def _elegir_carpeta(self):
        ruta = filedialog.askdirectory(title="Carpeta raíz de resultados")
        if ruta:
            self.var_carpeta.set(ruta)

    def _elegir_excel(self):
        ruta = filedialog.askopenfilename(
            title="Archivo Excel de resultados",
            filetypes=[("Excel", "*.xlsm *.xlsx"), ("Todos", "*.*")]
        )
        if ruta:
            self.var_excel.set(ruta)

    def _leer_datos_formulario(self):
        return {
            "regiones": [k for k, v in self.vars_regiones.items() if v.get()],
            "carpeta_raiz": self.var_carpeta.get(),
            "ruta_excel": self.var_excel.get(),
            "n_posiciones": self.var_n_posiciones.get(),
            "distancia_cm": self.var_distancia.get(),
            "frecuencia": self.var_frecuencia.get(),
            "technology": self.var_technology.get(),
            "wpt_client": self.var_wpt.get(),
            "battery": self.var_battery.get(),
        }

    def _on_iniciar_sesion(self):
        datos = self._leer_datos_formulario()
        try:
            config = validar_config_formulario(datos)
        except ErrorValidacion as e:
            self.lbl_error_form.configure(text=str(e))
            return

        self.lbl_error_form.configure(text="")
        self._mostrar_resumen_y_confirmar(config)

    def _precargar_perfil(self):
        perfil = config_store.cargar_perfil()
        if not perfil:
            return
        for k in perfil.get("regiones", []):
            if k in self.vars_regiones:
                self.vars_regiones[k].set(True)
        self.var_carpeta.set(perfil.get("carpeta_raiz", ""))
        self.var_excel.set(perfil.get("ruta_excel") or "")
        if perfil.get("distancia_cm") is not None:
            self.var_distancia.set(str(perfil["distancia_cm"]))
        if perfil.get("frecuencia") is not None:
            self.var_frecuencia.set(str(perfil["frecuencia"]))
        self.var_technology.set(perfil.get("technology", "") or "")
        self.var_wpt.set(perfil.get("wpt_client", "") or "")
        self.var_battery.set(perfil.get("battery", "") or "")

    def _rellenar_formulario_desde_config(self, config):
        for k in self.vars_regiones:
            self.vars_regiones[k].set(k in config.get("regiones", []))
        self.var_carpeta.set(config.get("carpeta_raiz", ""))
        self.var_excel.set(config.get("ruta_excel") or "")
        self.var_distancia.set(str(config.get("distancia_cm", "")))
        self.var_frecuencia.set(str(config.get("frecuencia", "")))
        self.var_technology.set(config.get("technology", "") or "")
        self.var_wpt.set(config.get("wpt_client", "") or "")
        self.var_battery.set(config.get("battery", "") or "")
        self.var_n_posiciones.set(str(config.get("n_posiciones", "1")))

    def _mostrar_resumen_y_confirmar(self, config):
        plan = config["plan"]
        duracion_total_min = (plan["duracion_rms_s"] + plan["duracion_peak_s"]) // 60
        detalle_fases = []
        for r in config["regiones"]:
            for fdef in core.CONFIG_REGION[r]["fases"]:
                detalle_fases.append(f"  [{r}]  {fdef['etiqueta']}  ({fdef['ms_ventana']//60000} min)")

        resumen = (
            f"Regiones          : {config['nombres_regiones']}\n"
            f"Posiciones        : {config['n_posiciones']}\n"
            f"Distancia         : {config['distancia_cm']} cm\n"
            f"Frecuencia        : {config['frecuencia']} kHz\n"
            f"Duración/posición : {duracion_total_min} min\n"
            f"Technology        : {config['technology']}\n"
            f"WPT Client        : {config['wpt_client']}\n"
            f"Battery           : {config['battery']}\n"
            f"Carpeta raíz      : {config['carpeta_raiz']}\n"
            f"Excel             : {config['ruta_excel'] or 'No configurado'}\n\n"
            f"Fases por región:\n" + "\n".join(detalle_fases)
        )

        if not messagebox.askokcancel("Confirmar configuración",
                                      resumen + "\n\n¿Iniciar la sesión con estos parámetros?"):
            return

        self.config_sesion = config
        self._iniciar_flujo_medida(config)

    # ── Pantalla 2: sesión de medida en curso ───────────────────────────

    def _construir_pantalla_sesion(self):
        scroll = ScrollArea(self)
        self.frame_sesion = scroll

        f = ctk.CTkFrame(scroll, fg_color="transparent")
        f.pack(fill="both", expand=True, padx=24, pady=20)

        cab = ctk.CTkFrame(f, fg_color="transparent")
        cab.pack(fill="x")
        _label(cab, "Sesión de medida en curso", seccion=True,
              font=FUENTE_TITULO_MODAL).pack(side="left")
        self.lbl_fase_actual = _label(cab, "")
        self.lbl_fase_actual.pack(side="right")

        self.frame_paso = SeccionCard(f, "Acción requerida en MAGpy3 / en campo")
        self.frame_paso.pack(fill="x", pady=(16, 12))

        cuerpo_paso = ctk.CTkFrame(self.frame_paso.cuerpo, fg_color="transparent")
        cuerpo_paso.pack(fill="x", padx=18, pady=16)
        cuerpo_paso.columnconfigure(0, weight=1)

        self.lbl_paso_titulo = _label(cuerpo_paso, "—", font=FUENTE_VALOR)
        self.lbl_paso_titulo.configure(text_color=TOKEN_TEXT_PRIMARY)
        self.lbl_paso_titulo.grid(row=0, column=0, sticky="w")
        self.lbl_paso_texto = _label(cuerpo_paso, "Sin acciones pendientes.", wraplength=820)
        self.lbl_paso_texto.grid(row=1, column=0, sticky="w", pady=(4, 10))

        frame_botones_paso = ctk.CTkFrame(cuerpo_paso, fg_color="transparent")
        frame_botones_paso.grid(row=2, column=0, sticky="e")
        self.btn_paso_abortar = _boton(frame_botones_paso, "Abortar sesión",
                                       command=self._on_abortar_sesion, alto=48)
        self.btn_paso_abortar.pack(side="right", padx=(0, 10))
        self.btn_paso_continuar = _boton(frame_botones_paso, "Continuar",
                                         command=self._on_paso_continuar,
                                         primario=True, alto=48)
        self.btn_paso_continuar.pack(side="right")
        self.btn_paso_continuar.configure(state="disabled")
        self.btn_paso_abortar.configure(state="disabled")

        self.frame_nombre_pos = ctk.CTkFrame(f, fg_color="transparent")
        _label(self.frame_nombre_pos, "Nombre de esta posición (ej: Front-Face, Right, 0cm):")\
            .pack(side="left")
        self.var_nombre_pos = ctk.StringVar()
        entry_nombre = _entrada(self.frame_nombre_pos, self.var_nombre_pos, width=260)
        entry_nombre.pack(side="left", padx=10)
        entry_nombre.bind("<Return>", lambda e: self._on_confirmar_nombre_posicion())
        _boton(self.frame_nombre_pos, "Confirmar",
              command=self._on_confirmar_nombre_posicion, alto=48).pack(side="left")

        frame_prog = ctk.CTkFrame(f, fg_color="transparent")
        frame_prog.pack(fill="x", pady=(12, 12))
        self.lbl_progreso_etiqueta = _label(frame_prog, "")
        self.lbl_progreso_etiqueta.pack(anchor="w")
        self.barra_progreso_widget = ctk.CTkProgressBar(
            frame_prog, height=14, corner_radius=7,
            fg_color=TOKEN_BG_INPUT, progress_color=TOKEN_ACCENT_YELLOW)
        self.barra_progreso_widget.pack(fill="x", pady=6)
        self.barra_progreso_widget.set(0)
        self.lbl_progreso_restante = _label(frame_prog, "")
        self.lbl_progreso_restante.pack(anchor="w")

        self.frame_fin_sesion = SeccionCard(f, "Sesión finalizada")
        cuerpo_fin = ctk.CTkFrame(self.frame_fin_sesion.cuerpo, fg_color="transparent")
        cuerpo_fin.pack(fill="x", padx=18, pady=16)
        _label(cuerpo_fin, "¿Qué quieres hacer ahora?").pack(anchor="w", pady=(0, 10))
        frame_botones_fin = ctk.CTkFrame(cuerpo_fin, fg_color="transparent")
        frame_botones_fin.pack(anchor="w")
        _boton(frame_botones_fin, "Nueva sesión (mismos datos)", primario=True, alto=44,
              command=lambda: self._on_nueva_sesion(reutilizar_datos=True)).pack(side="left")
        _boton(frame_botones_fin, "Nueva sesión (cambiar configuración)", alto=44,
              command=lambda: self._on_nueva_sesion(reutilizar_datos=False))\
            .pack(side="left", padx=(10, 0))

        _label(f, "Registro de la sesión:").pack(anchor="w", pady=(6, 6))
        frame_log = ctk.CTkFrame(f, fg_color=TOKEN_BG_INPUT, corner_radius=RADIO_PILL,
                                 border_width=BORDE_GROSOR, border_color=TOKEN_BORDER_SUBTLE)
        frame_log.pack(fill="both", expand=True)
        self.txt_log = ctk.CTkTextbox(
            frame_log, height=260, font=FUENTE_LOG, wrap="word", state="disabled",
            fg_color=TOKEN_BG_INPUT, text_color=TOKEN_TEXT_PRIMARY,
            corner_radius=RADIO_PILL, border_width=0,
            scrollbar_button_color=TOKEN_BG_HEADER,
            scrollbar_button_hover_color=TOKEN_BG_INPUT_HOVER,
        )
        self.txt_log.pack(fill="both", expand=True, padx=2, pady=2)
        for tipo, color in (("OK", COLOR_OK), ("AVISO", COLOR_AVISO),
                            ("ERROR", COLOR_ERROR), ("INFO", COLOR_INFO), ("ESPERA", COLOR_INFO)):
            self.txt_log.tag_config(tipo, foreground=color)

    def _iniciar_flujo_medida(self, config):
        config_store.guardar_perfil(config)
        self.frame_formulario.grid_forget()
        self.frame_sesion.grid(row=0, column=0, sticky="nsew")
        self._log("Iniciando sesión de medida…", "INFO")
        self.hilo = threading.Thread(target=hilo_sesion_medida, args=(config,), daemon=True)
        self.hilo.start()

    # ── Log ──────────────────────────────────────────────────────────────

    def _log(self, texto, tipo="INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{ts}] ", ("INFO",))
        self.txt_log.insert("end", f"{texto}\n", (tipo,))
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ── Panel de paso manual ─────────────────────────────────────────────

    def _mostrar_paso_manual(self, payload):
        self.lbl_paso_titulo.configure(text=payload["titulo"])
        self.lbl_paso_texto.configure(text="\n".join(payload["lineas"]))
        self.btn_paso_continuar.configure(state="normal")
        self.btn_paso_abortar.configure(state="normal")
        self.bell()

    def _on_paso_continuar(self):
        self.btn_paso_continuar.configure(state="disabled")
        self.btn_paso_abortar.configure(state="disabled")
        self.lbl_paso_titulo.configure(text="—")
        self.lbl_paso_texto.configure(text="Sin acciones pendientes.")
        respuesta_q.put(None)

    def _on_abortar_sesion(self):
        if not messagebox.askyesno("Abortar sesión",
                                   "¿Seguro que quieres abortar la sesión de medida en curso?"):
            return
        self.btn_paso_continuar.configure(state="disabled")
        self.btn_paso_abortar.configure(state="disabled")
        respuesta_q.put("ABORTAR")

    # ── Nombre de posición ───────────────────────────────────────────────

    def _pedir_nombre_posicion(self, payload):
        self.frame_nombre_pos.pack(fill="x", pady=(0, 10))
        self.var_nombre_pos.set("")
        self.lbl_fase_actual.configure(text=f"Posición {payload['num']} de {payload['total']}")
        self._log(f"Introduce el nombre de la posición {payload['num']}/{payload['total']}.", "INFO")

    def _on_confirmar_nombre_posicion(self):
        nombre = self.var_nombre_pos.get().strip()
        if not nombre:
            nombre = f"Posicion_{datetime.now().strftime('%H%M%S')}"
        self.frame_nombre_pos.pack_forget()
        respuesta_q.put(nombre)

    # ── Contraseña del Excel (popup modal) ──────────────────────────────

    def _pedir_password_excel(self):
        dialogo = DialogoModal(self, "Excel protegido", ancho=440)
        cont = dialogo.contenido

        _label(cont, "El archivo Excel de resultados está protegido con contraseña.",
              wraplength=380).pack(anchor="w")

        _label(cont, "Contraseña", secundario=True).pack(anchor="w", pady=(16, 4))
        var_password = ctk.StringVar()
        entry = _entrada(cont, var_password, show="•", width=380)
        entry.pack(anchor="w", fill="x")
        entry.focus_set()

        var_recordar = ctk.BooleanVar(value=True)
        cb = ctk.CTkCheckBox(
            cont, text="Recordar en este PC (cifrada, no se vuelve a pedir)",
            variable=var_recordar, font=FUENTE_LABEL, text_color=TOKEN_TEXT_PRIMARY,
            fg_color=TOKEN_ACCENT_YELLOW, hover_color=TOKEN_ACCENT_YELLOW,
            checkmark_color=TOKEN_BG_MODAL, border_color=TOKEN_BORDER_SUBTLE,
            corner_radius=5,
        )
        cb.pack(anchor="w", pady=(14, 0))

        lbl_error = _label(cont, "")
        lbl_error.configure(text_color=COLOR_ERROR)
        lbl_error.pack(anchor="w", pady=(6, 0))

        resultado = {}

        def _aceptar():
            pw = var_password.get()
            if not pw:
                lbl_error.configure(text="Introduce la contraseña o pulsa Cancelar.")
                return
            resultado["password"] = pw
            resultado["recordar"] = var_recordar.get()
            dialogo.destroy()

        def _cancelar():
            dialogo.destroy()

        entry.bind("<Return>", lambda e: _aceptar())

        frame_botones = ctk.CTkFrame(cont, fg_color="transparent")
        frame_botones.pack(fill="x", pady=(18, 0))
        _boton(frame_botones, "Cancelar", command=_cancelar, alto=46).pack(side="right")
        _boton(frame_botones, "Aceptar", command=_aceptar, primario=True, alto=46)\
            .pack(side="right", padx=(0, 8))

        dialogo.protocol("WM_DELETE_WINDOW", _cancelar)
        dialogo.ajustar_alto()
        self.wait_window(dialogo)

        respuesta_q.put(resultado if "password" in resultado else None)

    # ── Pedir valor numérico (distancia PEAK) ───────────────────────────

    def _pedir_numero(self, payload):
        texto, mn, mx = payload["texto"], payload["mn"], payload["mx"]
        actual = payload.get("actual")

        dialogo = DialogoModal(self, "Dato requerido", ancho=520)
        cont = dialogo.contenido

        _label(cont, texto, wraplength=470).pack(anchor="w", fill="x")
        lbl_rango = _label(cont, f"Rango válido: {mn} – {mx}", secundario=True)
        lbl_rango.pack(anchor="w", pady=(4, 14))

        var_valor = ctk.StringVar(value=str(actual) if actual is not None else "")
        entry = _entrada(cont, var_valor, width=470, justify="center")
        entry.pack(anchor="w", fill="x")
        entry.focus_set()
        entry.select_range(0, "end")

        lbl_error = _label(cont, "", wraplength=470)
        lbl_error.configure(text_color=COLOR_ERROR)
        lbl_error.pack(anchor="w", pady=(8, 0))

        resultado = {}

        def _aceptar():
            crudo = var_valor.get().strip().replace(",", ".")
            try:
                v = float(crudo)
            except ValueError:
                lbl_error.configure(text="Introduce un número válido.")
                return
            if not (mn <= v <= mx):
                lbl_error.configure(text=f"Debe estar entre {mn} y {mx}.")
                return
            resultado["valor"] = v
            dialogo.destroy()

        entry.bind("<Return>", lambda e: _aceptar())

        frame_botones = ctk.CTkFrame(cont, fg_color="transparent")
        frame_botones.pack(fill="x", pady=(18, 0))
        _boton(frame_botones, "Aceptar", command=_aceptar, primario=True, alto=46).pack(side="right")

        dialogo.protocol("WM_DELETE_WINDOW", lambda: None)
        dialogo.ajustar_alto()
        self.wait_window(dialogo)

        respuesta_q.put(resultado["valor"])

    def _posicion_fallida(self, payload):
        self.lbl_paso_titulo.configure(text="Posición no completada")
        self.lbl_paso_texto.configure(
            text=f"La posición '{payload['nombre']}' no se ha completado.\n"
                 f"No se pasará a la siguiente posición automáticamente."
        )
        self.btn_paso_continuar.configure(state="normal", text="Repetir posición")
        self.btn_paso_abortar.configure(state="normal")
        self._pendiente_repetir = True
        self.bell()

    def _on_paso_continuar_wrapper(self):
        if getattr(self, "_pendiente_repetir", False):
            self._pendiente_repetir = False
            self.btn_paso_continuar.configure(text="Continuar")
            self.btn_paso_continuar.configure(state="disabled")
            self.btn_paso_abortar.configure(state="disabled")
            respuesta_q.put("REPETIR")
        else:
            self._on_paso_continuar()

    # ── Progreso ─────────────────────────────────────────────────────────

    def _progreso_inicio(self, payload):
        self.lbl_progreso_etiqueta.configure(text=payload["etiqueta"])
        self.barra_progreso_widget.set(0)

    def _progreso_update(self, payload):
        self.barra_progreso_widget.set(payload["pct"] / 100)
        mr, sr = divmod(payload["restante_s"], 60)
        self.lbl_progreso_restante.configure(text=f"Restante: {mr:02d}:{sr:02d}   ({payload['pct']}%)")

    def _progreso_fin(self, _payload):
        self.lbl_progreso_restante.configure(text="Fase completada.")

    # ── Fin de sesión ────────────────────────────────────────────────────

    def _sesion_completada(self, payload):
        self.lbl_fase_actual.configure(text="Sesión completada", text_color=COLOR_OK)
        resumen = (
            f"Posiciones medidas : {payload['n_posiciones']}\n"
            f"Regiones           : {payload['nombres_regiones']}\n"
            f"Frecuencia         : {payload['frecuencia']} kHz\n"
            f"Hora inicio        : {payload['t_inicio'].strftime('%H:%M:%S')}\n"
            f"Hora fin           : {payload['t_fin'].strftime('%H:%M:%S')}\n"
            f"Duración total     : {payload['h']}h {payload['m']}min\n\n"
            f"Carpeta raíz  --> {payload['carpeta_raiz']}\n"
            f"CSV resumen   --> {payload['archivo_csv']}\n"
        )
        if payload["archivo_csv_peak"]:
            resumen += f"CSV Peak      --> {payload['archivo_csv_peak']}\n"
        resumen += f"Excel         --> {payload['ruta_excel'] or 'No configurado'}\n"
        self._log("Sesión completada.", "OK")
        messagebox.showinfo("Sesión completada", resumen)
        self._mostrar_panel_fin_sesion()

    def _sesion_abortada(self, _payload):
        self.lbl_fase_actual.configure(text="Sesión abortada", text_color=COLOR_ERROR)
        self._log("Sesión abortada por el operador.", "ERROR")
        self._mostrar_panel_fin_sesion()

    def _error_fatal(self, texto):
        self._log(texto, "ERROR")
        messagebox.showerror("Error", texto)
        self.lbl_fase_actual.configure(text="Sesión detenida por error", text_color=COLOR_ERROR)
        self._mostrar_panel_fin_sesion()

    def _mostrar_panel_fin_sesion(self):
        self.hilo = None
        self.btn_paso_continuar.configure(state="disabled")
        self.btn_paso_abortar.configure(state="disabled")
        self.frame_nombre_pos.pack_forget()
        self.frame_fin_sesion.pack(fill="x", pady=(12, 0))

    def _on_nueva_sesion(self, reutilizar_datos):
        if reutilizar_datos and getattr(self, "config_sesion", None):
            self._rellenar_formulario_desde_config(self.config_sesion)

        self.frame_fin_sesion.pack_forget()
        self.lbl_fase_actual.configure(text="")
        self.lbl_paso_titulo.configure(text="—")
        self.lbl_paso_texto.configure(text="Sin acciones pendientes.")
        self.barra_progreso_widget.set(0)
        self.lbl_progreso_etiqueta.configure(text="")
        self.lbl_progreso_restante.configure(text="")
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

        self.frame_sesion.grid_forget()
        self.lbl_error_form.configure(text="")
        self.frame_formulario.grid(row=0, column=0, sticky="nsew")

    # ── Bucle principal de consumo de eventos ───────────────────────────

    def _procesar_eventos(self):
        try:
            while True:
                tipo, payload = eventos_q.get_nowait()

                if tipo == "LOG":
                    texto, nivel = payload
                    self._log(texto, nivel)

                elif tipo == "FASE":
                    self.lbl_fase_actual.configure(text=payload, text_color=COLOR_INFO)

                elif tipo == "CUADRO":
                    self._log(f"{payload['titulo']}: " + " | ".join(payload["lineas"]), "AVISO")

                elif tipo == "PASO_MANUAL":
                    self._mostrar_paso_manual(payload)
                    self.btn_paso_continuar.configure(command=self._on_paso_continuar)

                elif tipo == "POSICION_FALLIDA":
                    self._posicion_fallida(payload)
                    self.btn_paso_continuar.configure(command=self._on_paso_continuar_wrapper)

                elif tipo == "PEDIR_NOMBRE_POSICION":
                    self._pedir_nombre_posicion(payload)

                elif tipo == "PEDIR_PASSWORD_EXCEL":
                    self._pedir_password_excel()

                elif tipo == "PEDIR_NUMERO":
                    self._pedir_numero(payload)

                elif tipo == "PROGRESO_INICIO":
                    self._progreso_inicio(payload)
                elif tipo == "PROGRESO":
                    self._progreso_update(payload)
                elif tipo == "PROGRESO_FIN":
                    self._progreso_fin(payload)

                elif tipo == "SESION_COMPLETADA":
                    self._sesion_completada(payload)
                elif tipo == "SESION_ABORTADA":
                    self._sesion_abortada(payload)
                elif tipo == "ERROR_FATAL":
                    self._error_fatal(payload)

        except queue.Empty:
            pass
        finally:
            self.after(150, self._procesar_eventos)


def instalar_monkeypatches():
    core._msg_original  = core.msg
    core._beep_original = core.beep

    core.msg            = gui_msg
    core.enter          = gui_enter
    core.cuadro         = gui_cuadro
    core.barra_progreso = gui_barra_progreso
    core.beep           = gui_beep
    core.input          = gui_input
    core.getpass.getpass = gui_getpass_excel
    core.pedir_numero   = gui_pedir_numero

    password_guardada = config_store.cargar_password_excel()
    if password_guardada:
        core._EXCEL_PASSWORD = password_guardada


def main():
    instalar_monkeypatches()

    if not core.DEBUG_FAST and os.environ.get("MAGPY_DEBUG_FAST") == "1":
        core.DEBUG_FAST = True

    app = AppMagpy3()
    app.mainloop()


if __name__ == "__main__":
    main()
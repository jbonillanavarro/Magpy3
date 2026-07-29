"""
config_store.py
────────────────
Persistencia entre sesiones para la GUI de MAGpy3:

  1. Perfil de configuración (regiones, carpetas, distancia, frecuencia, etc.)
     guardado en un JSON legible en texto plano. No contiene secretos.

  2. Contraseña del Excel protegido, guardada por separado, cifrada con
     DPAPI (Windows Data Protection API) atada al usuario/máquina actual.
     Un archivo copiado a otro PC o abierto por otro usuario de Windows
     no se puede descifrar: DPAPI deriva la clave del perfil del usuario
     de Windows, no de nada que viaje con el archivo.

     En sistemas no-Windows (p. ej. si se prueba este módulo en Linux/Mac)
     no hay DPAPI disponible: se hace un fallback explícito a "no guardar
     la contraseña" en vez de guardarla en claro, y se avisa por consola.

Ambos ficheros se guardan en:
    %APPDATA%\\MagpySAR\\   (Windows)
    ~/.magpy_sar/            (fallback en otros SO, solo para pruebas)
"""

import os
import json
import base64
import sys

NOMBRE_APP = "MagpySAR"


def _carpeta_config():
    base = os.environ.get("APPDATA")
    if base:
        carpeta = os.path.join(base, NOMBRE_APP)
    else:
        carpeta = os.path.join(os.path.expanduser("~"), f".{NOMBRE_APP.lower()}")
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


RUTA_PERFIL      = os.path.join(_carpeta_config(), "ultima_sesion.json")
RUTA_PASSWORD    = os.path.join(_carpeta_config(), "excel_pw.bin")

# Claves del perfil que SÍ se recuerdan entre sesiones. Deliberadamente
# no incluye n_posiciones (cada sesión suele medir un lote distinto de
# posiciones y forzar el número anterior invita a error) — se puede
# ampliar luego si se pide.
CLAVES_PERFIL = (
    "regiones", "carpeta_raiz", "ruta_excel",
    "distancia_cm", "frecuencia",
    "technology", "wpt_client", "battery",
)


# ═══════════════════════════════════════════════════════════════
# PERFIL DE CONFIGURACIÓN (JSON en texto plano, sin secretos)
# ═══════════════════════════════════════════════════════════════

def cargar_perfil():
    """Devuelve el dict de la última sesión guardada, o {} si no hay o está corrupto."""
    if not os.path.isfile(RUTA_PERFIL):
        return {}
    try:
        with open(RUTA_PERFIL, "r", encoding="utf-8") as f:
            datos = json.load(f)
        return {k: datos[k] for k in CLAVES_PERFIL if k in datos}
    except Exception:
        # Perfil corrupto o de un formato antiguo: no bloquear el arranque,
        # simplemente empezar en blanco.
        return {}


def guardar_perfil(config):
    """
    config: dict de configuración validada (el mismo formato que produce
    validar_config_formulario en gui_magpy3.py). Solo se guardan las
    CLAVES_PERFIL; el resto (archivo_csv, plan, etc.) se recalcula cada vez.
    """
    datos = {k: config.get(k) for k in CLAVES_PERFIL}
    try:
        with open(RUTA_PERFIL, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception as e:
        # No es crítico para la medida: si falla, simplemente no se recuerda
        # la próxima vez. Se deja constancia por consola.
        print(f"[AVISO] No se pudo guardar el perfil de sesión: {e}")


# ═══════════════════════════════════════════════════════════════
# CONTRASEÑA DEL EXCEL (cifrada con DPAPI, atada a usuario/máquina)
# ═══════════════════════════════════════════════════════════════

def _dpapi_disponible():
    return sys.platform == "win32"


def _dpapi_proteger(texto_plano: str) -> bytes:
    import ctypes
    import ctypes.wintypes as wt

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    datos_in  = texto_plano.encode("utf-8")
    blob_in   = DATA_BLOB(len(datos_in), ctypes.cast(ctypes.create_string_buffer(datos_in), ctypes.POINTER(ctypes.c_char)))
    blob_out  = DATA_BLOB()

    CRYPTPROTECT_UI_FORBIDDEN = 0x01
    ok = ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(blob_in), "MagpySAR - Excel password", None, None, None,
        CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out)
    )
    if not ok:
        raise OSError("CryptProtectData falló (DPAPI).")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def _dpapi_desproteger(datos_cifrados: bytes) -> str:
    import ctypes
    import ctypes.wintypes as wt

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wt.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    blob_in  = DATA_BLOB(len(datos_cifrados), ctypes.cast(ctypes.create_string_buffer(datos_cifrados), ctypes.POINTER(ctypes.c_char)))
    blob_out = DATA_BLOB()

    CRYPTPROTECT_UI_FORBIDDEN = 0x01
    ok = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None,
        CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out)
    )
    if not ok:
        raise OSError("CryptUnprotectData falló (DPAPI). ¿Archivo de otro usuario/PC?")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData).decode("utf-8")
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def guardar_password_excel(password: str):
    """Cifra y guarda la contraseña del Excel. No hace nada (con aviso) si no hay DPAPI."""
    if not _dpapi_disponible():
        print("[AVISO] DPAPI no disponible en este sistema operativo; "
              "la contraseña del Excel NO se guarda en disco (solo en memoria de esta sesión).")
        return
    try:
        cifrado = _dpapi_proteger(password)
        with open(RUTA_PASSWORD, "wb") as f:
            f.write(base64.b64encode(cifrado))
    except Exception as e:
        print(f"[AVISO] No se pudo guardar la contraseña del Excel de forma persistente: {e}")


def cargar_password_excel():
    """Devuelve la contraseña guardada (str) o None si no hay ninguna / no se puede leer."""
    if not _dpapi_disponible():
        return None
    if not os.path.isfile(RUTA_PASSWORD):
        return None
    try:
        with open(RUTA_PASSWORD, "rb") as f:
            cifrado = base64.b64decode(f.read())
        return _dpapi_desproteger(cifrado)
    except Exception as e:
        # Típicamente: archivo de otro usuario/PC, o perfil de Windows
        # reinstalado. No es un error fatal: simplemente se volverá a pedir.
        print(f"[AVISO] No se pudo leer la contraseña guardada del Excel: {e}")
        return None


def borrar_password_excel():
    """Elimina la contraseña guardada en disco (por si el operador quiere olvidarla)."""
    try:
        if os.path.isfile(RUTA_PASSWORD):
            os.remove(RUTA_PASSWORD)
    except Exception as e:
        print(f"[AVISO] No se pudo borrar la contraseña guardada: {e}")

"""app/excel_engine.py — Motor de actualización Excel (COM / openpyxl fallback)"""
from __future__ import annotations

import os
import sys
import time
import logging
import threading
import traceback
from pathlib import Path
from dataclasses import dataclass, field


# ── win32com opcional ─────────────────────────────────────────────────────────
try:
    import win32com.client
    import pythoncom
    import pywintypes
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

# HRESULT devuelto cuando Excel está ocupado y rechaza la llamada COM
_COM_CALL_REJECTED = -2147418111   # RPC_E_CALL_REJECTED
_COM_RETRY_ATTEMPTS = 6
_COM_RETRY_WAIT = 3  # segundos entre reintentos

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class EngineResult:
    success: bool = False
    connections_found: int = 0
    pivots_ok: int = 0
    pivots_err: int = 0
    error_msg: str = ""
    duration_s: float = 0.0


@dataclass
class EngineConfig:
    file_path: str
    refresh_connections: bool = True
    refresh_pivots: bool = True
    save_on_success: bool = True
    excel_visible: bool = False
    lock_timeout: int = 120
    lock_retry: int = 5
    lock_max_retries: int = 5
    refresh_timeout: int = 300
    refresh_check: int = 3
    stop_event: threading.Event = field(default_factory=threading.Event)


# ══════════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════════════════════

# Atributos de OneDrive Files On-Demand (placeholders en la nube)
_ONEDRIVE_CLOUD_ATTRS = 0x400000 | 0x40000  # RECALL_ON_DATA_ACCESS | RECALL_ON_OPEN


def resolve_path(raw: str) -> str:
    """Expande variables de entorno y normaliza la ruta sin cambiar el disco."""
    expanded = os.path.expandvars(raw.strip())
    p = Path(expanded)
    # Si la ruta es absoluta, usar os.path.normpath para no depender del CWD
    # (Path.resolve() en D: podría mutar rutas de C: en algunos entornos)
    if p.is_absolute():
        return os.path.normpath(str(p))
    # Si es relativa, resolver desde el directorio de trabajo del proceso
    return os.path.normpath(os.path.abspath(expanded))


def _is_onedrive_placeholder(path: str) -> bool:
    """Devuelve True si el archivo es un placeholder OneDrive (no descargado)."""
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        attrs = ctypes.windll.kernel32.GetFileAttributesW(path)
        return attrs != -1 and bool(attrs & _ONEDRIVE_CLOUD_ATTRS)
    except Exception:
        return False


def _trigger_onedrive_download(path: str, log: logging.Logger, timeout: int = 120):
    """
    Fuerza la descarga del placeholder OneDrive y bloquea hasta que el archivo
    esté físicamente disponible (o se agote el timeout).
    """
    import ctypes
    import subprocess

    # Disparar hidratación abriendo en lectura (Cloud Provider IRP)
    try:
        with open(path, "rb") as f:
            f.read(1)
    except Exception:
        pass

    # Marcar como "pinned" para que OneDrive lo mantenga local
    try:
        subprocess.run(["attrib", "+P", "-U", path],
                       capture_output=True, timeout=10)
    except Exception:
        pass

    log.info("Descarga de OneDrive iniciada. Esperando sincronización...")

    RECALL_FLAGS = 0x400000 | 0x40000  # RECALL_ON_DATA_ACCESS | RECALL_ON_OPEN
    deadline = time.time() + timeout
    while time.time() < deadline:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(path)
        if attrs != -1 and not (attrs & RECALL_FLAGS):
            log.info("Archivo OneDrive hidratado correctamente.")
            return
        time.sleep(2)

    log.warning(f"Timeout ({timeout}s) esperando descarga de OneDrive. Se intentará continuar.")


def _lock_reason(path: str) -> str:
    """Detecta el motivo probable por el que el archivo no está disponible."""
    p = Path(path)
    if _is_onedrive_placeholder(path):
        return "archivo OneDrive en la nube (descargando)"
    lock_file = p.parent / f"~${p.name}"
    if lock_file.exists():
        return "Excel tiene el archivo abierto (~$lock detectado)"
    tmp_files = list(p.parent.glob("*.tmp"))
    if tmp_files:
        return "posible sincronización de OneDrive en curso"
    return "proceso externo desconocido"


def wait_for_file(path: str, timeout: int, interval: int,
                  log: logging.Logger, max_retries: int = 0,
                  stop_event: threading.Event | None = None) -> bool:
    """
    Espera hasta que el archivo esté disponible.
    - timeout: tiempo máximo total en segundos.
    - max_retries: número máximo de intentos (0 = ilimitado, solo por timeout).
    - stop_event: si se activa, aborta la espera inmediatamente.
    """
    deadline = time.time() + timeout
    attempt = 0
    while time.time() < deadline:
        if stop_event and stop_event.is_set():
            log.warning("Ejecución detenida manualmente durante espera de archivo.")
            return False
        try:
            with open(path, "r+b"):
                return True
        except (PermissionError, OSError):
            attempt += 1
            elapsed = round(timeout - (deadline - time.time()))
            reason = _lock_reason(path)
            limit_info = f"/{max_retries}" if max_retries else ""
            log.warning(
                f"Archivo bloqueado (intento {attempt}{limit_info}, {elapsed}s/{timeout}s) "
                f"— {reason}. Reintentando en {interval}s…"
            )
            if max_retries and attempt >= max_retries:
                log.error(f"Se alcanzó el límite de {max_retries} reintentos. Tarea cancelada.")
                return False
            # Esperar en pequeños intervalos para responder rápido al stop_event
            end = time.time() + interval
            while time.time() < end:
                if stop_event and stop_event.is_set():
                    log.warning("Ejecución detenida manualmente durante espera de archivo.")
                    return False
                time.sleep(0.5)
    return False


# ══════════════════════════════════════════════════════════════════════════════
# MOTOR COM
# ══════════════════════════════════════════════════════════════════════════════

class ExcelCOMUpdater:
    def __init__(self, cfg: EngineConfig, log: logging.Logger):
        self.cfg = cfg
        self.log = log
        self.xl = None
        self.wb = None

    def open(self):
        pythoncom.CoInitialize()
        self.xl = win32com.client.DispatchEx("Excel.Application")
        self.xl.Visible = self.cfg.excel_visible
        self.xl.DisplayAlerts = False
        self.xl.AskToUpdateLinks = False
        self.wb = self.xl.Workbooks.Open(
            self.cfg.file_path, UpdateLinks=0, ReadOnly=False
        )
        # Esperar a que Excel termine de cargar el libro antes de seguir.
        # Sin este delay, el primer acceso a wb.Connections puede devolver
        # RPC_E_CALL_REJECTED (-2147418111) porque Excel aún no está listo.
        time.sleep(2)
        self.log.info("Libro abierto (COM).")

    def close(self, save: bool):
        if self.wb:
            try:
                if save:
                    self.wb.Save()
                    self.log.info("Guardado.")
                self.wb.Close(SaveChanges=False)
            except Exception as e:
                self.log.error(f"Error cerrando libro: {e}")
        if self.xl:
            try:
                self.xl.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    def _is_connection_refreshing(self, conn) -> bool:
        """Comprueba si una conexión COM está actualizando. Maneja tipos no soportados."""
        for attr in ("ODBCConnection", "OLEDBConnection",
                     "TextConnection", "WorksheetDataConnection"):
            try:
                # getattr en COM puede lanzar excepción si el atributo no existe
                sub = getattr(conn, attr)
                if sub is not None and sub.Refreshing:
                    return True
            except Exception:
                pass
        return False

    def _com_retry(self, func, retries: int = 5, delay: float = 2.0):
        """Reintenta una llamada COM si Excel devuelve RPC_E_CALL_REJECTED."""
        import pywintypes
        RPC_E_CALL_REJECTED = -2147418111
        for attempt in range(retries):
            try:
                return func()
            except pywintypes.com_error as e:
                if e.args[0] == RPC_E_CALL_REJECTED and attempt < retries - 1:
                    self.log.debug(f"Excel ocupado (intento {attempt + 1}/{retries}), reintentando en {delay}s...")
                    time.sleep(delay)
                else:
                    raise

    def _refresh_connections(self) -> tuple[bool, int]:
        conns = self._com_retry(lambda: self.wb.Connections)
        n = conns.Count
        self.log.info(f"Conexiones: {n}")
        if n == 0:
            return True, 0

        self.wb.RefreshAll()
        deadline = time.time() + self.cfg.refresh_timeout
        while time.time() < deadline:
            if self.cfg.stop_event.is_set():
                self.log.warning("Ejecución detenida manualmente durante actualización de conexiones.")
                return False, n
            time.sleep(self.cfg.refresh_check)
            still_refreshing = []
            for i in range(1, n + 1):
                try:
                    conn = conns.Item(i)
                    if self._is_connection_refreshing(conn):
                        still_refreshing.append(conn.Name)
                except Exception as e:
                    self.log.debug(f"No se pudo leer estado de conexión {i}: {e}")
            if not still_refreshing:
                self.log.info("Todas las conexiones finalizadas.")
                return True, n
            self.log.debug(f"Esperando conexiones: {still_refreshing}")
        self.log.error("Timeout esperando conexiones.")
        return False, n

    def _com_retry(self, fn, *args):
        """Ejecuta fn(*args) reintentando si Excel rechaza la llamada (RPC_E_CALL_REJECTED)."""
        for attempt in range(_COM_RETRY_ATTEMPTS):
            try:
                return fn(*args)
            except pywintypes.com_error as e:
                if e.hresult == _COM_CALL_REJECTED and attempt < _COM_RETRY_ATTEMPTS - 1:
                    self.log.debug(
                        f"COM rechazado por Excel (intento {attempt + 1}/{_COM_RETRY_ATTEMPTS}), "
                        f"reintentando en {_COM_RETRY_WAIT}s..."
                    )
                    time.sleep(_COM_RETRY_WAIT)
                else:
                    raise

    def _refresh_pivots(self) -> tuple[int, int]:
        ok = err = 0
        for s in range(1, self.wb.Sheets.Count + 1):
            sheet = self._com_retry(lambda idx=s: self.wb.Sheets(idx))
            for p in range(1, self._com_retry(lambda sh=sheet: sh.PivotTables().Count) + 1):
                pt = self._com_retry(lambda sh=sheet, idx=p: sh.PivotTables(idx))
                try:
                    self._com_retry(lambda t=pt: t.RefreshTable())
                    ok += 1
                    self.log.info(f"  ✔ PivotTable '{pt.Name}' en '{sheet.Name}'")
                except Exception as e:
                    err += 1
                    self.log.error(f"  ✘ PivotTable '{pt.Name}' en '{sheet.Name}': {e}")
                    # Pausa tras fallo de pivot para que Excel vuelva a estado estable
                    time.sleep(2)
        return ok, err

    def run(self) -> EngineResult:
        t0 = time.time()
        res = EngineResult()
        try:
            self.open()
            conn_ok, n_conn = (True, 0)
            if self.cfg.refresh_connections:
                conn_ok, n_conn = self._refresh_connections()
            res.connections_found = n_conn

            if self.cfg.refresh_pivots:
                res.pivots_ok, res.pivots_err = self._refresh_pivots()

            res.success = conn_ok and res.pivots_err == 0
            if not res.success and not res.error_msg:
                parts = []
                if not conn_ok:
                    parts.append("Error actualizando conexiones")
                if res.pivots_err:
                    parts.append(f"{res.pivots_err} tabla(s) dinámica(s) fallida(s)")
                res.error_msg = "; ".join(parts)
        except Exception:
            res.error_msg = traceback.format_exc()
            self.log.error(f"Excepción:\n{res.error_msg}")
        finally:
            self.close(save=res.success and self.cfg.save_on_success)
            res.duration_s = round(time.time() - t0, 2)
        return res


# ══════════════════════════════════════════════════════════════════════════════
# FALLBACK openpyxl
# ══════════════════════════════════════════════════════════════════════════════

class OpenpyxlUpdater:
    def __init__(self, cfg: EngineConfig, log: logging.Logger):
        self.cfg = cfg
        self.log = log

    def run(self) -> EngineResult:
        t0 = time.time()
        res = EngineResult()
        self.log.warning(
            "Modo openpyxl (sin COM): conexiones externas y tablas dinámicas "
            "NO se refrescan automáticamente."
        )
        try:
            wb = openpyxl.load_workbook(self.cfg.file_path, keep_vba=True)
            self.log.info(f"Hojas: {wb.sheetnames}")
            # ── PUNTO DE EXTENSIÓN: modificar datos aquí ──
            if self.cfg.save_on_success:
                wb.save(self.cfg.file_path)
                self.log.info("Guardado (openpyxl).")
            res.success = True
        except Exception:
            res.error_msg = traceback.format_exc()
            self.log.error(res.error_msg)
        finally:
            res.duration_s = round(time.time() - t0, 2)
        return res


# ══════════════════════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA PÚBLICO
# ══════════════════════════════════════════════════════════════════════════════

def run_update(cfg: EngineConfig, log: logging.Logger) -> EngineResult:
    """Valida el archivo y ejecuta el motor adecuado."""
    path = resolve_path(cfg.file_path)
    cfg.file_path = path
    log.info(f"Ruta resuelta: {path}")

    if not Path(path).exists():
        msg = f"Archivo no encontrado: {path}"
        log.error(msg)
        return EngineResult(error_msg=msg)

    # Si es un placeholder de OneDrive, disparar la descarga antes del loop
    if _is_onedrive_placeholder(path):
        log.warning(
            "El archivo está solo en la nube (OneDrive Files On-Demand). "
            "Iniciando descarga local..."
        )
        _trigger_onedrive_download(path, log)

    available = wait_for_file(path, cfg.lock_timeout, cfg.lock_retry, log, cfg.lock_max_retries, cfg.stop_event)
    if not available:
        limit = f"{cfg.lock_max_retries} intentos" if cfg.lock_max_retries else f"{cfg.lock_timeout}s"
        msg = f"Archivo bloqueado: se canceló la tarea tras {limit}."
        log.error(msg)
        return EngineResult(error_msg=msg)

    if WIN32_AVAILABLE:
        log.info("Motor: Excel COM (win32com)")
        return ExcelCOMUpdater(cfg, log).run()
    elif OPENPYXL_AVAILABLE:
        log.info("Motor: openpyxl (fallback)")
        return OpenpyxlUpdater(cfg, log).run()
    else:
        msg = "Sin motor disponible. Instala: pip install openpyxl pywin32"
        log.critical(msg)
        return EngineResult(error_msg=msg)

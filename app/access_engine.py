"""app/access_engine.py — Motor de automatización Access ETL (pipeline Excel → Access)

Por qué el VBScript fallaba:
  El VBScript usaba EsperarConexiones() que solo verificaba ODBCConnection.Refreshing.
  Las conexiones Power Query (cubos RMS) usan OLEDBConnection, por lo que el wait
  retornaba prematuramente: Excel guardaba el archivo aún sin los datos actualizados.
  Access importaba entonces los archivos con datos viejos.

  Esta implementación reutiliza ExcelCOMUpdater (que ya espera TODOS los tipos de
  conexión: ODBC, OLEDB, Text, WorksheetData) antes de guardar.
"""
from __future__ import annotations

import os
import shutil
import time
import logging
import traceback
from dataclasses import dataclass, field
from pathlib import Path

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════════════
# CONFIG / RESULT
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class PipelineConfig:
    """Configuración completa del pipeline Excel → Access."""
    # Archivos Excel FUENTE (cubos) a refrescar antes de importar.
    # Formato: [{"path": "...", "visible": false}, ...]
    excel_files: list[dict] = field(default_factory=list)
    # Base de datos Access
    access_db: str = ""
    access_visible: bool = False
    # Operaciones Access
    compact_before_import: bool = True   # legacy; respetado si compact_position == ""
    pre_import_macros: list[str] = field(default_factory=list)
    saved_imports: list[str] = field(default_factory=list)
    post_import_macros: list[str] = field(default_factory=list)
    # Cuándo hacer Compact & Repair. Valores:
    #   "before_macros"   → antes de abrir Access (más rápido pero compacta datos viejos)
    #   "after_pre_macros"→ entre pre_import_macros e imports (replica manual: borra→compact→importa)
    #   "skip"            → no compactar
    #   ""                → resuelve en runtime: "after_pre_macros" si compact_before_import else "skip"
    compact_position: str = ""
    # Archivos Excel CONSUMIDORES a refrescar después del ETL (paso 8 del manual:
    # tableros/herramientas que leen de Access vía Power Query/conexiones).
    post_refresh_excel_files: list[dict] = field(default_factory=list)
    # Timeouts Excel
    excel_refresh_timeout: int = 300
    excel_refresh_check: int = 3
    excel_lock_timeout: int = 120
    # Timeouts Access (lock + hidratación OneDrive)
    access_lock_timeout: int = 120
    access_lock_retry: int = 5
    # Si True, un fallo individual en macro/importación NO aborta el pipeline:
    # registra el fallo en steps_failed y continúa con los pasos restantes.
    # Solo aplica al bloque de operaciones Access; los Excel siguen fail-fast
    # porque un .xlsx corrupto/desactualizado contamina lo posterior.
    continue_on_error: bool = False

    def resolved_compact_position(self) -> str:
        """Resuelve compact_position considerando el flag legacy compact_before_import."""
        if self.compact_position:
            return self.compact_position
        return "after_pre_macros" if self.compact_before_import else "skip"


@dataclass
class PipelineResult:
    success: bool = False
    error_msg: str = ""
    duration_s: float = 0.0
    connections_found: int = 0   # compatibilidad con EngineResult
    pivots_ok: int = 0
    pivots_err: int = 0
    steps_done: list[str] = field(default_factory=list)
    steps_failed: list[str] = field(default_factory=list)


# ══════════════════════════════════════════════════════════════════════════════
# RUNNER
# ══════════════════════════════════════════════════════════════════════════════

class AccessPipelineRunner:
    """
    Ejecuta el pipeline completo de actualización Access/Excel.

    Pasos:
        1. Refrescar cada archivo Excel (espera correctamente todos los tipos
           de conexión COM: ODBC, OLEDB, Power Query, WorksheetData).
        2. Compact & Repair de la BD Access (opcional).
        3. Macros pre-importación (p.ej. 'Ejecutar Elimina Cubos').
        4. Importaciones guardadas (p.ej. 'Importación: Cubo_SKU_SUC').
        5. Macros post-importación (p.ej. 'Ejecutar ETL Procesos sin Top Venta Categoría').
    """

    def __init__(self, cfg: PipelineConfig, log: logging.Logger):
        self.cfg = cfg
        self.log = log

    # ── Paso 0: Preparar .accdb (OneDrive + lock) ──────────────────────────

    def _prepare_access_db(self) -> bool:
        """
        Garantiza que el .accdb esté localmente disponible y desbloqueado antes
        de cualquier apertura COM. Hereda los helpers de excel_engine:
          - Si es placeholder OneDrive → fuerza descarga y espera hidratación.
          - Espera a que el archivo no esté bloqueado (.laccdb / lock exclusivo).
        """
        import threading
        from app.excel_engine import (
            _is_onedrive_placeholder,
            _trigger_onedrive_download,
            wait_for_file,
        )

        path = self.cfg.access_db

        if _is_onedrive_placeholder(path):
            self.log.warning(
                "[Access] La BD está solo en la nube (OneDrive Files On-Demand). "
                "Iniciando descarga local..."
            )
            _trigger_onedrive_download(path, self.log)

        available = wait_for_file(
            path,
            timeout=self.cfg.access_lock_timeout,
            interval=self.cfg.access_lock_retry,
            log=self.log,
            max_retries=0,
            stop_event=threading.Event(),
        )
        if not available:
            self.log.error(
                f"[Access] BD bloqueada tras {self.cfg.access_lock_timeout}s. "
                "Probablemente Access la tiene abierta en otra sesión."
            )
        return available

    # ── Paso 1: Actualizar Excel ───────────────────────────────────────────

    def _refresh_excel_file(self, file_path: str, visible: bool) -> bool:
        """Refresca un archivo Excel usando COM, esperando TODOS los tipos de conexión.

        Delega en run_update() (no en ExcelCOMUpdater.run() directo) para heredar
        pre-hidratación OneDrive y wait_for_file (lock-wait) que solo viven en el
        wrapper público.
        """
        import threading
        from app.excel_engine import EngineConfig, run_update

        ecfg = EngineConfig(
            file_path=file_path,
            refresh_connections=True,
            refresh_pivots=False,
            save_on_success=True,
            excel_visible=visible,
            refresh_timeout=self.cfg.excel_refresh_timeout,
            refresh_check=self.cfg.excel_refresh_check,
            lock_timeout=self.cfg.excel_lock_timeout,
            lock_retry=5,
            lock_max_retries=0,
            stop_event=threading.Event(),
        )
        result = run_update(ecfg, self.log)
        name = Path(file_path).name
        if result.success:
            self.log.info(
                f"[Excel] '{name}' actualizado "
                f"({result.connections_found} conexiones, {result.duration_s}s)."
            )
        else:
            self.log.error(f"[Excel] Error en '{name}': {result.error_msg}")
        return result.success

    # ── Paso 2: Compact & Repair ──────────────────────────────────────────

    def _compact_repair(self, db_path: str) -> bool:
        """Ejecuta Compact & Repair de Access y reemplaza el archivo original."""
        temp_path = str(Path(db_path).parent / (Path(db_path).stem + "_compact_temp.accdb"))

        # Limpiar residuo de ejecución previa
        if Path(temp_path).exists():
            try:
                os.remove(temp_path)
            except Exception:
                pass

        compact_ok = False
        acc = None
        acc_pid = None
        try:
            pythoncom.CoInitialize()
            acc = win32com.client.DispatchEx("Access.Application")
            acc.Visible = False
            # Registrar PID en cuanto Access esté listo.
            try:
                import win32process
                from app.excel_engine import current_run_id_var
                from app import com_registry
                _, acc_pid = win32process.GetWindowThreadProcessId(acc.hWndAccessApp)
                com_registry.register(
                    acc_pid, "MSACCESS.EXE", db_path, current_run_id_var.get(),
                )
            except Exception:
                acc_pid = None
            # Bloquear prompts de macros firmadas/no firmadas (msoAutomationSecurityForceDisable = 3).
            # Property no disponible en Access < 2007; ignorar si COM la rechaza.
            try:
                acc.AutomationSecurity = 3
            except Exception:
                pass
            compact_ok = bool(acc.CompactRepair(db_path, temp_path, True))
            if not compact_ok:
                self.log.error("[Access] CompactRepair retornó False.")
        except Exception:
            self.log.error(f"[Access] Error en Compact & Repair:\n{traceback.format_exc()}")
        finally:
            if acc:
                try:
                    acc.Quit()
                except Exception:
                    pass
            try:
                from app import com_registry
                com_registry.unregister(acc_pid)
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        if compact_ok and Path(temp_path).exists():
            try:
                os.remove(db_path)
                shutil.move(temp_path, db_path)
                self.log.info("[Access] Compact & Repair completado.")
                return True
            except Exception:
                self.log.error(
                    f"[Access] Error al reemplazar archivo compactado:\n{traceback.format_exc()}"
                )

        # Limpiar temp si algo falló
        if Path(temp_path).exists():
            try:
                os.remove(temp_path)
            except Exception:
                pass
        return False

    # ── Pasos 3-5: Operaciones en Access abierto ──────────────────────────

    def _run_access_operations(
        self,
        db_path: str,
        res: "PipelineResult | None" = None,
        *,
        run_pre_macros: bool = True,
        run_imports: bool = True,
        run_post_macros: bool = True,
    ) -> bool:
        """
        Abre la BD Access y ejecuta el subset indicado por flags. Útil para
        intercalar Compact & Repair entre pre_macros e imports (manual real).

        Si cfg.continue_on_error es True, una macro/importación individual que
        falle se registra en res.steps_failed y la ejecución continúa. En modo
        estricto (default) cualquier fallo levanta excepción y aborta el bloque.
        """
        def _attempt(label: str, action) -> bool:
            """Ejecuta una operación COM; devuelve True si éxito.

            En modo continue_on_error captura excepciones individuales y registra
            el fallo sin propagarlo. En modo estricto re-lanza.
            """
            self.log.info(f"[Access] {label}")
            try:
                action()
                self.log.info(f"[Access] '{label}' OK.")
                if res is not None:
                    res.steps_done.append(label)
                time.sleep(1)
                return True
            except Exception:
                self.log.error(
                    f"[Access] Error en '{label}':\n{traceback.format_exc()}"
                )
                if not self.cfg.continue_on_error:
                    raise
                if res is not None:
                    res.steps_failed.append(label)
                return False

        acc = None
        acc_pid = None
        try:
            pythoncom.CoInitialize()
            acc = win32com.client.DispatchEx("Access.Application")
            acc.Visible = self.cfg.access_visible
            # Registrar PID antes de abrir la BD (AutoExec puede colgar y dejar huérfano).
            try:
                import win32process
                from app.excel_engine import current_run_id_var
                from app import com_registry
                _, acc_pid = win32process.GetWindowThreadProcessId(acc.hWndAccessApp)
                com_registry.register(
                    acc_pid, "MSACCESS.EXE", db_path, current_run_id_var.get(),
                )
            except Exception:
                acc_pid = None
            # Bloquear prompts de macros antes de OpenCurrentDatabase (que ya
            # puede disparar AutoExec). msoAutomationSecurityForceDisable = 3.
            try:
                acc.AutomationSecurity = 3
            except Exception:
                pass
            acc.OpenCurrentDatabase(db_path)
            # Dar tiempo a Access para inicializar completamente la BD
            time.sleep(2)

            # Silenciar confirmaciones de action queries / delete / update.
            # Aplica a TODA la sesión de operaciones; el Quit() lo descarta.
            try:
                acc.DoCmd.SetWarnings(False)
            except Exception:
                pass

            # ── Macros pre-importación ─────────────────────────────────────
            if run_pre_macros:
                for macro in self.cfg.pre_import_macros:
                    _attempt(
                        f"Macro pre-import: '{macro}'",
                        lambda m=macro: acc.DoCmd.RunMacro(m),
                    )

            # ── Importaciones guardadas ────────────────────────────────────
            # DoCmd.RunSavedImportExport es SINCRÓNICO: bloquea hasta completar.
            # La clave es que los archivos Excel ya están guardados con datos
            # actualizados gracias al paso anterior.
            if run_imports:
                for spec in self.cfg.saved_imports:
                    _attempt(
                        f"Importación guardada: '{spec}'",
                        lambda s=spec: acc.DoCmd.RunSavedImportExport(s),
                    )

            # ── Macros post-importación ────────────────────────────────────
            if run_post_macros:
                for macro in self.cfg.post_import_macros:
                    _attempt(
                        f"Macro post-import: '{macro}'",
                        lambda m=macro: acc.DoCmd.RunMacro(m),
                    )

            # En modo continue_on_error reportamos éxito parcial: el orquestador
            # decide si el run global es success/failed según steps_failed.
            return True
        except Exception:
            self.log.error(f"[Access] Excepción en operaciones:\n{traceback.format_exc()}")
            return False
        finally:
            if acc:
                try:
                    acc.CloseCurrentDatabase()
                except Exception:
                    pass
                try:
                    acc.Quit()
                except Exception:
                    pass
            try:
                from app import com_registry
                com_registry.unregister(acc_pid)
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # ── Orquestador principal ─────────────────────────────────────────────

    def run(self) -> PipelineResult:
        t0 = time.time()
        res = PipelineResult()

        if not WIN32_AVAILABLE:
            res.error_msg = "win32com no disponible. Instala: pip install pywin32"
            self.log.critical(res.error_msg)
            return res

        # ── Paso 1: Refrescar Excel ────────────────────────────────────────
        res.connections_found = len(self.cfg.excel_files)
        for item in self.cfg.excel_files:
            path = item.get("path", "")
            visible = item.get("visible", False)
            label = Path(path).name

            if not Path(path).exists():
                res.error_msg = f"Archivo Excel no encontrado: {path}"
                res.steps_failed.append(f"Excel: {label}")
                self.log.error(res.error_msg)
                res.duration_s = round(time.time() - t0, 2)
                return res

            self.log.info(f"[Excel] Procesando: {label}")
            if self._refresh_excel_file(path, visible):
                res.steps_done.append(f"Excel: {label}")
            else:
                res.error_msg = f"Error actualizando: {label}"
                res.steps_failed.append(f"Excel: {label}")
                res.duration_s = round(time.time() - t0, 2)
                return res

        # ── Paso 1.5: Preparar .accdb (hidratar OneDrive + esperar lock) ──
        self.log.info("[Access] Verificando disponibilidad de la BD...")
        if not self._prepare_access_db():
            res.error_msg = (
                f"BD Access no disponible (lock o sin hidratar): {self.cfg.access_db}"
            )
            res.steps_failed.append("Preparación .accdb")
            res.duration_s = round(time.time() - t0, 2)
            return res
        res.steps_done.append("Preparación .accdb")

        compact_pos = self.cfg.resolved_compact_position()

        def do_compact() -> bool:
            self.log.info("[Access] Iniciando Compact & Repair...")
            if self._compact_repair(self.cfg.access_db):
                res.steps_done.append("Compact & Repair")
                return True
            res.error_msg = "Error en Compact & Repair"
            res.steps_failed.append("Compact & Repair")
            return False

        def fail_with_ops_error() -> PipelineResult:
            res.steps_failed.append("Operaciones Access")
            if not res.error_msg:
                res.error_msg = "Error en operaciones Access"
            res.duration_s = round(time.time() - t0, 2)
            return res

        # ── Paso 2: Compact pre-macros (modo legacy) ──────────────────────
        if compact_pos == "before_macros":
            if not do_compact():
                res.duration_s = round(time.time() - t0, 2)
                return res

        # ── Pasos 3-7: Operaciones Access ─────────────────────────────────
        # _run_access_operations rellena steps_done/steps_failed por paso.
        # En modo continue_on_error puede devolver True con steps_failed > 0.
        if compact_pos == "after_pre_macros":
            # Sesión 1: solo pre_import_macros (ej. "Elimina Cubos")
            self.log.info("[Access] Operaciones — fase 1 (pre-macros)...")
            if not self._run_access_operations(
                self.cfg.access_db, res,
                run_pre_macros=True, run_imports=False, run_post_macros=False,
            ):
                return fail_with_ops_error()

            # Cerrar Access (lo hizo el finally interno) y compactar.
            # CompactRepair requiere que la BD no esté abierta por otra app.
            if not do_compact():
                res.duration_s = round(time.time() - t0, 2)
                return res

            # Re-preparar tras compact (el archivo fue reemplazado por shutil.move;
            # OneDrive puede re-marcarlo como placeholder en algunas configs).
            if not self._prepare_access_db():
                res.error_msg = "BD no disponible tras Compact & Repair"
                res.steps_failed.append("Re-preparación post-compact")
                res.duration_s = round(time.time() - t0, 2)
                return res

            # Sesión 2: imports + post_import_macros
            self.log.info("[Access] Operaciones — fase 2 (imports + post-macros)...")
            if not self._run_access_operations(
                self.cfg.access_db, res,
                run_pre_macros=False, run_imports=True, run_post_macros=True,
            ):
                return fail_with_ops_error()
        else:
            # "before_macros" o "skip": sesión única con todo
            self.log.info("[Access] Iniciando operaciones en BD...")
            if not self._run_access_operations(self.cfg.access_db, res):
                return fail_with_ops_error()

        # ── Paso 8: Refrescar tableros/herramientas consumidores ──────────
        # Equivalente al "Actualizar Tableros, Reportes y Herramientas" manual:
        # los .xlsm finales leen de Access vía Power Query/conexiones, por lo
        # que requieren refresh DESPUÉS del ETL.
        for item in self.cfg.post_refresh_excel_files:
            path = item.get("path", "")
            visible = item.get("visible", False)
            label = Path(path).name

            if not Path(path).exists():
                msg = f"Tablero post-refresh no encontrado: {path}"
                self.log.error(msg)
                res.steps_failed.append(f"Post-Excel: {label}")
                if not self.cfg.continue_on_error:
                    res.error_msg = msg
                    res.duration_s = round(time.time() - t0, 2)
                    return res
                continue

            self.log.info(f"[Post-Excel] Procesando: {label}")
            if self._refresh_excel_file(path, visible):
                res.steps_done.append(f"Post-Excel: {label}")
            else:
                res.steps_failed.append(f"Post-Excel: {label}")
                if not self.cfg.continue_on_error:
                    res.error_msg = f"Error refrescando tablero: {label}"
                    res.duration_s = round(time.time() - t0, 2)
                    return res

        # En modo continue_on_error: éxito global solo si NINGÚN paso falló;
        # de lo contrario marcamos éxito parcial (success=False) para que el
        # scheduler dispare notificaciones on_error.
        if res.steps_failed:
            res.error_msg = (
                f"Pipeline completado con {len(res.steps_failed)} fallo(s) parcial(es)."
            )
            res.success = False
        else:
            res.success = True

        res.duration_s = round(time.time() - t0, 2)
        self.log.info(
            f"[Pipeline] Terminado en {res.duration_s}s. "
            f"OK: {len(res.steps_done)} · Fallos: {len(res.steps_failed)}."
        )
        return res


# ══════════════════════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA PÚBLICO
# ══════════════════════════════════════════════════════════════════════════════

def run_pipeline(cfg: PipelineConfig, log: logging.Logger) -> PipelineResult:
    """Valida rutas y ejecuta el pipeline completo."""
    if not cfg.access_db or not Path(cfg.access_db).exists():
        msg = f"BD Access no encontrada: {cfg.access_db}"
        log.error(msg)
        return PipelineResult(error_msg=msg)

    return AccessPipelineRunner(cfg, log).run()

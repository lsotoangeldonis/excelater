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
    # Archivos Excel a refrescar en orden: [{"path": "...", "visible": false}, ...]
    excel_files: list[dict] = field(default_factory=list)
    # Base de datos Access
    access_db: str = ""
    access_visible: bool = False
    # Operaciones Access
    compact_before_import: bool = True
    pre_import_macros: list[str] = field(default_factory=list)
    saved_imports: list[str] = field(default_factory=list)
    post_import_macros: list[str] = field(default_factory=list)
    # Timeouts Excel
    excel_refresh_timeout: int = 300
    excel_refresh_check: int = 3
    excel_lock_timeout: int = 120


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
        2. Macros pre-importación (p.ej. 'Ejecutar Elimina Cubos').
        3. Compact & Repair de la BD Access (opcional).
        4. Importaciones guardadas (p.ej. 'Importación: Cubo_SKU_SUC').
        5. Macros post-importación (p.ej. 'Ejecutar ETL Procesos').
    """

    def __init__(self, cfg: PipelineConfig, log: logging.Logger):
        self.cfg = cfg
        self.log = log

    # ── Paso 1: Actualizar Excel ───────────────────────────────────────────

    def _refresh_excel_file(self, file_path: str, visible: bool) -> bool:
        """Refresca un archivo Excel usando COM, esperando TODOS los tipos de conexión."""
        import threading
        from app.excel_engine import ExcelCOMUpdater, EngineConfig

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
        result = ExcelCOMUpdater(ecfg, self.log).run()
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
        try:
            pythoncom.CoInitialize()
            acc = win32com.client.DispatchEx("Access.Application")
            acc.Visible = False
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
        run_pre_macros: bool = True,
        run_saved_imports: bool = True,
        run_post_macros: bool = True,
    ) -> bool:
        """
        Abre la BD Access y ejecuta en orden:
                    - pre_import_macros (opcional)
                    - saved_imports (opcional)
                    - post_import_macros (opcional)
        """
        acc = None
        try:
            pythoncom.CoInitialize()
            acc = win32com.client.DispatchEx("Access.Application")
            acc.Visible = self.cfg.access_visible
            acc.OpenCurrentDatabase(db_path)
            # Dar tiempo a Access para inicializar completamente la BD
            time.sleep(2)

            # ── Macros pre-importación ─────────────────────────────────────
            if run_pre_macros:
                for macro in self.cfg.pre_import_macros:
                    self.log.info(f"[Access] Macro pre-import: '{macro}'")
                    acc.DoCmd.RunMacro(macro)
                    self.log.info(f"[Access] '{macro}' completada.")
                    time.sleep(1)

            # ── Importaciones guardadas ────────────────────────────────────
            # DoCmd.RunSavedImportExport es SINCRÓNICO: bloquea hasta completar.
            # La clave es que los archivos Excel ya están guardados con datos
            # actualizados gracias al paso anterior.
            if run_saved_imports:
                for spec in self.cfg.saved_imports:
                    self.log.info(f"[Access] Importación guardada: '{spec}'")
                    acc.DoCmd.RunSavedImportExport(spec)
                    self.log.info(f"[Access] '{spec}' importada.")
                    time.sleep(1)

            # ── Macros post-importación ────────────────────────────────────
            if run_post_macros:
                for macro in self.cfg.post_import_macros:
                    self.log.info(f"[Access] Macro post-import: '{macro}'")
                    acc.DoCmd.RunMacro(macro)
                    self.log.info(f"[Access] '{macro}' completada.")
                    time.sleep(1)

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

        # ── Paso 2: Macros pre-importación ────────────────────────────────
        if self.cfg.pre_import_macros:
            self.log.info("[Access] Ejecutando macros pre-importación...")
            if self._run_access_operations(
                self.cfg.access_db,
                run_pre_macros=True,
                run_saved_imports=False,
                run_post_macros=False,
            ):
                for m in self.cfg.pre_import_macros:
                    res.steps_done.append(f"Macro pre: {m}")
            else:
                res.error_msg = "Error en macros pre-importación"
                res.steps_failed.append("Macros pre-importación")
                res.duration_s = round(time.time() - t0, 2)
                return res

        # ── Paso 3: Compact & Repair ───────────────────────────────────────
        if self.cfg.compact_before_import:
            self.log.info("[Access] Iniciando Compact & Repair...")
            if self._compact_repair(self.cfg.access_db):
                res.steps_done.append("Compact & Repair")
            else:
                res.error_msg = "Error en Compact & Repair"
                res.steps_failed.append("Compact & Repair")
                res.duration_s = round(time.time() - t0, 2)
                return res

        # ── Pasos 4-5: Importaciones + macros post ────────────────────────
        self.log.info("[Access] Ejecutando importaciones y macros post...")
        if self._run_access_operations(
            self.cfg.access_db,
            run_pre_macros=False,
            run_saved_imports=True,
            run_post_macros=True,
        ):
            for m in self.cfg.pre_import_macros:
                if f"Macro pre: {m}" not in res.steps_done:
                    res.steps_done.append(f"Macro pre: {m}")
            for s in self.cfg.saved_imports:
                res.steps_done.append(f"Import: {s}")
            for m in self.cfg.post_import_macros:
                res.steps_done.append(f"Macro post: {m}")
        else:
            res.steps_failed.append("Operaciones Access")
            if not res.error_msg:
                res.error_msg = "Error en operaciones Access"
            res.duration_s = round(time.time() - t0, 2)
            return res

        res.success = True
        res.duration_s = round(time.time() - t0, 2)
        self.log.info(
            f"[Pipeline] Completado en {res.duration_s}s. "
            f"Pasos OK: {len(res.steps_done)}."
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

"""app/workflows/weekly_excel_copy.py — Flujo semanal de copia y refresco de Excel

Lógica:
  - Lunes: copia el archivo de la semana anterior (Sem N-1) al nombre de la
    semana actual (Sem N) y lo refresca con Excel COM.
  - Otros días + daily_refresh=True: refresca el archivo de la semana actual.
  - Otros días + daily_refresh=False: no-op (retorna success sin hacer nada).

Configuración (pipeline_config JSON):
  {
    "workflow_type": "weekly_excel_copy",
    "folder": "C:\\...\\2. Análisis de Ventas",
    "file_patterns": ["Analisis Ventas The Box Sem {week}.xlsx"],
    "week_padding": 2,
    "daily_refresh": false,
    "fail_if_source_missing": true,
    "excel_visible": false,
    "refresh_timeout": 300
  }
"""
from __future__ import annotations

import shutil
import time
import threading
import traceback
from datetime import datetime, timedelta
from pathlib import Path

from app.excel_engine import EngineConfig, EngineResult, ExcelCOMUpdater, resolve_path
from app.workflows.base import BaseWorkflow

import logging


def _prev_iso_week(today: datetime) -> tuple[int, int]:
    """Devuelve (year, week) ISO de la semana anterior a today."""
    prev = today - timedelta(weeks=1)
    cal = prev.isocalendar()
    return cal.year, cal.week


def _format_week(pattern: str, week: int, padding: int) -> str:
    """Formatea el patrón reemplazando {week} con el número de semana con padding."""
    return pattern.replace("{week}", str(week).zfill(padding))


class WeeklyExcelCopyWorkflow(BaseWorkflow):
    """
    Flujo: duplicar Excel de semana anterior, refrescar y renombrar a semana actual.
    Soporta múltiples file_patterns (uno por archivo a procesar).
    """

    def run(self, config: dict, logger: logging.Logger) -> EngineResult:
        t0 = time.time()

        folder            = resolve_path(config.get("folder", ""))
        file_patterns     = config.get("file_patterns", [])
        week_padding      = int(config.get("week_padding", 2))
        daily_refresh     = bool(config.get("daily_refresh", False))
        fail_missing      = bool(config.get("fail_if_source_missing", True))
        excel_visible     = bool(config.get("excel_visible", False))
        refresh_timeout   = int(config.get("refresh_timeout", 300))
        pivot_max_retries = int(config.get("pivot_max_retries", 3))
        pivot_retry_delay = int(config.get("pivot_retry_delay_s", 60))
        # skip_pivots llega como lista de dicts desde config_overrides (JSON serializable)
        skip_pivots = {
            (d["sheet"], d["pivot"])
            for d in config.get("skip_pivots", [])
        }

        today    = datetime.now()
        cal      = today.isocalendar()
        cur_year, cur_week = cal.year, cal.week

        # force_weekday permite simular un día concreto sin esperar al lunes real
        force_weekday = config.get("force_weekday")
        effective_weekday = int(force_weekday) if force_weekday is not None else today.isoweekday()
        is_monday = effective_weekday == 1

        if force_weekday is not None:
            logger.info(
                f"[WeeklyExcelCopy] ⚠ Modo simulación: weekday forzado a {effective_weekday} "
                f"({'Lunes' if is_monday else 'día no-lunes'}). Fecha real: {today.strftime('%A %d/%m/%Y')}."
            )

        if not is_monday and not daily_refresh:
            logger.info(
                f"[WeeklyExcelCopy] Hoy es {today.strftime('%A %d/%m/%Y')} — "
                "sin tarea programada (daily_refresh=False). Nada que hacer."
            )
            return EngineResult(success=True, duration_s=round(time.time() - t0, 2))

        any_failed = False
        all_pivots_completed: list[dict] = []

        for pattern in file_patterns:
            pat_result = self._process_pattern(
                pattern=pattern,
                folder=folder,
                week_padding=week_padding,
                cur_week=cur_week,
                cur_year=cur_year,
                is_monday=is_monday,
                fail_missing=fail_missing,
                excel_visible=excel_visible,
                refresh_timeout=refresh_timeout,
                pivot_guards=config.get("pivot_guards", []),
                pivot_max_retries=pivot_max_retries,
                pivot_retry_delay=pivot_retry_delay,
                skip_pivots=skip_pivots,
                logger=logger,
            )
            all_pivots_completed.extend(pat_result.pivots_completed)
            if not pat_result.success:
                any_failed = True

        duration = round(time.time() - t0, 2)
        if any_failed:
            return EngineResult(
                success=False,
                error_msg="Uno o más archivos fallaron. Revisa el log para detalles.",
                duration_s=duration,
                pivots_completed=all_pivots_completed,
            )
        return EngineResult(success=True, duration_s=duration, pivots_completed=all_pivots_completed)

    # ── Lógica por patrón ─────────────────────────────────────────────────────

    def _process_pattern(
        self,
        pattern: str,
        folder: str,
        week_padding: int,
        cur_week: int,
        cur_year: int,
        is_monday: bool,
        fail_missing: bool,
        excel_visible: bool,
        refresh_timeout: int,
        pivot_guards: list,
        pivot_max_retries: int,
        pivot_retry_delay: int,
        skip_pivots: set,
        logger: logging.Logger,
    ) -> EngineResult:
        """Procesa un patrón de archivo. Retorna EngineResult."""
        cur_name = _format_week(pattern, cur_week, week_padding)
        cur_path = Path(folder) / cur_name

        if is_monday:
            return self._monday_copy_and_refresh(
                pattern=pattern,
                folder=folder,
                week_padding=week_padding,
                cur_week=cur_week,
                cur_name=cur_name,
                cur_path=cur_path,
                fail_missing=fail_missing,
                excel_visible=excel_visible,
                refresh_timeout=refresh_timeout,
                pivot_guards=pivot_guards,
                pivot_max_retries=pivot_max_retries,
                pivot_retry_delay=pivot_retry_delay,
                skip_pivots=skip_pivots,
                logger=logger,
            )
        else:
            # daily_refresh=True (ya validado antes de llegar aquí)
            if pivot_guards and cur_path.exists():
                self._expand_pivot_space(str(cur_path), pivot_guards, logger)
            return self._refresh_file(
                file_path=str(cur_path),
                label=cur_name,
                fail_missing=fail_missing,
                excel_visible=excel_visible,
                refresh_timeout=refresh_timeout,
                pivot_max_retries=pivot_max_retries,
                pivot_retry_delay=pivot_retry_delay,
                skip_pivots=skip_pivots,
                logger=logger,
            )

    def _monday_copy_and_refresh(
        self,
        pattern: str,
        folder: str,
        week_padding: int,
        cur_week: int,
        cur_name: str,
        cur_path: Path,
        fail_missing: bool,
        excel_visible: bool,
        refresh_timeout: int,
        pivot_guards: list,
        pivot_max_retries: int,
        pivot_retry_delay: int,
        skip_pivots: set,
        logger: logging.Logger,
    ) -> EngineResult:
        """Lunes: copia Sem N-1 → Sem N (si no existe) y refresca."""
        # Calcular semana anterior
        today = datetime.now()
        prev_year, prev_week = _prev_iso_week(today)
        prev_name = _format_week(pattern, prev_week, week_padding)
        prev_path = Path(folder) / prev_name

        logger.info(
            f"[WeeklyExcelCopy] Inicio de semana ISO {cur_week} ({today.strftime('%d/%m/%Y')})."
        )
        logger.info(f"  Origen : {prev_path}")
        logger.info(f"  Destino: {cur_path}")

        # Verificar si el archivo destino ya existe
        if cur_path.exists():
            logger.warning(
                f"[WeeklyExcelCopy] '{cur_name}' ya existe — se omite la copia, "
                "solo se refresca."
            )
        else:
            # Verificar fuente
            if not prev_path.exists():
                msg = (
                    f"[WeeklyExcelCopy] Archivo fuente no encontrado: '{prev_name}'. "
                    f"Se buscó en: {prev_path}"
                )
                if fail_missing:
                    logger.error(msg)
                    return EngineResult(success=False, error_msg=msg)
                else:
                    logger.warning(msg + " — se omite este archivo (fail_if_source_missing=False).")
                    return EngineResult(success=True)

            # Copiar
            try:
                shutil.copy2(str(prev_path), str(cur_path))
                logger.info(f"[WeeklyExcelCopy] Copia completada: '{prev_name}' → '{cur_name}'.")
            except Exception:
                err = traceback.format_exc()
                logger.error(f"[WeeklyExcelCopy] Error al copiar '{prev_name}':\n{err}")
                return EngineResult(success=False, error_msg=err)

        # Refrescar el archivo de la semana actual
        if pivot_guards:
            self._expand_pivot_space(str(cur_path), pivot_guards, logger)
        return self._refresh_file(
            file_path=str(cur_path),
            label=cur_name,
            fail_missing=fail_missing,
            excel_visible=excel_visible,
            refresh_timeout=refresh_timeout,
            pivot_max_retries=pivot_max_retries,
            pivot_retry_delay=pivot_retry_delay,
            skip_pivots=skip_pivots,
            logger=logger,
        )

    def _refresh_file(
        self,
        file_path: str,
        label: str,
        fail_missing: bool,
        excel_visible: bool,
        refresh_timeout: int,
        pivot_max_retries: int,
        pivot_retry_delay: int,
        skip_pivots: set,
        logger: logging.Logger,
    ) -> EngineResult:
        """Refresca un archivo Excel con ExcelCOMUpdater. Retorna EngineResult."""
        if not Path(file_path).exists():
            msg = f"[WeeklyExcelCopy] Archivo no encontrado para refresco: '{label}' ({file_path})"
            if fail_missing:
                logger.error(msg)
                return EngineResult(success=False, error_msg=msg)
            else:
                logger.warning(msg + " — se omite (fail_if_source_missing=False).")
                return EngineResult(success=True)

        logger.info(f"[WeeklyExcelCopy] Refrescando '{label}'...")
        cfg = EngineConfig(
            file_path=file_path,
            refresh_connections=True,
            refresh_pivots=True,
            save_on_success=True,
            excel_visible=excel_visible,
            refresh_timeout=refresh_timeout,
            refresh_check=3,
            lock_timeout=120,
            lock_retry=5,
            lock_max_retries=0,
            stop_event=threading.Event(),
            pivot_max_retries=pivot_max_retries,
            pivot_retry_delay_s=pivot_retry_delay,
            skip_pivots=skip_pivots,
        )
        try:
            result = ExcelCOMUpdater(cfg, logger).run()
        except Exception:
            err = traceback.format_exc()
            logger.error(f"[WeeklyExcelCopy] Error inesperado al refrescar '{label}':\n{err}")
            return EngineResult(success=False, error_msg=err)

        if result.success:
            logger.info(
                f"[WeeklyExcelCopy] '{label}' refrescado correctamente "
                f"({result.connections_found} conexiones, {result.duration_s}s)."
            )
        else:
            logger.error(
                f"[WeeklyExcelCopy] Error al refrescar '{label}': {result.error_msg}"
            )
        return result

    # ── Expansión automática de tabla dinámica ────────────────────────────────

    def _expand_pivot_space(
        self,
        file_path: str,
        guards: list,
        logger: logging.Logger,
    ) -> None:
        """
        Verifica que haya al menos `min_gap` columnas vacías a la derecha de cada
        tabla dinámica configurada en `guards`. Inserta columnas si es necesario.
        Los errores no son fatales: se loguean y el flujo continúa.

        Cada guard: {"sheet": str, "pivot": str, "min_gap": int (default 2)}
        """
        if not guards:
            return

        try:
            import pythoncom
            import win32com.client as win32
        except ImportError:
            logger.warning(
                "[WeeklyExcelCopy][expand] pywin32 no disponible — "
                "se omite la expansión de tabla dinámica."
            )
            return

        pythoncom.CoInitialize()
        xl = None
        wb = None
        try:
            xl = win32.Dispatch("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.AskToUpdateLinks = False

            wb = xl.Workbooks.Open(file_path, UpdateLinks=0, ReadOnly=False)
            modified = False

            for guard in guards:
                sheet_name = guard.get("sheet", "")
                pivot_name = guard.get("pivot", "")
                min_gap    = int(guard.get("min_gap", 2))

                if not sheet_name or not pivot_name:
                    logger.warning(
                        f"[WeeklyExcelCopy][expand] Guard incompleto "
                        f"(sheet='{sheet_name}', pivot='{pivot_name}') — se omite."
                    )
                    continue

                try:
                    ws = wb.Sheets(sheet_name)
                except Exception:
                    logger.warning(
                        f"[WeeklyExcelCopy][expand] Hoja '{sheet_name}' no encontrada — se omite."
                    )
                    continue

                try:
                    pt = ws.PivotTables(pivot_name)
                except Exception:
                    logger.warning(
                        f"[WeeklyExcelCopy][expand] Tabla dinámica '{pivot_name}' "
                        f"no encontrada en hoja '{sheet_name}' — se omite."
                    )
                    continue

                pivot_rng      = pt.TableRange2
                pivot_last_col = pivot_rng.Column + pivot_rng.Columns.Count - 1
                pivot_first_row = pivot_rng.Row
                pivot_last_row  = pivot_rng.Row + pivot_rng.Rows.Count - 1

                logger.info(
                    f"[WeeklyExcelCopy][expand] '{pivot_name}' ('{sheet_name}'): "
                    f"cols {pivot_rng.Column}–{pivot_last_col}, "
                    f"filas {pivot_first_row}–{pivot_last_row}."
                )

                # Medir cuántas columnas vacías hay a la derecha del pivot
                gap = 0
                for offset in range(1, 50):
                    col_idx = pivot_last_col + offset
                    check = ws.Range(
                        ws.Cells(pivot_first_row, col_idx),
                        ws.Cells(pivot_last_row, col_idx),
                    )
                    if xl.WorksheetFunction.CountA(check) > 0:
                        break
                    gap += 1

                cols_to_insert = max(0, min_gap - gap)
                logger.info(
                    f"[WeeklyExcelCopy][expand] Gap actual: {gap} col. "
                    f"Mínimo requerido: {min_gap}. "
                    + (f"Insertando {cols_to_insert} columna(s)." if cols_to_insert > 0
                       else "Sin cambios necesarios.")
                )

                if cols_to_insert > 0:
                    # Insertar antes del primer bloque bloqueador
                    insert_at = pivot_last_col + gap + 1
                    insert_rng = ws.Range(
                        ws.Cells(1, insert_at),
                        ws.Cells(1_048_576, insert_at + cols_to_insert - 1),
                    )
                    insert_rng.Insert(Shift=-4161)  # xlShiftToRight
                    logger.info(
                        f"[WeeklyExcelCopy][expand] {cols_to_insert} columna(s) insertada(s) "
                        f"en col {insert_at} de '{sheet_name}'."
                    )
                    modified = True

            if modified:
                wb.Save()
                logger.info("[WeeklyExcelCopy][expand] Archivo guardado tras expansión.")

            wb.Close(False)

        except Exception:
            logger.error(
                f"[WeeklyExcelCopy][expand] Error inesperado:\n{traceback.format_exc()}"
            )
            if wb:
                try:
                    wb.Close(False)
                except Exception:
                    pass
        finally:
            if xl:
                try:
                    xl.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

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

        folder        = resolve_path(config.get("folder", ""))
        file_patterns = config.get("file_patterns", [])
        week_padding  = int(config.get("week_padding", 2))
        daily_refresh = bool(config.get("daily_refresh", False))
        fail_missing  = bool(config.get("fail_if_source_missing", True))
        excel_visible = bool(config.get("excel_visible", False))
        refresh_timeout = int(config.get("refresh_timeout", 300))

        today    = datetime.now()
        cal      = today.isocalendar()
        cur_year, cur_week = cal.year, cal.week
        is_monday = today.isoweekday() == 1

        if not is_monday and not daily_refresh:
            logger.info(
                f"[WeeklyExcelCopy] Hoy es {today.strftime('%A %d/%m/%Y')} — "
                "sin tarea programada (daily_refresh=False). Nada que hacer."
            )
            return EngineResult(success=True, duration_s=round(time.time() - t0, 2))

        any_failed = False

        for pattern in file_patterns:
            ok = self._process_pattern(
                pattern=pattern,
                folder=folder,
                week_padding=week_padding,
                cur_week=cur_week,
                cur_year=cur_year,
                is_monday=is_monday,
                fail_missing=fail_missing,
                excel_visible=excel_visible,
                refresh_timeout=refresh_timeout,
                logger=logger,
                t0=t0,
            )
            if not ok:
                any_failed = True

        duration = round(time.time() - t0, 2)
        if any_failed:
            return EngineResult(
                success=False,
                error_msg="Uno o más archivos fallaron. Revisa el log para detalles.",
                duration_s=duration,
            )
        return EngineResult(success=True, duration_s=duration)

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
        logger: logging.Logger,
        t0: float,
    ) -> bool:
        """Procesa un patrón de archivo. Retorna True si tuvo éxito."""
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
                logger=logger,
            )
        else:
            # daily_refresh=True (ya validado antes de llegar aquí)
            return self._refresh_file(
                file_path=str(cur_path),
                label=cur_name,
                fail_missing=fail_missing,
                excel_visible=excel_visible,
                refresh_timeout=refresh_timeout,
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
        logger: logging.Logger,
    ) -> bool:
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
                    return False
                else:
                    logger.warning(msg + " — se omite este archivo (fail_if_source_missing=False).")
                    return True

            # Copiar
            try:
                shutil.copy2(str(prev_path), str(cur_path))
                logger.info(f"[WeeklyExcelCopy] Copia completada: '{prev_name}' → '{cur_name}'.")
            except Exception:
                logger.error(
                    f"[WeeklyExcelCopy] Error al copiar '{prev_name}':\n{traceback.format_exc()}"
                )
                return False

        # Refrescar el archivo de la semana actual
        return self._refresh_file(
            file_path=str(cur_path),
            label=cur_name,
            fail_missing=fail_missing,
            excel_visible=excel_visible,
            refresh_timeout=refresh_timeout,
            logger=logger,
        )

    def _refresh_file(
        self,
        file_path: str,
        label: str,
        fail_missing: bool,
        excel_visible: bool,
        refresh_timeout: int,
        logger: logging.Logger,
    ) -> bool:
        """Refresca un archivo Excel con ExcelCOMUpdater. Retorna True si tuvo éxito."""
        if not Path(file_path).exists():
            msg = f"[WeeklyExcelCopy] Archivo no encontrado para refresco: '{label}' ({file_path})"
            if fail_missing:
                logger.error(msg)
                return False
            else:
                logger.warning(msg + " — se omite (fail_if_source_missing=False).")
                return True

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
        )
        try:
            result = ExcelCOMUpdater(cfg, logger).run()
        except Exception:
            logger.error(
                f"[WeeklyExcelCopy] Error inesperado al refrescar '{label}':\n"
                f"{traceback.format_exc()}"
            )
            return False

        if result.success:
            logger.info(
                f"[WeeklyExcelCopy] '{label}' refrescado correctamente "
                f"({result.connections_found} conexiones, {result.duration_s}s)."
            )
        else:
            logger.error(
                f"[WeeklyExcelCopy] Error al refrescar '{label}': {result.error_msg}"
            )
        return result.success

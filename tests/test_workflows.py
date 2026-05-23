"""tests/test_workflows.py — Tests unitarios de workflows personalizados"""
import shutil
import threading
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from app.excel_engine import EngineResult
from app.workflows import registry
from app.workflows.weekly_excel_copy import (
    WeeklyExcelCopyWorkflow,
    _format_week,
    _prev_iso_week,
)


# ══════════════════════════════════════════════════════════════════════════════
# Helpers ISO week
# ══════════════════════════════════════════════════════════════════════════════

def test_format_week_padding():
    assert _format_week("Analisis Ventas Sem {week}.xlsx", 5, 2) == "Analisis Ventas Sem 05.xlsx"
    assert _format_week("Analisis Ventas Sem {week}.xlsx", 20, 2) == "Analisis Ventas Sem 20.xlsx"
    assert _format_week("Analisis Ventas Sem {week}.xlsx", 1, 1) == "Analisis Ventas Sem 1.xlsx"


def test_prev_iso_week_normal():
    # Semana 20 de 2026 → anterior es semana 19 de 2026
    monday_week20 = datetime(2026, 5, 11)  # Lunes semana 20
    prev_year, prev_week = _prev_iso_week(monday_week20)
    assert prev_year == 2026
    assert prev_week == 19


def test_prev_iso_week_year_boundary():
    # Lunes 5 de enero 2026 = semana ISO 2 de 2026
    # Semana anterior = semana 1 de 2026
    monday_week2 = datetime(2026, 1, 5)
    prev_year, prev_week = _prev_iso_week(monday_week2)
    assert prev_year == 2026
    assert prev_week == 1


def test_prev_iso_week_crosses_year():
    # Lunes 29 de diciembre 2025: isocalendar → 2026 semana 1
    # La semana anterior debería ser semana 52 de 2025
    monday_week1_2026 = datetime(2025, 12, 29)
    cal = monday_week1_2026.isocalendar()
    assert cal.year == 2026 and cal.week == 1, "Precondición: debe ser sem 1 de 2026"

    prev_year, prev_week = _prev_iso_week(monday_week1_2026)
    assert prev_year == 2025
    assert prev_week == 52


def test_prev_iso_week_week53():
    # 2015 tuvo semana 53; lunes 4 de enero 2016 = semana 1 de 2016
    # La semana anterior debería ser semana 53 de 2015
    monday_week1_2016 = datetime(2016, 1, 4)
    cal = monday_week1_2016.isocalendar()
    assert cal.year == 2016 and cal.week == 1

    prev_year, prev_week = _prev_iso_week(monday_week1_2016)
    assert prev_year == 2015
    assert prev_week == 53


# ══════════════════════════════════════════════════════════════════════════════
# Registro
# ══════════════════════════════════════════════════════════════════════════════

def test_registry_contains_weekly_excel_copy():
    assert "weekly_excel_copy" in registry.available()
    assert registry.get("weekly_excel_copy") is WeeklyExcelCopyWorkflow


def test_registry_unknown_returns_none():
    assert registry.get("no_existe") is None


# ══════════════════════════════════════════════════════════════════════════════
# Workflow: no-op en día no-lunes con daily_refresh=False
# ══════════════════════════════════════════════════════════════════════════════

def test_noop_non_monday_daily_refresh_false(tmp_path):
    logger = MagicMock()
    cfg = {
        "workflow_type": "weekly_excel_copy",
        "folder": str(tmp_path),
        "file_patterns": ["Sem {week}.xlsx"],
        "week_padding": 2,
        "daily_refresh": False,
        "fail_if_source_missing": True,
        "excel_visible": False,
        "refresh_timeout": 300,
    }
    # Forzar un día que NO sea lunes (miércoles)
    wednesday = datetime(2026, 5, 20)   # miércoles semana 21
    with patch("app.workflows.weekly_excel_copy.datetime") as mock_dt:
        mock_dt.now.return_value = wednesday
        result = WeeklyExcelCopyWorkflow().run(cfg, logger)

    assert result.success is True
    logger.info.assert_called()  # Debe loguear el mensaje "sin tarea programada"


# ══════════════════════════════════════════════════════════════════════════════
# Workflow: falla cuando falta archivo fuente (fail_if_source_missing=True)
# ══════════════════════════════════════════════════════════════════════════════

def test_fail_when_source_missing(tmp_path):
    logger = MagicMock()
    cfg = {
        "workflow_type": "weekly_excel_copy",
        "folder": str(tmp_path),
        "file_patterns": ["Analisis Sem {week}.xlsx"],
        "week_padding": 2,
        "daily_refresh": False,
        "fail_if_source_missing": True,
        "excel_visible": False,
        "refresh_timeout": 300,
    }
    # Forzar lunes semana 20 de 2026 (no existe el archivo sem 19)
    monday = datetime(2026, 5, 11)  # lunes semana 20
    with patch("app.workflows.weekly_excel_copy.datetime") as mock_dt:
        mock_dt.now.return_value = monday
        result = WeeklyExcelCopyWorkflow().run(cfg, logger)

    assert result.success is False
    logger.error.assert_called()


# ══════════════════════════════════════════════════════════════════════════════
# Workflow: continúa cuando falta archivo fuente (fail_if_source_missing=False)
# ══════════════════════════════════════════════════════════════════════════════

def test_skip_when_source_missing_and_not_fail(tmp_path):
    logger = MagicMock()
    cfg = {
        "workflow_type": "weekly_excel_copy",
        "folder": str(tmp_path),
        "file_patterns": ["Analisis Sem {week}.xlsx"],
        "week_padding": 2,
        "daily_refresh": False,
        "fail_if_source_missing": False,
        "excel_visible": False,
        "refresh_timeout": 300,
    }
    monday = datetime(2026, 5, 11)
    with patch("app.workflows.weekly_excel_copy.datetime") as mock_dt:
        mock_dt.now.return_value = monday
        result = WeeklyExcelCopyWorkflow().run(cfg, logger)

    assert result.success is True
    logger.warning.assert_called()


# ══════════════════════════════════════════════════════════════════════════════
# Workflow: copia correctamente el archivo fuente al destino en lunes
# ══════════════════════════════════════════════════════════════════════════════

def test_copy_on_monday(tmp_path):
    """Verifica que en lunes se copia el archivo semana N-1 a semana N."""
    # Crear archivo semana 19
    src = tmp_path / "Analisis Sem 19.xlsx"
    src.write_bytes(b"fake excel content")

    logger = MagicMock()
    cfg = {
        "workflow_type": "weekly_excel_copy",
        "folder": str(tmp_path),
        "file_patterns": ["Analisis Sem {week}.xlsx"],
        "week_padding": 2,
        "daily_refresh": False,
        "fail_if_source_missing": True,
        "excel_visible": False,
        "refresh_timeout": 300,
    }

    monday_week20 = datetime(2026, 5, 11)  # Lunes semana 20 de 2026

    # Mockear ExcelCOMUpdater para no necesitar Excel instalado
    mock_result = EngineResult(success=True, connections_found=3, duration_s=1.5)
    with patch("app.workflows.weekly_excel_copy.datetime") as mock_dt, \
         patch("app.workflows.weekly_excel_copy.ExcelCOMUpdater") as mock_updater:
        mock_dt.now.return_value = monday_week20
        mock_updater.return_value.run.return_value = mock_result

        result = WeeklyExcelCopyWorkflow().run(cfg, logger)

    assert result.success is True
    dest = tmp_path / "Analisis Sem 20.xlsx"
    assert dest.exists(), "El archivo de la semana actual debe haberse creado"


def test_skip_copy_if_dest_already_exists(tmp_path):
    """Si el archivo destino ya existe, no sobreescribe y solo refresca."""
    src = tmp_path / "Analisis Sem 19.xlsx"
    src.write_bytes(b"old content")
    dst = tmp_path / "Analisis Sem 20.xlsx"
    dst.write_bytes(b"existing content")

    logger = MagicMock()
    cfg = {
        "workflow_type": "weekly_excel_copy",
        "folder": str(tmp_path),
        "file_patterns": ["Analisis Sem {week}.xlsx"],
        "week_padding": 2,
        "daily_refresh": False,
        "fail_if_source_missing": True,
        "excel_visible": False,
        "refresh_timeout": 300,
    }

    monday_week20 = datetime(2026, 5, 11)
    mock_result = EngineResult(success=True, connections_found=2, duration_s=1.0)

    with patch("app.workflows.weekly_excel_copy.datetime") as mock_dt, \
         patch("app.workflows.weekly_excel_copy.ExcelCOMUpdater") as mock_updater, \
         patch("app.workflows.weekly_excel_copy.shutil") as mock_shutil:
        mock_dt.now.return_value = monday_week20
        mock_updater.return_value.run.return_value = mock_result

        result = WeeklyExcelCopyWorkflow().run(cfg, logger)

    assert result.success is True
    # No debe haberse llamado a shutil.copy2 (destino ya existía)
    mock_shutil.copy2.assert_not_called()
    logger.warning.assert_called()

"""tests/test_routes.py — Tests de endpoints REST"""
import pytest
from httpx import AsyncClient

TASK_PAYLOAD = {
    "name": "Test Task",
    "description": "Tarea de prueba",
    "file_path": __file__,  # Usamos este mismo archivo como ruta válida
    "schedule_type": "once_daily",
    "schedule_config": {"hour": 8, "minute": 0},
    "refresh_connections": False,
    "refresh_pivots": False,
    "save_on_success": False,
    "excel_visible": False,
    "max_retries": 0,
    "retry_delay_s": 60,
}

REPOSICION_PAYLOAD = {
    "name": "Reposicion Auto",
    "description": "Pipeline de reposicion",
    "schedule_type": "once_daily",
    "schedule_config": {"hour": 7, "minute": 15},
    "access_db": __file__,
    "cubo_sku_suc_maestro": __file__,
    "cubo_sku_suc": __file__,
    "cubo_sku_suc_transferencias": __file__,
    "access_visible": False,
    "compact_before_import": True,
    "excel_refresh_timeout": 300,
    "max_retries": 1,
    "retry_delay_s": 30,
}


async def test_health_ok(client: AsyncClient):
    r = await client.get("/health")
    assert r.status_code == 200
    data = r.json()
    assert data["status"] == "ok"
    assert "scheduler_running" in data
    assert "jobs" in data


async def test_list_tasks_empty(client: AsyncClient):
    r = await client.get("/api/tasks")
    assert r.status_code == 200
    assert r.json() == []


async def test_create_task(client: AsyncClient):
    r = await client.post("/api/tasks", json=TASK_PAYLOAD)
    assert r.status_code == 201
    data = r.json()
    assert data["name"] == "Test Task"
    assert data["status"] == "active"
    assert "id" in data


async def test_create_task_invalid_path(client: AsyncClient):
    payload = {**TASK_PAYLOAD, "file_path": "C:/no/existe/archivo.xlsx"}
    r = await client.post("/api/tasks", json=payload)
    assert r.status_code == 400
    assert "no encontrado" in r.json()["detail"].lower()


async def test_get_task(client: AsyncClient):
    created = (await client.post("/api/tasks", json=TASK_PAYLOAD)).json()
    r = await client.get(f"/api/tasks/{created['id']}")
    assert r.status_code == 200
    assert r.json()["id"] == created["id"]


async def test_get_task_not_found(client: AsyncClient):
    r = await client.get("/api/tasks/no-existe")
    assert r.status_code == 404


async def test_update_task(client: AsyncClient):
    created = (await client.post("/api/tasks", json=TASK_PAYLOAD)).json()
    r = await client.put(f"/api/tasks/{created['id']}", json={"name": "Nombre Actualizado"})
    assert r.status_code == 200
    assert r.json()["name"] == "Nombre Actualizado"


async def test_pause_and_resume_task(client: AsyncClient):
    created = (await client.post("/api/tasks", json=TASK_PAYLOAD)).json()
    task_id = created["id"]

    r = await client.post(f"/api/tasks/{task_id}/pause")
    assert r.status_code == 200
    assert r.json()["status"] == "paused"

    r = await client.post(f"/api/tasks/{task_id}/resume")
    assert r.status_code == 200
    assert r.json()["status"] == "active"


async def test_delete_task_soft(client: AsyncClient):
    created = (await client.post("/api/tasks", json=TASK_PAYLOAD)).json()
    task_id = created["id"]

    r = await client.delete(f"/api/tasks/{task_id}")
    assert r.status_code == 204

    # Ya no debe aparecer en el listado
    tasks = (await client.get("/api/tasks")).json()
    assert all(t["id"] != task_id for t in tasks)

    # Tampoco debe ser accesible por ID
    r = await client.get(f"/api/tasks/{task_id}")
    assert r.status_code == 404


async def test_restore_task(client: AsyncClient):
    created = (await client.post("/api/tasks", json=TASK_PAYLOAD)).json()
    task_id = created["id"]

    await client.delete(f"/api/tasks/{task_id}")
    r = await client.post(f"/api/tasks/{task_id}/restore")
    assert r.status_code == 200
    assert r.json()["status"] == "active"

    # Debe volver a aparecer en el listado
    tasks = (await client.get("/api/tasks")).json()
    assert any(t["id"] == task_id for t in tasks)


async def test_stats(client: AsyncClient):
    await client.post("/api/tasks", json=TASK_PAYLOAD)
    r = await client.get("/api/stats")
    assert r.status_code == 200
    data = r.json()
    assert data["total_tasks"] == 1
    assert data["active_tasks"] == 1


async def test_logs_empty(client: AsyncClient):
    r = await client.get("/api/logs")
    assert r.status_code == 200
    data = r.json()
    assert data["total"] == 0
    assert data["items"] == []


async def test_logs_export_csv(client: AsyncClient):
    r = await client.get("/api/logs/export")
    assert r.status_code == 200
    assert "text/csv" in r.headers["content-type"]
    # Debe contener al menos la cabecera del CSV
    assert "task_name" in r.text


async def test_create_reposicion_pipeline_task(client: AsyncClient):
    r = await client.post("/api/tasks/pipeline/reposicion", json=REPOSICION_PAYLOAD)
    assert r.status_code == 201
    data = r.json()
    assert data["name"] == "Reposicion Auto"
    assert data["task_type"] == "pipeline"
    assert data["status"] == "active"

    pipeline_cfg = data["pipeline_config"]
    assert pipeline_cfg["pre_import_macros"] == ["Ejecutar Elimina Cubos"]
    assert pipeline_cfg["saved_imports"] == [
        "Importación: Cubo_SKU_SUC_Maestro",
        "Importación: Cubo_SKU_SUC",
        "Importación: Cubo_SKU_SUC_Transferencias",
    ]
    assert pipeline_cfg["post_import_macros"] == ["Ejecutar ETL Procesos"]
    assert len(pipeline_cfg["excel_files"]) == 3


async def test_create_reposicion_pipeline_task_invalid_file(client: AsyncClient):
    payload = dict(REPOSICION_PAYLOAD)
    payload["cubo_sku_suc"] = "C:/no/existe/cubo.xlsx"
    r = await client.post("/api/tasks/pipeline/reposicion", json=payload)
    assert r.status_code == 400
    assert "archivo excel no encontrado" in r.json()["detail"].lower()


# ══════════════════════════════════════════════════════════════════════════════
# WORKFLOW: weekly-excel-copy
# ══════════════════════════════════════════════════════════════════════════════

import os

WEEKLY_COPY_PAYLOAD = {
    "name": "Analisis Ventas The Box",
    "description": "Copia y refresco semanal de análisis de ventas",
    "schedule_type": "cron",
    "schedule_config": {"cron": "0 7 * * 1"},
    "folder": os.path.dirname(__file__),   # carpeta de tests, siempre existe
    "file_patterns": ["Analisis Ventas The Box Sem {week}.xlsx"],
    "week_padding": 2,
    "daily_refresh": False,
    "fail_if_source_missing": True,
    "excel_visible": False,
    "refresh_timeout": 300,
    "max_retries": 1,
    "retry_delay_s": 60,
}


async def test_create_weekly_excel_copy_task(client: AsyncClient):
    r = await client.post("/api/tasks/workflow/weekly-excel-copy", json=WEEKLY_COPY_PAYLOAD)
    assert r.status_code == 201
    data = r.json()
    assert data["name"] == "Analisis Ventas The Box"
    assert data["task_type"] == "workflow"
    assert data["status"] == "active"

    cfg = data["pipeline_config"]
    assert cfg["workflow_type"] == "weekly_excel_copy"
    assert cfg["file_patterns"] == ["Analisis Ventas The Box Sem {week}.xlsx"]
    assert cfg["week_padding"] == 2
    assert cfg["daily_refresh"] is False
    assert cfg["fail_if_source_missing"] is True


async def test_create_weekly_excel_copy_invalid_folder(client: AsyncClient):
    payload = {**WEEKLY_COPY_PAYLOAD, "folder": "C:/carpeta/que/no/existe/jamas"}
    r = await client.post("/api/tasks/workflow/weekly-excel-copy", json=payload)
    assert r.status_code == 400
    assert "carpeta no encontrada" in r.json()["detail"].lower()


async def test_create_weekly_excel_copy_invalid_pattern(client: AsyncClient):
    payload = {**WEEKLY_COPY_PAYLOAD, "file_patterns": ["Analisis Ventas Sin Placeholder.xlsx"]}
    r = await client.post("/api/tasks/workflow/weekly-excel-copy", json=payload)
    assert r.status_code == 400
    assert "{week}" in r.json()["detail"]


async def test_create_weekly_excel_copy_multiple_patterns(client: AsyncClient):
    payload = {
        **WEEKLY_COPY_PAYLOAD,
        "file_patterns": [
            "Analisis Ventas The Box Sem {week}.xlsx",
            "Analisis Ventas Lima Sem {week}.xlsx",
        ],
    }
    r = await client.post("/api/tasks/workflow/weekly-excel-copy", json=payload)
    assert r.status_code == 201
    assert len(r.json()["pipeline_config"]["file_patterns"]) == 2

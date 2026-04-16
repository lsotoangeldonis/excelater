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

"""app/routes.py — Endpoints REST del dashboard"""
from __future__ import annotations

import asyncio
import csv
import io
import json
import logging
import subprocess
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Depends, File, Header, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
from sqlalchemy import select, func, desc
from sqlalchemy.ext.asyncio import AsyncSession

from app.config import settings
from app.database import (
    get_db, Task, RunLog, TaskStatus, RunStatus, ScheduleType,
    NotificationRule, ReportSchedule, TriggerType, ChannelType,
)
from app.excel_engine import EngineConfig, run_update, resolve_path
from app.scheduler import (
    add_or_replace_job, remove_job, pause_job, resume_job,
    execute_task, scheduler, cancel_run, add_report_job, remove_report_job,
)
from app.auth import require_reader, require_admin

logger = logging.getLogger(__name__)

router = APIRouter()


# ══════════════════════════════════════════════════════════════════════════════
# AUTENTICACIÓN LEGACY (API Key — mantenida para compatibilidad)
# ══════════════════════════════════════════════════════════════════════════════

async def verify_api_key(
    x_api_key: str = Header(default=""),
    api_key: str = Query(default=""),
):
    """Si API_KEY está configurada, exige que coincida en header o query param."""
    if not settings.api_key:
        return  # Sin auth configurada
    key = x_api_key or api_key
    if key != settings.api_key:
        raise HTTPException(status_code=401, detail="API key inválida o ausente")


# ══════════════════════════════════════════════════════════════════════════════
# SCHEMAS
# ══════════════════════════════════════════════════════════════════════════════

class ScheduleConfig(BaseModel):
    # ONCE_DAILY
    hour: Optional[int] = None
    minute: Optional[int] = 0
    # INTERVAL
    hours: Optional[int] = None
    minutes: Optional[int] = None
    start_hour: Optional[int] = 0
    start_minute: Optional[int] = 0
    # CRON
    cron: Optional[str] = None


class TaskCreate(BaseModel):
    name: str
    description: str = ""
    file_path: str
    schedule_type: ScheduleType
    schedule_config: ScheduleConfig
    refresh_connections: bool = True
    refresh_pivots: bool = True
    save_on_success: bool = True
    excel_visible: bool = False
    max_retries: int = 0
    retry_delay_s: int = 60


class TaskUpdate(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None
    file_path: Optional[str] = None
    schedule_type: Optional[ScheduleType] = None
    schedule_config: Optional[ScheduleConfig] = None
    refresh_connections: Optional[bool] = None
    refresh_pivots: Optional[bool] = None
    save_on_success: Optional[bool] = None
    excel_visible: Optional[bool] = None
    max_retries: Optional[int] = None
    retry_delay_s: Optional[int] = None
    # Pipeline Access ETL
    pipeline_config: Optional[dict] = None


class ExcelFileConfig(BaseModel):
    path: str
    visible: bool = False


class PipelineTaskCreate(BaseModel):
    name: str
    description: str = ""
    schedule_type: ScheduleType
    schedule_config: ScheduleConfig
    # Pipeline-specific
    access_db: str
    excel_files: list[ExcelFileConfig] = []
    access_visible: bool = False
    compact_before_import: bool = True
    pre_import_macros: list[str] = []
    saved_imports: list[str] = []
    post_import_macros: list[str] = []
    excel_refresh_timeout: int = 300
    max_retries: int = 0
    retry_delay_s: int = 60


class ReposicionTaskCreate(BaseModel):
    name: str = "Actualizacion Herramienta Reposicion"
    description: str = "Pipeline automatico de reposicion (Access ETL)"
    schedule_type: ScheduleType
    schedule_config: ScheduleConfig
    access_db: str
    cubo_sku_suc_maestro: str
    cubo_sku_suc: str
    cubo_sku_suc_transferencias: str
    access_visible: bool = False
    compact_before_import: bool = True
    excel_refresh_timeout: int = 300
    max_retries: int = 0
    retry_delay_s: int = 60


def _build_reposicion_pipeline_cfg(body: ReposicionTaskCreate) -> dict:
    return {
        "excel_files": [
            {"path": body.cubo_sku_suc_maestro, "visible": False},
            {"path": body.cubo_sku_suc, "visible": False},
            {"path": body.cubo_sku_suc_transferencias, "visible": False},
        ],
        "access_db": body.access_db,
        "access_visible": body.access_visible,
        "compact_before_import": body.compact_before_import,
        "pre_import_macros": ["Ejecutar Elimina Cubos"],
        "saved_imports": [
            "Importación: Cubo_SKU_SUC_Maestro",
            "Importación: Cubo_SKU_SUC",
            "Importación: Cubo_SKU_SUC_Transferencias",
        ],
        "post_import_macros": ["Ejecutar ETL Procesos"],
        "excel_refresh_timeout": body.excel_refresh_timeout,
    }


def _task_to_dict(task: Task) -> dict:
    cfg = json.loads(task.schedule_config or "{}")
    pipeline_cfg = json.loads(task.pipeline_config or "{}") if getattr(task, "pipeline_config", None) else {}
    return {
        "id": task.id,
        "name": task.name,
        "description": task.description,
        "file_path": task.file_path,
        "task_type": getattr(task, "task_type", "excel") or "excel",
        "pipeline_config": pipeline_cfg,
        "schedule_type": task.schedule_type,
        "schedule_config": cfg,
        "refresh_connections": task.refresh_connections,
        "refresh_pivots": task.refresh_pivots,
        "save_on_success": task.save_on_success,
        "excel_visible": task.excel_visible,
        "max_retries": task.max_retries,
        "retry_delay_s": task.retry_delay_s,
        "retry_count": task.retry_count,
        "status": task.status,
        "created_at": task.created_at,
        "updated_at": task.updated_at,
        "last_run_at": task.last_run_at,
        "last_run_status": getattr(task, "last_run_status", None),
        "next_run_at": task.next_run_at,
    }


def _validate_file_path(raw_path: str) -> str:
    """Resuelve variables de entorno y verifica que el archivo exista."""
    resolved = resolve_path(raw_path)
    if not Path(resolved).exists():
        raise HTTPException(
            status_code=400,
            detail=f"Archivo no encontrado: {resolved}. "
                   "Verifica la ruta o que el archivo esté sincronizado.",
        )
    return raw_path  # Guardar la ruta original (puede tener variables de entorno)


# ══════════════════════════════════════════════════════════════════════════════
# TASKS
# ══════════════════════════════════════════════════════════════════════════════

@router.get("/tasks", dependencies=[Depends(verify_api_key)])
async def list_tasks(db: AsyncSession = Depends(get_db)):
    result = await db.execute(
        select(Task)
        .where(Task.deleted_at.is_(None))
        .order_by(desc(Task.created_at))
    )
    tasks = result.scalars().all()
    return [_task_to_dict(t) for t in tasks]


@router.post("/tasks", status_code=201, dependencies=[Depends(verify_api_key)])
async def create_task(body: TaskCreate, db: AsyncSession = Depends(get_db)):
    _validate_file_path(body.file_path)

    task = Task(
        id=str(uuid.uuid4()),
        name=body.name,
        description=body.description,
        file_path=body.file_path,
        schedule_type=body.schedule_type,
        schedule_config=json.dumps(body.schedule_config.model_dump(exclude_none=True)),
        refresh_connections=body.refresh_connections,
        refresh_pivots=body.refresh_pivots,
        save_on_success=body.save_on_success,
        excel_visible=body.excel_visible,
        max_retries=body.max_retries,
        retry_delay_s=body.retry_delay_s,
        status=TaskStatus.ACTIVE,
    )
    db.add(task)
    await db.flush()

    nxt = add_or_replace_job(task)
    if nxt:
        task.next_run_at = nxt.replace(tzinfo=None)

    await db.commit()
    return _task_to_dict(task)


@router.post("/tasks/pipeline", status_code=201, dependencies=[Depends(verify_api_key)])
async def create_pipeline_task(body: PipelineTaskCreate, db: AsyncSession = Depends(get_db)):
    """Crea una tarea de tipo Pipeline Access ETL (Excel → Access)."""
    # Validar BD Access
    from app.excel_engine import resolve_path
    access_db_resolved = resolve_path(body.access_db)
    if not Path(access_db_resolved).exists():
        raise HTTPException(
            status_code=400,
            detail=f"BD Access no encontrada: {access_db_resolved}",
        )
    # Validar archivos Excel
    for ef in body.excel_files:
        ep = resolve_path(ef.path)
        if not Path(ep).exists():
            raise HTTPException(
                status_code=400,
                detail=f"Archivo Excel no encontrado: {ep}",
            )

    pipeline_cfg = {
        "excel_files": [{"path": ef.path, "visible": ef.visible} for ef in body.excel_files],
        "access_db": body.access_db,
        "access_visible": body.access_visible,
        "compact_before_import": body.compact_before_import,
        "pre_import_macros": body.pre_import_macros,
        "saved_imports": body.saved_imports,
        "post_import_macros": body.post_import_macros,
        "excel_refresh_timeout": body.excel_refresh_timeout,
    }

    task = Task(
        id=str(uuid.uuid4()),
        name=body.name,
        description=body.description,
        file_path=body.access_db,       # BD Access como "archivo principal"
        task_type="pipeline",
        pipeline_config=json.dumps(pipeline_cfg),
        schedule_type=body.schedule_type,
        schedule_config=json.dumps(body.schedule_config.model_dump(exclude_none=True)),
        refresh_connections=False,
        refresh_pivots=False,
        save_on_success=False,
        excel_visible=False,
        max_retries=body.max_retries,
        retry_delay_s=body.retry_delay_s,
        status=TaskStatus.ACTIVE,
    )
    db.add(task)
    await db.flush()

    nxt = add_or_replace_job(task)
    if nxt:
        task.next_run_at = nxt.replace(tzinfo=None)

    await db.commit()
    return _task_to_dict(task)


@router.post("/tasks/pipeline/reposicion", status_code=201, dependencies=[Depends(verify_api_key)])
async def create_reposicion_pipeline_task(body: ReposicionTaskCreate, db: AsyncSession = Depends(get_db)):
    """Crea una tarea pipeline de Reposición con pasos preconfigurados de Access."""
    from app.excel_engine import resolve_path

    access_db_resolved = resolve_path(body.access_db)
    if not Path(access_db_resolved).exists():
        raise HTTPException(
            status_code=400,
            detail=f"BD Access no encontrada: {access_db_resolved}",
        )

    excel_files = [
        body.cubo_sku_suc_maestro,
        body.cubo_sku_suc,
        body.cubo_sku_suc_transferencias,
    ]
    for ef in excel_files:
        ep = resolve_path(ef)
        if not Path(ep).exists():
            raise HTTPException(
                status_code=400,
                detail=f"Archivo Excel no encontrado: {ep}",
            )

    pipeline_cfg = _build_reposicion_pipeline_cfg(body)

    task = Task(
        id=str(uuid.uuid4()),
        name=body.name,
        description=body.description,
        file_path=body.access_db,
        task_type="pipeline",
        pipeline_config=json.dumps(pipeline_cfg),
        schedule_type=body.schedule_type,
        schedule_config=json.dumps(body.schedule_config.model_dump(exclude_none=True)),
        refresh_connections=False,
        refresh_pivots=False,
        save_on_success=False,
        excel_visible=False,
        max_retries=body.max_retries,
        retry_delay_s=body.retry_delay_s,
        status=TaskStatus.ACTIVE,
    )
    db.add(task)
    await db.flush()

    nxt = add_or_replace_job(task)
    if nxt:
        task.next_run_at = nxt.replace(tzinfo=None)

    await db.commit()
    return _task_to_dict(task)


# ══════════════════════════════════════════════════════════════════════════════
# WORKFLOWS PERSONALIZADOS
# ══════════════════════════════════════════════════════════════════════════════

class WeeklyExcelCopyTaskCreate(BaseModel):
    name: str
    description: str = ""
    schedule_type: ScheduleType
    schedule_config: ScheduleConfig
    # Carpeta base donde residen los archivos
    folder: str
    # Lista de patrones de nombre con placeholder {week} (ej: "Análisis Ventas Sem {week}.xlsx")
    file_patterns: list[str]
    # Número de dígitos del número de semana (2 → "05", 1 → "5")
    week_padding: int = 2
    # Si True, refresca el archivo de la semana actual también en días no-lunes
    daily_refresh: bool = False
    # Si True, la tarea falla cuando el archivo fuente no existe; False → warning y continúa
    fail_if_source_missing: bool = True
    excel_visible: bool = False
    refresh_timeout: int = 300
    max_retries: int = 0
    retry_delay_s: int = 60
    # Lista de guardas de tabla dinámica. Cada elemento: {"sheet": str, "pivot": str, "min_gap": int}
    pivot_guards: list[dict] = []


@router.post(
    "/tasks/workflow/weekly-excel-copy",
    status_code=201,
    dependencies=[Depends(verify_api_key)],
)
async def create_weekly_excel_copy_task(
    body: WeeklyExcelCopyTaskCreate,
    db: AsyncSession = Depends(get_db),
):
    """Crea una tarea de tipo workflow: copia semanal de Excel por semana ISO."""
    folder_resolved = resolve_path(body.folder)
    if not Path(folder_resolved).exists():
        raise HTTPException(
            status_code=400,
            detail=f"Carpeta no encontrada: {folder_resolved}",
        )

    for pattern in body.file_patterns:
        if "{week}" not in pattern:
            raise HTTPException(
                status_code=400,
                detail=f"El patrón '{pattern}' no contiene el placeholder {{week}}.",
            )

    workflow_cfg = {
        "workflow_type": "weekly_excel_copy",
        "folder": body.folder,
        "file_patterns": body.file_patterns,
        "week_padding": body.week_padding,
        "daily_refresh": body.daily_refresh,
        "fail_if_source_missing": body.fail_if_source_missing,
        "excel_visible": body.excel_visible,
        "refresh_timeout": body.refresh_timeout,
        "pivot_guards": body.pivot_guards,
    }

    task = Task(
        id=str(uuid.uuid4()),
        name=body.name,
        description=body.description,
        file_path=body.folder,
        task_type="workflow",
        pipeline_config=json.dumps(workflow_cfg),
        schedule_type=body.schedule_type,
        schedule_config=json.dumps(body.schedule_config.model_dump(exclude_none=True)),
        refresh_connections=False,
        refresh_pivots=False,
        save_on_success=False,
        excel_visible=body.excel_visible,
        max_retries=body.max_retries,
        retry_delay_s=body.retry_delay_s,
        status=TaskStatus.ACTIVE,
    )
    db.add(task)
    await db.flush()

    nxt = add_or_replace_job(task)
    if nxt:
        task.next_run_at = nxt.replace(tzinfo=None)

    await db.commit()
    return _task_to_dict(task)


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT / IMPORT DE TAREAS
# ══════════════════════════════════════════════════════════════════════════════

@router.get("/tasks/export", dependencies=[Depends(verify_api_key)])
async def export_tasks(
    ids: Optional[str] = Query(default=None, description="IDs de tarea separados por coma (vacío = todas)"),
    db: AsyncSession = Depends(get_db),
):
    """Exporta la configuración de tareas activas como JSON descargable.
    Solo incluye campos de configuración; omite estado de ejecución e historial."""
    q = select(Task).where(Task.deleted_at.is_(None))
    if ids:
        id_list = [i.strip() for i in ids.split(",") if i.strip()]
        q = q.where(Task.id.in_(id_list))
    tasks = (await db.execute(q)).scalars().all()

    export_data = {
        "version": "1",
        "exported_at": datetime.now(timezone.utc).isoformat(),
        "tasks": [
            {
                "name": t.name,
                "description": t.description,
                "file_path": t.file_path,
                "task_type": getattr(t, "task_type", "excel") or "excel",
                "pipeline_config": json.loads(t.pipeline_config) if getattr(t, "pipeline_config", None) else {},
                "schedule_type": t.schedule_type,
                "schedule_config": json.loads(t.schedule_config or "{}"),
                "refresh_connections": t.refresh_connections,
                "refresh_pivots": t.refresh_pivots,
                "save_on_success": t.save_on_success,
                "excel_visible": t.excel_visible,
                "max_retries": t.max_retries,
                "retry_delay_s": t.retry_delay_s,
            }
            for t in tasks
        ],
    }

    filename = f"excelater_tasks_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.json"
    content = json.dumps(export_data, ensure_ascii=False, indent=2, default=str)
    return StreamingResponse(
        iter([content]),
        media_type="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.post("/tasks/import", dependencies=[Depends(verify_api_key)])
async def import_tasks(
    file: UploadFile = File(...),
    validate_paths: bool = Query(default=True, description="Verificar que los archivos existan en disco"),
    db: AsyncSession = Depends(get_db),
):
    """Importa tareas desde un JSON exportado por Excelater.
    Cada tarea importada recibe un ID nuevo; las existentes no se modifican."""
    try:
        raw = await file.read()
        data = json.loads(raw)
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise HTTPException(400, f"Archivo JSON inválido: {exc}")

    if not isinstance(data, dict) or "tasks" not in data:
        raise HTTPException(400, "Formato inválido: se espera un objeto con clave 'tasks'")
    if data.get("version", "1") != "1":
        raise HTTPException(400, f"Versión de formato no soportada: {data.get('version')}")

    created: list[dict] = []
    skipped: list[dict] = []

    for raw_task in data["tasks"]:
        name = raw_task.get("name")
        if not name:
            skipped.append({"name": "(sin nombre)", "reason": "Falta el campo 'name'"})
            continue

        schedule_type = raw_task.get("schedule_type")
        if not schedule_type:
            skipped.append({"name": name, "reason": "Falta 'schedule_type'"})
            continue

        file_path = raw_task.get("file_path", "")
        if validate_paths and file_path:
            resolved = resolve_path(file_path)
            if not Path(resolved).exists():
                skipped.append({"name": name, "reason": f"Archivo no encontrado: {resolved}"})
                continue

        pipeline_cfg = raw_task.get("pipeline_config") or {}
        task = Task(
            id=str(uuid.uuid4()),
            name=name,
            description=raw_task.get("description", ""),
            file_path=file_path,
            task_type=raw_task.get("task_type", "excel"),
            pipeline_config=json.dumps(pipeline_cfg) if pipeline_cfg else None,
            schedule_type=schedule_type,
            schedule_config=json.dumps(raw_task.get("schedule_config", {})),
            refresh_connections=raw_task.get("refresh_connections", True),
            refresh_pivots=raw_task.get("refresh_pivots", True),
            save_on_success=raw_task.get("save_on_success", True),
            excel_visible=raw_task.get("excel_visible", False),
            max_retries=raw_task.get("max_retries", 0),
            retry_delay_s=raw_task.get("retry_delay_s", 60),
            status=TaskStatus.ACTIVE,
        )
        db.add(task)
        await db.flush()

        nxt = add_or_replace_job(task)
        if nxt:
            task.next_run_at = nxt.replace(tzinfo=None)

        created.append(_task_to_dict(task))

    await db.commit()
    return {
        "imported": len(created),
        "skipped": len(skipped),
        "tasks": created,
        "errors": skipped,
    }


@router.get("/tasks/{task_id}", dependencies=[Depends(verify_api_key)])
async def get_task(task_id: str, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404, "Tarea no encontrada")
    return _task_to_dict(task)


@router.put("/tasks/{task_id}", dependencies=[Depends(verify_api_key)])
async def update_task(task_id: str, body: TaskUpdate, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404, "Tarea no encontrada")

    if body.file_path is not None:
        _validate_file_path(body.file_path)
        task.file_path = body.file_path
    if body.name is not None:
        task.name = body.name
    if body.description is not None:
        task.description = body.description
    if body.schedule_type is not None:
        task.schedule_type = body.schedule_type
    if body.schedule_config is not None:
        task.schedule_config = json.dumps(body.schedule_config.model_dump(exclude_none=True))
    if body.refresh_connections is not None:
        task.refresh_connections = body.refresh_connections
    if body.refresh_pivots is not None:
        task.refresh_pivots = body.refresh_pivots
    if body.save_on_success is not None:
        task.save_on_success = body.save_on_success
    if body.excel_visible is not None:
        task.excel_visible = body.excel_visible
    if body.max_retries is not None:
        task.max_retries = body.max_retries
    if body.retry_delay_s is not None:
        task.retry_delay_s = body.retry_delay_s
    if body.pipeline_config is not None:
        task.pipeline_config = json.dumps(body.pipeline_config)
        # Actualizar también file_path al nuevo access_db si se incluye
        new_access_db = body.pipeline_config.get("access_db")
        if new_access_db:
            task.file_path = new_access_db

    task.updated_at = datetime.now()
    nxt = add_or_replace_job(task)
    if nxt:
        task.next_run_at = nxt.replace(tzinfo=None)

    await db.commit()
    return _task_to_dict(task)


@router.delete("/tasks/{task_id}", status_code=204, dependencies=[Depends(verify_api_key)])
async def delete_task(task_id: str, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404, "Tarea no encontrada")
    remove_job(task_id)
    task.deleted_at = datetime.now()
    task.status = TaskStatus.DISABLED
    await db.commit()


@router.post("/tasks/{task_id}/restore", dependencies=[Depends(verify_api_key)])
async def restore_task(task_id: str, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task:
        raise HTTPException(404, "Tarea no encontrada")
    if task.deleted_at is None:
        raise HTTPException(400, "La tarea no está eliminada")
    task.deleted_at = None
    task.status = TaskStatus.ACTIVE
    task.updated_at = datetime.now()
    nxt = add_or_replace_job(task)
    if nxt:
        task.next_run_at = nxt.replace(tzinfo=None)
    await db.commit()
    return _task_to_dict(task)


@router.post("/tasks/{task_id}/pause", dependencies=[Depends(verify_api_key)])
async def pause_task(task_id: str, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404)
    task.status = TaskStatus.PAUSED
    task.updated_at = datetime.now()
    pause_job(task_id)
    await db.commit()
    return {"status": "paused"}


@router.post("/tasks/{task_id}/resume", dependencies=[Depends(verify_api_key)])
async def resume_task(task_id: str, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404)
    task.status = TaskStatus.ACTIVE
    task.updated_at = datetime.now()
    nxt = add_or_replace_job(task)
    if nxt:
        task.next_run_at = nxt.replace(tzinfo=None)
    await db.commit()
    return {"status": "active"}


@router.post("/tasks/{task_id}/run-now", dependencies=[Depends(verify_api_key)])
async def run_now(task_id: str, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404)
    asyncio.create_task(execute_task(task_id))
    return {"message": "Ejecución iniciada"}


class WorkflowTestRunBody(BaseModel):
    # 1=Lunes … 7=Domingo (ISO). Omitir para usar el día real.
    force_weekday: Optional[int] = None


@router.post("/tasks/{task_id}/test-run", dependencies=[Depends(verify_api_key)])
async def test_run(task_id: str, body: WorkflowTestRunBody, db: AsyncSession = Depends(get_db)):
    """Ejecución de prueba para tareas de tipo workflow, con día de la semana simulable."""
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404)
    task_type = getattr(task, "task_type", None) or "excel"
    if task_type != "workflow":
        raise HTTPException(400, detail="Este endpoint solo está disponible para tareas de tipo workflow.")

    overrides: dict | None = None
    if body.force_weekday is not None:
        if not 1 <= body.force_weekday <= 7:
            raise HTTPException(400, detail="force_weekday debe ser un valor entre 1 (lunes) y 7 (domingo).")
        overrides = {"force_weekday": body.force_weekday}

    asyncio.create_task(execute_task(task_id, config_overrides=overrides))
    return {"message": "Simulación iniciada", "force_weekday": body.force_weekday}


@router.post("/admin/cleanup-stuck-runs", dependencies=[Depends(verify_api_key)])
async def cleanup_stuck_runs(db: AsyncSession = Depends(get_db)):
    """Marca como 'failed' todos los RunLog que quedaron en 'running' sin estar activos en memoria."""
    from app.scheduler import _running_tasks
    result = await db.execute(
        select(RunLog).where(RunLog.status == RunStatus.RUNNING)
    )
    stuck = [r for r in result.scalars().all() if r.id not in _running_tasks]
    now = datetime.now()
    for run in stuck:
        run.status = RunStatus.FAILED
        run.finished_at = now
        run.error_msg = "Ejecucion interrumpida (proceso finalizado sin actualizar estado)"
    await db.commit()
    return {"fixed": len(stuck), "ids": [r.id for r in stuck]}


@router.post("/logs/{run_id}/stop", dependencies=[Depends(verify_api_key)])
async def stop_run(run_id: int, db: AsyncSession = Depends(get_db)):
    run = await db.get(RunLog, run_id)
    if not run:
        raise HTTPException(404, "Ejecucion no encontrada")
    if run.status != RunStatus.RUNNING:
        raise HTTPException(400, "La ejecucion no esta en curso")
    if cancel_run(run_id):
        return {"message": "Detencion solicitada"}
    # El proceso ya no está en memoria (ej. servidor reiniciado): regularizar directamente
    run.status = RunStatus.FAILED
    run.finished_at = datetime.now()
    run.error_msg = "Ejecucion interrumpida (proceso finalizado sin actualizar estado)"
    await db.commit()
    return {"message": "Ejecucion regularizada"}


@router.post("/tasks/{task_id}/dry-run", dependencies=[Depends(verify_api_key)])
async def dry_run(task_id: str, db: AsyncSession = Depends(get_db)):
    """Ejecuta la tarea sin guardar ni registrar en RunLog. Útil para probar configuración."""
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404, "Tarea no encontrada")

    logger = logging.getLogger(f"dry_run.{task_id}")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    log_buffer: list[str] = []

    class ListHandler(logging.Handler):
        def emit(self, record):
            log_buffer.append(self.format(record))

    lh = ListHandler()
    lh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S"))
    logger.addHandler(lh)

    cfg = EngineConfig(
        file_path=task.file_path,
        refresh_connections=task.refresh_connections,
        refresh_pivots=task.refresh_pivots,
        save_on_success=False,  # Nunca guardar en dry-run
        excel_visible=task.excel_visible,
        lock_timeout=settings.lock_timeout_s,
        lock_retry=settings.lock_retry_s,
        refresh_timeout=settings.refresh_timeout_s,
        refresh_check=settings.refresh_check_s,
    )

    result = await asyncio.to_thread(run_update, cfg, logger)
    return {
        "dry_run": True,
        "success": result.success,
        "duration_s": result.duration_s,
        "connections_found": result.connections_found,
        "pivots_ok": result.pivots_ok,
        "pivots_err": result.pivots_err,
        "error_msg": result.error_msg or None,
        "log": "\n".join(log_buffer),
    }


# ══════════════════════════════════════════════════════════════════════════════
# LOGS / HISTORIAL
# ══════════════════════════════════════════════════════════════════════════════

@router.get("/logs", dependencies=[Depends(verify_api_key)])
async def list_logs(
    task_id: Optional[str] = None,
    status: Optional[RunStatus] = None,
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    db: AsyncSession = Depends(get_db),
):
    q = select(RunLog).order_by(desc(RunLog.started_at))
    if task_id:
        q = q.where(RunLog.task_id == task_id)
    if status:
        q = q.where(RunLog.status == status)

    total_q = select(func.count()).select_from(q.subquery())
    total = (await db.execute(total_q)).scalar_one()

    q = q.offset((page - 1) * page_size).limit(page_size)
    rows = (await db.execute(q)).scalars().all()

    return {
        "total": total,
        "page": page,
        "page_size": page_size,
        "items": [
            {
                "id": r.id,
                "task_id": r.task_id,
                "task_name": r.task_name,
                "status": r.status,
                "started_at": r.started_at,
                "finished_at": r.finished_at,
                "duration_s": r.duration_s,
                "log_file": r.log_file,
                "error_msg": r.error_msg,
                "connections": r.connections,
                "pivots_ok": r.pivots_ok,
                "pivots_err": r.pivots_err,
                "retry_attempt": getattr(r, "retry_attempt", 0) or 0,
            }
            for r in rows
        ],
    }


@router.delete("/logs", dependencies=[Depends(verify_api_key)])
async def clear_logs(
    task_id: Optional[str] = None,
    status: Optional[RunStatus] = None,
    db: AsyncSession = Depends(get_db),
):
    """Elimina todos los RunLog (o solo los de una tarea/estado) y sus archivos de log en disco."""
    q = select(RunLog).where(RunLog.status != RunStatus.RUNNING)
    if task_id:
        q = q.where(RunLog.task_id == task_id)
    if status:
        q = q.where(RunLog.status == status)
    rows = (await db.execute(q)).scalars().all()
    deleted = 0
    for r in rows:
        if r.log_file:
            try:
                Path(r.log_file).unlink(missing_ok=True)
            except Exception:
                pass
        await db.delete(r)
        deleted += 1
    await db.commit()
    return {"deleted": deleted}


@router.get("/logs/export", dependencies=[Depends(verify_api_key)])
async def export_logs_csv(
    task_id: Optional[str] = None,
    status: Optional[RunStatus] = None,
    db: AsyncSession = Depends(get_db),
):
    """Exporta el historial de ejecuciones como archivo CSV."""
    q = select(RunLog).order_by(desc(RunLog.started_at))
    if task_id:
        q = q.where(RunLog.task_id == task_id)
    if status:
        q = q.where(RunLog.status == status)
    rows = (await db.execute(q)).scalars().all()

    def generate():
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow([
            "id", "task_id", "task_name", "status",
            "started_at", "finished_at", "duration_s",
            "error_msg", "connections", "pivots_ok", "pivots_err",
        ])
        for r in rows:
            writer.writerow([
                r.id, r.task_id, r.task_name, r.status,
                r.started_at, r.finished_at, r.duration_s,
                r.error_msg or "", r.connections, r.pivots_ok, r.pivots_err,
            ])
        yield buf.getvalue()

    filename = f"excelater_logs_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}.csv"
    return StreamingResponse(
        generate(),
        media_type="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/logs/{run_id}/download", dependencies=[Depends(verify_api_key)])
async def download_log(run_id: int, db: AsyncSession = Depends(get_db)):
    run = await db.get(RunLog, run_id)
    if not run or not run.log_file:
        raise HTTPException(404, "Log no encontrado")
    path = Path(run.log_file)
    if not path.exists():
        raise HTTPException(404, "Archivo de log no existe en disco")
    return FileResponse(
        path,
        media_type="text/plain",
        filename=path.name,
    )


@router.get("/logs/{run_id}/content", dependencies=[Depends(verify_api_key)])
async def view_log(run_id: int, db: AsyncSession = Depends(get_db)):
    run = await db.get(RunLog, run_id)
    if not run or not run.log_file:
        raise HTTPException(404)
    path = Path(run.log_file)
    if not path.exists():
        return {"content": "(archivo de log no disponible)"}
    return {"content": path.read_text(encoding="utf-8", errors="replace")}


@router.get("/logs/{run_id}/tail", dependencies=[Depends(verify_api_key)])
async def tail_log(run_id: int, offset: int = 0, db: AsyncSession = Depends(get_db)):
    """Devuelve el contenido del log desde `offset` bytes hasta el final.
    Permite al frontend hacer polling incremental para ver logs en tiempo real."""
    run = await db.get(RunLog, run_id)
    if not run:
        raise HTTPException(404, "Ejecución no encontrada")
    if not run.log_file:
        return {"content": "", "offset": 0, "status": run.status}
    path = Path(run.log_file)
    if not path.exists():
        return {"content": "", "offset": 0, "status": run.status}
    size = path.stat().st_size
    if offset >= size:
        return {"content": "", "offset": size, "status": run.status}
    with path.open("rb") as f:
        f.seek(offset)
        chunk = f.read(size - offset)
    return {
        "content": chunk.decode("utf-8", errors="replace"),
        "offset": size,
        "status": run.status,
    }


# ══════════════════════════════════════════════════════════════════════════════
# ESTADÍSTICAS
# ══════════════════════════════════════════════════════════════════════════════

@router.get("/stats", dependencies=[Depends(verify_api_key)])
async def get_stats(db: AsyncSession = Depends(get_db)):
    total_tasks = (await db.execute(
        select(func.count(Task.id)).where(Task.deleted_at.is_(None))
    )).scalar_one()
    active_tasks = (await db.execute(
        select(func.count(Task.id)).where(Task.status == TaskStatus.ACTIVE, Task.deleted_at.is_(None))
    )).scalar_one()
    total_runs = (await db.execute(select(func.count(RunLog.id)))).scalar_one()
    success_runs = (await db.execute(
        select(func.count(RunLog.id)).where(RunLog.status == RunStatus.SUCCESS)
    )).scalar_one()
    failed_runs = (await db.execute(
        select(func.count(RunLog.id)).where(RunLog.status == RunStatus.FAILED)
    )).scalar_one()
    running_now = (await db.execute(
        select(func.count(RunLog.id)).where(RunLog.status == RunStatus.RUNNING)
    )).scalar_one()

    return {
        "total_tasks": total_tasks,
        "active_tasks": active_tasks,
        "paused_tasks": total_tasks - active_tasks,
        "total_runs": total_runs,
        "success_runs": success_runs,
        "failed_runs": failed_runs,
        "running_now": running_now,
        "success_rate": round(success_runs / total_runs * 100, 1) if total_runs else 0,
    }


# ══════════════════════════════════════════════════════════════════════════════
# FILE BROWSER (selector nativo de Windows)
# ══════════════════════════════════════════════════════════════════════════════

_FILTER_MAP = {
    "excel":  "Archivos Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Todos los archivos (*.*)|*.*",
    "access": "Bases de datos Access (*.accdb;*.mdb)|*.accdb;*.mdb|Todos los archivos (*.*)|*.*",
    "any":    "Todos los archivos (*.*)|*.*",
}


@router.get("/browse-file", dependencies=[Depends(verify_api_key)])
async def browse_file(filter: str = Query(default="any")):
    """Abre el diálogo de apertura de archivos de Windows y devuelve la ruta seleccionada.
    Solo funciona cuando el servidor corre en Windows en la misma máquina que el usuario."""
    if sys.platform != "win32":
        raise HTTPException(400, "El selector de archivos solo está disponible en Windows")

    file_filter = _FILTER_MAP.get(filter, _FILTER_MAP["any"])
    ps_script = (
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        "Add-Type -AssemblyName System.Windows.Forms; "
        "$f = New-Object System.Windows.Forms.Form; "
        "$f.TopMost = $true; "
        "$f.Opacity = 0; "
        "$f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; "
        "$f.Show(); "
        "$d = New-Object System.Windows.Forms.OpenFileDialog; "
        f"$d.Filter = '{file_filter}'; "
        "$d.Title = 'Seleccionar archivo'; "
        "if ($d.ShowDialog($f) -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.FileName } "
        "$f.Dispose();"
    )
    # Ruta absoluta para evitar FileNotFoundError cuando el servidor no tiene powershell en PATH
    ps_exe = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    try:
        result = await asyncio.get_running_loop().run_in_executor(
            None,
            lambda: subprocess.run(
                [ps_exe, "-Sta", "-NoProfile", "-Command", ps_script],
                capture_output=True, text=True, encoding="utf-8", timeout=120,
            ),
        )
        path = result.stdout.strip()
        if not path:
            if result.returncode != 0:
                logger.warning("[browse-file] PowerShell stderr: %s", result.stderr.strip())
            return {"path": None}
        return {"path": path}
    except subprocess.TimeoutExpired:
        return {"path": None}
    except Exception as exc:
        logger.error("[browse-file] Error: %s", exc)
        raise HTTPException(500, f"Error al abrir el selector: {exc}")


@router.get("/browse-folder", dependencies=[Depends(verify_api_key)])
async def browse_folder():
    """Abre el diálogo de selección de carpetas de Windows y devuelve la ruta seleccionada."""
    if sys.platform != "win32":
        raise HTTPException(400, "El selector de carpetas solo está disponible en Windows")

    ps_script = (
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        "Add-Type -AssemblyName System.Windows.Forms; "
        "$f = New-Object System.Windows.Forms.Form; "
        "$f.TopMost = $true; "
        "$f.Opacity = 0; "
        "$f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; "
        "$f.Show(); "
        "$d = New-Object System.Windows.Forms.FolderBrowserDialog; "
        "$d.Description = 'Seleccionar carpeta'; "
        "$d.ShowNewFolderButton = $false; "
        "if ($d.ShowDialog($f) -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.SelectedPath } "
        "$f.Dispose();"
    )
    ps_exe = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    try:
        result = await asyncio.get_running_loop().run_in_executor(
            None,
            lambda: subprocess.run(
                [ps_exe, "-Sta", "-NoProfile", "-Command", ps_script],
                capture_output=True, text=True, encoding="utf-8", timeout=120,
            ),
        )
        path = result.stdout.strip()
        if not path:
            if result.returncode != 0:
                logger.warning("[browse-folder] PowerShell stderr: %s", result.stderr.strip())
            return {"path": None}
        return {"path": path}
    except subprocess.TimeoutExpired:
        return {"path": None}
    except Exception as exc:
        logger.error("[browse-folder] Error: %s", exc)
        raise HTTPException(500, f"Error al abrir el selector de carpetas: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# NOTIFICATION RULES
# ══════════════════════════════════════════════════════════════════════════════

class NotificationRuleCreate(BaseModel):
    trigger: TriggerType
    channel: ChannelType
    recipients: list  # list[str] para email, list[dict] para whatsapp
    enabled: bool = True


class NotificationRuleUpdate(BaseModel):
    trigger: Optional[TriggerType] = None
    channel: Optional[ChannelType] = None
    recipients: Optional[list] = None
    enabled: Optional[bool] = None


def _rule_to_dict(r: NotificationRule) -> dict:
    return {
        "id": r.id,
        "task_id": r.task_id,
        "trigger": r.trigger,
        "channel": r.channel,
        "recipients": json.loads(r.recipients or "[]"),
        "enabled": r.enabled,
        "created_at": r.created_at,
    }


@router.get("/tasks/{task_id}/notifications", dependencies=[Depends(verify_api_key)])
async def list_notification_rules(task_id: str, db: AsyncSession = Depends(get_db)):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404, "Tarea no encontrada")
    result = await db.execute(
        select(NotificationRule).where(NotificationRule.task_id == task_id)
    )
    return [_rule_to_dict(r) for r in result.scalars()]


@router.post("/tasks/{task_id}/notifications", status_code=201, dependencies=[Depends(verify_api_key)])
async def create_notification_rule(
    task_id: str, body: NotificationRuleCreate, db: AsyncSession = Depends(get_db)
):
    task = await db.get(Task, task_id)
    if not task or task.deleted_at is not None:
        raise HTTPException(404, "Tarea no encontrada")
    rule = NotificationRule(
        task_id=task_id,
        trigger=body.trigger,
        channel=body.channel,
        recipients=json.dumps(body.recipients),
        enabled=body.enabled,
    )
    db.add(rule)
    await db.commit()
    await db.refresh(rule)
    return _rule_to_dict(rule)


@router.put("/notifications/{rule_id}", dependencies=[Depends(verify_api_key)])
async def update_notification_rule(
    rule_id: int, body: NotificationRuleUpdate, db: AsyncSession = Depends(get_db)
):
    rule = await db.get(NotificationRule, rule_id)
    if not rule:
        raise HTTPException(404, "Regla no encontrada")
    if body.trigger is not None:
        rule.trigger = body.trigger
    if body.channel is not None:
        rule.channel = body.channel
    if body.recipients is not None:
        rule.recipients = json.dumps(body.recipients)
    if body.enabled is not None:
        rule.enabled = body.enabled
    await db.commit()
    return _rule_to_dict(rule)


@router.delete("/notifications/{rule_id}", status_code=204, dependencies=[Depends(verify_api_key)])
async def delete_notification_rule(rule_id: int, db: AsyncSession = Depends(get_db)):
    rule = await db.get(NotificationRule, rule_id)
    if not rule:
        raise HTTPException(404, "Regla no encontrada")
    await db.delete(rule)
    await db.commit()


# ══════════════════════════════════════════════════════════════════════════════
# REPORT SCHEDULES
# ══════════════════════════════════════════════════════════════════════════════

class ReportScheduleCreate(BaseModel):
    name: str
    schedule_type: ScheduleType
    schedule_config: ScheduleConfig
    lookback_hours: int = 24
    channel: ChannelType
    recipients: list
    task_ids: Optional[list[str]] = None  # None = todas las tareas
    enabled: bool = True


class ReportScheduleUpdate(BaseModel):
    name: Optional[str] = None
    schedule_type: Optional[ScheduleType] = None
    schedule_config: Optional[ScheduleConfig] = None
    lookback_hours: Optional[int] = None
    channel: Optional[ChannelType] = None
    recipients: Optional[list] = None
    task_ids: Optional[list[str]] = None
    enabled: Optional[bool] = None


def _report_to_dict(s: ReportSchedule) -> dict:
    return {
        "id": s.id,
        "name": s.name,
        "schedule_type": s.schedule_type,
        "schedule_config": json.loads(s.schedule_config or "{}"),
        "lookback_hours": s.lookback_hours,
        "channel": s.channel,
        "recipients": json.loads(s.recipients or "[]"),
        "task_ids": json.loads(s.task_ids) if s.task_ids else None,
        "enabled": s.enabled,
        "created_at": s.created_at,
    }


@router.get("/reports", dependencies=[Depends(verify_api_key)])
async def list_reports(db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(ReportSchedule).order_by(desc(ReportSchedule.created_at)))
    return [_report_to_dict(s) for s in result.scalars()]


@router.post("/reports", status_code=201, dependencies=[Depends(verify_api_key)])
async def create_report(body: ReportScheduleCreate, db: AsyncSession = Depends(get_db)):
    s = ReportSchedule(
        name=body.name,
        schedule_type=body.schedule_type,
        schedule_config=json.dumps(body.schedule_config.model_dump(exclude_none=True)),
        lookback_hours=body.lookback_hours,
        channel=body.channel,
        recipients=json.dumps(body.recipients),
        task_ids=json.dumps(body.task_ids) if body.task_ids is not None else None,
        enabled=body.enabled,
    )
    db.add(s)
    await db.commit()
    await db.refresh(s)
    if s.enabled:
        add_report_job(s)
    return _report_to_dict(s)


@router.put("/reports/{report_id}", dependencies=[Depends(verify_api_key)])
async def update_report(
    report_id: int, body: ReportScheduleUpdate, db: AsyncSession = Depends(get_db)
):
    s = await db.get(ReportSchedule, report_id)
    if not s:
        raise HTTPException(404, "Reporte no encontrado")
    if body.name is not None:
        s.name = body.name
    if body.schedule_type is not None:
        s.schedule_type = body.schedule_type
    if body.schedule_config is not None:
        s.schedule_config = json.dumps(body.schedule_config.model_dump(exclude_none=True))
    if body.lookback_hours is not None:
        s.lookback_hours = body.lookback_hours
    if body.channel is not None:
        s.channel = body.channel
    if body.recipients is not None:
        s.recipients = json.dumps(body.recipients)
    if body.task_ids is not None:
        s.task_ids = json.dumps(body.task_ids)
    if body.enabled is not None:
        s.enabled = body.enabled
    await db.commit()
    if s.enabled:
        add_report_job(s)
    else:
        remove_report_job(s.id)
    return _report_to_dict(s)


@router.delete("/reports/{report_id}", status_code=204, dependencies=[Depends(verify_api_key)])
async def delete_report(report_id: int, db: AsyncSession = Depends(get_db)):
    s = await db.get(ReportSchedule, report_id)
    if not s:
        raise HTTPException(404, "Reporte no encontrado")
    remove_report_job(report_id)
    await db.delete(s)
    await db.commit()


@router.post("/reports/{report_id}/run-now", dependencies=[Depends(verify_api_key)])
async def run_report_now(report_id: int, db: AsyncSession = Depends(get_db)):
    s = await db.get(ReportSchedule, report_id)
    if not s:
        raise HTTPException(404, "Reporte no encontrado")
    from app.notifications import send_report
    asyncio.create_task(send_report(s, db))
    return {"message": "Reporte en proceso"}

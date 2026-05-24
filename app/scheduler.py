"""app/scheduler.py — Gestión de tareas con APScheduler"""
from __future__ import annotations

import asyncio
import json
import logging
import logging.handlers
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path

import httpx
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.date import DateTrigger
from sqlalchemy import select

from app.config import settings
from app.database import (
    AsyncSessionLocal, Task, RunLog, RunStatus, ScheduleType, TaskStatus,
    NotificationRule, ReportSchedule,
)
from app.excel_engine import EngineConfig, run_update


scheduler = AsyncIOScheduler(timezone=settings.timezone)

# Mapeo run_id -> asyncio.Task para poder cancelar ejecuciones en curso
_running_tasks: dict[int, asyncio.Task] = {}
# Mapeo run_id -> threading.Event para señalizar parada real al hilo
_stop_events: dict[int, threading.Event] = {}


# ══════════════════════════════════════════════════════════════════════════════
# LOGGING POR TAREA
# ══════════════════════════════════════════════════════════════════════════════

def make_task_logger(task_id: str, task_name: str, run_id: int) -> tuple[logging.Logger, Path]:
    log_path = settings.logs_path / f"task_{task_id}_{run_id}.log"
    logger = logging.getLogger(f"task.{task_id}.{run_id}")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S")

    fh = logging.handlers.RotatingFileHandler(
        log_path,
        maxBytes=settings.max_log_size_mb * 1024 * 1024,
        backupCount=settings.log_backup_count,
        encoding="utf-8",
    )
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    logger.info(f"=== INICIO TAREA: {task_name} (id={task_id}) ===")
    return logger, log_path


# ══════════════════════════════════════════════════════════════════════════════
# WEBHOOK
# ══════════════════════════════════════════════════════════════════════════════

async def send_webhook(payload: dict):
    """Envía una notificación POST al webhook configurado. Silencia errores."""
    if not settings.webhook_url:
        return
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            await client.post(settings.webhook_url, json=payload)
    except Exception as exc:
        logging.getLogger("excelater").warning(f"Webhook falló: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# EJECUCIÓN DE TAREA
# ══════════════════════════════════════════════════════════════════════════════

async def execute_task(task_id: str, config_overrides: dict | None = None):
    async with AsyncSessionLocal() as db:
        task = await db.get(Task, task_id)
        if not task or task.status != TaskStatus.ACTIVE:
            return

        # Crear RunLog inicial
        run = RunLog(
            task_id=task_id,
            task_name=task.name,
            status=RunStatus.RUNNING,
            started_at=datetime.now(),
            retry_attempt=task.retry_count,  # 0=original, 1=primer reintento, etc.
        )
        db.add(run)
        await db.flush()
        run_id = run.id

        logger, log_path = make_task_logger(task_id, task.name, run_id)
        run.log_file = str(log_path)

        # Si es un reintento, cargar pivots ya completados del run fallido anterior
        prev_pivots_completed: list[dict] = []
        if task.retry_count > 0:
            prev_run = await db.scalar(
                select(RunLog)
                .where(RunLog.task_id == task_id, RunLog.status == RunStatus.FAILED)
                .order_by(RunLog.id.desc())
                .limit(1)
            )
            if prev_run and prev_run.pivots_completed:
                try:
                    prev_pivots_completed = json.loads(prev_run.pivots_completed)
                except Exception:
                    prev_pivots_completed = []
            if prev_pivots_completed:
                logger.info(
                    f"[Retry] Saltando {len(prev_pivots_completed)} pivot(s) ya completados "
                    "en el run anterior."
                )

        await db.commit()

    # Ejecutar fuera de la sesión DB para no bloquearla
    task_type = getattr(task, "task_type", None) or "excel"

    _running_tasks[run_id] = asyncio.current_task()
    t0 = time.time()
    cancelled = False

    if task_type == "pipeline":
        # ── Pipeline Access ETL ──────────────────────────────────────────
        from app.access_engine import PipelineConfig, run_pipeline
        pipeline_cfg_dict = json.loads(task.pipeline_config or "{}")
        pipe_cfg = PipelineConfig(
            excel_files=pipeline_cfg_dict.get("excel_files", []),
            access_db=pipeline_cfg_dict.get("access_db", ""),
            access_visible=pipeline_cfg_dict.get("access_visible", False),
            compact_before_import=pipeline_cfg_dict.get("compact_before_import", True),
            pre_import_macros=pipeline_cfg_dict.get("pre_import_macros", []),
            saved_imports=pipeline_cfg_dict.get("saved_imports", []),
            post_import_macros=pipeline_cfg_dict.get("post_import_macros", []),
            excel_refresh_timeout=pipeline_cfg_dict.get("excel_refresh_timeout", settings.refresh_timeout_s),
            excel_refresh_check=pipeline_cfg_dict.get("excel_refresh_check", settings.refresh_check_s),
            excel_lock_timeout=pipeline_cfg_dict.get("excel_lock_timeout", settings.lock_timeout_s),
        )
        try:
            result = await asyncio.to_thread(run_pipeline, pipe_cfg, logger)
        except asyncio.CancelledError:
            cancelled = True
            logger.warning("Ejecución detenida manualmente.")
        finally:
            _running_tasks.pop(run_id, None)
    elif task_type == "workflow":
        # ── Workflow personalizado (registry) ────────────────────────────
        from app.workflows import registry
        workflow_cfg = json.loads(task.pipeline_config or "{}")
        # Aplicar overrides en tiempo de ejecución (ej: force_weekday para pruebas)
        if config_overrides:
            workflow_cfg = {**workflow_cfg, **config_overrides}
        # Inyectar pivots completados del run anterior para skip en retry
        if prev_pivots_completed:
            workflow_cfg["skip_pivots"] = prev_pivots_completed
        workflow_type_name = workflow_cfg.get("workflow_type", "")
        handler_cls = registry.get(workflow_type_name)
        if handler_cls is None:
            result = type(
                "EngineResult", (),
                {
                    "success": False,
                    "error_msg": f"Workflow desconocido: '{workflow_type_name}'. "
                                 f"Disponibles: {registry.available()}",
                    "duration_s": 0.0,
                    "connections_found": 0,
                    "pivots_ok": 0,
                    "pivots_err": 0,
                },
            )()
            _running_tasks.pop(run_id, None)
        else:
            try:
                result = await asyncio.to_thread(handler_cls().run, workflow_cfg, logger)
            except asyncio.CancelledError:
                cancelled = True
                logger.warning("Ejecución detenida manualmente.")
            finally:
                _running_tasks.pop(run_id, None)
    else:
        # ── Tarea Excel estándar ─────────────────────────────────────────
        stop_event = threading.Event()
        cfg = EngineConfig(
            file_path=task.file_path,
            refresh_connections=task.refresh_connections,
            refresh_pivots=task.refresh_pivots,
            save_on_success=task.save_on_success,
            excel_visible=task.excel_visible,
            lock_timeout=settings.lock_timeout_s,
            lock_retry=settings.lock_retry_s,
            lock_max_retries=settings.lock_max_retries,
            refresh_timeout=settings.refresh_timeout_s,
            refresh_check=settings.refresh_check_s,
            stop_event=stop_event,
        )
        _stop_events[run_id] = stop_event
        try:
            # Ejecutar en un hilo separado para no bloquear el event loop de asyncio
            result = await asyncio.to_thread(run_update, cfg, logger)
        except asyncio.CancelledError:
            cancelled = True
            logger.warning("Ejecución detenida manualmente.")
        finally:
            _running_tasks.pop(run_id, None)
            _stop_events.pop(run_id, None)

    if cancelled:
        async with AsyncSessionLocal() as db:
            run = await db.get(RunLog, run_id)
            if run:
                run.status = RunStatus.CANCELLED
                run.finished_at = datetime.now()
                run.duration_s = round(time.time() - t0, 2)
                run.error_msg = "Detenida manualmente"
                await db.commit()
                task = await db.get(Task, task_id)
                if task:
                    task.last_run_status = RunStatus.CANCELLED.value
                    await db.commit()
        return

    finished = datetime.now()

    async with AsyncSessionLocal() as db:
        run = await db.get(RunLog, run_id)
        run.status = RunStatus.SUCCESS if result.success else RunStatus.FAILED
        run.finished_at = finished
        run.duration_s = result.duration_s
        run.error_msg = result.error_msg or None
        run.connections = result.connections_found
        run.pivots_ok = result.pivots_ok
        run.pivots_err = result.pivots_err
        completed = getattr(result, "pivots_completed", None)
        run.pivots_completed = json.dumps(completed) if completed else None

        task = await db.get(Task, task_id)
        task.last_run_at = finished
        task.last_run_status = (RunStatus.SUCCESS if result.success else RunStatus.FAILED).value

        # ── Retry automático ──────────────────────────────────────────────
        if not result.success and task.max_retries > 0:
            if task.retry_count < task.max_retries:
                task.retry_count += 1
                delay = task.retry_delay_s
                retry_time = datetime.now() + timedelta(seconds=delay)
                scheduler.add_job(
                    execute_task,
                    trigger=DateTrigger(run_date=retry_time),
                    id=f"{task_id}_retry_{task.retry_count}",
                    name=f"{task.name} (reintento {task.retry_count}/{task.max_retries})",
                    kwargs={"task_id": task_id},
                    replace_existing=True,
                )
                # Próxima ejecución = momento del reintento (tiene prioridad sobre schedule regular)
                task.next_run_at = retry_time
                logger.warning(
                    f"Reintento {task.retry_count}/{task.max_retries} "
                    f"programado en {delay}s ({retry_time.strftime('%H:%M:%S')})."
                )
            else:
                task.retry_count = 0
                logger.error("Se agotaron todos los reintentos.")
                # Restaurar próxima ejecución al schedule regular
                job = scheduler.get_job(task_id)
                if job and job.next_run_time:
                    task.next_run_at = job.next_run_time.replace(tzinfo=None)
        else:
            # Éxito o sin retries: próxima ejecución = schedule regular
            task.retry_count = 0 if result.success else task.retry_count
            job = scheduler.get_job(task_id)
            if job and job.next_run_time:
                task.next_run_at = job.next_run_time.replace(tzinfo=None)

        await db.commit()

    logger.info(
        f"=== FIN: {'ÉXITO' if result.success else 'FALLIDO'} "
        f"({result.duration_s}s) ==="
    )

    # ── Notificación webhook global ───────────────────────────────────────
    should_notify = (
        (not result.success and settings.notify_on_failure) or
        (result.success and settings.notify_on_success)
    )
    if should_notify:
        await send_webhook({
            "task_id": task_id,
            "task_name": task.name,
            "status": "success" if result.success else "failed",
            "duration_s": result.duration_s,
            "error_msg": result.error_msg or None,
            "connections": result.connections_found,
            "pivots_ok": result.pivots_ok,
            "pivots_err": result.pivots_err,
            "timestamp": finished.isoformat(),
        })

    # ── Notificaciones por regla (email / whatsapp) ───────────────────────
    from app.notifications import dispatch_notification  # import tardío para evitar circular
    async with AsyncSessionLocal() as db:
        run_for_notify = await db.get(RunLog, run_id)
        rules_result = await db.execute(
            select(NotificationRule).where(
                NotificationRule.task_id == task_id,
                NotificationRule.enabled.is_(True),
            )
        )
        for rule in rules_result.scalars():
            try:
                await dispatch_notification(rule, run_for_notify, db)
            except Exception as exc:
                logging.getLogger("excelater").warning(f"Error en notificación regla {rule.id}: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DE TRIGGERS
# ══════════════════════════════════════════════════════════════════════════════

def build_trigger(schedule_type: ScheduleType, schedule_config: dict):
    """
    schedule_config según tipo:

    ONCE_DAILY:
        {"hour": 6, "minute": 0}

    INTERVAL:
        {"hours": 1, "start_hour": 6, "start_minute": 0}
        {"minutes": 30, "start_hour": 8, "start_minute": 0}

    CRON:
        {"cron": "0 6 * * 1-5"}   → expresión cron estándar de 5 campos
    """
    if schedule_type == ScheduleType.ONCE_DAILY:
        return CronTrigger(
            hour=schedule_config.get("hour", 6),
            minute=schedule_config.get("minute", 0),
            timezone=settings.timezone,
        )

    if schedule_type == ScheduleType.INTERVAL:
        hours = schedule_config.get("hours", 0)
        minutes = schedule_config.get("minutes", 60)
        sh = schedule_config.get("start_hour", 0)
        sm = schedule_config.get("start_minute", 0)
        if hours >= 1:
            return CronTrigger(
                hour=f"{sh}-23/{int(hours)}",
                minute=sm,
                timezone=settings.timezone,
            )
        else:
            step = int(minutes) if minutes else 60
            return CronTrigger(
                minute=f"*/{step}",
                hour=f"{sh}-23",
                timezone=settings.timezone,
            )

    if schedule_type == ScheduleType.CRON:
        expr = schedule_config.get("cron", "0 6 * * *")
        parts = expr.strip().split()
        return CronTrigger(
            minute=parts[0],
            hour=parts[1],
            day=parts[2],
            month=parts[3],
            day_of_week=parts[4],
            timezone=settings.timezone,
        )

    raise ValueError(f"Tipo de programación desconocido: {schedule_type}")


# ══════════════════════════════════════════════════════════════════════════════
# API DEL SCHEDULER
# ══════════════════════════════════════════════════════════════════════════════

def add_or_replace_job(task: Task):
    config = json.loads(task.schedule_config or "{}")
    trigger = build_trigger(task.schedule_type, config)

    if scheduler.get_job(task.id):
        scheduler.remove_job(task.id)

    if task.status == TaskStatus.ACTIVE:
        job = scheduler.add_job(
            execute_task,
            trigger=trigger,
            id=task.id,
            name=task.name,
            kwargs={"task_id": task.id},
            replace_existing=True,
            misfire_grace_time=300,
        )
        return job.next_run_time


def remove_job(task_id: str):
    if scheduler.get_job(task_id):
        scheduler.remove_job(task_id)


def pause_job(task_id: str):
    job = scheduler.get_job(task_id)
    if job:
        job.pause()


def resume_job(task_id: str):
    job = scheduler.get_job(task_id)
    if job:
        job.resume()


def cancel_run(run_id: int) -> bool:
    """Cancela una ejecución en curso. Devuelve True si se encontró y canceló."""
    stop_event = _stop_events.get(run_id)
    if stop_event:
        stop_event.set()  # Señal al hilo para que detenga el trabajo en el próximo ciclo

    task = _running_tasks.get(run_id)
    if task and not task.done():
        task.cancel()
        return True
    return False


async def load_all_tasks():
    """Carga todas las tareas activas y reportes programados al iniciar el servicio."""
    async with AsyncSessionLocal() as db:
        result = await db.execute(
            select(Task).where(Task.status == TaskStatus.ACTIVE, Task.deleted_at.is_(None))
        )
        tasks = result.scalars().all()
        for task in tasks:
            nxt = add_or_replace_job(task)
            if nxt:
                task.next_run_at = nxt.replace(tzinfo=None)
        await db.commit()
    print(f"[Scheduler] {len(tasks)} tarea(s) cargada(s).")

    await load_all_reports()


# ══════════════════════════════════════════════════════════════════════════════
# REPORTES PROGRAMADOS
# ══════════════════════════════════════════════════════════════════════════════

def add_report_job(schedule: ReportSchedule):
    """Registra o reemplaza el job de un reporte programado."""
    config = json.loads(schedule.schedule_config or "{}")
    trigger = build_trigger(schedule.schedule_type, config)
    scheduler.add_job(
        _run_report,
        trigger=trigger,
        id=f"report_{schedule.id}",
        name=f"Reporte: {schedule.name}",
        kwargs={"schedule_id": schedule.id},
        replace_existing=True,
        misfire_grace_time=300,
    )


def remove_report_job(schedule_id: int):
    job_id = f"report_{schedule_id}"
    if scheduler.get_job(job_id):
        scheduler.remove_job(job_id)


async def _run_report(schedule_id: int):
    """Ejecuta un reporte programado."""
    from app.notifications import send_report  # import tardío para evitar circular
    async with AsyncSessionLocal() as db:
        schedule = await db.get(ReportSchedule, schedule_id)
        if schedule and schedule.enabled:
            try:
                await send_report(schedule, db)
            except Exception as exc:
                logging.getLogger("excelater").error(f"Error en reporte {schedule_id}: {exc}")


async def load_all_reports():
    """Carga todos los reportes programados habilitados."""
    async with AsyncSessionLocal() as db:
        result = await db.execute(
            select(ReportSchedule).where(ReportSchedule.enabled.is_(True))
        )
        schedules = result.scalars().all()
        for s in schedules:
            add_report_job(s)
    print(f"[Scheduler] {len(schedules)} reporte(s) programado(s) cargado(s).")

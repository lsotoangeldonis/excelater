"""app/notifications.py — Servicio de notificaciones: Email (SMTP) y WhatsApp (CallMeBot)"""
from __future__ import annotations

import json
import logging
import urllib.parse
from datetime import datetime, timezone, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import httpx
from sqlalchemy import select, func
from sqlalchemy.ext.asyncio import AsyncSession

from app.config import settings
from app.database import (
    AsyncSessionLocal, NotificationRule, ReportSchedule,
    RunLog, RunStatus, Task, TriggerType, ChannelType,
)

log = logging.getLogger("excelater.notifications")


# ══════════════════════════════════════════════════════════════════════════════
# ENVÍO EMAIL
# ══════════════════════════════════════════════════════════════════════════════

async def send_email(to_list: list[str], subject: str, body_html: str):
    """Envía un email via SMTP async. Si SMTP no está configurado, avisa y retorna."""
    if not settings.smtp_host:
        log.warning("Email no configurado (SMTP_HOST vacío). Notificación omitida.")
        return
    if not to_list:
        return

    try:
        import aiosmtplib
    except ImportError:
        log.error("aiosmtplib no instalado. Ejecuta: poetry install")
        return

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = settings.smtp_from or settings.smtp_user
    msg["To"] = ", ".join(to_list)
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    try:
        await aiosmtplib.send(
            msg,
            hostname=settings.smtp_host,
            port=settings.smtp_port,
            username=settings.smtp_user or None,
            password=settings.smtp_password or None,
            start_tls=settings.smtp_tls,
        )
        log.info(f"Email enviado a: {', '.join(to_list)}")
    except Exception as exc:
        log.error(f"Error enviando email: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# ENVÍO WHATSAPP (CallMeBot)
# ══════════════════════════════════════════════════════════════════════════════

async def send_whatsapp(recipients: list[dict], message: str):
    """
    Envía un mensaje WhatsApp via CallMeBot a cada destinatario.
    recipients: [{"phone": "51999123456", "apikey": "abc123"}, ...]
    """
    if not recipients:
        return

    encoded = urllib.parse.quote(message)
    async with httpx.AsyncClient(timeout=15) as client:
        for r in recipients:
            phone = r.get("phone", "").strip()
            apikey = r.get("apikey", "").strip()
            if not phone or not apikey:
                log.warning(f"Destinatario WhatsApp inválido: {r}")
                continue
            url = f"https://api.callmebot.com/whatsapp.php?phone={phone}&text={encoded}&apikey={apikey}"
            try:
                resp = await client.get(url)
                if resp.status_code == 200:
                    log.info(f"WhatsApp enviado a +{phone}")
                else:
                    log.warning(f"CallMeBot respondió {resp.status_code} para +{phone}: {resp.text[:200]}")
            except Exception as exc:
                log.error(f"Error enviando WhatsApp a +{phone}: {exc}")


# ══════════════════════════════════════════════════════════════════════════════
# FORMATO DE MENSAJES
# ══════════════════════════════════════════════════════════════════════════════

def _format_task_message_text(run: RunLog) -> str:
    icon = "OK" if run.status == RunStatus.SUCCESS else "FALLO"
    lines = [
        f"[{icon}] {run.task_name}",
        f"Duracion: {run.duration_s}s | Conexiones: {run.connections} | Pivots: {run.pivots_ok} ok, {run.pivots_err} err",
        f"Hora: {run.finished_at.strftime('%Y-%m-%d %H:%M:%S') if run.finished_at else 'N/A'}",
    ]
    if run.error_msg:
        lines.append(f"Error: {run.error_msg[:300]}")
    return "\n".join(lines)


def _format_task_message_html(run: RunLog) -> str:
    color = "#22c55e" if run.status == RunStatus.SUCCESS else "#ef4444"
    label = "EXITO" if run.status == RunStatus.SUCCESS else "FALLO"
    error_row = f"<tr><td><b>Error</b></td><td style='color:#ef4444'>{run.error_msg[:500]}</td></tr>" if run.error_msg else ""
    return f"""
<html><body style="font-family:sans-serif;max-width:600px">
  <h2 style="color:{color};">[{label}] {run.task_name}</h2>
  <table border="0" cellpadding="6" style="border-collapse:collapse;width:100%">
    <tr><td><b>Estado</b></td><td style="color:{color}"><b>{label}</b></td></tr>
    <tr><td><b>Duracion</b></td><td>{run.duration_s}s</td></tr>
    <tr><td><b>Conexiones</b></td><td>{run.connections}</td></tr>
    <tr><td><b>Pivots</b></td><td>{run.pivots_ok} correctas, {run.pivots_err} con error</td></tr>
    <tr><td><b>Finalizo</b></td><td>{run.finished_at.strftime('%Y-%m-%d %H:%M:%S') if run.finished_at else 'N/A'}</td></tr>
    {error_row}
  </table>
  <p style="color:#9ca3af;font-size:12px">Enviado por Excelater</p>
</body></html>"""


def _format_report_text(runs: list[RunLog], schedule_name: str, lookback_hours: int,
                         desde: datetime, hasta: datetime) -> str:
    total = len(runs)
    success = sum(1 for r in runs if r.status == RunStatus.SUCCESS)
    failed = sum(1 for r in runs if r.status == RunStatus.FAILED)
    rate = round(success / total * 100, 1) if total else 0

    # Resumen por tarea
    por_tarea: dict[str, dict] = {}
    for r in runs:
        if r.task_name not in por_tarea:
            por_tarea[r.task_name] = {"ok": 0, "err": 0}
        if r.status == RunStatus.SUCCESS:
            por_tarea[r.task_name]["ok"] += 1
        elif r.status == RunStatus.FAILED:
            por_tarea[r.task_name]["err"] += 1

    lines = [
        f"[REPORTE] {schedule_name}",
        f"Periodo: ultimas {lookback_hours}h ({desde.strftime('%d/%m %H:%M')} - {hasta.strftime('%d/%m %H:%M')})",
        f"Total: {total} | OK: {success} | Error: {failed} | Tasa: {rate}%",
        "",
        "Por tarea:",
    ]
    for nombre, counts in por_tarea.items():
        lines.append(f"  {nombre}: {counts['ok']} ok, {counts['err']} err")

    return "\n".join(lines)


def _format_report_html(runs: list[RunLog], schedule_name: str, lookback_hours: int,
                         desde: datetime, hasta: datetime) -> str:
    total = len(runs)
    success = sum(1 for r in runs if r.status == RunStatus.SUCCESS)
    failed = sum(1 for r in runs if r.status == RunStatus.FAILED)
    rate = round(success / total * 100, 1) if total else 0

    por_tarea: dict[str, dict] = {}
    for r in runs:
        if r.task_name not in por_tarea:
            por_tarea[r.task_name] = {"ok": 0, "err": 0}
        if r.status == RunStatus.SUCCESS:
            por_tarea[r.task_name]["ok"] += 1
        elif r.status == RunStatus.FAILED:
            por_tarea[r.task_name]["err"] += 1

    rows = "".join(
        f"<tr><td>{n}</td><td style='color:#22c55e'>{c['ok']}</td><td style='color:#ef4444'>{c['err']}</td></tr>"
        for n, c in por_tarea.items()
    )

    return f"""
<html><body style="font-family:sans-serif;max-width:600px">
  <h2>Reporte Excelater</h2>
  <p><b>{schedule_name}</b> &mdash; Ultimas {lookback_hours}h
    ({desde.strftime('%d/%m %H:%M')} &rarr; {hasta.strftime('%d/%m %H:%M')})</p>
  <table border="0" cellpadding="6" style="border-collapse:collapse;width:100%">
    <tr><td><b>Total ejecuciones</b></td><td>{total}</td></tr>
    <tr><td><b>Exitosas</b></td><td style="color:#22c55e"><b>{success}</b></td></tr>
    <tr><td><b>Fallidas</b></td><td style="color:#ef4444"><b>{failed}</b></td></tr>
    <tr><td><b>Tasa de exito</b></td><td><b>{rate}%</b></td></tr>
  </table>
  <h3>Por tarea</h3>
  <table border="1" cellpadding="6" style="border-collapse:collapse;width:100%">
    <tr><th>Tarea</th><th>OK</th><th>Error</th></tr>
    {rows}
  </table>
  <p style="color:#9ca3af;font-size:12px">Enviado por Excelater</p>
</body></html>"""


# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA DE DISPARO
# ══════════════════════════════════════════════════════════════════════════════

async def _should_notify(rule: NotificationRule, run: RunLog, db: AsyncSession) -> bool:
    t = rule.trigger
    if t == TriggerType.ALWAYS:
        return True
    if t == TriggerType.ON_ERROR:
        return run.status == RunStatus.FAILED
    if t == TriggerType.ON_SUCCESS:
        return run.status == RunStatus.SUCCESS
    if t == TriggerType.FIRST_RUN_OF_DAY:
        # Primera ejecución del día si no hay otros runs completados hoy
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)  # hora local
        count = (await db.execute(
            select(func.count(RunLog.id)).where(
                RunLog.task_id == run.task_id,
                RunLog.status.in_([RunStatus.SUCCESS, RunStatus.FAILED]),
                RunLog.started_at >= today,
                RunLog.id != run.id,
            )
        )).scalar_one()
        return count == 0
    if t == TriggerType.ON_FINAL_FAILURE:
        # Solo si el fallo ocurrió en el último intento permitido por max_retries.
        # Si max_retries=0 (sin retries), el fallo original ya es el final.
        if run.status != RunStatus.FAILED:
            return False
        task = await db.get(Task, run.task_id)
        if task is None:
            return False
        return (run.retry_attempt or 0) >= (task.max_retries or 0)
    return False


async def dispatch_notification(rule: NotificationRule, run: RunLog, db: AsyncSession):
    """Evalúa el trigger y envía la notificación si aplica."""
    if not await _should_notify(rule, run, db):
        return

    recipients = json.loads(rule.recipients or "[]")
    if not recipients:
        log.warning(f"Regla {rule.id} sin destinatarios, omitida.")
        return

    if rule.channel == ChannelType.EMAIL:
        subject = f"[Excelater] {'OK' if run.status == RunStatus.SUCCESS else 'FALLO'} — {run.task_name}"
        await send_email(recipients, subject, _format_task_message_html(run))
    elif rule.channel == ChannelType.WHATSAPP:
        await send_whatsapp(recipients, _format_task_message_text(run))


# ══════════════════════════════════════════════════════════════════════════════
# REPORTES PROGRAMADOS
# ══════════════════════════════════════════════════════════════════════════════

async def send_report(schedule: ReportSchedule, db: AsyncSession):
    """Genera y envía el reporte de resumen del schedule dado."""
    hasta = datetime.now()
    desde = hasta - timedelta(hours=schedule.lookback_hours)

    q = select(RunLog).where(RunLog.started_at >= desde)
    if schedule.task_ids:
        ids = json.loads(schedule.task_ids)
        if ids:
            q = q.where(RunLog.task_id.in_(ids))

    runs = (await db.execute(q)).scalars().all()
    recipients = json.loads(schedule.recipients or "[]")

    if not recipients:
        log.warning(f"ReportSchedule {schedule.id} sin destinatarios.")
        return

    if schedule.channel == ChannelType.EMAIL:
        subject = f"[Excelater] Reporte: {schedule.name}"
        html = _format_report_html(runs, schedule.name, schedule.lookback_hours, desde, hasta)
        await send_email(recipients, subject, html)
    elif schedule.channel == ChannelType.WHATSAPP:
        text = _format_report_text(runs, schedule.name, schedule.lookback_hours, desde, hasta)
        await send_whatsapp(recipients, text)

    log.info(f"Reporte '{schedule.name}' enviado ({len(runs)} ejecuciones).")

"""app/main.py — Punto de entrada de la aplicación"""
import asyncio
import logging
import uvicorn
from contextlib import asynccontextmanager
from datetime import datetime
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pathlib import Path
from sqlalchemy import select

from app.config import settings
from app.database import init_db, engine, AsyncSessionLocal, RunLog, RunStatus
from app.scheduler import scheduler, load_all_tasks
from app.routes import router
from app.auth_routes import auth_router

logger = logging.getLogger("excelater")


async def _cleanup_stuck_runs():
    """Al iniciar, marca como 'failed' cualquier RunLog que quedó en 'running' de una sesión anterior."""
    async with AsyncSessionLocal() as db:
        result = await db.execute(select(RunLog).where(RunLog.status == RunStatus.RUNNING))
        stuck = result.scalars().all()
        if stuck:
            now = datetime.utcnow()
            for run in stuck:
                run.status = RunStatus.FAILED
                run.finished_at = now
                run.error_msg = "Ejecucion interrumpida (servidor reiniciado)"
            await db.commit()
            logger.warning(f"[Startup] {len(stuck)} ejecucion(es) colgada(s) marcada(s) como fallidas: {[r.id for r in stuck]}")


# Tiempo máximo (segundos) para esperar que las tareas en curso terminen al apagar
SHUTDOWN_TIMEOUT = 30


@asynccontextmanager
async def lifespan(app: FastAPI):
    # ── Startup ───────────────────────────────────────────────────────────────
    await init_db()
    await _cleanup_stuck_runs()
    # Limpiar PIDs huérfanos del registro COM (procesos que ya no existen).
    try:
        from app import com_registry
        removed = com_registry.prune_dead()
        if removed:
            print(f"[Server] com_registry: {removed} entrada(s) muerta(s) purgadas.")
    except Exception as e:
        print(f"[Server] com_registry.prune_dead falló: {e}")
    scheduler.start()
    await load_all_tasks()
    print(f"[Server] Dashboard en http://{settings.host}:{settings.port}")
    print("[Server] Presiona Ctrl+C para detener.")

    yield

    # ── Shutdown ──────────────────────────────────────────────────────────────
    print("\n[Server] Deteniendo... espera que terminen las tareas en curso.")

    # 1. Detener el scheduler esperando hasta SHUTDOWN_TIMEOUT segundos
    try:
        await asyncio.wait_for(
            asyncio.to_thread(scheduler.shutdown, True),  # wait=True
            timeout=SHUTDOWN_TIMEOUT,
        )
        print("[Server] Scheduler detenido correctamente.")
    except asyncio.TimeoutError:
        # Si tardó demasiado, forzar cierre
        scheduler.shutdown(wait=False)
        logger.warning(
            f"[Server] El scheduler no terminó en {SHUTDOWN_TIMEOUT}s. "
            "Algunas tareas podrían haber quedado incompletas."
        )

    # 2. Cerrar el pool de conexiones de la base de datos
    await engine.dispose()
    print("[Server] Base de datos cerrada.")
    print("[Server] Servicio detenido. ¡Hasta luego!")


app = FastAPI(
    title="Excel OneDrive Scheduler",
    version="1.0.0",
    lifespan=lifespan,
)

# CORS configurable vía CORS_ORIGINS en .env
_origins = [o.strip() for o in settings.cors_origins.split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_origins,
    allow_methods=["*"],
    allow_headers=["*"],
)

# API
app.include_router(auth_router, prefix="/api")
app.include_router(router, prefix="/api")


# ── Health check ──────────────────────────────────────────────────────────────
@app.get("/health")
async def health():
    return {
        "status": "ok",
        "scheduler_running": scheduler.running,
        "jobs": len(scheduler.get_jobs()),
        "version": "1.0.0",
        "timezone": settings.timezone,
    }


# ── Servir frontend estático ──────────────────────────────────────────────────
static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/static", StaticFiles(directory=static_dir), name="static")

    @app.get("/{full_path:path}")
    async def serve_spa(full_path: str):
        # No interceptar rutas de la API — dejar que el router las maneje
        if full_path.startswith("api") or full_path == "api":
            from fastapi import HTTPException
            raise HTTPException(status_code=404, detail="Not found")
        # Ruta explícita para login
        if full_path in ("login", "login.html"):
            login_page = static_dir / "login.html"
            if login_page.exists():
                return FileResponse(login_page)
        index = static_dir / "index.html"
        return FileResponse(index)


def start():
    uvicorn.run(
        "app.main:app",
        host=settings.host,
        port=settings.port,
        reload=settings.debug,
        timeout_graceful_shutdown=SHUTDOWN_TIMEOUT + 5,  # dar margen extra a uvicorn
    )


if __name__ == "__main__":
    start()

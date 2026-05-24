"""app/database.py — Modelos SQLAlchemy + setup de base de datos SQLite async"""
from datetime import datetime
from sqlalchemy import (
    Column, String, Integer, Boolean, DateTime, Text, Float, Enum as _SAEnum
)

# SQLAlchemy 2.x usa enum.name (mayúsculas) por defecto; forzar enum.value (minúsculas)
# para compatibilidad con los datos existentes en la DB.
def SAEnum(enum_cls):
    return _SAEnum(enum_cls, values_callable=lambda obj: [e.value for e in obj])
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession, async_sessionmaker
from sqlalchemy.orm import DeclarativeBase
import enum

from app.config import settings


# ── Motor async ───────────────────────────────────────────────────────────────
engine = create_async_engine(settings.db_url, echo=False)
AsyncSessionLocal = async_sessionmaker(engine, expire_on_commit=False)


# ── Base ORM ──────────────────────────────────────────────────────────────────
class Base(DeclarativeBase):
    pass


# ── Enums ─────────────────────────────────────────────────────────────────────
class TaskStatus(str, enum.Enum):
    ACTIVE   = "active"
    PAUSED   = "paused"
    DISABLED = "disabled"


class RunStatus(str, enum.Enum):
    RUNNING   = "running"
    SUCCESS   = "success"
    FAILED    = "failed"
    SKIPPED   = "skipped"
    CANCELLED = "cancelled"


class ScheduleType(str, enum.Enum):
    ONCE_DAILY   = "once_daily"    # Una vez al día a una hora
    INTERVAL     = "interval"      # Repetir cada N minutos/horas
    CRON         = "cron"          # Expresión cron libre


class TriggerType(str, enum.Enum):
    ALWAYS           = "always"           # Toda ejecución (éxito o fallo)
    ON_ERROR         = "on_error"         # Solo fallos (cada intento)
    ON_SUCCESS       = "on_success"       # Solo éxitos
    FIRST_RUN_OF_DAY = "first_run_of_day" # Solo la primera ejecución del día
    ON_FINAL_FAILURE = "on_final_failure" # Solo cuando se agotaron los reintentos


class ChannelType(str, enum.Enum):
    EMAIL    = "email"
    WHATSAPP = "whatsapp"


class UserRole(str, enum.Enum):
    SUPERUSER = "superuser"   # Acceso total + gestión de admins
    ADMIN     = "admin"       # Acceso total + gestión de usuarios lectores
    READER    = "reader"      # Solo lectura


# ── Tablas ────────────────────────────────────────────────────────────────────
class Task(Base):
    __tablename__ = "tasks"

    id               = Column(String, primary_key=True)
    name             = Column(String, nullable=False)
    description      = Column(Text, default="")
    file_path        = Column(String, nullable=False)

    # Programación
    schedule_type    = Column(SAEnum(ScheduleType), nullable=False)
    schedule_config  = Column(Text, default="{}")   # JSON serializado

    # Comportamiento
    refresh_connections = Column(Boolean, default=True)
    refresh_pivots      = Column(Boolean, default=True)
    save_on_success     = Column(Boolean, default=True)
    excel_visible       = Column(Boolean, default=False)

    # Pipeline Access ETL
    task_type        = Column(String, default="excel")  # "excel" | "pipeline"
    pipeline_config  = Column(Text, default="{}")       # JSON con config del pipeline

    # Retry
    max_retries      = Column(Integer, default=0)   # 0 = sin retry
    retry_delay_s    = Column(Integer, default=60)  # Segundos entre reintentos
    retry_count      = Column(Integer, default=0)   # Contador de reintentos actuales

    # Estado
    status           = Column(SAEnum(TaskStatus), default=TaskStatus.ACTIVE)
    created_at       = Column(DateTime, default=datetime.now)
    updated_at       = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    last_run_at      = Column(DateTime, nullable=True)
    last_run_status  = Column(String, nullable=True)   # "success" | "failed" | "cancelled"
    next_run_at      = Column(DateTime, nullable=True)
    deleted_at       = Column(DateTime, nullable=True)  # Soft-delete


class RunLog(Base):
    __tablename__ = "run_logs"

    id             = Column(Integer, primary_key=True, autoincrement=True)
    task_id        = Column(String, nullable=False)
    task_name      = Column(String, nullable=False)
    status         = Column(SAEnum(RunStatus), nullable=False)
    started_at     = Column(DateTime, default=datetime.now)
    finished_at    = Column(DateTime, nullable=True)
    duration_s     = Column(Float, nullable=True)
    log_file       = Column(String, nullable=True)   # ruta al .log
    error_msg      = Column(Text, nullable=True)
    connections        = Column(Integer, default=0)
    pivots_ok          = Column(Integer, default=0)
    pivots_err         = Column(Integer, default=0)
    pivots_completed   = Column(Text, nullable=True)  # JSON: [{"sheet": .., "pivot": ..}, ...]
    retry_attempt      = Column(Integer, default=0)  # 0 = ejecución original, 1+ = reintento N


class NotificationRule(Base):
    __tablename__ = "notification_rules"

    id         = Column(Integer, primary_key=True, autoincrement=True)
    task_id    = Column(String, nullable=False)
    trigger    = Column(SAEnum(TriggerType), nullable=False)
    channel    = Column(SAEnum(ChannelType), nullable=False)
    # email:    ["correo@ejemplo.com", ...]
    # whatsapp: [{"phone": "51999...", "apikey": "abc123"}, ...]
    recipients = Column(Text, nullable=False, default="[]")
    enabled    = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)


class ReportSchedule(Base):
    __tablename__ = "report_schedules"

    id             = Column(Integer, primary_key=True, autoincrement=True)
    name           = Column(String, nullable=False)
    schedule_type  = Column(SAEnum(ScheduleType), nullable=False)
    schedule_config = Column(Text, default="{}")
    lookback_hours = Column(Integer, default=24)  # Ventana del resumen
    channel        = Column(SAEnum(ChannelType), nullable=False)
    recipients     = Column(Text, nullable=False, default="[]")
    task_ids       = Column(Text, nullable=True)  # JSON list de task_id; None = todas
    enabled        = Column(Boolean, default=True)
    created_at     = Column(DateTime, default=datetime.now)


class FsBrowseRootDB(Base):
    """Raíces autorizadas para el navegador de archivos remoto.
    Gestionado desde la UI (solo superuser). Si esta tabla está vacía, el sistema
    cae al fallback definido en FS_BROWSE_ROOTS del .env."""
    __tablename__ = "fs_browse_roots"

    id           = Column(Integer, primary_key=True, autoincrement=True)
    label        = Column(String, nullable=False)
    path         = Column(String, nullable=False)
    allow_upload = Column(Boolean, default=False)   # Permite subir archivos dentro de esta raíz
    sort_order   = Column(Integer, default=0)
    created_at   = Column(DateTime, default=datetime.now)
    updated_at   = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    created_by   = Column(String, nullable=True)   # username del superuser que la creó


class User(Base):
    __tablename__ = "users"

    id           = Column(Integer, primary_key=True, autoincrement=True)
    username     = Column(String, nullable=False, unique=True, index=True)
    full_name    = Column(String, nullable=False, default="")
    email        = Column(String, nullable=True, unique=True, index=True)
    hashed_pw    = Column(String, nullable=False)
    role         = Column(SAEnum(UserRole), nullable=False, default=UserRole.READER)
    is_active    = Column(Boolean, default=True)
    created_at   = Column(DateTime, default=datetime.now)
    updated_at   = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    last_login   = Column(DateTime, nullable=True)


# ── Inicialización ────────────────────────────────────────────────────────────
async def init_db():
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)
        # Migración manual para columnas nuevas en DBs existentes
        await conn.run_sync(_migrate_existing_db)


def _migrate_existing_db(conn):
    """Añade columnas nuevas a DBs existentes (SQLite no soporta IF NOT EXISTS en ALTER)."""
    from sqlalchemy import text
    migrations = [
        "ALTER TABLE tasks ADD COLUMN deleted_at DATETIME",
        "ALTER TABLE tasks ADD COLUMN max_retries INTEGER DEFAULT 0",
        "ALTER TABLE tasks ADD COLUMN retry_delay_s INTEGER DEFAULT 60",
        "ALTER TABLE tasks ADD COLUMN retry_count INTEGER DEFAULT 0",
        "ALTER TABLE tasks ADD COLUMN task_type TEXT DEFAULT 'excel'",
        "ALTER TABLE tasks ADD COLUMN pipeline_config TEXT DEFAULT '{}'",
        "ALTER TABLE tasks ADD COLUMN last_run_status TEXT",
        "ALTER TABLE run_logs ADD COLUMN retry_attempt INTEGER DEFAULT 0",
        "ALTER TABLE fs_browse_roots ADD COLUMN allow_upload INTEGER DEFAULT 0",
    ]
    for sql in migrations:
        try:
            conn.execute(text(sql))
        except Exception:
            pass  # La columna ya existe

    # Tabla fs_browse_roots (creada con create_all; CREATE IF NOT EXISTS por si
    # la DB existente es anterior).
    try:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS fs_browse_roots (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                label        TEXT NOT NULL,
                path         TEXT NOT NULL,
                allow_upload INTEGER DEFAULT 0,
                sort_order   INTEGER DEFAULT 0,
                created_at   DATETIME,
                updated_at   DATETIME,
                created_by   TEXT
            )
        """))
    except Exception:
        pass

    # Tabla users (creada con create_all; esta migración es por si la DB
    # existente fue creada antes de que se añadiera la tabla)
    try:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS users (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                username   TEXT NOT NULL UNIQUE,
                full_name  TEXT NOT NULL DEFAULT '',
                email      TEXT UNIQUE,
                hashed_pw  TEXT NOT NULL,
                role       TEXT NOT NULL DEFAULT 'reader',
                is_active  INTEGER NOT NULL DEFAULT 1,
                created_at DATETIME,
                updated_at DATETIME,
                last_login DATETIME
            )
        """))
    except Exception:
        pass

    # Normalizar valores de columnas enum a lowercase (versiones anteriores
    # guardaban los nombres del enum en mayúsculas, p.ej. "INTERVAL" en vez de "interval")
    for sql in [
        "UPDATE tasks SET schedule_type = LOWER(schedule_type) WHERE schedule_type != LOWER(schedule_type)",
        "UPDATE tasks SET status = LOWER(status) WHERE status != LOWER(status)",
        "UPDATE run_logs SET status = LOWER(status) WHERE status != LOWER(status)",
    ]:
        try:
            conn.execute(text(sql))
        except Exception:
            pass


async def get_db():
    async with AsyncSessionLocal() as session:
        yield session

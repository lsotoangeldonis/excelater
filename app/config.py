"""app/config.py — Configuración centralizada vía variables de entorno o .env"""
import json
import logging
from pathlib import Path
from typing import List

from pydantic import AliasChoices, BaseModel, Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

_cfg_logger = logging.getLogger(__name__)


class FsBrowseRoot(BaseModel):
    """Una raíz autorizada para el navegador de archivos remoto (formato .env, fallback).
    `allow_upload` también se acepta como `allowUpload` (camelCase) en el JSON."""
    label: str
    path: str
    allow_upload: bool = Field(
        default=False,
        validation_alias=AliasChoices("allow_upload", "allowUpload"),
    )

    model_config = {"populate_by_name": True}

    @field_validator("label", "path")
    @classmethod
    def _strip(cls, v: str) -> str:
        return v.strip()


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    # Servidor
    host: str = "0.0.0.0"
    port: int = 8000
    debug: bool = False

    # Rutas base
    base_dir: Path = Path(__file__).parent.parent
    db_path: str = "scheduler.db"
    logs_dir: str = "logs"

    # Límites
    max_log_size_mb: int = 10
    log_backup_count: int = 5
    lock_timeout_s: int = 120
    lock_retry_s: int = 5
    lock_max_retries: int = 5   # 0 = sin límite (solo por timeout)
    refresh_timeout_s: int = 300
    refresh_check_s: int = 3

    # Autenticación legacy (vacío = sin auth)
    api_key: str = ""

    # Autenticación JWT
    jwt_secret: str = ""          # Se genera al crear el superadmin si está vacío
    jwt_algorithm: str = "HS256"
    jwt_expire_minutes: int = 480  # 8 horas por defecto
    auth_enabled: bool = True      # False = deshabilita auth (solo desarrollo)

    # Notificaciones webhook
    webhook_url: str = ""           # URL a la que se enviará un POST en cada ejecución
    notify_on_failure: bool = True  # Notificar cuando una tarea falla
    notify_on_success: bool = False # Notificar cuando una tarea tiene éxito

    # Timezone del scheduler
    timezone: str = "America/Lima"

    # Retry automático
    retry_max: int = 0      # 0 = sin retry global; la tarea puede sobreescribirlo
    retry_delay_s: int = 60 # Segundos de espera entre reintentos

    # CORS (separar múltiples orígenes con coma; "*" = todos)
    cors_origins: str = "*"

    # Navegador de archivos remoto (clientes que no son el host)
    # JSON con la whitelist de raíces accesibles. Ejemplo:
    #   FS_BROWSE_ROOTS=[{"label":"OneDrive","path":"C:\\Users\\luis\\OneDrive"}]
    # Si está vacío, el navegador remoto queda deshabilitado.
    fs_browse_roots: List[FsBrowseRoot] = []
    fs_browse_allow_hidden: bool = False

    @field_validator("fs_browse_roots", mode="before")
    @classmethod
    def _parse_fs_roots(cls, v):
        # Acepta string JSON (caso .env) o lista ya parseada.
        if v in (None, "", []):
            return []
        if isinstance(v, str):
            try:
                v = json.loads(v)
            except json.JSONDecodeError as exc:
                _cfg_logger.warning("[config] FS_BROWSE_ROOTS no es JSON válido: %s", exc)
                return []
        return v

    # Email SMTP (vacío = notificaciones por email deshabilitadas)
    smtp_host: str = ""
    smtp_port: int = 587
    smtp_user: str = ""
    smtp_password: str = ""
    smtp_from: str = ""   # Ej: "Excelater <no-reply@empresa.com>"
    smtp_tls: bool = True

    @property
    def db_url(self) -> str:
        return f"sqlite+aiosqlite:///{self.db_path}"

    @property
    def logs_path(self) -> Path:
        p = Path(self.logs_dir)
        p.mkdir(parents=True, exist_ok=True)
        return p

    @property
    def fs_browse_resolved_roots(self) -> List[dict]:
        """Devuelve las raíces válidas con su Path resuelto (en absoluto).
        Las entradas con path inexistente se descartan con warning."""
        out: List[dict] = []
        for r in self.fs_browse_roots:
            try:
                p = Path(r.path).resolve(strict=False)
                if not p.exists():
                    _cfg_logger.warning("[config] FS_BROWSE_ROOTS: ruta inexistente '%s' (descartada)", r.path)
                    continue
                out.append({
                    "label": r.label or str(p),
                    "path": p,
                    "raw": r.path,
                    "allow_upload": bool(r.allow_upload),
                })
            except Exception as exc:
                _cfg_logger.warning("[config] FS_BROWSE_ROOTS: ruta inválida '%s' (%s)", r.path, exc)
        return out


settings = Settings()

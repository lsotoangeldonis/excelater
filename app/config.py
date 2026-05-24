"""app/config.py — Configuración centralizada vía variables de entorno o .env"""
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict


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


settings = Settings()

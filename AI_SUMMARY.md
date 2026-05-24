# AI_SUMMARY — Excelater (ultra compact)

> Resumen denso del proyecto optimizado para uso por IA. Pegar al inicio de sesiones nuevas para que el modelo entienda el proyecto sin re-escanear el repo.
> Si necesitas más contexto humano: `PROJECT_CONTEXT.md`. Si necesitas plantillas: `PROMPTS.md`.

---

## IDENTIDAD
- **Nombre:** Excelater (paquete: `execelater`).
- **Tipo:** servicio web local en Windows, monolito Python.
- **Propósito:** programar y ejecutar refresco automático de archivos Excel/Access (OneDrive/SharePoint) vía COM. Dashboard SPA, cron, reintentos, notificaciones, workflows custom.
- **Despliegue:** Scheduled Task de Windows bajo usuario interactivo (no servicio sesión-0, porque COM y OneDrive lo requieren).

## STACK (versiones exactas)
- Python `>=3.10,<3.14`. Poetry ≥1.8. Entry: `poetry run excelater` → `app.main:start`.
- FastAPI 0.111, uvicorn[standard] 0.29, lifespan async.
- APScheduler 3.10 AsyncIOScheduler, tz default `America/Lima`.
- SQLAlchemy 2.0 async + aiosqlite. DB única: `scheduler.db`.
- Pydantic 2 + pydantic-settings (Settings desde `.env`).
- python-jose JWT HS256, bcrypt 5 (NO passlib — fue migrado).
- pywin32 306 (Windows only, marker `sys_platform == 'win32'`).
- openpyxl 3.1 (uso puntual sin COM), aiosmtplib 3, httpx 0.27, python-multipart 0.0.9.
- Tests: pytest 8 + pytest-asyncio 0.23 (mode=auto), anyio[trio].
- Frontend: HTML + JS vanilla + Font Awesome CDN, IBM Plex Mono. SIN build step.

## LAYOUT
```
app/
  main.py            141   FastAPI + lifespan + SPA fallback + start()
  config.py           75   Settings(BaseSettings) — env vars
  database.py        230   Modelos ORM + init_db + _migrate_existing_db
  auth.py            127   bcrypt + JWT + require_reader/admin/superuser
  auth_routes.py     273   /api/auth/* (login, users, me)
  routes.py         1298   /api/* (tasks, logs, reports, notifications, browse)
  scheduler.py       502   APScheduler + execute_task + jobs
  excel_engine.py    448   COM Excel — refresh, save, lock retry
  access_engine.py   313   COM Access — macros, compact, imports
  notifications.py   273   Email SMTP + WhatsApp CallMeBot
  workflows/__init__.py     registry (singleton) — register("name", Cls)
  workflows/base.py         BaseWorkflow ABC (.run(config, logger) -> EngineResult)
  workflows/weekly_excel_copy.py  450
  static/index.html 2928   SPA monolítica
  static/login.html  250
scripts/             CLI: create_superadmin.py, reset_password.py, test_bcrypt.py
tests/               conftest.py + test_routes.py + test_scheduler.py + test_workflows.py  (36 tests)
install-service.ps1  registra Scheduled Task Windows
deploy.ps1           hot-update del Scheduled Task
pyproject.toml       Poetry, pytest asyncio_mode=auto
.env                 (gitignored)  ej: AUTH_ENABLED, JWT_SECRET, SMTP_*
scheduler.db         (gitignored)
logs/                (gitignored) excelater.log + task_<id>_<run>.log
```

## MODELOS (`app/database.py`)
- Enums = `str, enum.Enum`. SAEnum custom usa `values_callable` (valores lowercase).
- **Task**(`tasks`): id (uuid str), name, description, file_path, schedule_type, schedule_config (JSON TEXT), refresh_connections, refresh_pivots, save_on_success, excel_visible, **task_type** (`excel`/`pipeline`/`workflow`), **pipeline_config** (JSON TEXT), max_retries (default 0), retry_delay_s (60), retry_count, status (`active`/`paused`/`disabled`), created_at, updated_at, last_run_at, last_run_status (str), next_run_at, **deleted_at** (soft delete).
- **RunLog**(`run_logs`): id (autoinc int), task_id, task_name, status (`running`/`success`/`failed`/`skipped`/`cancelled`), started_at, finished_at, duration_s (float), log_file, error_msg, connections (int), pivots_ok, pivots_err, pivots_completed (JSON TEXT), retry_attempt (0 = original).
- **NotificationRule**: id, task_id, trigger (`always`/`on_error`/`on_success`/`first_run_of_day`), channel (`email`/`whatsapp`), recipients (JSON TEXT: email = `[str]`, whatsapp = `[{phone, apikey}]`), enabled.
- **ReportSchedule**: id, name, schedule_type, schedule_config, lookback_hours (24), channel, recipients, task_ids (JSON or null = todas), enabled.
- **User**: id, username (unique idx), full_name, email (unique idx), hashed_pw, role (`superuser`/`admin`/`reader`, default reader), is_active, created/updated_at, last_login.
- **Migraciones:** sin Alembic. `_migrate_existing_db()` ejecuta `ALTER TABLE ... ADD COLUMN ...` en try/except. Para añadir columna: 1) modelo, 2) ALTER ahí.

## ENDPOINTS (todos `/api/...` + auth excepto `/auth/login`)
### Tasks
- `GET /tasks` — listado (con next_run_at)
- `POST /tasks` — Excel | `POST /tasks/pipeline` — Access ETL | `POST /tasks/pipeline/reposicion` — preset
- `POST /tasks/workflow/weekly-excel-copy` — workflow
- `GET/PUT/DELETE /tasks/{id}` (DELETE = soft)
- `POST /tasks/{id}/pause` `/resume` `/run-now`
- `POST /tasks/{id}/test-run` body `{force_weekday: 1..7}` — solo workflow
- `GET /tasks/export` (JSON), `POST /tasks/import` (multipart file)
### Logs
- `GET /logs?page&page_size&task_id&status`
- `DELETE /logs?task_id&status`
- `GET /logs/{run_id}/tail?offset=N` (live polling), `/content`, `/download`
- `POST /logs/{run_id}/stop`
### Reports / Notifications
- `GET/POST /reports`, `GET/PUT/DELETE /reports/{id}`, `POST /reports/{id}/run-now`
- `GET/POST /tasks/{id}/notifications`, `DELETE /notifications/{rule_id}`
### Misc
- `GET /stats` (active/total/success/failed/running, success_rate)
- `GET /browse-file?filter=excel|access|any` (Windows-only, PowerShell OpenFileDialog)
- `GET /browse-folder` (Windows-only, FolderBrowserDialog)
- `POST /admin/cleanup-stuck-runs`
### Auth (`/api/auth/`)
- `POST /login` (form data: username, password) → `{access_token, token_type, expires_at, user}`
- `GET/POST /users`, `GET/PATCH/DELETE /users/{id}`, `POST /users/{id}/reset-password`
- `GET /me`, `POST /me/change-password`
### Sistema
- `GET /health` → `{status, scheduler_running, jobs, version, timezone}`
- `GET /login` → sirve `login.html`
- `GET /{full_path:path}` catch-all → `index.html` (NO si empieza con `api`)

## FLUJOS

### Crear tarea
UI → POST → pydantic valida → DB insert (`Task`) → `add_or_replace_job(task)` registra trigger en APScheduler → retorna dict con `next_run_at`.

### Ejecución programada → `scheduler.py::execute_task(task_id, config_overrides=None)`
1. `AsyncSessionLocal()` lee Task; si `deleted_at` o no existe → log warning + return.
2. Crea `RunLog(status=running, started_at=now)`.
3. `logger, log_path = make_task_logger(...)` → rotating file (10 MB × 5).
4. Despacha por `task.task_type`:
   - `excel` → construye `EngineConfig` desde Task fields → `await asyncio.to_thread(run_update, cfg)` (COM síncrono).
   - `pipeline` → orquesta `excel_engine` (refresh cubos en orden) + `access_engine` (compact+repair → pre-macros → saved-imports → post-macros).
   - `workflow` → `workflow_type = pipeline_config["workflow_type"]` → `cls = workflows.registry.get(workflow_type)` → `cls().run(config, logger)`. `config_overrides` se mergea aquí (ej: `force_weekday` para test-run).
5. Si falla y `max_retries > 0`: `asyncio.sleep(retry_delay_s)` y reintenta hasta agotar; cada reintento es un `RunLog` con `retry_attempt += 1`.
6. Cierra `RunLog` (status final, finished_at, duration_s, métricas). `try/finally` garantiza cierre incluso si revienta.
7. `dispatch_notifications(task, runlog)` filtra `NotificationRule` por `trigger` y manda email/WhatsApp.
8. Si `settings.webhook_url`: POST silencioso al webhook con `payload = {task_id, status, ...}`.

### Auth
- `POST /api/auth/login` form data → `auth_routes.login` → `verify_password(plain, hashed)` (bcrypt) → `create_access_token(username, role)` → JWT `{sub: username, role, exp, iat}`, default 8h (`jwt_expire_minutes=480`).
- Frontend almacena en `localStorage['excelater_token']` como JSON `{token, expires_at, username, full_name, role}`.
- Cada request: `Authorization: Bearer <token>` (lo inyecta el helper `api(method, path, body)` en `index.html`).
- `auth.py::_get_current_user_optional` decodifica token → carga User → valida `is_active`. Si `auth_enabled=false` retorna `None` (auth deshabilitada).
- Dependencies: `require_reader` (todos), `require_admin` (admin+superuser), `require_superuser`.
- 401 en frontend → `doLogout()` → `localStorage.removeItem` → redirect `/login`.
- Si `JWT_SECRET` vacío: `_secret()` genera uno en memoria (los tokens mueren al reiniciar). `create_superadmin.py` persiste un secret en `.env`.

## CONFIG (`app/config.py` — Settings)
Todos cargan de `.env` (case-insensitive, prefix vacío).
- Server: `HOST=0.0.0.0`, `PORT=8000`, `DEBUG=false`
- DB: `DB_PATH=scheduler.db` (URL = `sqlite+aiosqlite:///{db_path}`)
- Logs: `LOGS_DIR=logs`, `MAX_LOG_SIZE_MB=10`, `LOG_BACKUP_COUNT=5`
- Timing: `LOCK_TIMEOUT_S=120`, `LOCK_RETRY_S=5`, `LOCK_MAX_RETRIES=5`, `REFRESH_TIMEOUT_S=300`, `REFRESH_CHECK_S=3`
- Auth legacy: `API_KEY=""` (vacío = sin auth API-key)
- JWT: `JWT_SECRET=""` (autogenera si vacío), `JWT_ALGORITHM=HS256`, `JWT_EXPIRE_MINUTES=480`, `AUTH_ENABLED=true`
- Webhook: `WEBHOOK_URL=""`, `NOTIFY_ON_FAILURE=true`, `NOTIFY_ON_SUCCESS=false`
- Scheduler: `TIMEZONE=America/Lima`
- Retry global: `RETRY_MAX=0`, `RETRY_DELAY_S=60`
- CORS: `CORS_ORIGINS="*"` (separar por coma)
- SMTP: `SMTP_HOST/PORT/USER/PASSWORD/FROM/TLS` (vacío = email deshabilitado)

## CONVENCIONES
- **Async end-to-end**. COM (sync) → `asyncio.to_thread()` o `run_in_executor`.
- **Pydantic** para todo request body. Nunca aceptar dict crudo.
- **Soft delete** vía `deleted_at`. Filtros siempre `WHERE deleted_at IS NULL`.
- **Type hints modernos**: `from __future__ import annotations`, `list[X]`, `X | None`, `Optional[X]`.
- **Logging**: cada ejecución abre su propio logger con rotating handler.
- **PowerShell scripts**: ASCII puro (PS5 Windows-1252 compat). NO em-dashes `—`, NO `?.Source`, NO operadores PS7-only.
- **Frontend JS**: helper `api()` único; `esc()`/`escJs()` para escape; `toast(msg, type)` para feedback; páginas `<div id="page-X" class="page">` mostradas por `goPage('X')`; admin-only via clase `.admin-only` con `display:none` inicial.
- **Commits**: español, prefijo `feat:`/`fix:`/`refactor:`/`test:`/`docs:`.

## RIESGOS / GOTCHAS
1. COM Excel se cuelga con diálogos modales — timeout 300s, sin watchdog.
2. Excel ya abierto → lock; reintento configurable.
3. OneDrive Files On-Demand: si archivo no descargado → FileNotFoundError, sin pre-hidratación.
4. Sesión interactiva caída → Scheduled Task no corre (por diseño).
5. Migraciones SQLite con ALTER + try/except — fácil olvidar la ALTER al añadir columna.
6. `JWT_SECRET` vacío → tokens mueren al reiniciar.
7. `index.html` es 2900+ líneas monolítico — alto riesgo de merge conflict (ya ocurrió: commit `0a7b3b4` borró ~50% de features, recuperación 2026-05-24).
8. CORS abierto (`*`) por defecto — endurecer con `CORS_ORIGINS` en prod.
9. `.ps1` con encoding wrong → falla en PS5. ASCII puro siempre.
10. Workflows: añadir uno requiere tocar 3 archivos (`workflows/<x>.py`, `__init__.py`, `routes.py`) + opcionalmente `index.html`.

## COMANDOS
```powershell
# Setup
poetry install
copy .env.example .env
poetry run python scripts/create_superadmin.py   # genera JWT_SECRET + crea superadmin
# Dev
poetry run excelater                              # http://localhost:8000
poetry run pytest                                 # 36+ tests, < 5s
poetry run pytest -k workflow -x
# Scripts
poetry run python scripts/reset_password.py <username>
poetry run python scripts/test_bcrypt.py
# Deploy (PS admin)
.\install-service.ps1                             # primera vez
.\deploy.ps1                                      # hot-update
Get-ScheduledTask -TaskName Excelater
Start-/Stop-/Unregister-ScheduledTask -TaskName Excelater
# Logs
Get-Content -Tail 100 -Wait logs\excelater.log
Get-Content logs\task_<id>_<run>.log
# DB
sqlite3 scheduler.db "SELECT id, name, status FROM tasks WHERE deleted_at IS NULL"
```

## API CONTRACT (frontend ↔ backend — claves que NO se pueden renombrar sin coordinación)
- Task dict (`_task_to_dict`): `id, name, description, file_path, schedule_type, schedule_config, refresh_connections, refresh_pivots, save_on_success, excel_visible, task_type, pipeline_config, max_retries, retry_delay_s, retry_count, status, last_run_at, last_run_status, next_run_at`.
- RunLog dict: `id, task_id, task_name, status, started_at, finished_at, duration_s, log_file, error_msg, connections, pivots_ok, pivots_err, retry_attempt`.
- Tail response: `{content, offset, status}` (status = `running` mantiene polling vivo).
- Auth login response: `{access_token, token_type, expires_at, user: {username, full_name, email, role}}`.
- Frontend espera `pipeline_config` con shape específica por `task_type`:
  - `pipeline`: `{access_db, excel_files: [{path, visible}], compact_before_import, pre_import_macros: [str], saved_imports: [str], post_import_macros: [str], access_visible, excel_refresh_timeout}`.
  - `workflow weekly_excel_copy`: `{workflow_type: "weekly_excel_copy", folder, file_patterns: [str con {week}], week_padding, daily_refresh, fail_if_source_missing, excel_visible, refresh_timeout, pivot_guards: [{sheet, pivot, min_gap}]}`.

## TESTS (`tests/`)
- `conftest.py` (68 líneas): fixtures `client_async` (httpx AsyncClient sobre `app`), DB SQLite en memoria, ` overrides_dependency` para auth.
- `test_routes.py` (244): happy paths de CRUD tareas, logs, stats.
- `test_scheduler.py` (41): registro/cancelación de jobs.
- `test_workflows.py` (231): workflow weekly_excel_copy con mocks de COM.
- Convención: tests async sin decorador (mode=auto). Nombre `test_<accion>_<resultado>`.
- Total: 36 tests, ~1 segundo de wall time.

## REGLAS PARA CLAUDE
1. Antes de modificar `index.html`: grep el ID/función exacta; el archivo es enorme y los selectores cruzan secciones.
2. Antes de tocar schema: confirmar que añadiste el ALTER en `_migrate_existing_db`.
3. Antes de añadir endpoint: incluir `Depends(require_*)` o `Depends(verify_api_key)`.
4. NUNCA usar: `passlib`, `asyncio.get_event_loop()`, em-dashes en `.ps1`, `DELETE` directo en `tasks` (usa soft delete).
5. Para nuevos workflows: clase + register + endpoint + UI (4 lugares).
6. Para bugs de COM: revisar `logs/task_<id>_<run>.log` específico, no sólo el log global.
7. No commitear `scheduler.db`, `.env`, `logs/`, `.venv/`, `*.lock` (excepto `poetry.lock`).
8. Antes de operaciones destructivas (reset DB, force push, install/uninstall del Scheduled Task): pedir confirmación.

## ARCHIVOS DOCUMENTALES ADICIONALES
- `PROJECT_CONTEXT.md` — versión humana extendida.
- `CLAUDE_CONTEXT.md` — versión corta para pegar al inicio de sesiones.
- `PROMPTS.md` — plantillas listas para tareas comunes.
- `DEPENDENCY_MAP.md` — grafo de deps + archivos peligrosos + contratos UI↔backend.
- `ONBOARDING.md` — setup en <30 min para dev nuevo.
- `KEY_FILES.md` — top 20 archivos del repo, explicados.
- `CHANGELOG_AI.md` — bitácora de cambios mayores (mantener actualizada).
- `README.md` — quickstart de usuario final.

---

_Versión: 1.0 — 2026-05-24. ~250 líneas. Diseñado para minimizar tokens en sesiones con Claude._

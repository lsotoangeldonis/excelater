# CLAUDE_CONTEXT — Excelater

> **Pega o adjunta este archivo al inicio de cualquier conversación nueva.**
> Le da a Claude lo mínimo indispensable para trabajar sin re-escanear el repo.
> Versión densa, optimizada para tokens. Para el documento humano completo: `PROJECT_CONTEXT.md`.

---

## TL;DR del proyecto

Servicio web local en Windows que **programa y ejecuta actualizaciones de archivos Excel/Access** (OneDrive/SharePoint) vía COM. FastAPI + APScheduler + SQLite async + frontend HTML/JS vanilla. Auth JWT con bcrypt. Deploy como **Scheduled Task de Windows** (no servicio puro; COM necesita sesión interactiva).

---

## Stack y versiones críticas

- Python 3.10–3.13, Poetry (≥1.8) — `poetry run excelater`
- FastAPI 0.111 + uvicorn 0.29 (async, lifespan en `app/main.py`)
- APScheduler 3.10 (AsyncIOScheduler, tz `America/Lima` por defecto)
- SQLAlchemy 2.0 async + aiosqlite — DB única: `scheduler.db`
- Pydantic 2 + pydantic-settings (config desde `.env`)
- python-jose (JWT HS256) + bcrypt 5 (NO usar passlib)
- pywin32 306 (Windows only; `DispatchEx` para Excel/Access)
- aiosmtplib + httpx + python-multipart (form login)
- Tests: pytest + pytest-asyncio (mode=auto)
- Frontend: HTML+JS vanilla en `app/static/index.html` (~2900 líneas, sin build)

---

## Estructura mínima a recordar

```
app/
  main.py              # FastAPI bootstrap + lifespan + SPA fallback
  config.py            # Settings(BaseSettings) — toda env var nueva va aquí
  database.py          # Modelos + init_db + _migrate_existing_db (ALTER manual)
  auth.py              # bcrypt + JWT + require_reader/admin/superuser
  auth_routes.py       # /api/auth/login, /auth/users, /auth/me
  routes.py            # /api/* — endpoints principales (1300 líneas)
  scheduler.py         # APScheduler + execute_task() + jobs
  excel_engine.py      # COM Excel — refresh conexiones/pivots, save, lock retry
  access_engine.py     # COM Access — macros, compact&repair, importaciones
  notifications.py     # Email (SMTP) + WhatsApp (CallMeBot)
  workflows/
    __init__.py        # registry (singleton) — register("name", Cls)
    base.py            # BaseWorkflow (ABC, método run)
    weekly_excel_copy.py
  static/
    index.html         # Dashboard SPA monolítico
    login.html         # Login
scripts/               # CLI: create_superadmin.py, reset_password.py
tests/                 # pytest-asyncio
install-service.ps1    # Registra Scheduled Task
deploy.ps1             # Hot-update
```

---

## Modelos clave (`app/database.py`)

- **Task** (`tasks`): id (uuid), name, file_path, schedule_type, schedule_config (JSON), refresh_*, save_on_success, excel_visible, **task_type** (`excel`/`pipeline`/`workflow`), **pipeline_config** (JSON), max_retries, retry_delay_s, retry_count, status (`active`/`paused`/`disabled`), last_run_at, last_run_status, next_run_at, **deleted_at** (soft delete).
  - `pipeline_config` (cuando `task_type == "pipeline"`): `excel_files[]` (fuentes), `access_db`, `pre_import_macros[]`, `saved_imports[]`, `post_import_macros[]`, `post_refresh_excel_files[]` (paso 8 = tableros consumidores), `compact_position` (`"" | "before_macros" | "after_pre_macros" | "skip"`, default resuelto a `after_pre_macros`), `continue_on_error`, `compact_before_import` (legacy, respetado si `compact_position == ""`), timeouts Excel/Access.
- **RunLog** (`run_logs`): id, task_id, status (`running`/`success`/`failed`/`skipped`/`cancelled`), started/finished_at, duration_s, log_file, error_msg, connections, pivots_ok/err, retry_attempt.
- **NotificationRule**: por tarea; trigger (`always`/`on_error`/`on_success`/`first_run_of_day`), channel (`email`/`whatsapp`), recipients (JSON).
- **ReportSchedule**: reportes resumen programados; mismo schedule_type que Task.
- **User**: username (unique), full_name, email, hashed_pw, **role** (`superuser`/`admin`/`reader`), is_active, last_login.

**Enums** son `str, enum.Enum`. SAEnum custom usa `values_callable` → valores en minúsculas (legacy).

---

## Flujos críticos

1. **Crear tarea** → `routes.py` (POST `/tasks` o `/tasks/pipeline` o `/tasks/workflow/weekly-excel-copy`) → DB insert → `scheduler.add_or_replace_job(task)` → respuesta con `next_run_at`.
2. **Ejecución** → APScheduler dispara → `execute_task(task_id, config_overrides=None)`:
   - Crea RunLog (running) + logger por-tarea (rotating en `logs/`).
   - Despacha por `task_type`: excel_engine / pipeline (excel+access) / workflows.registry.
   - Reintentos si `max_retries > 0`.
   - Cierra RunLog + dispara notificaciones por reglas.

   **Pipeline Access ETL** (`AccessPipelineRunner.run()`, en `app/access_engine.py`) — orden real:
   1. Refrescar `excel_files[]` (fuentes / cubos) vía `run_engine()` (hereda hidratación OneDrive + lock-wait).
   2. **Preparar `.accdb`**: hidratar OneDrive si es placeholder + `wait_for_file` (detecta `.laccdb`).
   3. Según `compact_position`:
      - `before_macros`: Compact → abrir Access → pre-macros → imports → post-macros → cerrar.
      - `after_pre_macros` (**default**): abrir Access → pre-macros → cerrar → Compact → re-preparar → reabrir → imports → post-macros → cerrar.
      - `skip`: sin Compact.
   4. Refrescar `post_refresh_excel_files[]` (paso 8 manual: tableros consumidores que leen de Access).

   Cada apertura COM aplica `AutomationSecurity = 3` (silencia prompts de macros) + `DoCmd.SetWarnings(False)` (silencia confirmaciones de action queries). Si `continue_on_error == True`, un fallo individual en macro/import/tablero no aborta; el run termina con `success=False` y `error_msg` con conteo.
3. **Run-now manual** → `POST /tasks/{id}/run-now` → mismo `execute_task` en `asyncio.create_task`.
4. **Workflow test** → `POST /tasks/{id}/test-run {force_weekday: 1..7}` (sólo task_type=workflow).
5. **Auth**: login → JWT (8h) → frontend guarda en `localStorage['excelater_token']` → cada request lleva `Authorization: Bearer ...` → 401 redirect `/login`.

---

## Convenciones obligatorias

- **Async end-to-end**. COM (síncrono) se ejecuta vía `asyncio.to_thread()` o `run_in_executor`.
- **Pydantic models** para todos los request bodies.
- **Soft delete** vía `deleted_at` (nunca DELETE en `tasks`).
- **Cambios de schema** → añade `ALTER TABLE` en `_migrate_existing_db()` (no hay Alembic).
- **Nuevo workflow** → 3 lugares: clase en `app/workflows/<x>.py` (hereda BaseWorkflow) + register en `workflows/__init__.py` + endpoint en `routes.py` + sección + radio en `index.html`.
- **Logging**: cada ejecución abre su propio logger en `logs/task_{id}_{run}.log` (rotating, 10 MB × 5).
- **PowerShell scripts** (`.ps1`): compatibles con PS5; ASCII puro (sin em-dashes ni `?.Source`).
- **Frontend**: helper `api(method, path, body)` auto-inyecta Bearer y maneja 401. Usar `esc()` / `escJs()` en templates, `toast(msg, type)` para feedback. Páginas son divs `#page-X` mostrados por `goPage('X')`.

---

## Riesgos a tener en mente

1. **Excel/COM**: si Excel está abierto por el usuario → lock retry (`lock_timeout`, default 120s). Si Excel muestra modal → cuelga hasta `refresh_timeout` (300s). En pipeline, `AutomationSecurity=3` + `SetWarnings(False)` mitigan modales de Access.
2. **OneDrive Files On-Demand**: pre-hidratación implementada para `.xlsx` fuente, `.accdb`, y `.xlsx` consumidores en pipeline (vía `_is_onedrive_placeholder` + `_trigger_onedrive_download`). Para tareas Excel puras el wrapper `run_engine()` también lo cubre.
3. **Migraciones SQLite**: `_migrate_existing_db` usa ALTER + try/except. No olvidar al cambiar modelos.
4. **JWT_SECRET vacío**: tokens se invalidan al reiniciar. `create_superadmin.py` lo genera y persiste.
5. **`index.html` monolítico (2900+ líneas)**: alto riesgo de conflicto en merges grandes. Ya pasó una vez (`0a7b3b4` borró ~50% de features; recuperación 2026-05-24).
6. **CORS abierto (`*`) por defecto** — configurable vía `CORS_ORIGINS`.

---

## Comandos esenciales

```powershell
poetry install
poetry run python scripts/create_superadmin.py   # primera vez
poetry run excelater                              # http://localhost:8000
poetry run pytest                                 # todos los tests
poetry run pytest -k workflow                     # filtrado
.\install-service.ps1                             # registra Scheduled Task (admin)
.\deploy.ps1                                      # hot-update
Get-ScheduledTask -TaskName Excelater
Get-Content -Tail 100 -Wait logs\excelater.log
```

---

## Cosas que NO hacer

- ❌ No agregar `passlib` — se migró a `bcrypt` puro (commits 995fca5 → 3a39f8a).
- ❌ No usar `asyncio.get_event_loop()` — usa `get_running_loop()` (deprecated en 3.12+).
- ❌ No abrir Excel sin `DispatchEx` (que aísla instancias).
- ❌ No commitear `scheduler.db`, `.env`, `logs/`, `.venv/`.
- ❌ No usar em-dashes (`—`) en archivos `.ps1` — rompen PS5 con Windows-1252.
- ❌ No `DELETE` directo en `tasks` — usa soft delete (`deleted_at`).
- ❌ No retirar campos del modelo sin grep previo en `routes.py` + `index.html`.

---

## Para Claude: cómo trabajar aquí

- Para **explorar código**: usa `Grep`/`Read` directos. El proyecto es chico (~8000 líneas Python + 2900 HTML).
- Para **agregar features**: lee `PROMPTS.md` (ya hay plantillas listas).
- Para **cambios riesgosos** (deploy, schema, COM): preguntar antes de actuar.
- Para **errores de COM/Excel**: siempre revisar `logs/task_<id>_<run>.log`, no sólo `logs/excelater.log`.
- Antes de **modificar `index.html`**: confirmar qué sección y qué función JS toca; el archivo es enorme y los IDs se referencian cruzados.

---

_Versión: 1.1 — 2026-05-24_

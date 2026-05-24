# Excelater — PROJECT_CONTEXT

> Resumen maestro técnico del proyecto. Fuente de verdad para arquitectura, convenciones y operación.
> Actualizar cuando cambien decisiones de arquitectura, dependencias mayores o flujos críticos.

---

## 1. Propósito

Excelater es un **servicio web local para Windows** que programa y ejecuta la **actualización automática de archivos Excel/Access** ubicados en OneDrive/SharePoint (sincronizados localmente). Incluye dashboard SPA, programación tipo cron, ejecución forzada, reintentos, notificaciones (email + WhatsApp), reportes programados y workflows personalizados (ej: copia semanal de archivos).

Usuario objetivo: equipo interno de analítica que mantiene cubos/dashboards Excel + pipelines Access ETL en SharePoint.

---

## 2. Arquitectura general

```
┌─────────────────────────────────────────────────────────────────┐
│                       Navegador (SPA)                            │
│        app/static/index.html  +  app/static/login.html           │
└────────────────────────────────┬────────────────────────────────┘
                                 │ HTTP + JWT Bearer
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                FastAPI app (app/main.py — lifespan)              │
│  /api/auth/*  (auth_routes.py)   /api/*  (routes.py)             │
└────┬──────────────────┬─────────────────────┬───────────────────┘
     │                  │                     │
     ▼                  ▼                     ▼
┌──────────┐   ┌─────────────────┐   ┌────────────────────┐
│ Database │   │   Scheduler     │   │ Notifications      │
│ SQLite   │   │ APScheduler     │   │ SMTP / CallMeBot   │
│ async    │   │ (AsyncIO)       │   │ + webhook genérico │
└──────────┘   └────────┬────────┘   └────────────────────┘
                        │ ejecuta
                        ▼
        ┌──────────────────────────────────┐
        │  Engines (COM via pywin32)        │
        │  • excel_engine.py — Excel.App    │
        │  • access_engine.py — Access.App  │
        │  • workflows/ — flujos compuestos │
        └──────────────────────────────────┘
```

**Estilo de despliegue:** monolito Python que corre como **Scheduled Task de Windows** (no servicio Win32 puro, porque Excel COM + OneDrive Files On-Demand requieren sesión interactiva). El instalador (`install-service.ps1`) registra la tarea bajo el usuario actual con trigger "at logon".

---

## 3. Tecnologías

| Capa | Tecnología | Notas |
|------|-----------|-------|
| Lenguaje | Python 3.10–3.13 | `pyproject.toml` define `>=3.10,<3.14` |
| Gestor pkgs | **Poetry** (≥1.8) | `poetry install`, `poetry run excelater` |
| Web framework | **FastAPI 0.111** + uvicorn[standard] 0.29 | Lifespan async, CORS configurable |
| Scheduler | **APScheduler 3.10** (AsyncIOScheduler) | Triggers: `once_daily`, `interval`, `cron` |
| ORM/DB | **SQLAlchemy 2.0** async + **aiosqlite** | DB única: `scheduler.db` (SQLite) |
| Modelos/Config | **Pydantic v2** + pydantic-settings | `.env` por defecto |
| Auth | **JWT** (python-jose) + **bcrypt 5** | Roles: superuser/admin/reader |
| Excel/Access | **pywin32** (sólo Windows, COM) | `win32com.client.DispatchEx` |
| HTTP cliente | httpx | Webhooks + tests |
| Email | aiosmtplib | SMTP opcional, deshabilitado si vacío |
| WhatsApp | CallMeBot (vía HTTP) | Por destinatario: `{phone, apikey}` |
| Tests | pytest + pytest-asyncio | `tests/`, asyncio_mode=auto |
| Frontend | HTML + JS vanilla + Font Awesome CDN | Sin build step; servido desde `app/static/` |
| Deploy | PowerShell 5+ | `install-service.ps1`, `deploy.ps1` |

---

## 4. Flujo principal de datos

### 4.1 Crear tarea (UI → API → DB → Scheduler)
1. Usuario abre "Nueva tarea" en dashboard.
2. JS envía `POST /api/tasks` (Excel) / `/api/tasks/pipeline` / `/api/tasks/workflow/weekly-excel-copy` con Bearer JWT.
3. `routes.py` valida payload (pydantic) y crea `Task` en DB.
4. `add_or_replace_job(task)` registra el trigger en APScheduler y retorna `next_run_at`.
5. Respuesta incluye el dict de tarea con `next_run_at` y `status`.

### 4.2 Ejecución programada (Scheduler → Engine → DB/Log → Notify)
1. APScheduler dispara `execute_task(task_id, config_overrides=None)`.
2. Crea `RunLog` con `status=running`, abre logger por-tarea (rotating file en `logs/`).
3. Según `task_type`:
   - `excel`  → `excel_engine.run_update(EngineConfig)` (COM, refresh conexiones/pivots, save).
   - `pipeline` → orquesta `excel_engine` + `access_engine` (cubos Excel → Compact&Repair → importaciones guardadas → macros pre/post).
   - `workflow` → `workflows/registry.get(name)` → `WeeklyExcelCopyWorkflow.run(...)` (copia semana N-1 a N, refresca, guarda).
4. Reintentos si `max_retries > 0` y falló.
5. Cierra `RunLog` con `success/failed/cancelled`, duración, métricas (conexiones, pivots_ok/err).
6. Dispara notificaciones via `notifications.py` según `NotificationRule` filtradas por `TriggerType`.

### 4.3 Ejecución manual ("Ejecutar ahora")
- UI llama `POST /api/tasks/{id}/run-now` → mismo `execute_task` en background asyncio.
- Workflow simulado: `POST /api/tasks/{id}/test-run {force_weekday: 1..7}` permite simular cualquier día sin esperar al lunes.

### 4.4 Auth
- Login: `POST /api/auth/login` (form data) → JWT (8h por defecto).
- Frontend almacena en `localStorage['excelater_token']` y envía `Authorization: Bearer ...` en cada request.
- 401 → redirect a `/login`.
- Roles enforced vía dependencias `require_reader`/`require_admin`/`require_superuser` en `auth.py`.

---

## 5. Módulos críticos

| Módulo | Tamaño | Responsabilidad | Por qué es crítico |
|--------|--------|-----------------|---------------------|
| [app/main.py](app/main.py) | 141 | Bootstrap FastAPI, lifespan (init_db, cleanup_stuck_runs, start scheduler), SPA fallback | Único entry point; cualquier cambio en startup/shutdown afecta todo |
| [app/routes.py](app/routes.py) | 1298 | Endpoints REST (tasks, logs, reports, notifications, browse, stats) | Superficie API — el dashboard depende íntegramente de aquí |
| [app/scheduler.py](app/scheduler.py) | 502 | APScheduler, `execute_task`, registro/cancelación de jobs, webhook | Orquesta TODA la ejecución; bugs aquí significan tareas no corren |
| [app/excel_engine.py](app/excel_engine.py) | 448 | Motor COM Excel: refresh conexiones, pivots, save, lock/retry | Bloqueo principal; toca COM directamente |
| [app/access_engine.py](app/access_engine.py) | 313 | Motor COM Access: macros, compact&repair, importaciones | Pipeline ETL depende de esto |
| [app/auth.py](app/auth.py) | 127 | Hash bcrypt, JWT encode/decode, dependencies por rol | Toda autorización pasa por aquí |
| [app/auth_routes.py](app/auth_routes.py) | 273 | Login, gestión de usuarios, change/reset password | Surface de admin de usuarios |
| [app/database.py](app/database.py) | 230 | Modelos SQLAlchemy + `_migrate_existing_db` (ALTER TABLE manual) | Migraciones se aplican en cada arranque — añadir columnas aquí |
| [app/workflows/weekly_excel_copy.py](app/workflows/weekly_excel_copy.py) | 450 | Workflow: copia archivo Sem N-1 → Sem N cada lunes, refresca, guarda | Lógica de negocio del workflow estrella |
| [app/notifications.py](app/notifications.py) | 273 | Email (SMTP) + WhatsApp (CallMeBot) + dispatch por trigger | Único punto de salida de notificaciones |
| [app/config.py](app/config.py) | 75 | Settings (pydantic) cargado de `.env` | Cualquier env var nuevo va aquí |
| [app/static/index.html](app/static/index.html) | 2928 | SPA completa (HTML+CSS+JS inline), modales, tablas ordenables, auto-refresh 3 min | El "todo" del frontend; merge cuidadoso |
| [app/static/login.html](app/static/login.html) | 250 | Pantalla de login (independiente) | Único acceso si auth_enabled=true |
| [install-service.ps1](install-service.ps1) | — | Registra Scheduled Task que arranca al login | Único método soportado de deploy en prod |
| [deploy.ps1](deploy.ps1) | — | Actualización in-place del servicio ya instalado | Usado para hot-updates |

---

## 6. Convenciones de código

### Python
- **Type hints** modernos (`from __future__ import annotations`, `list[dict]`, `Optional[X]`, `X | None`).
- **Async** end-to-end: rutas async, SQLAlchemy async, scheduler en asyncio. Operaciones COM (sincrónicas) se ejecutan con `asyncio.to_thread()` o `run_in_executor`.
- **Enums** como `str, enum.Enum` para serialización JSON nativa.
- **Soft delete** vía `deleted_at` (no DELETE real en `tasks`).
- **Logging por tarea**: cada ejecución crea su propio logger con rotating file handler en `logs/task_{id}_{run}.log`.
- **Comments**: en español, breves. Docstrings sólo en funciones públicas no triviales.
- **Pydantic models** para todos los request bodies; nunca aceptar dict crudo.
- **No SQLAlchemy 1.x patterns**: usar `select()` + `await db.execute(...)`.

### Frontend
- **Sin framework**: vanilla JS + funciones top-level. Sin bundler.
- **Estilo**: variables CSS (`--accent`, `--bg`, etc.), monospace IBM Plex Mono para datos.
- **Convenciones JS**:
  - `api(method, path, body)` helper único — auto-inyecta Bearer y maneja 401.
  - `esc()` / `escJs()` para escape de strings en templates.
  - `toast(msg, type)` para feedback.
  - Tablas con `tableSortBy(tableId, col, renderFn)`.
- **Páginas**: divs con id `page-X` y clase `page`, mostrados por `goPage('X')`.
- **Admin-only**: clase `.admin-only` con `display:none` por defecto; se muestra al inicializar sesión.

### Git / Commits
- Mensaje en español, prefijo `feat:` / `fix:` / `refactor:` / `test:` / `docs:`.
- Commit `f3b9e8b Init` marca el origen.
- Hito mayor reciente: `0a7b3b4 feat(auth): implement JWT authentication and user management endpoints` — introdujo todo el sistema de auth.

### PowerShell (scripts)
- Compatible con **PowerShell 5** (Windows-1252 friendly): nada de em-dashes (`—`), ni `?.Source`, ni operadores PS7-only en archivos ASCII.
- Usar `& "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"` con ruta absoluta cuando se invoca desde Python.

---

## 7. Dependencias importantes

### Runtime
- **fastapi 0.111**, **uvicorn[standard] 0.29** — servidor web.
- **apscheduler 3.10** — scheduler (AsyncIO).
- **sqlalchemy 2.0 + aiosqlite** — ORM async + driver SQLite.
- **pydantic 2 + pydantic-settings** — modelos y config.
- **pywin32 306** (Windows only, marker en `pyproject.toml`) — COM Excel/Access.
- **openpyxl 3.1** — lectura/escritura de Excel sin COM (usado puntualmente).
- **python-jose[cryptography] 3.5** — JWT.
- **bcrypt 5** — hash passwords (migrado desde passlib — no usar passlib).
- **aiosmtplib 3** — SMTP async.
- **httpx 0.27** — cliente HTTP (webhooks + tests).
- **python-multipart 0.0.9** — necesario para form login (`OAuth2PasswordRequestForm`).

### Dev
- **pytest 8 + pytest-asyncio 0.23 + anyio[trio]** — testing async.

### Sistema operativo
- **Windows 10 / Server 2019+** (obligatorio): COM requiere Windows.
- **Microsoft Excel de escritorio** instalado (Office 2016+).
- **Microsoft Access** (sólo si se usan tareas pipeline).
- **PowerShell 5+** (preinstalado en Windows).

---

## 8. Posibles riesgos técnicos

| Riesgo | Severidad | Mitigación actual / pendiente |
|--------|-----------|-------------------------------|
| COM Excel se cuelga (ej: diálogo modal de Excel) | Alta | `refresh_timeout_s=300` + polling; sin watchdog que mate proceso |
| Excel ya abierto por el usuario (lock) | Alta | Reintento configurable (`lock_retry_s`, `lock_max_retries`); abrir read-write conflict |
| OneDrive Files On-Demand: archivo no descargado | Media | Sin pre-hidratación; depende del cliente OneDrive |
| Sesión interactiva caída → Scheduled Task no corre | Alta | Por diseño (necesario para COM); requiere usuario logueado |
| Migraciones SQLite ad-hoc (`_migrate_existing_db` con ALTER + except pass) | Media | No usa Alembic; cualquier cambio de schema = añadir ALTER ahí |
| JWT_SECRET vacío → tokens invalidos al reiniciar | Media | `create_superadmin.py` genera y persiste el secret; auth deshabilitable vía `auth_enabled=false` |
| `index.html` es 2900+ líneas en un solo archivo | Media | Riesgo de merge conflicts grandes (ya ocurrió, ver historial commit `0a7b3b4`) |
| Sin paginación en `/api/tasks` | Baja | OK por escala actual (decenas de tareas) |
| CORS abierto por defecto (`*`) | Baja | Configurable en `.env` (`CORS_ORIGINS`) |
| Scripts PowerShell asumen UTF-8 en algunos sitios y Windows-1252 en otros | Baja | Hay incidentes documentados (commits `88054f0`, `11a11a0`); usar ASCII puro en `.ps1` |
| Workflows: la única manera de añadir uno nuevo requiere tocar 3 archivos (`workflows/`, `routes.py`, `index.html`) | Baja | Documentado en `workflows/__init__.py` |
| `recipients` se almacena como JSON en TEXT (sin validación a nivel DB) | Baja | Validación pydantic al entrar; suficiente |

---

## 9. Comandos frecuentes

```powershell
# Setup inicial
poetry install
copy .env.example .env  # editar variables

# Crear superadmin (también genera JWT_SECRET si falta)
poetry run python scripts/create_superadmin.py

# Levantar servidor en desarrollo (auto-reload si DEBUG=true)
poetry run excelater
# → http://localhost:8000  (login → admin creado arriba)

# Tests
poetry run pytest                       # todo
poetry run pytest -k workflow           # filtrado
poetry run pytest tests/test_routes.py -x  # uno, parar al primer fallo

# Reset de contraseña sin UI
poetry run python scripts/reset_password.py <username>

# Validar hash bcrypt
poetry run python scripts/test_bcrypt.py

# Instalar / actualizar como Scheduled Task (PowerShell admin)
.\install-service.ps1                   # instala desde cero
.\deploy.ps1                            # actualiza in-place

# Gestión del Scheduled Task
Get-ScheduledTask -TaskName Excelater
Start-ScheduledTask -TaskName Excelater
Stop-ScheduledTask -TaskName Excelater
Unregister-ScheduledTask -TaskName Excelater -Confirm:$false

# Logs
Get-Content -Tail 100 -Wait logs\excelater.log       # log del servicio
Get-Content -Tail 100 logs\task_<id>_<run>.log       # log de ejecución individual

# DB inspection (necesita sqlite3.exe en PATH)
sqlite3 scheduler.db "SELECT id, name, status, last_run_status FROM tasks WHERE deleted_at IS NULL"
```

---

## 10. Estructura de carpetas explicada

```
Excelater/
├── app/                              # Código del servicio
│   ├── __init__.py                   # (vacío)
│   ├── main.py                       # FastAPI + lifespan + start()
│   ├── config.py                     # Settings(BaseSettings) — todo .env
│   ├── database.py                   # Modelos ORM + init_db + migraciones SQLite
│   ├── auth.py                       # Hash, JWT, dependencias por rol
│   ├── auth_routes.py                # /auth/login, /auth/users, /auth/me
│   ├── routes.py                     # Endpoints REST principales
│   ├── scheduler.py                  # APScheduler + execute_task + jobs
│   ├── excel_engine.py               # COM Excel (refresh, save, lock retry)
│   ├── access_engine.py              # COM Access (macros, importaciones)
│   ├── notifications.py              # Email + WhatsApp + dispatch por trigger
│   ├── workflows/                    # Workflows personalizados
│   │   ├── __init__.py               # registry + register()
│   │   ├── base.py                   # BaseWorkflow (ABC)
│   │   └── weekly_excel_copy.py      # Copia semanal Sem N-1 → Sem N
│   └── static/                       # Frontend (sin build)
│       ├── index.html                # Dashboard SPA (2900+ líneas)
│       └── login.html                # Pantalla de login
│
├── scripts/                          # Utilidades CLI (no se ejecutan automáticamente)
│   ├── create_superadmin.py          # Crea superadmin + genera JWT_SECRET
│   ├── reset_password.py             # Resetea password de un usuario
│   └── test_bcrypt.py                # Diagnóstico de hash bcrypt
│
├── tests/                            # pytest-asyncio
│   ├── conftest.py                   # Fixtures: client async, db en memoria
│   ├── test_routes.py                # Endpoints REST
│   ├── test_scheduler.py             # Scheduler + ejecución
│   └── test_workflows.py             # Workflows (weekly_excel_copy)
│
├── logs/                             # Archivos .log generados (gitignored)
│   ├── excelater.log                 # Log global del servicio
│   ├── task_<id>_<run>.log           # Log por ejecución (rotating)
│   └── index_pre_auth*.html          # Backups históricos (limpieza pendiente)
│
├── install-service.ps1               # Instala Scheduled Task de Windows
├── deploy.ps1                        # Actualiza Scheduled Task sin reinstalar
├── pyproject.toml                    # Poetry, dependencias, pytest config
├── README.md                         # Quickstart de usuario
├── scheduler.db                      # SQLite (gitignored)
└── .env                              # Variables de entorno (gitignored)
```

---

## 11. Patrones a respetar al modificar

1. **No agregar endpoints sin dependencia de auth**: usar `Depends(verify_api_key)` o, mejor, `Depends(require_reader/admin/superuser)`.
2. **Si cambias el schema**: añade el `ALTER TABLE` en `_migrate_existing_db` (en `database.py`). No olvides la columna también en el modelo.
3. **Si añades un workflow**: clase en `app/workflows/<nombre>.py` que herede `BaseWorkflow`, registrarla en `workflows/__init__.py`, exponer endpoint en `routes.py`, añadir sección de form + radio button en `index.html`.
4. **Si cambias el frontend**: cuidado — `index.html` es monolítico y editarlo en paralelo crea conflictos. Hacer commits pequeños y revisar diff antes de push.
5. **Antes de eliminar un campo del modelo**: confirmar que ningún endpoint lo lee/escribe y que el frontend tampoco lo usa (grep `campo_x` en todo el repo).
6. **Tests**: cualquier nuevo endpoint debería tener al menos un caso happy-path en `tests/test_routes.py`.
7. **Logs**: las ejecuciones de tarea **siempre** deben dejar un `RunLog` cerrado (`success`/`failed`/`cancelled`); usar try/finally.

---

## 12. Histórico relevante

- **`f3b9e8b Init`** — commit inicial.
- **`0a7b3b4 feat(auth): JWT`** — introducido sistema de auth; sobreescribió `index.html` borrando ~50% de features de la UI. **Recuperación documentada en `CHANGELOG_AI.md`** (entrada del 2026-05-24).
- **`995fca5 → 3a39f8a`** — migración passlib → bcrypt puro.
- **`e831c94`+** — agregados `deploy.ps1` e `install-service.ps1` idempotente.
- **`d89cad8`** — fix: servir `index.html` en `/` para evitar loop de redirect a `/login`.

---

_Última actualización: 2026-05-24_

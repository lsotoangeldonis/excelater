# DEPENDENCY_MAP — Excelater

> Mapa real de dependencias entre módulos del repo. Útil para entender el blast radius antes de modificar cualquier cosa.

---

## 1. Grafo de dependencias internas (módulo → de quién depende)

```
app/main.py
  ├── app/config.py
  ├── app/database.py        (init_db, AsyncSessionLocal, RunLog, RunStatus, engine)
  ├── app/scheduler.py       (scheduler, load_all_tasks)
  ├── app/routes.py          (router)
  └── app/auth_routes.py     (auth_router)

app/routes.py
  ├── app/config.py
  ├── app/database.py        (get_db + todos los modelos + enums)
  ├── app/excel_engine.py    (EngineConfig, run_update, resolve_path)
  ├── app/scheduler.py       (add_or_replace_job, remove_job, pause_job, resume_job,
  │                           execute_task, scheduler, cancel_run,
  │                           add_report_job, remove_report_job)
  └── app/auth.py            (require_reader, require_admin)

app/auth_routes.py
  ├── app/database.py        (User, UserRole, get_db, AsyncSessionLocal)
  ├── app/auth.py            (hash_password, verify_password, create_access_token,
  │                           require_admin, require_superuser, get_current_user)
  └── app/config.py

app/scheduler.py
  ├── app/config.py
  ├── app/database.py        (todos los modelos + enums)
  ├── app/excel_engine.py    (EngineConfig, run_update)
  ├── app/access_engine.py   (implícito vía pipeline)
  ├── app/workflows/         (registry — para resolver workflow_type)
  └── app/notifications.py   (dispatch_notifications)

app/excel_engine.py
  ├── app/config.py
  └── pywin32 (win32com.client) — Windows only

app/access_engine.py
  ├── app/config.py
  └── pywin32 (win32com.client) — Windows only

app/workflows/__init__.py
  ├── app/workflows/base.py
  └── app/workflows/weekly_excel_copy.py

app/workflows/weekly_excel_copy.py
  ├── app/workflows/base.py
  └── app/excel_engine.py    (run_update directamente)

app/notifications.py
  ├── app/config.py
  └── app/database.py        (NotificationRule, TriggerType, ChannelType)

app/auth.py
  ├── app/config.py
  └── app/database.py        (User, UserRole, AsyncSessionLocal)

app/database.py
  └── app/config.py          (settings.db_url)

app/config.py
  └── (sin deps internas — sólo pydantic-settings)
```

**Direcciones a recordar:**
- `config.py` está al fondo del grafo: nadie depende hacia abajo de él. Cambios aquí casi siempre seguros (sólo se añade).
- `database.py` es el nodo más referenciado: cualquier cambio de schema toca todo lo de arriba.
- `routes.py` y `scheduler.py` son hubs simétricos: `routes` invoca `scheduler` y `scheduler` actualiza tablas que `routes` lee.

---

## 2. Acoplamiento crítico

| Acoplamiento | Tipo | Notas / riesgo |
|--------------|------|----------------|
| `routes.py` ↔ `scheduler.py` | Mutuo | `routes` crea Tasks que `scheduler` ejecuta. Si renombras `add_or_replace_job` o cambias firma de `execute_task`, hay que actualizar ambos. |
| `database.py` ↔ todo | One-to-many | Cualquier columna nueva exige tocar también `_migrate_existing_db` y, según el caso, el dict serializador en `routes._task_to_dict`. |
| `index.html` ↔ `routes.py` | API contract | El frontend conoce paths y shapes de respuesta. Renombrar un endpoint o cambiar una key del JSON rompe la UI silenciosamente. Grep `fetch(`/`api(` en `index.html` antes de cambiar. |
| `index.html` ↔ `database.py` | Indirecto | Campos como `task.task_type`, `t.retry_count`, `r.retry_attempt`, `r.pivots_ok` aparecen directamente en la UI. Renombrar columna = romper UI. |
| `auth.py` → `database.User` | Dura | `_get_current_user_optional` hace `select(User).where(username == ...)`. Cambiar `User.username` impacta todo el login. |
| `excel_engine.py` ↔ `pywin32` | Externa, frágil | COM no tiene contrato estable; comportamiento depende de versión de Office. Documentado en `excel_engine.py`. |
| `scheduler.py` → `workflows.registry` | Dinámica | Resolución por nombre en runtime. Si el nombre en `pipeline_config["workflow_type"]` no está en el registry → falla en ejecución, no al guardar. |
| `notifications.py` ↔ servicios externos | SMTP + CallMeBot HTTP | Sin reintento ni circuit breaker; un proveedor caído puede retrasar shutdown de la tarea. |

---

## 3. Servicios externos / dependencias de runtime

| Servicio | Donde se usa | Comportamiento si falla |
|----------|--------------|--------------------------|
| **Microsoft Excel** (COM) | `excel_engine.py` | Tarea Excel falla; pipeline ETL falla. Reintento por `lock_max_retries`. |
| **Microsoft Access** (COM) | `access_engine.py` | Tareas pipeline fallan; tareas Excel puras siguen OK. |
| **OneDrive Files On-Demand** | Indirecto vía `file_path` | `FileNotFoundError`; sin pre-hidratación. |
| **SMTP server** | `notifications.py` (email) | Notificación email se loggea como warning, no bloquea la tarea. |
| **CallMeBot HTTPS** | `notifications.py` (WhatsApp) | Igual: warning, no bloquea. |
| **Webhook genérico** (`WEBHOOK_URL`) | `scheduler.send_webhook` | Silenciado vía try/except. |
| **APScheduler internal store** | `scheduler.py` | En memoria; al reiniciar los jobs se reconstruyen desde DB en `load_all_tasks()`. |

---

## 4. Archivos peligrosos de modificar

| Archivo | Por qué es peligroso | Reglas para tocarlo |
|---------|---------------------|----------------------|
| **`app/static/index.html`** (2900+ líneas) | Monolítico, IDs cruzados, JS y HTML mezclados. Ya hubo una catástrofe (commit `0a7b3b4` borró ~50% de la UI). | Edits quirúrgicos, nunca sed masivo. Diff completo antes de commit. Confirmar que los `colspan="N"` siguen cuadrando con las columnas. |
| **`app/database.py`** | Cambios de schema = migración manual via ALTER. Sin Alembic. | Añadir columna: 1) en modelo, 2) en `_migrate_existing_db`. Para borrar/renombrar: NO — añadir columna nueva, deprecar, eventualmente limpiar (raramente vale la pena). |
| **`app/scheduler.py::execute_task`** | Es el bucle central. Bugs aquí = tareas no corren o RunLog huérfanos. | Mantener `try/finally` que cierra el RunLog. No agregar `awaits` largos sin timeout. |
| **`app/excel_engine.py`** | COM frágil + recursos no liberados → memory leak en Excel. | Mantener `try/finally` con `app.Quit()` / `wb.Close(False)`. No abrir Excel sin `DispatchEx`. |
| **`app/access_engine.py`** | Mismo riesgo que Excel + Access es aún más temperamental. | Idem. |
| **`install-service.ps1`** | Scripts PS5 con encoding Windows-1252; em-dashes o caracteres no-ASCII rompen ejecución. | ASCII puro. Ya hubo fix `11a11a0` / `88054f0` por esto. |
| **`pyproject.toml`** + `poetry.lock` | Cambios de versiones grandes pueden romper en sistemas con Python distinto. | Actualizar una dep a la vez; correr `poetry install && pytest`. |
| **`.env`** (no en repo) | Reconfiguración silenciosa: cambiar `JWT_SECRET` invalida todos los tokens activos. | Validar con `create_superadmin.py` después de cambiar secrets. |
| **`scheduler.db`** | DB en uso por el servicio. Editar a mano puede corromper. | Detener servicio (`Stop-ScheduledTask`) antes de cualquier escritura. |

---

## 5. Dependencias externas (PyPI) — usadas dónde

| Package | Usado en |
|---------|----------|
| fastapi | `app/main.py`, `app/routes.py`, `app/auth_routes.py` |
| uvicorn | `app/main.py::start()` |
| apscheduler | `app/scheduler.py` |
| sqlalchemy + aiosqlite | `app/database.py` (motor), `routes/auth_routes/scheduler/notifications` (queries) |
| pydantic + pydantic-settings | `app/config.py` (Settings), routes/auth_routes (models) |
| python-jose | `app/auth.py` (JWT encode/decode) |
| bcrypt | `app/auth.py` (hash/verify) |
| python-multipart | `app/auth_routes.py` (OAuth2PasswordRequestForm) |
| aiosmtplib | `app/notifications.py` (email) |
| httpx | `app/scheduler.py` (webhook), `tests/` (TestClient async) |
| openpyxl | `app/excel_engine.py` (puntual, sin COM) |
| pywin32 | `app/excel_engine.py`, `app/access_engine.py` (Windows only) |

---

## 6. Endpoints (superficie pública)

> Todos bajo `/api` + dependencia auth (excepto `/auth/login`).

### Tareas
- `GET    /tasks`                                — listar (con next_run_at)
- `POST   /tasks`                                — crear Excel
- `POST   /tasks/pipeline`                       — crear pipeline Access ETL
- `POST   /tasks/pipeline/reposicion`            — preset Reposición (atajo)
- `POST   /tasks/workflow/weekly-excel-copy`     — crear workflow semanal
- `GET    /tasks/{id}` / `PUT` / `DELETE` (soft) — CRUD
- `POST   /tasks/{id}/pause` / `/resume`
- `POST   /tasks/{id}/run-now`
- `POST   /tasks/{id}/test-run`                  — workflow only, `force_weekday`
- `GET    /tasks/export` / `POST /tasks/import`  — JSON

### Logs
- `GET    /logs?page&page_size&task_id&status`
- `DELETE /logs?task_id&status`
- `GET    /logs/{run_id}/tail?offset=N`          — live polling
- `GET    /logs/{run_id}/content`                — full
- `GET    /logs/{run_id}/download`
- `POST   /logs/{run_id}/stop`

### Reports
- `GET/POST /reports`, `GET/PUT/DELETE /reports/{id}`, `POST /reports/{id}/run-now`

### Notifications
- `GET/POST /tasks/{id}/notifications`, `DELETE /notifications/{rule_id}`

### Misc
- `GET    /stats`
- `GET    /browse-file?filter=excel|access|any`  — selector nativo Windows
- `GET    /browse-folder`
- `POST   /admin/cleanup-stuck-runs`

### Auth (`/api/auth/`)
- `POST   /login` (form)
- `GET/POST /users`, `GET/PATCH/DELETE /users/{id}`, `POST /users/{id}/reset-password`
- `GET    /me`, `POST /me/change-password`

### Sistema (raíz)
- `GET    /health`
- `GET    /login` → sirve `login.html`
- `GET    /` y catch-all SPA → sirve `index.html` (no si empieza con `api`)

---

## 7. Puntos de acoplamiento entre frontend y backend

Estos son los "contratos implícitos" que rompes si renombras algo en el backend:

| Frontend (JS) | Backend (clave esperada) |
|---------------|--------------------------|
| `t.task_type`, `t.last_run_at`, `t.last_run_status`, `t.retry_count`, `t.max_retries`, `t.next_run_at`, `t.status` | `_task_to_dict` en `routes.py` |
| `r.duration_s`, `r.retry_attempt`, `r.pivots_ok`, `r.pivots_err`, `r.connections`, `r.log_file` | columnas de `RunLog` |
| `data.items`, `data.total`, `data.offset`, `data.content`, `data.status` (en `/logs/.../tail`) | paginación + tail estructura |
| `me.role`, `me.username`, `me.full_name`, `me.email` | `/auth/me` shape |
| `u.is_active`, `u.last_login` | `/auth/users` shape |
| `pipeline_config.folder`, `.file_patterns`, `.week_padding`, `.daily_refresh`, `.fail_if_source_missing`, `.excel_visible`, `.refresh_timeout`, `.pivot_guards` | workflow weekly_excel_copy |
| `pipeline_config.access_db`, `.excel_files`, `.compact_before_import`, `.pre_import_macros`, `.saved_imports`, `.post_import_macros`, `.access_visible` | pipeline Access |

Antes de tocar cualquier campo, grep en `app/static/index.html` la string exacta.

---

_Última actualización: 2026-05-24_

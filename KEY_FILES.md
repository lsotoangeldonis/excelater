# KEY_FILES — Los 20 archivos más importantes de Excelater

> Si tuvieras que entender el proyecto leyendo sólo 20 archivos, estos son y en este orden.
> El número entre paréntesis es el conteo de líneas aproximado al 2026-05-24.

---

## Tier 1 — Imprescindibles (si lees 5, lee estos)

### 1. `app/main.py` (141)
**Por qué:** punto de entrada del servicio. Aquí está el `lifespan` de FastAPI (init DB, cleanup de runs colgados, arranque del scheduler, carga de tareas), el montaje del SPA, el catch-all que sirve `index.html`, y `start()` que arranca uvicorn. Cualquier cosa que falle al arrancar pasa por aquí.

### 2. `app/database.py` (230)
**Por qué:** define todo el vocabulario del proyecto. Modelos (`Task`, `RunLog`, `NotificationRule`, `ReportSchedule`, `User`) + enums + la función `_migrate_existing_db` (importantísima: añade columnas a DBs existentes vía ALTER + try/except, porque no hay Alembic).

### 3. `app/routes.py` (1298)
**Por qué:** superficie API completa. Endpoints REST para tareas, logs, reportes, notificaciones, browse, stats. Es el archivo más largo del backend; léelo de arriba abajo escaneando los `@router.X(...)` para mapear la API mental.

### 4. `app/scheduler.py` (502)
**Por qué:** el corazón ejecutor. Contiene `execute_task(task_id, config_overrides=None)` — la función que efectivamente corre las tareas. También: registro/cancelación de jobs APScheduler, `load_all_tasks` (rehidratación al arranque), webhook, y el dispatch por `task_type`.

### 5. `app/static/index.html` (2928)
**Por qué:** el dashboard. HTML + CSS + JS vanilla en un único archivo. Toda la UX vive aquí: tablas ordenables, modales, auto-refresh, manejo de JWT, role-based hide/show, formularios de creación de los 3 tipos de tarea. Es enorme y monolítico — leer sólo cuando vayas a tocarlo, y siempre con grep antes.

---

## Tier 2 — Críticos para áreas específicas

### 6. `app/auth.py` (127)
**Por qué:** todo lo relativo a JWT, bcrypt y dependencias por rol. Función clave: `_get_current_user_optional` (decodifica token + carga User + valida `is_active`). Si `auth_enabled=false` corto-circuita y retorna `None`.

### 7. `app/auth_routes.py` (273)
**Por qué:** endpoints de `/api/auth/*` — login (form OAuth2PasswordRequestForm), CRUD de usuarios, change/reset password, `/me`. Lee junto con `auth.py` para entender el flujo completo.

### 8. `app/excel_engine.py` (448)
**Por qué:** la capa COM contra Excel. Función pública: `run_update(EngineConfig)`. Maneja apertura con `DispatchEx`, refresh de conexiones externas, refresh de PivotTables, save_on_success, lock retry (`lock_max_retries`, `lock_retry_s`), visibilidad. Es donde están los bugs reales de COM cuando algo se cuelga.

### 9. `app/access_engine.py` (313)
**Por qué:** equivalente a `excel_engine.py` pero para Access. Macros, Compact&Repair, importaciones guardadas. Sólo usado por tareas tipo `pipeline`.

### 10. `app/workflows/weekly_excel_copy.py` (450)
**Por qué:** ejemplo de workflow completo. Implementa `BaseWorkflow.run`, hace la lógica de "copia archivo Sem N-1 a Sem N cada lunes, refresca, guarda con nombre nuevo". Soporta `force_weekday` para test-run y `pivot_guards` para evitar bloqueos de tablas dinámicas. Usar como plantilla para nuevos workflows.

### 11. `app/notifications.py` (273)
**Por qué:** dispatch de notificaciones post-ejecución. Email (SMTP via aiosmtplib) + WhatsApp (CallMeBot via HTTPS). Filtros por `TriggerType` (always/on_error/on_success/first_run_of_day).

### 12. `app/config.py` (75)
**Por qué:** todas las variables de entorno en un solo lugar. Si vas a añadir una env nueva, va aquí. Cualquier setting con default seguro va con un valor por defecto en la clase `Settings`.

---

## Tier 3 — Importantes pero específicos

### 13. `app/static/login.html` (250)
**Por qué:** página de login independiente. Envía form a `/api/auth/login`, guarda token en `localStorage`, redirige a `/`. Si quieres rebrandear la pantalla de login, está acá.

### 14. `app/workflows/__init__.py` (33)
**Por qué:** el `WorkflowRegistry` singleton. Si añades un workflow, lo registras aquí con `registry.register("nombre", Clase)`. Pequeño pero crítico — sin el register, el workflow no es descubrible.

### 15. `app/workflows/base.py` (31)
**Por qué:** `BaseWorkflow` (ABC). Define el contrato `run(config: dict, logger) -> EngineResult`. Todo workflow nuevo debe heredar de aquí.

### 16. `tests/conftest.py` (68)
**Por qué:** fixtures compartidos: client async, DB SQLite en memoria, overrides de auth dependencies. Mira esto antes de escribir un test nuevo.

### 17. `tests/test_routes.py` (244)
**Por qué:** patrón canónico de tests de endpoints. Copiar la estructura cuando añadas un endpoint nuevo.

### 18. `pyproject.toml` (44)
**Por qué:** declara stack y versiones. Si alguien pregunta "qué versión usa X", la respuesta está aquí. También el entry point `excelater = "app.main:start"`.

---

## Tier 4 — Operación y deploy

### 19. `install-service.ps1` (~17 KB)
**Por qué:** único método soportado de instalación en producción. Registra Scheduled Task de Windows, detecta venv de Poetry, instala dependencies, hace check de superadmin. Es idempotente — re-correr es seguro y fuerza re-registro si detecta `cmd.exe` legacy.

### 20. `deploy.ps1` (~7.4 KB)
**Por qué:** hot-update del servicio ya instalado, sin re-registrar el Scheduled Task. Lo más usado en operación día a día tras un cambio en main.

---

## Bonus: archivos a NO ignorar pero que no entran en el top 20

- **`scripts/create_superadmin.py`** (165) — corre la primera vez para crear el superadmin y persistir `JWT_SECRET`.
- **`scripts/reset_password.py`** (81) — recuperar acceso si se pierde una contraseña.
- **`README.md`** — quickstart de usuario final (no de developer; para eso usar `ONBOARDING.md`).
- **`.env.example`** (no en el listing porque no fue leído, pero debe existir) — template de configuración.

---

## Qué archivos NO leer (gitignored o autogenerados)

- `.venv/` — venv de Poetry.
- `logs/*` — logs generados.
- `scheduler.db` — DB SQLite.
- `.env` — secretos locales.
- `poetry.lock` — léelo solo si depuras versions exactas.
- `app/__init__.py` — vacío.

---

_Última actualización: 2026-05-24_

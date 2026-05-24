# ONBOARDING — Excelater

> Objetivo: que un developer nuevo entienda el proyecto y lo tenga corriendo localmente en **menos de 30 minutos**.

---

## 0. Antes de empezar (verifica que tienes)

- **Windows 10 / Server 2019+** — no funciona en macOS/Linux (usa COM).
- **Python 3.10 a 3.13** — verifica con `python --version`.
- **Poetry ≥ 1.8** — instalar: https://python-poetry.org/docs/#installation
- **Microsoft Excel** de escritorio (cualquier versión 2016+) — obligatorio para tareas que refrescan archivos reales. Para sólo navegar el dashboard / correr tests no es necesario.
- **Acceso al repo** (`git clone ...`).

Opcional (sólo si vas a tocar pipelines Access ETL):
- **Microsoft Access** instalado.

---

## 1. Setup local (≈10 min)

```powershell
# 1. Clonar
git clone <repo-url> Excelater
cd Excelater

# 2. Instalar dependencias en venv aislado
poetry install

# 3. Configuración
copy .env.example .env
# Editar .env si necesitas cambiar HOST/PORT, deshabilitar auth, etc.
# Para desarrollo rápido, basta con poner: AUTH_ENABLED=false
# (Si lo dejas true, hace falta crear superadmin en el siguiente paso.)

# 4a. (Si auth_enabled=true) Crear superadmin + generar JWT_SECRET persistente
poetry run python scripts/create_superadmin.py
# Pide username, full_name, email, password. Persiste JWT_SECRET en .env.

# 4b. (Si auth_enabled=false) Saltar este paso.

# 5. Verificar que arranca
poetry run excelater
# Output esperado:
#   [Scheduler] N tarea(s) cargada(s).
#   [Server] Dashboard en http://0.0.0.0:8000
# Abrir http://localhost:8000 — debería pedirte login (o entrar directo si auth_enabled=false).
```

---

## 2. Lectura mínima para entender el proyecto (≈10 min)

Lee en este orden:

1. **`CLAUDE_CONTEXT.md`** — TL;DR técnico (5 min).
2. **`app/main.py`** (141 líneas) — bootstrap, lifespan, montaje SPA.
3. **`app/database.py`** — modelos (Task, RunLog, NotificationRule, ReportSchedule, User) + enums. Aquí está el "vocabulario" del sistema.
4. **`app/scheduler.py::execute_task`** — el corazón de la ejecución. Lee solo esa función (no todo el archivo).
5. **`app/routes.py`** — escanea índice (busca `@router.`); no leer todo.

Si vas a tocar UI:
6. **`app/static/index.html`** — abre, lee solo el `<script>` del final. Busca `function api(`, `function goPage`, `function loadTasks`, `function saveTask`. Suficiente para orientarte.

---

## 3. Correr tests (≈2 min)

```powershell
poetry run pytest                  # debería pasar 36+ tests en < 5 segundos
poetry run pytest -k workflow      # filtrado por keyword
poetry run pytest --tb=short -x    # parar en el primer fallo, traceback corto
```

Los tests usan DB SQLite en memoria y mockean COM — no abren Excel real.

---

## 4. Hacer el primer cambio (≈5 min) — smoke test

Verifica que tu setup funciona haciendo un cambio trivial:

1. Edita `app/main.py` línea 49:
   ```python
   print(f"[Server] Dashboard en http://{settings.host}:{settings.port}")
   ```
   Cámbialo a:
   ```python
   print(f"[Server] [{tu_inicial}] Dashboard en http://{settings.host}:{settings.port}")
   ```
2. Reinicia con `Ctrl+C` y `poetry run excelater`.
3. Verifica que tu marca aparece en el log de arranque.
4. **Revierte** el cambio (`git checkout app/main.py`).

Si llegaste aquí, tienes todo funcionando.

---

## 5. Cómo está organizado el código (mental model)

```
Request HTTP  →  routes.py  →  Pydantic validator
                                 │
                                 ▼
                            database.py (SQLAlchemy async)
                                 │
                                 ▼
                            scheduler.py (add/remove jobs)
                                 │
                            APScheduler dispara
                                 │
                                 ▼
                          execute_task(task_id)
                                 │
                ┌────────────────┼────────────────┐
                ▼                ▼                ▼
       excel_engine.py   access_engine.py   workflows/<x>.py
            (COM)             (COM)         (compone los anteriores)
                                 │
                                 ▼
                          RunLog (DB) + log file
                                 │
                                 ▼
                       notifications.py (email/WhatsApp)
```

Características conceptuales:
- **Auth y datos** son ortogonales: `auth.py` provee `Depends(require_*)` que se inyecta donde haga falta; no hay middleware global de auth.
- **Tareas** tienen 3 tipos: `excel` (un archivo), `pipeline` (Access ETL), `workflow` (flujo personalizado registrado en `workflows/`).
- **Schedule** es independiente del tipo: una tarea workflow puede usar cron, interval o once_daily.
- **Notificaciones** son por-tarea (`NotificationRule`) o globales por reporte (`ReportSchedule`).

---

## 6. Cosas que vas a necesitar saber pronto

### 6.1 Logs
- **`logs/excelater.log`** — log del servicio (startup, shutdown, eventos globales).
- **`logs/task_<task_id>_<run_id>.log`** — log de cada ejecución (rotating, 10 MB × 5 backups).
- Para ver vivo: `Get-Content -Tail 100 -Wait logs\excelater.log`

### 6.2 Migraciones de schema
**NO hay Alembic.** Si añades una columna:
1. Añádela al modelo en `database.py`.
2. Añádela a la lista en `_migrate_existing_db` con `ALTER TABLE ... ADD COLUMN ...`.
3. El try/except silencia "ya existe", así que es seguro re-ejecutar.

### 6.3 Reset / borrar DB local
```powershell
# Detener cualquier instancia primero
Remove-Item scheduler.db
poetry run excelater  # se recrea vacía
poetry run python scripts/create_superadmin.py
```

### 6.4 Frontend
- No hay build. Editas `app/static/index.html` y refrescas el navegador.
- Cambios al `index.html` no requieren reiniciar el servidor (uvicorn re-sirve el archivo).
- Para inspeccionar requests: DevTools → Network. Cada uno lleva `Authorization: Bearer <jwt>`.

### 6.5 Entrar al dashboard sin auth (desarrollo)
```bash
# en .env
AUTH_ENABLED=false
```
Reiniciar. El login deja de exigirse en endpoints; la UI sigue mostrando widgets de usuario con valores vacíos.

---

## 7. Convenciones rápidas

- **Commits:** español, prefijo (`feat:`, `fix:`, `refactor:`, `test:`, `docs:`).
- **Python:** type hints modernos (`list[X]`, `X | None`), async end-to-end, pydantic para todo body.
- **Frontend:** vanilla JS; helper `api()` para llamadas; `esc()` antes de `innerHTML`; `toast()` para feedback.
- **PowerShell scripts:** ASCII puro (PS5 compat).
- **No usar:** passlib (migrado a bcrypt), `asyncio.get_event_loop()` (usar `get_running_loop()`).

---

## 8. Documentación complementaria

Cuando necesites ir más profundo:
- **`PROJECT_CONTEXT.md`** — referencia exhaustiva (arquitectura, riesgos, comandos).
- **`DEPENDENCY_MAP.md`** — quién depende de quién, archivos peligrosos.
- **`PROMPTS.md`** — plantillas para tareas comunes (con o sin Claude).
- **`KEY_FILES.md`** — los 20 archivos más importantes, en orden.
- **`README.md`** — quickstart de usuario final (no de developer).

---

## 9. Si te trabas

| Síntoma | Posible causa / qué revisar |
|---------|------------------------------|
| `poetry install` falla en `pywin32` | Estás en macOS/Linux — el proyecto sólo corre en Windows. El marker `sys_platform == 'win32'` debería omitirlo, pero puede haber errores en otras deps. |
| `poetry run excelater` arranca pero `/` da 404 | `app/static/index.html` no existe o está vacío. Verifica `git status`. |
| Login funciona pero después todo da 401 | `JWT_SECRET` cambió o está vacío sin haber corrido `create_superadmin.py`. Tokens emitidos antes ya no validan. Vuelve a hacer login. |
| "No module named app.X" | Estás corriendo Python fuera del venv. Usa `poetry run python ...`. |
| Una tarea de prueba no se ejecuta | Revisa `Get-ScheduledTaskInfo -TaskName Excelater` o el log; puede que el scheduler no esté arrancado, o la próxima ejecución sea en el futuro. Usa "Ejecutar ahora" desde la UI. |
| Excel se queda colgado | Probablemente otro proceso (Excel del usuario) tiene el archivo abierto, o Excel mostró un diálogo modal. Cerrar Excel, mirar `logs/task_*.log`. |

---

## 10. Próximos pasos sugeridos (tu primera semana)

1. **Día 1:** completar este onboarding. Familiarizarse con la UI.
2. **Día 2:** elegir un bug pequeño o feature menor de `PROMPTS.md` y aplicarlo end-to-end (código + test + verificar en UI).
3. **Día 3:** leer `app/scheduler.py` completo y trazar mentalmente una ejecución desde el cron hasta el cierre del `RunLog`.
4. **Día 4:** lectura completa de `app/routes.py` (es largo pero plano).
5. **Día 5:** explorar `excel_engine.py` y `workflows/weekly_excel_copy.py` — donde está la lógica de negocio real.

---

_Última actualización: 2026-05-24_

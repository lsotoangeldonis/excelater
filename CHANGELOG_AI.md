# CHANGELOG_AI — Bitácora de cambios significativos

> Registro de cambios mayores hecho con/para asistencia de IA.
> Cada entrada documenta el "porqué" y el "blast radius", no sólo el diff (que el commit ya tiene).
>
> **Cómo se actualiza:** este archivo no se actualiza solo. Para cada cambio significativo, pídele a Claude:
> _"Añade entrada en `CHANGELOG_AI.md` con los cambios de los últimos N commits y, si corresponde, actualiza `CLAUDE_CONTEXT.md` / `AI_SUMMARY.md`."_
> Si quieres automatizarlo en un hook (que se dispare al hacer commit), pídeme: _"configura un hook para que recuerde actualizar CHANGELOG_AI.md después de cada git commit"_ y lo dejo en `settings.json` vía el skill `update-config`.

---

## Formato de entrada

```markdown
## YYYY-MM-DD — Título corto

**Commits:** `<sha1>` `<sha2>` …  
**Motivo:** por qué se hizo.  
**Qué cambió:** resumen funcional (no diff).  
**Blast radius:** módulos/archivos afectados, contratos rotos o respetados.  
**Documentación actualizada:** lista de .md tocados (o "ninguna").  
**Notas para futuro:** trampas, decisiones que un cambio futuro podría querer revertir.
```

---

## 2026-05-24 — Feat: trigger `on_final_failure` para alertas tras agotar reintentos

**Commits relacionados:** sin commitear aún (working tree).

**Motivo:** con `max_retries=3`, una regla `on_error` enviaba 4 emails (uno por intento). El operador pedía una alerta única cuando la tarea sigue fallando tras agotar todos los reintentos, manteniendo `on_error` como opción para quienes prefieren la señal inmediata por intento.

**Qué cambió:**

- **`app/database.py`**: nuevo valor `ON_FINAL_FAILURE = "on_final_failure"` en el enum `TriggerType`. No requiere migración SQL: la columna `notification_rules.trigger` es `VARCHAR(16)` sin CHECK constraint (verificado vía `sqlite_master`), y `"on_final_failure"` cabe en 16 chars exactos.
- **`app/notifications.py`**:
  - Import: añadido `Task`.
  - `_should_notify`: nueva rama para `ON_FINAL_FAILURE`. Devuelve `True` sólo si `run.status == FAILED` y `(run.retry_attempt or 0) >= (task.max_retries or 0)`. Carga `Task` con `db.get()` por id.
  - Diseño: con `max_retries=0` (sin reintentos) la condición sigue cumpliéndose, así que el trigger se comporta como `on_error` y es seguro como reemplazo directo.
- **`app/static/index.html`**:
  - Modal `modal-notif-rule` (línea ~1260): nueva opción `"Solo errores (tras agotar reintentos)"` en el `<select id="nr-trigger">`. Etiqueta de `on_error` aclarada a "Solo errores (cada intento)" para diferenciarlas.
  - `TRIGGER_LABELS` (línea ~2888): añadida entrada `on_final_failure: 'Tras agotar reintentos'`.

**Blast radius:**

- Triggers existentes (`always`, `on_error`, `on_success`, `first_run_of_day`) se conservan sin cambios — backward-compatible.
- El despacho ocurre en `scheduler.py:336-349` después de la lógica de retry: cuando se agotan los reintentos, `task.retry_count` se resetea a 0 pero `RunLog.retry_attempt` ya quedó congelado al valor del último intento (igual a `max_retries`), por lo que la comparación funciona.
- UI: el frontend no requiere recarga forzada porque las reglas se cargan vía API; usuarios con caché agresiva pueden necesitar Ctrl+F5 para ver la nueva opción del select.

**Documentación actualizada:** `CLAUDE_CONTEXT.md` (lista de triggers en la sección de modelos clave).

**Notas para futuro:**

- Si en el futuro se quiere que `on_error` dispare sólo en intermedios (y excluya el final), añadir simetría: comparar `retry_attempt < max_retries`. Hoy `on_error` sigue disparando en TODOS los intentos para no romper integraciones existentes.
- El email del trigger `on_final_failure` no incluye conteo de reintentos en el subject ni el body. Si se quiere visualizar "falló 4 veces", añadir `run.retry_attempt` al template `_format_task_message_html` y `_format_task_message_text` en `notifications.py`.
- No hay UI dedicada para distinguir reglas duplicadas (p.ej. usuario configura `on_error` + `on_final_failure` y recibe 5 emails); considerar warning en `saveNotifRule()` si se detecta esa combinación en la misma tarea.

---

## 2026-05-24 — Fix: pipeline colgaba RunLog en "RUNNING" para siempre (ImportError silencioso)

**Commits relacionados:** sin commitear aún (working tree).

**Motivo:** una ejecución de la tarea "Herramienta de Reposición" (pipeline Access ETL) quedaba marcada como **EN VIVO** indefinidamente en el UI tras pegar:

```
=== INICIO TAREA: Herramienta de Reposición (id=58a58d1d-...) ===
[Excel] Procesando: Cubo_SKU_SUC_Maestro.xlsm
```

No había `EXCEL.exe` vivo. Diagnóstico: 3 bugs encadenados.

**Bug 1 — Import roto** (`app/access_engine.py:160`):
- El commit del Pipeline Access ETL hardening (entrada de arriba) introdujo `from app.excel_engine import EngineConfig, run_engine`. Pero `excel_engine.py` solo expone `run_update`. `run_engine` nunca existió.
- Al primer `_refresh_excel_file()` se lanza `ImportError: cannot import name 'run_engine' from 'app.excel_engine'`.

**Bug 2 — Excepción tragada** (`app/scheduler.py:154-160` y dos sitios más):
- `execute_task` solo capturaba `asyncio.CancelledError`. Cualquier `Exception` (incluido el `ImportError` de arriba) burbujeaba al hilo de APScheduler, que la loggueaba en `excelater.log` y abortaba `execute_task` antes de la rama que actualiza el `RunLog`.
- Resultado: `RunLog.status` quedaba en `RUNNING` para siempre, sin `finished_at`, sin `error_msg`. El UI lo mostraba "EN VIVO" eternamente.

**Bug 3 — `excelater.log` ilegible** (`install-service.ps1:319`):
- El operador `*>>` de PowerShell 5.1 escribe **UTF-16 LE**. Cada caracter ASCII pasa a ser 2 bytes (el segundo `0x00` se renderiza como espacio en `Get-Content`). El traceback del Bug 1 estaba ahí pero ilegible — por eso costó tanto encontrar root cause.

**Qué cambió:**

- **`app/access_engine.py`** (líneas 155, 160, 175): `run_engine` → `run_update`.
- **`app/scheduler.py`**:
  - Nuevo import: `traceback`.
  - Nuevo helper local `_unhandled_error_result(exc)` dentro de `execute_task` que construye un "EngineResult" sintético (`success=False`, `error_msg`, `duration_s`, etc.) cuando una excepción no esperada llega al try.
  - Los **3 bloques** (`pipeline`, `workflow`, Excel estándar) añaden `except Exception as exc: result = _unhandled_error_result(exc)`. Garantiza que el flujo de cierre de `RunLog` (status, finished_at, last_run_status) siempre se ejecute.
- **`install-service.ps1`**:
  - Línea de comando: `*>> '$log'` → `2>&1 | Out-File -FilePath '$log' -Encoding utf8 -Append`. Stdout y stderr van en UTF-8.
  - Pre-check: si el `excelater.log` existente está en UTF-16 (BOM `FF FE` o `byte[1] == 0x00`), lo rota a `excelater.log.utf16.bak` para que el nuevo handler arranque limpio. Abre con `FileShare.ReadWrite` por si el servicio aún lo tiene abierto.

**Blast radius:**
- `execute_task` ahora **siempre** finaliza el RunLog correctamente. Tareas con excepciones inesperadas pasan de `RUNNING` zombie a `FAILED` con `error_msg` legible.
- El cambio en `install-service.ps1` solo aplica al **siguiente** `install-service.ps1`. La Scheduled Task ya instalada sigue corriendo con la línea de comando vieja hasta que se re-registre.
- Logs anteriores en UTF-16 se preservan como `.utf16.bak` (no se pierden).

**Documentación actualizada:** este `CHANGELOG_AI.md`. Corregida también la mención errónea a `run_engine()` en la entrada inmediatamente posterior.

**Acciones manuales pendientes en PROD (E:\Automation\excelater)** — el dev (D:) sí está fixed:
1. `git pull` (o `.\deploy.ps1`) para traer los fixes de código.
2. Re-ejecutar `.\install-service.ps1` para que la Scheduled Task use la nueva redirección UTF-8 y rote el log viejo.
3. Marcar el RunLog huérfano (`id=18`) como `failed` para limpiar el "EN VIVO" del UI:
   ```sql
   UPDATE run_logs SET status='failed', finished_at=datetime('now'),
     error_msg='ImportError run_engine (fix en commit XXXX)' WHERE id=18 AND status='running';
   UPDATE tasks SET last_run_status='failed' WHERE id='58a58d1d-8e7f-4340-86bf-160bb27485c3';
   ```

**Notas para futuro:**
- Si en algún momento se agrega un cuarto camino de ejecución en `execute_task`, recordar replicar el `except Exception → _unhandled_error_result`.
- La detección de UTF-16 en `install-service.ps1` es heurística (BOM + null en byte 1). Si el log contiene primero un caracter ASCII real codificado como UTF-8 (1 byte) seguido de un null binario por accidente, falsearía positivo. Improbable en la práctica con logs de uvicorn.
- Largo plazo: lo ideal sería configurar `uvicorn --log-config logging.yaml` para que Python escriba el log directamente (sin redirección PowerShell), evitando todo el tema de encoding. Fuera de scope acá.

---

## 2026-05-24 — Pipeline Access ETL: hardening + paso 8 (post-refresh tableros)

**Commits relacionados:** sin commitear aún (working tree).

**Motivo:** el `AccessPipelineRunner` no replicaba fielmente el manual de actualización (Sneakers / Non-Sneakers). Detectados:
- **Bug**: `_refresh_excel_file` invocaba `ExcelCOMUpdater(...).run()` directo y se saltaba `run_engine()`, perdiendo hidratación OneDrive + `wait_for_file` que el flujo Excel-puro sí tiene.
- **Gap orden**: Compact & Repair se ejecutaba ANTES del pre_import_macro "Elimina Cubos", compactando datos viejos. El manual hace `Elimina Cubos → Compact → Importar`.
- **Gap paso 8**: el pipeline terminaba en `post_import_macros`. El paso 8 manual ("Actualizar Tableros, Reportes y Herramientas") nunca se ejecutaba, dejando los `.xlsm` consumidores con datos viejos hasta refresh manual.
- **Riesgos COM no mitigados**: `.accdb` en OneDrive sin pre-hidratación, BD bloqueada por sesión interactiva sin lock-wait, prompts modales (AutoExec / SetWarnings) que cuelgan COM hasta el timeout de 300s.

**Qué cambió:**

- **`app/access_engine.py`:**
  - Bug fix (intencional): `_refresh_excel_file` debía usar el wrapper público de `excel_engine` (no `ExcelCOMUpdater().run()`) para heredar hidratación OneDrive + `wait_for_file`. **Ojo**: este commit introdujo además un bug regresivo — referenciaba `run_engine` (símbolo inexistente) en vez de `run_update`. Ver entrada del 2026-05-24 más arriba para el fix.
  - Nuevo `_prepare_access_db()` que detecta placeholder OneDrive, dispara descarga, y espera lock del `.accdb` reutilizando helpers de `excel_engine`.
  - `_run_access_operations` parametrizado (`run_pre_macros` / `run_imports` / `run_post_macros`) para poder ejecutarse en dos sesiones cuando hay Compact en medio.
  - `_compact_repair` y `_run_access_operations` ahora aplican `AutomationSecurity = 3` (msoAutomationSecurityForceDisable) y `DoCmd.SetWarnings(False)` para silenciar prompts.
  - `PipelineConfig` extendido con: `compact_position` (`"" | "before_macros" | "after_pre_macros" | "skip"`), `post_refresh_excel_files`, `continue_on_error`, `access_lock_timeout`, `access_lock_retry`.
  - Default nuevo: `compact_position` resuelve a `"after_pre_macros"` cuando `compact_before_import=True` (replica orden manual). Tareas que dependen del orden viejo deben setear `compact_position: "before_macros"`.
  - Orquestador `run()`: nuevo Paso 1.5 (preparar .accdb), tres ramas según `compact_position` (incluye cerrar/reabrir Access cuando compact va entre pre-macros e imports), nuevo Paso 8 que refresca `post_refresh_excel_files` reutilizando `_refresh_excel_file`.
  - Helper `_attempt()` en `_run_access_operations` para encapsular semántica fail-fast vs `continue_on_error` por macro/import individual.
  - Cuando `continue_on_error` deja fallos parciales, `success=False` + `error_msg` con conteo → notificaciones `on_error` se disparan correctamente.

- **`app/scheduler.py`** ([líneas 137-152](app/scheduler.py#L137-L152)): cableado de los campos nuevos del JSON al `PipelineConfig`.

- **`app/routes.py`:**
  - `PipelineTaskCreate` y `ReposicionTaskCreate` extendidos con `compact_position`, `post_refresh_excel_files`, `continue_on_error`.
  - `create_pipeline_task` y `create_reposicion_pipeline_task` validan existencia de los archivos `post_refresh_excel_files`.
  - `_build_reposicion_pipeline_cfg` ahora propaga los 3 campos nuevos al JSON.

- **`app/static/index.html`** (sección pipeline):
  - Nuevo textarea `f-post-refresh-files` (TABLEROS A REFRESCAR DESPUÉS DEL ETL).
  - Nuevo select `f-compact-position` (auto / after_pre_macros / before_macros / skip).
  - Nuevo toggle `f-continue-on-error`.
  - `openEditModal` carga los 3 campos nuevos.
  - `saveTask` los envía en `pipelineCfg`.

**Blast radius:**
- API `POST /api/tasks/pipeline`: contrato extendido con campos opcionales. Backward compatible — clientes viejos siguen funcionando.
- **CAMBIO DE COMPORTAMIENTO**: tareas pipeline ya creadas en producción que tenían `compact_before_import=true` (default) ahora ejecutan Compact DESPUÉS de pre_import_macros en vez de antes. Es lo correcto según el manual, pero cambia timing y duración. Para forzar orden legacy explícito: `compact_position: "before_macros"`.
- Cuando `compact_position == "after_pre_macros"`, Access se abre/cierra DOS veces (sesión 1: pre-macros, sesión 2: imports + post-macros) en vez de una. ~2-4s extra por run.

**Documentación actualizada:** este `CHANGELOG_AI.md`, `CLAUDE_CONTEXT.md` (sección Modelos/Flujos).

**Notas para futuro:**
- `DoCmd.RunSavedImportExport` es case-sensitive y falla silenciosamente si el nombre no matchea. Verificar en Access → Datos externos → Importaciones guardadas.
- `CompactRepair` requiere BD cerrada (por eso la sesión-en-dos-partes para `after_pre_macros`).
- `post_refresh_excel_files` se valida en `create_pipeline_task` (existencia al crear). Si los tableros se mueven después, la validación de existencia se vuelve a hacer al ejecutar (en `_refresh_excel_file` → `run_engine`).
- El endpoint atajo `POST /api/tasks/pipeline/reposicion` también recibió los campos nuevos. Macros e importaciones siguen hardcoded (es justamente su valor agregado vs. el genérico); el body solo expone rutas y flags.
- UI: dependencia visual entre `f-compact-position` y `f-compact` no se refleja con disabled — si el usuario selecciona "Auto" el toggle aplica; si elige posición explícita el toggle se ignora silenciosamente. El hint lo dice pero podría confundir; futuro: deshabilitar el toggle dinámicamente cuando el select no esté en "Auto".

---

## 2026-05-24 — Recuperación masiva de UI/backend perdidos en el commit de auth

**Commits relacionados:** rebase del estado de trabajo previo a `0a7b3b4` sobre `HEAD` (sin commitear aún; cambios en working tree).

**Motivo:** el commit `0a7b3b4 feat(auth): implement JWT authentication and user management endpoints` sobreescribió `app/static/index.html` borrando ~50% de la UI existente y dejó tres endpoints sin restaurar en `app/routes.py`. La regresión pasó inadvertida durante varios commits posteriores. El usuario detectó UI revertida y faltantes en formularios.

**Qué cambió:**
- **Frontend (`app/static/index.html`):** reconstruido tomando la versión pre-auth (commit `2e79e57`) como base y reinyectando quirúrgicamente las adiciones de auth. Recuperadas: tipo de tarea Workflow Semanal con su sección completa, modal de simulación (`modal-wf-test`), auto-refresh cada 3 min con animación, columnas perdidas en Dashboard (Tipo / Última ejecución con badge / Reintentos), cabeceras ordenables, filtros de tareas extendidos, export/import JSON, split-button "Limpiar historial", live log polling, badge "Ejecutando", reintentos en form, browse-folder, estilos `input-browse-wrap`. Conservadas: sidebar widget de usuario, topbar chip, JWT Bearer, admin-only, páginas Users + Profile, modales user/reset-pw/change-pw/run-now.
- **Backend (`app/routes.py`):**
  - Restaurado `POST /api/tasks/workflow/weekly-excel-copy` + modelo `WeeklyExcelCopyTaskCreate`.
  - Restaurado `POST /api/tasks/{task_id}/test-run` + modelo `WorkflowTestRunBody` (con `force_weekday`).
  - Restaurado `GET /api/browse-folder`.
  - `/api/browse-file` arreglado: ruta absoluta de `powershell.exe`, encoding UTF-8 explícito, truco de form `TopMost` para traer el diálogo al frente, logging de errores, `get_running_loop()` (no deprecado).
  - Añadido `logger = logging.getLogger(__name__)` a nivel de módulo (faltaba; era NameError silenciado en pre-auth).

**Blast radius:** afecta a `app/static/index.html` (2207 → 2928 líneas) y `app/routes.py` (1127 → 1298 líneas). API contract: vuelven endpoints que el frontend ya esperaba pero el backend había dejado de exponer.

**Documentación actualizada:** primera entrada en este `CHANGELOG_AI.md`. Generación inicial de todo el corpus documental:
- `PROJECT_CONTEXT.md`, `CLAUDE_CONTEXT.md`, `AI_SUMMARY.md`, `PROMPTS.md`, `DEPENDENCY_MAP.md`, `ONBOARDING.md`, `KEY_FILES.md`, `CHANGELOG_AI.md`.

**Notas para futuro:**
- `index.html` siendo monolítico (2900+ líneas) es un riesgo permanente. Considerar separar a `app/static/js/` y `app/static/css/` cuando haya tiempo.
- El backup pre-auth quedó en `logs/index_pre_auth.html` y `logs/index_pre_auth_clean.html` — limpiar cuando la recuperación esté commiteada y verificada en producción.
- El frontend envía `{ skip_retry: bool }` al hacer run-now, pero el endpoint backend no lo procesa todavía. Es una feature UI a medio implementar (introducida por el commit de auth). Decidir: borrar el toggle o implementarlo end-to-end.

---

## Entradas anteriores (reconstruidas desde git log para contexto histórico)

> Estas no se redactaron en su momento; se incluyen aquí como referencia para que las búsquedas futuras encuentren el "porqué" sin necesidad de leer cada commit.

### 2026-05-23 — Migración passlib → bcrypt puro

**Commits:** `995fca5` `23e11d7` `90d3c58` `3a39f8a`

**Motivo:** problemas de empaquetado/incompatibilidad de `passlib` con la versión de bcrypt requerida; simplificación de la dependencia.

**Qué cambió:** `app/auth.py` usa `bcrypt.hashpw` / `bcrypt.checkpw` directamente. Se quitó `passlib` del `pyproject.toml`. Scripts auxiliares (`scripts/test_bcrypt.py`) confirman el contrato.

**Blast radius:** sólo `auth.py` y `pyproject.toml`. Hashes existentes en DB siguen válidos (mismo algoritmo bcrypt).

**Notas:** **NO** reintroducir `passlib`. Si Claude lo sugiere, recordarle esta entrada.

---

### 2026-05-23 — Instalación como Scheduled Task + scripts de deploy

**Commits:** `e831c94` `a8ba88e` `e941b13` `11a11a0` `88054f0` `4463283` `adbcd12` `5c62394` `6d8233c` `7c62a05` `b43f665`

**Motivo:** producción necesitaba un método robusto de arranque automático. Servicio Win32 puro descartado por incompatibilidad con COM (Excel) y OneDrive Files On-Demand.

**Qué cambió:** `install-service.ps1` registra Scheduled Task con trigger "at logon", detecta venv de Poetry, ejecuta `poetry install`, limpia instalaciones previas. `deploy.ps1` permite hot-update. Validación de superadmin al instalar/desplegar. Numerosos fixes de encoding PowerShell (Windows-1252 vs UTF-8) y de quoting de argumentos.

**Blast radius:** infraestructura/deploy, no afecta runtime del servicio.

**Notas:** archivos `.ps1` deben permanecer en ASCII puro. Histórico de bugs por encoding sirve de cautionary tale.

---

### 2026-05-23 — Fix SPA: redirect loop login

**Commits:** `d89cad8`

**Motivo:** `/` servía `login.html` que redirigía a `/login` → infinito.

**Qué cambió:** ruta catch-all en `app/main.py::serve_spa` ahora sirve `index.html` salvo cuando el path es explícitamente `login` o `login.html`. `index.html` se encarga de redirigir a `/login` cliente-side si no hay sesión.

**Blast radius:** flujo de auth en producción.

---

### 2026-05-22 — JWT auth + gestión de usuarios (commit "grande")

**Commits:** `0a7b3b4` `478b6b0`

**Motivo:** habilitar acceso multi-usuario con roles y autenticación adecuada.

**Qué cambió:** introducción de `app/auth.py`, `app/auth_routes.py`, modelo `User`, dependencias por rol, login UI, gestión de usuarios desde el dashboard. Generación/persistencia de `JWT_SECRET` y creación de superadmin via `scripts/create_superadmin.py`.

**Blast radius:** **alto**. Además del trabajo deliberado, este commit sobreescribió accidentalmente `app/static/index.html` y eliminó endpoints de workflow en `app/routes.py`. Ver entrada del 2026-05-24 para la recuperación.

**Notas:** ojo con commits "grandes" que reescriben archivos monolíticos. Considerar dividir features así en commits más chicos.

---

_Última entrada: 2026-05-24._
_Para actualizar: pedir a Claude resumen de los últimos N commits + actualización de `CLAUDE_CONTEXT.md`/`AI_SUMMARY.md` si los cambios afectan stack/convenciones._

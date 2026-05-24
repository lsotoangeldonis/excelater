# PROMPTS — Plantillas reutilizables para trabajar con Claude en Excelater

> Copia/pega el bloque que necesites. Reemplaza los `{{placeholders}}`.
> Asume que la conversación ya tiene cargado `CLAUDE_CONTEXT.md`. Si no, antepón:
> _"Lee primero `CLAUDE_CONTEXT.md` para tener el contexto del proyecto."_

---

## 1. Agregar feature — Endpoint nuevo

```
Agrega un endpoint {{METHOD}} {{/api/path}} que {{descripción funcional}}.

Requisitos:
- Body validado con pydantic (modelo nuevo en `app/routes.py`)
- Dependencia `Depends(require_{{reader|admin|superuser}})` para auth
- Si toca la DB: usar AsyncSession + `select(...)` (SQLAlchemy 2.x async)
- Si crea/modifica una tarea: llamar `add_or_replace_job(task)` después del commit
- Agregar un test happy-path en `tests/test_routes.py` (sigue el patrón de los existentes)
- Si necesita nuevo campo en DB: añadir al modelo + ALTER en `_migrate_existing_db`

Antes de codear:
1. Confirma qué modelo de Task/RunLog/etc se toca.
2. Lista las funciones de scheduler.py que vayas a usar.
3. Muéstrame el diff propuesto en `routes.py` antes de aplicarlo.
```

---

## 2. Agregar feature — Workflow nuevo

```
Crea un nuevo workflow llamado `{{nombre_snake}}` que {{descripción del flujo}}.

Sigue estos pasos exactamente:
1. Nuevo archivo `app/workflows/{{nombre_snake}}.py` con clase `{{NombreCamel}}Workflow(BaseWorkflow)`.
2. Implementa `run(self, config: dict, logger) -> EngineResult` (success/fail + duration).
3. Registra en `app/workflows/__init__.py`: `registry.register("{{nombre_snake}}", {{NombreCamel}}Workflow)`.
4. Nuevo endpoint en `app/routes.py`: `POST /api/tasks/workflow/{{nombre_kebab}}` (con `WeeklyExcelCopyTaskCreate` como referencia de modelo).
5. En `app/static/index.html`:
   - Añade radio button con `value="{{nombre_snake}}"` en la sección "Tipo de tarea" del modal-task.
   - Crea `<div id="section-{{nombre_snake}}">` con los campos del form.
   - Extiende `setTaskType()` y `saveTask()`/`openEditTask()` para manejar el nuevo tipo.
6. Test en `tests/test_workflows.py` con mock de la lógica COM.

Importante: NO toques `excel_engine.py` ni `access_engine.py` salvo que sea estrictamente necesario.
```

---

## 3. Corregir bug

```
Bug: {{síntoma observado, incluye pasos para reproducir + log esperado vs real}}.

Antes de proponer fix:
1. Lee el archivo de log relevante: `logs/task_<id>_<run>.log` o `logs/excelater.log`.
2. Identifica la función exacta donde ocurre (file:line).
3. Hipotetiza la causa raíz (no el síntoma).
4. Confirma la hipótesis con grep/lectura de código antes de codear.

Cuando propongas la solución:
- Mínima necesaria (no refactorices código aledaño).
- Sin try/except amplios que oculten el error.
- Con un test de regresión en `tests/` si el bug es testeable sin COM.

Si la causa raíz está en COM (Excel/Access), explícame qué condición la dispara y propón mitigación, no oculta-el-error.
```

---

## 4. Refactorizar

```
Refactoriza {{módulo o función}} con el objetivo de {{objetivo concreto: legibilidad / reducir duplicación / separar concerns}}.

Restricciones:
- Cero cambio de comportamiento observable (API ni efectos secundarios).
- Mantener firmas públicas; si tienes que cambiar una, justifica por qué.
- Conservar (y reutilizar) los tests existentes; si alguno se rompe, es señal de cambio de comportamiento.
- No introducir nuevas dependencias.

Entrega:
1. Lista de los cambios concretos (archivo + líneas + qué cambia).
2. Aplica los cambios.
3. Corre `poetry run pytest` y reporta resultado.
```

---

## 5. Revisar performance

```
Analiza el endpoint/función `{{nombre}}` desde perspectiva de performance.

Considera:
- Queries N+1 en SQLAlchemy (uso de `selectinload`/`joinedload`).
- Bloqueos del event loop (operaciones sync no envueltas en `to_thread`).
- COM Excel/Access: tiempo de apertura es caro, evitar reabrir.
- Lecturas de archivo: rotating log handler, OneDrive Files On-Demand.

Entrega:
1. Hot path identificado (con números si es posible: `time.perf_counter` o `cProfile`).
2. Top 3 cuellos de botella con evidencia (archivo:línea).
3. Recomendaciones priorizadas por impacto/esfuerzo, sin implementar nada todavía.

Después de mi OK, implementa solo la #1.
```

---

## 6. Crear tests

```
Escribe tests para `{{módulo o endpoint}}`.

Patrones del proyecto:
- pytest + pytest-asyncio (mode=auto, no decorador necesario en async).
- Fixtures en `tests/conftest.py` (DB en memoria, client async).
- Mockear COM (pywin32) — nunca tocar Excel real en CI.
- Para endpoints: usar `TestClient` (síncrono) o `httpx.AsyncClient` con `app`.
- Tests pequeños, un caso por función, nombres `test_<accion>_<resultado_esperado>`.

Cubre:
- Happy path principal.
- Validación de input (400 si payload inválido).
- Auth/rol (401 sin token, 403 con rol insuficiente).
- Edge cases relevantes que conozcas.

NO testees:
- Comportamiento de FastAPI/SQLAlchemy en sí.
- Casos imposibles dados los validadores pydantic.
```

---

## 7. Analizar seguridad

```
Revisa {{módulo/endpoint/feature}} desde perspectiva de seguridad.

Checklist específico de este proyecto:
- ¿Endpoint sin `Depends(require_*)`? (regresión de auth)
- ¿Inyección SQL? (usar `select()` parametrizado, no f-strings con valores user)
- ¿Path traversal en `file_path` o `resolve_path`?
- ¿Inyección en `subprocess.run([...])` para PowerShell — argumentos no shell-quoted dentro del ps_script?
- ¿XSS en `index.html`? — todo string user-controlled debe ir por `esc()` antes de `innerHTML`.
- ¿Secrets en logs? (passwords, tokens, JWT_SECRET, apikeys de CallMeBot)
- ¿CORS demasiado permisivo en producción?
- ¿bcrypt rounds suficientes? (default 12 está OK)
- ¿JWT con expire razonable? (`jwt_expire_minutes`, default 480)

Para cada hallazgo: severidad (crit/alta/media/baja), evidencia (archivo:línea), fix sugerido.
NO apliques fixes sin mi aprobación.
```

---

## 8. Migrar código (Python / framework / dep)

```
Migra {{de X versión a Y versión}} en el área {{archivo o módulo}}.

Antes:
1. Revisa el changelog/breaking changes de la versión destino.
2. Lista las llamadas / patrones afectados en el repo (con grep).
3. Estima el blast radius.

Durante:
- Cambios atómicos por commit (un patrón a la vez).
- Correr `poetry run pytest` entre commits.
- Si una API se renombró: usar alias o sed reproducible, no reemplazo a mano.

Después:
- Actualizar `PROJECT_CONTEXT.md` y `CHANGELOG_AI.md` si la migración cambia stack o convenciones.
```

---

## 9. Investigar regresión (post-deploy)

```
Algo dejó de funcionar después de `{{commit/deploy}}`.

Síntoma: {{qué falla}}
Esperado: {{qué debería pasar}}

Plan:
1. `git log --oneline {{commit_anterior_OK}}..HEAD` — qué entró.
2. `git diff {{commit_anterior_OK}} HEAD -- <archivos_sospechosos>` — qué cambió.
3. Reproducir el síntoma localmente (poetry run pytest o curl al endpoint).
4. Identificar el commit culpable con `git bisect` o lectura del diff.
5. Proponer fix mínimo; no hacer git reset --hard sin confirmar.
```

---

## 10. Auditar la deuda técnica

```
Haz un sweep de deuda técnica acumulada en `{{módulo o todo el repo}}`.

Categorías a marcar:
- TODO/FIXME/XXX hardcodeados.
- `except Exception: pass` (catches silenciosos).
- Imports sin usar.
- Funciones >100 líneas.
- Strings duplicados (rutas, mensajes, constantes mágicas).
- Endpoints sin tests.
- Modelos con columnas no usadas.

Entrega un informe priorizado: top 10 ítems por impacto, con archivo:línea, sin tocar código.
```

---

## 11. Generar / actualizar documentación

```
Actualiza `{{PROJECT_CONTEXT.md | CLAUDE_CONTEXT.md | DEPENDENCY_MAP.md}}` basándote en los cambios desde {{fecha o commit}}.

Pasos:
1. `git log --oneline {{since}}..HEAD` — qué cambió.
2. Para cada cambio relevante (no fixes triviales): identifica si toca arquitectura, dependencias o convenciones documentadas.
3. Actualiza solo las secciones afectadas; conserva el resto.
4. Si hay un cambio que merece entrada propia: añadirla en `CHANGELOG_AI.md`.

No reescribas todo el archivo; entrega un diff de las secciones tocadas.
```

---

## 12. Operación en producción (Scheduled Task)

```
{{Verificar estado / reiniciar / actualizar / desinstalar}} el servicio Excelater en Windows.

Comandos PowerShell (admin):
- Estado:        `Get-ScheduledTask -TaskName Excelater | Get-ScheduledTaskInfo`
- Detener:       `Stop-ScheduledTask -TaskName Excelater`
- Iniciar:       `Start-ScheduledTask -TaskName Excelater`
- Actualizar:    `.\deploy.ps1` (hot-update sin recrear)
- Reinstalar:    `.\install-service.ps1` (es idempotente; force re-register si detecta cmd.exe)
- Desinstalar:   `Unregister-ScheduledTask -TaskName Excelater -Confirm:$false`
- Logs vivos:    `Get-Content -Tail 100 -Wait logs\excelater.log`

NUNCA toques el servicio en producción sin confirmar conmigo primero.
```

---

## Convenciones para invocar prompts

- **Sé específico** en `{{}}`. Si dices "fix el bug del modal", Claude tiene que adivinar; mejor: "fix bug: al editar workflow, no se cargan los pivot_guards".
- **Una sola tarea por prompt**. Para tareas grandes, usa `/plan` o pide un plan antes.
- **Adjunta `CLAUDE_CONTEXT.md`** o reférencialo al inicio de sesiones nuevas.
- **Pide confirmación** explícita antes de operaciones destructivas: deploy, reset de DB, force push, install/uninstall del servicio.

---

_Última actualización: 2026-05-24_

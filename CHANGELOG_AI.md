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

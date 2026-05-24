# CLAUDE.md — Excelater

Este archivo lo carga Claude Code automáticamente al iniciar cualquier sesión en este repo. Mantenlo corto.

## Contexto del proyecto

@CLAUDE_CONTEXT.md

## Documentación complementaria (disponible bajo demanda)

Estos archivos NO se cargan automáticamente; léelos cuando la tarea lo amerite:

- **PROJECT_CONTEXT.md** — referencia exhaustiva: arquitectura, riesgos, comandos completos.
- **AI_SUMMARY.md** — versión ultra densa (~220 líneas) para pegar a otras IAs sin contexto.
- **PROMPTS.md** — 12 plantillas reutilizables (feature, bug, refactor, perf, tests, security, migrar, ops).
- **DEPENDENCY_MAP.md** — grafo de dependencias internas + archivos peligrosos de modificar + contratos UI↔backend.
- **KEY_FILES.md** — top 20 archivos del repo explicados en tiers.
- **ONBOARDING.md** — guía de setup en <30 min (orientada a humanos).
- **CHANGELOG_AI.md** — bitácora de cambios mayores (formato + entradas históricas).

## Reglas duras (no romper nunca)

1. **NO usar `passlib`** — el proyecto migró a `bcrypt` puro (commits `995fca5` → `3a39f8a`). Si lo sugieres, primero lee `app/auth.py`.
2. **NO usar `asyncio.get_event_loop()`** — está deprecado en Python 3.12+. Usar `asyncio.get_running_loop()`.
3. **NO hacer `DELETE` directo en `tasks`** — siempre soft delete vía `deleted_at IS NULL` / `IS NOT NULL`.
4. **NO usar caracteres no-ASCII en `.ps1`** — PowerShell 5 los lee como Windows-1252 y rompe. Ya hubo fixes (`88054f0`, `11a11a0`). Nada de em-dashes (`—`) ni `?.Source`.
5. **NO modificar `app/static/index.html` sin grep previo** — es monolítico (2900+ líneas), los IDs cruzan secciones, y ya hubo una catástrofe (commit `0a7b3b4` borró ~50% de la UI).
6. **Cambios de schema** → añadir el `ALTER TABLE` en `_migrate_existing_db` de `app/database.py` además del modelo. No hay Alembic.
7. **Nuevo workflow** → tocar 4 sitios: clase en `app/workflows/<x>.py` (hereda `BaseWorkflow`) + `registry.register(...)` en `workflows/__init__.py` + endpoint en `routes.py` + sección + radio button en `index.html`.
8. **Antes de operaciones destructivas** (reset DB, force push, install/uninstall Scheduled Task, drop columns) → pedir confirmación al usuario.

## Hook activo en este repo

`.claude/settings.json` define un hook `PostToolUse` para `Bash(git commit*)` que me recuerda actualizar `CHANGELOG_AI.md` (y `CLAUDE_CONTEXT.md` / `AI_SUMMARY.md` si aplica) después de cada commit significativo. No lo ignores.

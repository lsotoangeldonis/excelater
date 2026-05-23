"""app/workflows/__init__.py — Registro central de workflows personalizados

Para añadir un nuevo workflow:
  1. Crear app/workflows/mi_workflow.py con una clase que herede BaseWorkflow.
  2. Importarla aquí y registrarla con registry.register("nombre", MiWorkflow).
  3. Añadir el endpoint correspondiente en app/routes.py.
"""
from __future__ import annotations

from app.workflows.base import BaseWorkflow
from app.workflows.weekly_excel_copy import WeeklyExcelCopyWorkflow


class WorkflowRegistry:
    """Mapa nombre → clase de workflow. Instancia única (singleton)."""

    def __init__(self) -> None:
        self._registry: dict[str, type[BaseWorkflow]] = {}

    def register(self, name: str, cls: type[BaseWorkflow]) -> None:
        self._registry[name] = cls

    def get(self, name: str) -> type[BaseWorkflow] | None:
        return self._registry.get(name)

    def available(self) -> list[str]:
        return list(self._registry.keys())


registry = WorkflowRegistry()

# ── Registro de workflows disponibles ────────────────────────────────────────
registry.register("weekly_excel_copy", WeeklyExcelCopyWorkflow)

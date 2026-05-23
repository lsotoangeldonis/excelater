"""app/workflows/base.py — Clase base para flujos de trabajo personalizados"""
from __future__ import annotations

import logging
from abc import ABC, abstractmethod

from app.excel_engine import EngineResult


class BaseWorkflow(ABC):
    """
    Clase base que deben heredar todos los workflows personalizados.

    Implementar el método `run` con la lógica del flujo.
    Siempre retornar un EngineResult; nunca lanzar excepciones sin capturar.
    """

    @abstractmethod
    def run(self, config: dict, logger: logging.Logger) -> EngineResult:
        """
        Ejecuta el workflow.

        Args:
            config: Diccionario con la configuración almacenada en pipeline_config
                    (sin el campo workflow_type).
            logger: Logger de la tarea para trazabilidad.

        Returns:
            EngineResult con success=True/False, error_msg y duration_s.
        """
        ...

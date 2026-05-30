"""app/com_registry.py — Registro persistente de procesos COM que abrimos via DispatchEx.

Mantiene un JSON en `logs/com_pids.json` con los PIDs de Excel/Access que arrancamos,
asociados al `run_id` que los lanzó. Permite:

1. **Identificar con certeza** si un proceso bloqueante es nuestro (PID en el registro).
2. **Heurística de respaldo** vía visibilidad de ventana cuando no hay registro
   (procesos anteriores al cambio o casos borde).
3. **Terminar forzosamente** procesos huérfanos al cancelar un run.

El registro es best-effort: nunca rompe el flujo principal si falla I/O.
"""
from __future__ import annotations

import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path


_REGISTRY_PATH = Path("logs") / "com_pids.json"


# ── I/O atómico ───────────────────────────────────────────────────────────────

def _read() -> list[dict]:
    try:
        return json.loads(_REGISTRY_PATH.read_text("utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return []
    except Exception:
        return []


def _write(entries: list[dict]) -> None:
    _REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = _REGISTRY_PATH.with_name(_REGISTRY_PATH.name + ".tmp")
    tmp.write_text(json.dumps(entries, indent=2, default=str), "utf-8")
    os.replace(str(tmp), str(_REGISTRY_PATH))


# ── Estado del SO ─────────────────────────────────────────────────────────────

def pid_alive(pid: int | None) -> bool:
    """True si el PID corresponde a un proceso vivo en Windows."""
    if not pid or sys.platform != "win32":
        return False
    try:
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        STILL_ACTIVE = 259
        h = ctypes.windll.kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION, False, pid,
        )
        if not h:
            return False
        try:
            code = ctypes.c_ulong()
            ok = ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
            return bool(ok) and code.value == STILL_ACTIVE
        finally:
            ctypes.windll.kernel32.CloseHandle(h)
    except Exception:
        return False


def has_visible_window(pid: int | None) -> bool:
    """True si el proceso tiene alguna ventana top-level visible.
    Diferencia Office abierto por un usuario (visible) vs lanzado por nosotros
    con DispatchEx + Visible=False (headless)."""
    if not pid or sys.platform != "win32":
        return False
    try:
        import win32gui
        import win32process
    except Exception:
        return False
    found = [False]

    def cb(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd):
                _, p = win32process.GetWindowThreadProcessId(hwnd)
                if p == pid:
                    found[0] = True
                    return False
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return found[0]


def kill_pid(pid: int | None) -> bool:
    """TerminateProcess sobre `pid`. Devuelve True si tuvo éxito."""
    if not pid or sys.platform != "win32":
        return False
    try:
        import ctypes
        PROCESS_TERMINATE = 0x0001
        h = ctypes.windll.kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
        if not h:
            return False
        try:
            return bool(ctypes.windll.kernel32.TerminateProcess(h, 1))
        finally:
            ctypes.windll.kernel32.CloseHandle(h)
    except Exception:
        return False


# ── API del registro ──────────────────────────────────────────────────────────

def register(pid: int | None, name: str, path: str, run_id: int | None) -> None:
    """Añade (o reemplaza) una entrada en el registro. Best-effort."""
    if not pid:
        return
    try:
        entries = [e for e in _read() if e.get("pid") != pid]
        entries.append({
            "pid": pid,
            "name": name,
            "path": path,
            "run_id": run_id,
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "owner_pid": os.getpid(),
        })
        _write(entries)
    except Exception:
        pass


def unregister(pid: int | None) -> None:
    """Elimina la entrada correspondiente a `pid`. Best-effort."""
    if not pid:
        return
    try:
        entries = [e for e in _read() if e.get("pid") != pid]
        _write(entries)
    except Exception:
        pass


def find_by_pid(pid: int | None) -> dict | None:
    if not pid:
        return None
    try:
        for e in _read():
            if e.get("pid") == pid:
                return e
    except Exception:
        pass
    return None


def find_by_run(run_id: int | None) -> list[dict]:
    if run_id is None:
        return []
    try:
        return [e for e in _read() if e.get("run_id") == run_id]
    except Exception:
        return []


def prune_dead() -> int:
    """Quita del registro entradas cuyos PIDs ya no existen. Devuelve cuántas removió."""
    try:
        entries = _read()
        alive = [e for e in entries if pid_alive(e.get("pid"))]
        removed = len(entries) - len(alive)
        if removed:
            _write(alive)
        return removed
    except Exception:
        return 0


# ── Clasificación de procesos bloqueantes ────────────────────────────────────

_OFFICE_HINTS = ("excel", "access", "word", "outlook", "powerpoint")


def _is_office_name(name: str) -> bool:
    if not name:
        return False
    n = name.lower()
    return any(s in n for s in _OFFICE_HINTS)


def classify_locker(pid: int | None, name: str = "") -> dict:
    """Clasifica un PID bloqueante. Devuelve dict con:
      - category: 'ours' | 'maybe_ours' | 'external' | 'unknown'
      - source: cómo se llegó a la conclusión
      - entry: dict del registro si aplica
    """
    entry = find_by_pid(pid)
    if entry:
        return {"category": "ours", "source": "registry", "entry": entry}

    # Sin registro: heurística A (visibilidad de ventana) sólo para apps Office.
    if pid and _is_office_name(name) and pid_alive(pid):
        if not has_visible_window(pid):
            return {"category": "maybe_ours", "source": "no_visible_window", "entry": None}
        return {"category": "external", "source": "visible_window", "entry": None}

    return {"category": "unknown", "source": None, "entry": None}


# ── Terminación forzosa por run ──────────────────────────────────────────────

def kill_run_processes(run_id: int | None, log: logging.Logger | None = None) -> list[int]:
    """Mata todos los PIDs registrados para `run_id`. Devuelve los PIDs que se mataron."""
    killed: list[int] = []
    for entry in find_by_run(run_id):
        pid = entry.get("pid")
        name = entry.get("name", "?")
        if pid and pid_alive(pid):
            if kill_pid(pid):
                killed.append(pid)
                if log:
                    log.warning(f"Proceso {name} PID={pid} terminado forzosamente.")
            elif log:
                log.warning(f"No se pudo terminar {name} PID={pid}.")
        unregister(pid)
    return killed

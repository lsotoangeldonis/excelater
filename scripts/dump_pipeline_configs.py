"""Dump pipeline_config de todas las tareas tipo 'pipeline' no borradas.

Uso:
    poetry run python scripts/dump_pipeline_configs.py
"""
import json
import sqlite3
from pathlib import Path

DB = Path(__file__).resolve().parent.parent / "scheduler.db"

con = sqlite3.connect(str(DB))
cur = con.execute(
    "SELECT id, name, task_type, pipeline_config "
    "FROM tasks "
    "WHERE task_type = 'pipeline' AND deleted_at IS NULL "
    "ORDER BY name"
)

rows = cur.fetchall()
if not rows:
    print("No hay tareas pipeline activas.")
else:
    for tid, name, ttype, cfg_json in rows:
        print(f"=== {name}  (id={tid}, type={ttype}) ===")
        try:
            cfg = json.loads(cfg_json or "{}")
        except Exception as e:
            print(f"  <pipeline_config no parseable: {e}>")
            print(f"  raw: {cfg_json!r}")
        else:
            print(json.dumps(cfg, indent=2, ensure_ascii=False))
        print()

con.close()

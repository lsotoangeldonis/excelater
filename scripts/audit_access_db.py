"""Audita un .accdb exportando macros, queries, VBA y tablas vinculadas a texto.

Uso:
    poetry run python scripts/audit_access_db.py "C:\\ruta\\a\\DWH OYM.accdb"

Genera ./access_audit/<nombre>/:
    tables.txt         → tablas locales vs linked (con connect string)
    queries.txt        → SQL de todas las queries guardadas (marca action queries)
    macros/*.txt       → cada macro exportada
    modules/*.bas      → cada modulo VBA
    imports.txt        → saved import/export specs
    summary.txt        → resumen ejecutivo (lo que conviene pegarle a Claude)

Solo lee. No corre macros, no guarda, no compacta. Bloquea prompts de macros
(AutomationSecurity=3) y no dispara warnings.
"""
from __future__ import annotations

import sys
from pathlib import Path

try:
    import pythoncom
    import win32com.client
except ImportError:
    print("Falta pywin32. Ejecuta: poetry add pywin32")
    sys.exit(1)


# Constantes de Access COM
ACC_QUERY = 1
ACC_FORM = 2
ACC_REPORT = 3
ACC_MACRO = 4
ACC_MODULE = 5

# QueryDef.Type
QUERY_TYPES = {
    0: "SELECT",
    16: "CROSSTAB",
    32: "DELETE",
    48: "UPDATE",
    64: "APPEND",
    80: "MAKE_TABLE",
    96: "DDL",
    112: "PASSTHROUGH",
    128: "UNION",
}
ACTION_QUERY_TYPES = {32, 48, 64, 80, 96, 112}  # las que escriben


def safe_write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", errors="replace")


def dump_tables(acc, out_dir: Path) -> tuple[int, int, list[str]]:
    """Lista tablas. Devuelve (n_locales, n_linked, alertas_sql_server)."""
    db = acc.CurrentDb()
    n_local = 0
    n_linked = 0
    sql_server_hits: list[str] = []
    lines = ["# Tablas del .accdb\n"]
    lines.append("Columnas: NOMBRE | TIPO | CONNECT_STRING\n")
    lines.append("=" * 80 + "\n\n")

    for i in range(db.TableDefs.Count):
        tdef = db.TableDefs(i)
        name = tdef.Name
        if name.startswith("MSys") or name.startswith("~"):
            continue
        connect = tdef.Connect or ""
        if connect:
            n_linked += 1
            kind = "LINKED"
            # Detectar SQL Server explícitamente
            upper = connect.upper()
            if any(tok in upper for tok in ("SQL SERVER", "SQLSERVER", "SQLNCLI", "ODBC DRIVER 17", "ODBC DRIVER 18", "MSODBCSQL")):
                sql_server_hits.append(name)
                kind = "LINKED-SQLSERVER"
        else:
            n_local += 1
            kind = "LOCAL"
        lines.append(f"{name}\t{kind}\t{connect}\n")

    lines.append("\n" + "=" * 80 + "\n")
    lines.append(f"Total locales: {n_local}\n")
    lines.append(f"Total linked: {n_linked}\n")
    lines.append(f"Linked a SQL Server: {len(sql_server_hits)}\n")
    if sql_server_hits:
        lines.append("\n*** TABLAS VINCULADAS A SQL SERVER ***\n")
        for t in sql_server_hits:
            lines.append(f"  - {t}\n")

    safe_write(out_dir / "tables.txt", "".join(lines))
    return n_local, n_linked, sql_server_hits


def dump_queries(acc, out_dir: Path) -> tuple[int, list[tuple[str, str]]]:
    """Exporta SQL de todas las queries. Devuelve (total, [(name, type)] de las action)."""
    db = acc.CurrentDb()
    actions: list[tuple[str, str]] = []
    lines = ["# Queries guardadas\n"]
    lines.append("=" * 80 + "\n\n")
    total = 0

    for i in range(db.QueryDefs.Count):
        q = db.QueryDefs(i)
        name = q.Name
        if name.startswith("~"):
            continue
        total += 1
        qtype = QUERY_TYPES.get(q.Type, f"UNKNOWN({q.Type})")
        is_action = q.Type in ACTION_QUERY_TYPES
        if is_action:
            actions.append((name, qtype))
        try:
            sql = q.SQL
        except Exception as e:
            sql = f"<error leyendo SQL: {e}>"
        marker = " *** ACTION ***" if is_action else ""
        lines.append(f"--- [{qtype}]{marker} {name}\n")
        lines.append(sql)
        if not sql.endswith("\n"):
            lines.append("\n")
        lines.append("\n")

    safe_write(out_dir / "queries.txt", "".join(lines))
    return total, actions


def dump_named_objects(acc, out_dir: Path, obj_type: int, subdir: str, ext: str) -> int:
    """Exporta macros (obj_type=4) o modulos VBA (obj_type=5) via SaveAsText."""
    target = out_dir / subdir
    target.mkdir(parents=True, exist_ok=True)
    # AllMacros / AllModules
    coll = acc.CurrentProject.AllMacros if obj_type == ACC_MACRO else acc.CurrentProject.AllModules
    count = 0
    for i in range(coll.Count):
        obj = coll.Item(i)
        name = obj.Name
        if name.startswith("~"):
            continue
        # Sanitizar para FS
        safe_name = "".join(c if c.isalnum() or c in " -_.()" else "_" for c in name)
        out = target / f"{safe_name}.{ext}"
        try:
            acc.SaveAsText(obj_type, name, str(out))
            count += 1
        except Exception as e:
            (target / f"{safe_name}.ERROR.txt").write_text(
                f"No se pudo exportar '{name}': {e}", encoding="utf-8"
            )
    return count


def dump_import_specs(acc, out_dir: Path) -> int:
    """Exporta saved import/export specs (los que dispara RunSavedImportExport)."""
    lines = ["# Saved Import/Export specifications\n"]
    lines.append("=" * 80 + "\n\n")
    count = 0
    try:
        specs = acc.CurrentProject.ImportExportSpecifications
        for i in range(specs.Count):
            spec = specs.Item(i)
            count += 1
            lines.append(f"--- {spec.Name}\n")
            try:
                lines.append(spec.XML)
            except Exception as e:
                lines.append(f"<error leyendo XML: {e}>")
            lines.append("\n\n")
    except Exception as e:
        lines.append(f"<error accediendo a ImportExportSpecifications: {e}>\n")
    safe_write(out_dir / "imports.txt", "".join(lines))
    return count


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    db_path = Path(sys.argv[1])
    if not db_path.exists():
        print(f"No existe: {db_path}")
        return 1

    out_root = Path("access_audit") / db_path.stem
    out_root.mkdir(parents=True, exist_ok=True)
    print(f"[+] Auditando: {db_path}")
    print(f"[+] Salida:    {out_root.resolve()}")

    pythoncom.CoInitialize()
    acc = None
    try:
        acc = win32com.client.DispatchEx("Access.Application")
        acc.Visible = False
        try:
            acc.AutomationSecurity = 3  # bloquea macros (no corre AutoExec)
        except Exception:
            pass
        acc.OpenCurrentDatabase(str(db_path))
        try:
            acc.DoCmd.SetWarnings(False)
        except Exception:
            pass

        print("[+] Listando tablas...")
        n_local, n_linked, sql_hits = dump_tables(acc, out_root)
        print(f"    {n_local} locales, {n_linked} linked, {len(sql_hits)} a SQL Server")

        print("[+] Exportando queries...")
        n_q, actions = dump_queries(acc, out_root)
        print(f"    {n_q} queries totales, {len(actions)} action queries")

        print("[+] Exportando macros...")
        n_m = dump_named_objects(acc, out_root, ACC_MACRO, "macros", "txt")
        print(f"    {n_m} macros")

        print("[+] Exportando modulos VBA...")
        n_vba = dump_named_objects(acc, out_root, ACC_MODULE, "modules", "bas")
        print(f"    {n_vba} modulos")

        print("[+] Exportando saved imports/exports...")
        n_specs = dump_import_specs(acc, out_root)
        print(f"    {n_specs} specs")

        # Resumen ejecutivo
        summary = [
            "# Resumen ejecutivo de auditoria\n",
            f"DB: {db_path}\n\n",
            f"- Tablas locales:     {n_local}\n",
            f"- Tablas linked:      {n_linked}\n",
            f"- Linked a SQL Server: {len(sql_hits)}\n",
        ]
        if sql_hits:
            summary.append("\n## TABLAS VINCULADAS A SQL SERVER\n")
            for t in sql_hits:
                summary.append(f"  - {t}\n")
            summary.append(
                "\n*** Cualquier action query (DELETE/UPDATE/INSERT) contra una\n"
                "    de estas tablas se traduce a una escritura real en SQL Server. ***\n"
            )
        summary.append(f"\n- Queries totales:    {n_q}\n")
        summary.append(f"- Action queries:     {len(actions)}\n")
        if actions:
            summary.append("\n## ACTION QUERIES (potenciales escrituras)\n")
            for name, qt in actions:
                summary.append(f"  - [{qt}] {name}\n")
        summary.append(f"\n- Macros:             {n_m}\n")
        summary.append(f"- Modulos VBA:        {n_vba}\n")
        summary.append(f"- Saved imp/exp specs: {n_specs}\n")
        safe_write(out_root / "summary.txt", "".join(summary))

        print(f"\n[OK] Auditoria completa en: {out_root.resolve()}")
        print("     Empieza por summary.txt y tables.txt")
        return 0
    finally:
        if acc:
            try:
                acc.CloseCurrentDatabase()
            except Exception:
                pass
            try:
                acc.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())

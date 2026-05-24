#!/usr/bin/env python
"""scripts/test_bcrypt.py — Verifica que bcrypt funciona correctamente.

Uso:
    poetry run python scripts/test_bcrypt.py
"""
from __future__ import annotations

import getpass
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))


def main() -> None:
    from app.auth import hash_password, verify_password

    print()
    print("══════════════════════════════════════")
    print("  Excelater — Test de hash bcrypt      ")
    print("══════════════════════════════════════")
    print()

    try:
        pw = getpass.getpass("  Contraseña de prueba: ")
    except (KeyboardInterrupt, EOFError):
        print()
        sys.exit(0)

    h = hash_password(pw)
    ok = verify_password(pw, h)

    print(f"  Hash generado : {h[:25]}...")
    print(f"  Verificación  : {'OK ✓' if ok else 'FALLO ✗'}")
    print()

    if not ok:
        print("  ERROR: hash_password/verify_password no coinciden.")
        sys.exit(1)


if __name__ == "__main__":
    main()

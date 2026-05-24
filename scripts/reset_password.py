#!/usr/bin/env python
"""scripts/reset_password.py — Resetea la contraseña de un usuario existente.

Uso:
    poetry run python scripts/reset_password.py
    poetry run python scripts/reset_password.py --username lsoto
"""
from __future__ import annotations

import asyncio
import getpass
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))


async def main() -> None:
    from app.auth import hash_password
    from app.database import AsyncSessionLocal, User, init_db
    from sqlalchemy import select

    print()
    print("══════════════════════════════════════════")
    print("  Excelater — Reseteo de contraseña       ")
    print("══════════════════════════════════════════")
    print()

    await init_db()

    # Determinar username (argumento o interactivo)
    if len(sys.argv) >= 3 and sys.argv[1] == "--username":
        username = sys.argv[2].strip().lower()
    else:
        try:
            username = input("  Username: ").strip().lower()
        except (KeyboardInterrupt, EOFError):
            print()
            sys.exit(0)

    async with AsyncSessionLocal() as db:
        result = await db.execute(select(User).where(User.username == username))
        user = result.scalar_one_or_none()

    if not user:
        print(f"\n  ERROR: No existe el usuario '{username}'.\n")
        sys.exit(1)

    print(f"  Usuario : {user.username}  ({user.role.value})")
    print(f"  Activo  : {'sí' if user.is_active else 'no'}")
    print()

    # Nueva contraseña
    while True:
        try:
            pw1 = getpass.getpass("  Nueva contraseña (mín. 8 caracteres): ")
            if len(pw1) < 8:
                print("  La contraseña debe tener al menos 8 caracteres.")
                continue
            pw2 = getpass.getpass("  Confirmar contraseña: ")
        except (KeyboardInterrupt, EOFError):
            print()
            sys.exit(0)

        if pw1 != pw2:
            print("  Las contraseñas no coinciden. Intenta de nuevo.")
            continue
        break

    async with AsyncSessionLocal() as db:
        result = await db.execute(select(User).where(User.username == username))
        user = result.scalar_one_or_none()
        user.hashed_pw = hash_password(pw1)
        await db.commit()

    print(f"\n  ✓ Contraseña de '{username}' actualizada correctamente.\n")


if __name__ == "__main__":
    asyncio.run(main())

#!/usr/bin/env python
"""scripts/create_superadmin.py — Script interactivo para crear el primer superusuario.

Uso:
    poetry run python scripts/create_superadmin.py
    python scripts/create_superadmin.py  (dentro del venv)

Puede ejecutarse varias veces; solo fallará si el username ya existe.
"""
from __future__ import annotations

import asyncio
import getpass
import re
import secrets
import sys
from pathlib import Path

# Añadir raíz del proyecto al path para importar app.*
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))


# ── helpers UI ────────────────────────────────────────────────────────────────
def c(text: str, color: str) -> str:
    codes = {"cyan": "\033[96m", "green": "\033[92m", "red": "\033[91m",
             "yellow": "\033[93m", "bold": "\033[1m", "reset": "\033[0m"}
    return f"{codes.get(color,'')}{text}{codes['reset']}"


def prompt(label: str, default: str = "", secret: bool = False) -> str:
    suffix = f" [{default}]" if default else ""
    while True:
        try:
            if secret:
                val = getpass.getpass(f"  {label}{suffix}: ").strip()
            else:
                val = input(f"  {label}{suffix}: ").strip()
        except (KeyboardInterrupt, EOFError):
            print()
            sys.exit(0)
        if val == "" and default:
            return default
        if val:
            return val
        print(c("  Campo requerido.", "red"))


# ── validaciones ──────────────────────────────────────────────────────────────
def valid_username(v: str) -> bool:
    return bool(re.match(r"^[a-z0-9_.-]{3,50}$", v))


def valid_password(v: str) -> bool:
    return len(v) >= 8


# ── lógica principal ──────────────────────────────────────────────────────────
async def main() -> None:
    print()
    print(c("══════════════════════════════════════════", "cyan"))
    print(c("  Excelater — Creación de Superusuario    ", "bold"))
    print(c("══════════════════════════════════════════", "cyan"))
    print()

    # Importar aquí para asegurar que el path está configurado
    from app.config import settings
    from app.database import init_db, AsyncSessionLocal, User, UserRole
    from app.auth import hash_password
    from sqlalchemy import select

    # Inicializar la base de datos (crea tablas si no existen)
    await init_db()

    # ── JWT_SECRET ────────────────────────────────────────────────────────────
    if not settings.jwt_secret:
        generated = secrets.token_hex(32)
        print(c("  ADVERTENCIA: JWT_SECRET no está configurado en .env", "yellow"))
        print(f"  Se ha generado un secret aleatorio: {c(generated, 'cyan')}")
        print()
        print("  Agrégalo a tu .env para que los tokens sean persistentes:")
        print(f"    JWT_SECRET={generated}")
        print()
        input("  Presiona Enter para continuar de todas formas... ")

    # ── Datos del superusuario ────────────────────────────────────────────────
    print()
    print(c("  Datos del superusuario:", "bold"))
    print()

    # Username
    while True:
        username = prompt("Username (a-z, 0-9, _ . -)").lower()
        if not valid_username(username):
            print(c("  Username inválido. Usa 3-50 caracteres: letras, números, _ . -", "red"))
            continue
        # Verificar existencia
        async with AsyncSessionLocal() as db:
            result = await db.execute(select(User).where(User.username == username))
            existing = result.scalar_one_or_none()
        if existing:
            print(c(f"  El usuario '{username}' ya existe.", "red"))
            if existing.role.value == "superuser":
                print(c("  Ya existe un superusuario con este nombre.", "yellow"))
            ans = input("  ¿Usar otro nombre? [S/n]: ").strip().lower()
            if ans == "n":
                sys.exit(0)
            continue
        break

    full_name = prompt("Nombre completo", default=username.capitalize())
    email_raw = prompt("Email (opcional, Enter para omitir)", default="")
    email = email_raw if email_raw and "@" in email_raw else None

    # Contraseña
    while True:
        pw1 = prompt("Contraseña (mín. 8 caracteres)", secret=True)
        if not valid_password(pw1):
            print(c("  La contraseña debe tener al menos 8 caracteres.", "red"))
            continue
        pw2 = prompt("Confirmar contraseña", secret=True)
        if pw1 != pw2:
            print(c("  Las contraseñas no coinciden.", "red"))
            continue
        break

    # ── Confirmación ──────────────────────────────────────────────────────────
    print()
    print(c("  Resumen:", "bold"))
    print(f"    Username  : {c(username, 'cyan')}")
    print(f"    Nombre    : {full_name}")
    print(f"    Email     : {email or '(sin email)'}")
    print(f"    Rol       : {c('superuser', 'yellow')}")
    print()
    ans = input("  ¿Crear usuario? [S/n]: ").strip().lower()
    if ans == "n":
        print(c("  Cancelado.", "yellow"))
        sys.exit(0)

    # ── Persistir ─────────────────────────────────────────────────────────────
    user = User(
        username=username,
        full_name=full_name,
        email=email,
        hashed_pw=hash_password(pw1),
        role=UserRole.SUPERUSER,
        is_active=True,
    )
    async with AsyncSessionLocal() as db:
        db.add(user)
        await db.commit()
        await db.refresh(user)

    print()
    print(c("  ✓ Superusuario creado correctamente.", "green"))
    print(f"    ID       : {user.id}")
    print(f"    Username : {user.username}")
    print(f"    Rol      : {user.role.value}")
    print()
    print(c("  Ahora puedes iniciar sesión en el dashboard.", "cyan"))
    print()


if __name__ == "__main__":
    asyncio.run(main())

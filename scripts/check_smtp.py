#!/usr/bin/env python
"""scripts/check_smtp.py — Diagnóstico del envío SMTP de Excelater.

Carga la misma configuración (`.env` -> `app.config.settings`) que usa el
servicio en producción, verifica que el password no se haya corrompido por
interpolación, prueba conectividad TCP y luego intenta un envío real con
aiosmtplib (la misma librería que usa `app/notifications.py`).

Uso:
    poetry run python scripts/check_smtp.py destinatario@ejemplo.com
    poetry run python scripts/check_smtp.py destinatario@ejemplo.com --verbose
"""
from __future__ import annotations

import argparse
import asyncio
import logging
import socket
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))


def banner(title: str) -> None:
    print()
    print("=" * 70)
    print(f"  {title}")
    print("=" * 70)


def ok(msg: str) -> None:
    print(f"[ OK   ] {msg}")


def warn(msg: str) -> None:
    print(f"[ WARN ] {msg}")


def fail(msg: str) -> None:
    print(f"[ FAIL ] {msg}")


def info(msg: str) -> None:
    print(f"[ INFO ] {msg}")


def mask(value: str) -> str:
    if not value:
        return "(vacio)"
    if len(value) <= 4:
        return "*" * len(value) + f"  len={len(value)}"
    return f"{value[0]}***{value[-1]}  len={len(value)}"


def read_raw_password() -> str | None:
    """Lee el valor crudo de SMTP_PASSWORD directamente del .env (sin interpolar)."""
    env_path = ROOT / ".env"
    if not env_path.exists():
        return None
    try:
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("SMTP_PASSWORD"):
                _, _, val = line.partition("=")
                return val
    except Exception:
        pass
    return None


async def main(recipient: str, verbose: bool) -> int:
    if verbose:
        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s | %(name)s | %(levelname)s | %(message)s",
        )

    banner("1) Configuración cargada por Excelater")
    try:
        from app.config import settings
    except Exception as exc:
        fail(f"No se pudo importar app.config: {exc}")
        return 1

    print(f"  SMTP_HOST     = {settings.smtp_host!r}")
    print(f"  SMTP_PORT     = {settings.smtp_port}")
    print(f"  SMTP_USER     = {settings.smtp_user!r}")
    print(f"  SMTP_PASSWORD = {mask(settings.smtp_password)}")
    print(f"  SMTP_FROM     = {settings.smtp_from!r}")
    print(f"  SMTP_TLS      = {settings.smtp_tls}")

    if not settings.smtp_host:
        fail("SMTP_HOST vacío. Excelater no enviará correos.")
        return 1

    raw = read_raw_password()
    if raw is not None:
        # Quita comillas envolventes si las hay
        if (raw.startswith("'") and raw.endswith("'")) or (raw.startswith('"') and raw.endswith('"')):
            raw_clean = raw[1:-1]
        else:
            raw_clean = raw
        if raw_clean and settings.smtp_password != raw_clean:
            warn(
                "El password cargado NO coincide con el del .env. "
                f"En .env hay {len(raw_clean)} chars, en memoria {len(settings.smtp_password)}."
            )
            warn(
                "Probablemente la interpolación de python-dotenv "
                "expandió un '$VAR' inexistente. Envuelve el valor en comillas simples:"
            )
            warn("  SMTP_PASSWORD='valor con $ literal $'")
        elif "$" in raw_clean and not (raw.startswith("'") or raw.startswith('"')):
            warn(
                "El password contiene '$' y no está entre comillas en .env. "
                "Si en algún momento ves errores de auth, prueba envolverlo en comillas simples."
            )

    banner("2) Conectividad TCP al servidor SMTP")
    try:
        with socket.create_connection((settings.smtp_host, settings.smtp_port), timeout=10):
            ok(f"Conectado a {settings.smtp_host}:{settings.smtp_port}")
    except Exception as exc:
        fail(f"No se pudo conectar a {settings.smtp_host}:{settings.smtp_port} -> {exc}")
        info("Causas: firewall corporativo, proxy, DNS, o el servidor SMTP está caído.")
        return 1

    banner("3) Envío real via aiosmtplib (igual que Excelater)")
    try:
        import aiosmtplib
    except ImportError:
        fail("aiosmtplib no instalado. Ejecuta: poetry install")
        return 1

    sender_addr = settings.smtp_from or settings.smtp_user
    info(f"De:    {sender_addr}")
    info(f"Para:  {recipient}")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[Excelater] Test SMTP {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    msg["From"] = sender_addr
    msg["To"] = recipient
    msg.attach(MIMEText(
        "<html><body style='font-family:sans-serif'>"
        "<h3>Test SMTP Excelater</h3>"
        "<p>Si recibes este correo, el envío SMTP funciona desde el servidor donde corre Excelater.</p>"
        f"<p style='color:#888;font-size:12px'>Enviado a las {datetime.now().isoformat()}</p>"
        "</body></html>",
        "html", "utf-8"
    ))

    try:
        await aiosmtplib.send(
            msg,
            hostname=settings.smtp_host,
            port=settings.smtp_port,
            username=settings.smtp_user or None,
            password=settings.smtp_password or None,
            start_tls=settings.smtp_tls,
            timeout=30,
        )
        ok(f"Correo enviado a {recipient}. Verifica la bandeja (puede tardar unos segundos).")
        return 0

    except aiosmtplib.SMTPAuthenticationError as exc:
        fail(f"Autenticación rechazada por el servidor: code={exc.code} msg={exc.message!r}")
        msg_text = (exc.message or "")
        if "5.7.139" in msg_text or "SmtpClientAuthentication is disabled" in msg_text:
            info("Diagnóstico: SMTP AUTH está deshabilitado en el buzón.")
            info("Solución (PowerShell, como Global Admin):")
            info("  Connect-ExchangeOnline")
            info(f"  Set-CASMailbox -Identity {settings.smtp_user} -SmtpClientAuthenticationDisabled $false")
        elif "5.7.3" in msg_text or "tenant" in msg_text.lower():
            info("Diagnóstico: SMTP AUTH puede estar deshabilitado a nivel tenant.")
            info("Verificar: Get-TransportConfig | fl SmtpClientAuthenticationDisabled")
        elif "5.7.57" in msg_text or "not authenticated" in msg_text.lower() or "535" in str(exc.code):
            info("Diagnóstico: credenciales rechazadas. Revisa una a una:")
            info("  (a) ¿Password correcto en .env? Compáralo con el campo SMTP_PASSWORD impreso arriba.")
            info("  (b) ¿La cuenta tiene MFA activado? Si sí, necesitas App Password o quitar MFA.")
            info("  (c) ¿Security Defaults está ON en Entra ID? Bloquea legacy auth.")
            info("  (d) ¿Conditional Access bloqueando esta IP / esta app?")
        else:
            info("Mensaje sin patrón conocido. Copia el código de error y búscalo en docs.microsoft.com.")
        return 1

    except aiosmtplib.SMTPConnectError as exc:
        fail(f"No se pudo establecer la sesión SMTP: {exc}")
        return 1

    except aiosmtplib.SMTPException as exc:
        fail(f"Error SMTP: {type(exc).__name__}: {exc}")
        return 1

    except Exception as exc:
        fail(f"Error inesperado ({type(exc).__name__}): {exc}")
        return 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Diagnóstico SMTP de Excelater.")
    parser.add_argument("recipient", help="Email destinatario del correo de prueba.")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Activa el log DEBUG de aiosmtplib (muestra conversación SMTP).")
    args = parser.parse_args()
    sys.exit(asyncio.run(main(args.recipient, args.verbose)))

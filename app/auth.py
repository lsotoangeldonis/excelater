"""app/auth.py — JWT authentication, password hashing, role-based dependencies."""
from __future__ import annotations

import secrets
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from jose import JWTError, jwt
from passlib.context import CryptContext
from pydantic import BaseModel
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.config import settings
from app.database import AsyncSessionLocal, User, UserRole

# ── Crypto ────────────────────────────────────────────────────────────────────
_pwd_ctx = CryptContext(schemes=["bcrypt"], deprecated="auto")

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login", auto_error=False)


def hash_password(plain: str) -> str:
    return _pwd_ctx.hash(plain)


def verify_password(plain: str, hashed: str) -> bool:
    return _pwd_ctx.verify(plain, hashed)


# ── JWT ───────────────────────────────────────────────────────────────────────

class TokenData(BaseModel):
    sub: str            # username
    role: str
    exp: Optional[int] = None


def _secret() -> str:
    """Devuelve el secret JWT. Genera uno temporal si no está configurado."""
    if settings.jwt_secret:
        return settings.jwt_secret
    # Fallback: genera un secret en memoria (tokens invalidan al reiniciar)
    return secrets.token_hex(32)


def create_access_token(username: str, role: str) -> tuple[str, datetime]:
    """Crea un JWT. Devuelve (token, expires_at_utc)."""
    secret = _secret()
    expire = datetime.now(timezone.utc) + timedelta(minutes=settings.jwt_expire_minutes)
    payload = {
        "sub": username,
        "role": role,
        "exp": expire,
        "iat": datetime.now(timezone.utc),
    }
    token = jwt.encode(payload, secret, algorithm=settings.jwt_algorithm)
    return token, expire


def decode_token(token: str) -> TokenData:
    """Decodifica y valida el JWT. Lanza HTTPException si es inválido/expirado."""
    try:
        payload = jwt.decode(
            token,
            _secret(),
            algorithms=[settings.jwt_algorithm],
            options={"verify_exp": True},
        )
        username: str = payload.get("sub", "")
        role: str = payload.get("role", "")
        if not username:
            raise HTTPException(status_code=401, detail="Token inválido")
        return TokenData(sub=username, role=role)
    except JWTError:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token expirado o inválido",
            headers={"WWW-Authenticate": "Bearer"},
        )


# ── Dependencies ──────────────────────────────────────────────────────────────

async def _get_current_user_optional(token: Optional[str] = Depends(oauth2_scheme)) -> Optional[User]:
    """Obtiene el usuario actual; devuelve None si no hay token (auth deshabilitada)."""
    if not settings.auth_enabled:
        return None
    if not token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="No autenticado",
            headers={"WWW-Authenticate": "Bearer"},
        )
    data = decode_token(token)
    async with AsyncSessionLocal() as db:
        result = await db.execute(select(User).where(User.username == data.sub))
        user = result.scalar_one_or_none()
    if user is None or not user.is_active:
        raise HTTPException(status_code=401, detail="Usuario inactivo o no encontrado")
    return user


async def get_current_user(user: Optional[User] = Depends(_get_current_user_optional)) -> Optional[User]:
    return user


def require_roles(*roles: UserRole):
    """Factory de dependencia que exige que el usuario tenga uno de los roles dados."""
    async def dependency(user: Optional[User] = Depends(_get_current_user_optional)) -> Optional[User]:
        if not settings.auth_enabled:
            return user
        if user is None:
            raise HTTPException(status_code=401, detail="No autenticado")
        if user.role not in roles:
            raise HTTPException(status_code=403, detail="Permisos insuficientes")
        return user
    return dependency


# Shortcuts comunes
require_reader    = require_roles(UserRole.READER, UserRole.ADMIN, UserRole.SUPERUSER)
require_admin     = require_roles(UserRole.ADMIN, UserRole.SUPERUSER)
require_superuser = require_roles(UserRole.SUPERUSER)

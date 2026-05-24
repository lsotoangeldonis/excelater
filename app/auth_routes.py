"""app/auth_routes.py — Endpoints de autenticación y gestión de usuarios."""
from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordRequestForm
from pydantic import BaseModel, EmailStr, field_validator
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.auth import (
    create_access_token, hash_password, verify_password,
    require_admin, require_superuser, get_current_user,
)
from app.database import AsyncSessionLocal, User, UserRole, get_db

auth_router = APIRouter(prefix="/auth", tags=["auth"])

# ── Schemas ───────────────────────────────────────────────────────────────────

class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_at: datetime
    username: str
    role: str
    full_name: str


class UserOut(BaseModel):
    id: int
    username: str
    full_name: str
    email: Optional[str]
    role: UserRole
    is_active: bool
    created_at: datetime
    last_login: Optional[datetime]

    model_config = {"from_attributes": True}


class UserCreate(BaseModel):
    username: str
    full_name: str
    email: Optional[str] = None
    password: str
    role: UserRole = UserRole.READER

    @field_validator("username")
    @classmethod
    def username_valid(cls, v: str) -> str:
        v = v.strip().lower()
        if not re.match(r"^[a-z0-9_.-]{3,50}$", v):
            raise ValueError("El username debe tener 3-50 caracteres: letras, números, _ . -")
        return v

    @field_validator("password")
    @classmethod
    def password_strength(cls, v: str) -> str:
        if len(v) < 8:
            raise ValueError("La contraseña debe tener al menos 8 caracteres")
        return v


class UserUpdate(BaseModel):
    full_name: Optional[str] = None
    email: Optional[str] = None
    role: Optional[UserRole] = None
    is_active: Optional[bool] = None


class PasswordChange(BaseModel):
    current_password: str
    new_password: str

    @field_validator("new_password")
    @classmethod
    def password_strength(cls, v: str) -> str:
        if len(v) < 8:
            raise ValueError("La contraseña debe tener al menos 8 caracteres")
        return v


# ── Helper ────────────────────────────────────────────────────────────────────

async def _get_user_by_username(username: str, db: AsyncSession) -> Optional[User]:
    result = await db.execute(select(User).where(User.username == username))
    return result.scalar_one_or_none()


async def _get_user_by_id(user_id: int, db: AsyncSession) -> User:
    result = await db.execute(select(User).where(User.id == user_id))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    return user


# ── Endpoints ─────────────────────────────────────────────────────────────────

@auth_router.post("/login", response_model=TokenResponse)
async def login(form: OAuth2PasswordRequestForm = Depends()):
    """Autentica y devuelve un JWT Bearer token."""
    async with AsyncSessionLocal() as db:
        user = await _get_user_by_username(form.username.strip().lower(), db)

    if not user or not verify_password(form.password, user.hashed_pw):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Credenciales incorrectas",
            headers={"WWW-Authenticate": "Bearer"},
        )
    if not user.is_active:
        raise HTTPException(status_code=403, detail="Cuenta desactivada")

    # Actualizar last_login
    async with AsyncSessionLocal() as db:
        db_user = await _get_user_by_username(user.username, db)
        db_user.last_login = datetime.now(timezone.utc)
        await db.commit()

    token, expires_at = create_access_token(user.username, user.role.value)
    return TokenResponse(
        access_token=token,
        expires_at=expires_at,
        username=user.username,
        role=user.role.value,
        full_name=user.full_name,
    )


@auth_router.get("/me", response_model=UserOut)
async def me(current_user: Optional[User] = Depends(get_current_user)):
    """Devuelve los datos del usuario autenticado."""
    if current_user is None:
        raise HTTPException(status_code=401, detail="No autenticado")
    return current_user


@auth_router.post("/me/change-password", status_code=204)
async def change_my_password(
    body: PasswordChange,
    current_user: Optional[User] = Depends(get_current_user),
):
    """Permite al usuario cambiar su propia contraseña."""
    if current_user is None:
        raise HTTPException(status_code=401, detail="No autenticado")
    if not verify_password(body.current_password, current_user.hashed_pw):
        raise HTTPException(status_code=400, detail="Contraseña actual incorrecta")
    async with AsyncSessionLocal() as db:
        user = await _get_user_by_id(current_user.id, db)
        user.hashed_pw = hash_password(body.new_password)
        await db.commit()


# ── User management (admin+) ──────────────────────────────────────────────────

@auth_router.get("/users", response_model=list[UserOut])
async def list_users(_: Optional[User] = Depends(require_admin)):
    """Lista todos los usuarios (admin / superuser)."""
    async with AsyncSessionLocal() as db:
        result = await db.execute(select(User).order_by(User.created_at))
        return result.scalars().all()


@auth_router.post("/users", response_model=UserOut, status_code=201)
async def create_user(
    body: UserCreate,
    current_user: Optional[User] = Depends(require_admin),
):
    """Crea un usuario nuevo. Admin puede crear solo lectores; superuser puede crear cualquier rol."""
    from app.database import UserRole as R
    # Admin solo puede crear readers
    if current_user and current_user.role == R.ADMIN and body.role != R.READER:
        raise HTTPException(status_code=403, detail="Admin solo puede crear usuarios lectores")

    async with AsyncSessionLocal() as db:
        existing = await _get_user_by_username(body.username, db)
        if existing:
            raise HTTPException(status_code=409, detail="El username ya existe")
        if body.email:
            r = await db.execute(select(User).where(User.email == body.email))
            if r.scalar_one_or_none():
                raise HTTPException(status_code=409, detail="El email ya está registrado")

        user = User(
            username=body.username,
            full_name=body.full_name,
            email=body.email,
            hashed_pw=hash_password(body.password),
            role=body.role,
        )
        db.add(user)
        await db.commit()
        await db.refresh(user)
        return user


@auth_router.get("/users/{user_id}", response_model=UserOut)
async def get_user(user_id: int, _: Optional[User] = Depends(require_admin)):
    async with AsyncSessionLocal() as db:
        return await _get_user_by_id(user_id, db)


@auth_router.patch("/users/{user_id}", response_model=UserOut)
async def update_user(
    user_id: int,
    body: UserUpdate,
    current_user: Optional[User] = Depends(require_admin),
):
    """Actualiza datos de usuario. Admin no puede cambiar el rol a admin/superuser."""
    from app.database import UserRole as R
    async with AsyncSessionLocal() as db:
        user = await _get_user_by_id(user_id, db)

        # Un admin no puede editar a otro admin/superuser
        if current_user and current_user.role == R.ADMIN and user.role in (R.ADMIN, R.SUPERUSER):
            raise HTTPException(status_code=403, detail="No tienes permisos para editar este usuario")

        # Admin no puede elevar roles
        if current_user and current_user.role == R.ADMIN and body.role in (R.ADMIN, R.SUPERUSER):
            raise HTTPException(status_code=403, detail="Admin no puede asignar este rol")

        if body.full_name is not None:
            user.full_name = body.full_name
        if body.email is not None:
            user.email = body.email or None
        if body.role is not None:
            user.role = body.role
        if body.is_active is not None:
            user.is_active = body.is_active

        await db.commit()
        await db.refresh(user)
        return user


@auth_router.post("/users/{user_id}/reset-password", status_code=204)
async def reset_password(
    user_id: int,
    body: dict,
    current_user: Optional[User] = Depends(require_admin),
):
    """Resetea la contraseña de un usuario (admin/superuser)."""
    new_password = body.get("new_password", "")
    if len(new_password) < 8:
        raise HTTPException(status_code=400, detail="La contraseña debe tener al menos 8 caracteres")

    from app.database import UserRole as R
    async with AsyncSessionLocal() as db:
        user = await _get_user_by_id(user_id, db)
        if current_user and current_user.role == R.ADMIN and user.role in (R.ADMIN, R.SUPERUSER):
            raise HTTPException(status_code=403, detail="No tienes permisos sobre este usuario")
        user.hashed_pw = hash_password(new_password)
        await db.commit()


@auth_router.delete("/users/{user_id}", status_code=204)
async def delete_user(
    user_id: int,
    current_user: Optional[User] = Depends(require_superuser),
):
    """Elimina un usuario (solo superuser)."""
    if current_user and current_user.id == user_id:
        raise HTTPException(status_code=400, detail="No puedes eliminarte a ti mismo")
    async with AsyncSessionLocal() as db:
        user = await _get_user_by_id(user_id, db)
        await db.delete(user)
        await db.commit()

"""tests/conftest.py — Fixtures compartidas para los tests de Excelater"""
import asyncio
import pytest
from httpx import AsyncClient, ASGITransport
from sqlalchemy.ext.asyncio import create_async_engine, async_sessionmaker

from app.database import Base, get_db
from app.main import app
from app.scheduler import scheduler


# ── Base de datos en memoria para tests ───────────────────────────────────────
TEST_DB_URL = "sqlite+aiosqlite:///:memory:"
test_engine = create_async_engine(TEST_DB_URL, echo=False)
TestSessionLocal = async_sessionmaker(test_engine, expire_on_commit=False)


async def override_get_db():
    async with TestSessionLocal() as session:
        yield session


@pytest.fixture(autouse=True)
async def setup_db():
    """Crea las tablas antes de cada test y las limpia al terminar."""
    async with test_engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)
    yield
    async with test_engine.begin() as conn:
        await conn.run_sync(Base.metadata.drop_all)


@pytest.fixture
async def client():
    """Cliente HTTP de prueba con DB en memoria y scheduler desactivado."""
    app.dependency_overrides[get_db] = override_get_db

    # Reiniciar el scheduler en el loop actual para evitar "Event loop is closed"
    if scheduler.running:
        try:
            scheduler.shutdown(wait=False)
        except Exception:
            pass
    scheduler._eventloop = asyncio.get_running_loop()
    scheduler.start()

    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as ac:
        yield ac

    if scheduler.running:
        try:
            scheduler.shutdown(wait=False)
        except Exception:
            pass

    app.dependency_overrides.clear()

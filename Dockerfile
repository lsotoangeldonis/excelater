# ─────────────────────────────────────────────────────────────────────────────
# NOTA: Este Dockerfile ejecuta Excelater en modo openpyxl (sin Excel COM).
# El motor COM (win32com) requiere Windows con Microsoft Excel instalado y
# NO funciona dentro de contenedores Linux/Docker.
# Usa este Dockerfile para entornos donde solo necesites abrir/guardar archivos
# sin refrescar conexiones externas ni tablas dinámicas.
# ─────────────────────────────────────────────────────────────────────────────
FROM python:3.11-slim

WORKDIR /app

# Instalar Poetry
RUN pip install --no-cache-dir poetry==1.8.3

# Copiar dependencias primero para aprovechar la caché de capas
COPY pyproject.toml poetry.lock* ./
RUN poetry config virtualenvs.create false \
    && poetry install --without dev --no-root --no-interaction

# Copiar el resto del código
COPY . .
RUN poetry install --without dev --no-interaction

EXPOSE 8000

CMD ["poetry", "run", "scheduler"]

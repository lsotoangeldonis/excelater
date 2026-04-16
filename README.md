# Excel OneDrive Scheduler

Servicio web para programar y monitorear la actualización automática de archivos
Excel sincronizados desde SharePoint a través de OneDrive.

## Requisitos

- Python 3.10+
- [Poetry](https://python-poetry.org/docs/#installation)
- Windows con Microsoft Excel instalado *(para refresco de conexiones y tablas dinámicas)*

## Instalación

```bash
# 1. Clonar / descomprimir el proyecto
cd excel-onedrive-scheduler

# 2. Instalar dependencias con Poetry
poetry install

# 3. (Opcional) Crear archivo .env
cp .env.example .env
# Editar .env con tus valores
```

## Uso

```bash
# Iniciar el servidor (dashboard en http://localhost:8000)
poetry run scheduler

# O directamente
poetry run python -m app.main
```

El dashboard quedará disponible en **http://localhost:8000**

## Configuración de una tarea

Desde el dashboard → **Tareas** → **+ Nueva**:

| Campo | Descripción |
|-------|-------------|
| Nombre | Nombre descriptivo de la tarea |
| Ruta del archivo | Ruta completa al `.xlsx`, soporta `%USERPROFILE%`, `%OneDrive%` |
| Programación | Ver tipos abajo |

### Tipos de programación

**Una vez al día**
> Ejecuta a una hora fija todos los días.  
> Ej: `06:00` → actualiza cada día a las 6am.

**Repetir cada…**
> Define hora de inicio y frecuencia.  
> Ej: Desde `06:00` cada `1h` → ejecuta a las 6, 7, 8, 9… hasta las 23h.

**Expresión Cron**
> Control total con sintaxis cron estándar de 5 campos.
> ```
> minuto  hora  día  mes  día_semana
>    0      6    *    *      1-5      → lunes a viernes a las 6:00
>    0    6,12   *    *       *       → 6am y 12pm todos los días
>   30     8     *    *       1       → lunes 8:30am
> ```

## Instalar como servicio de Windows (opcional)

```bash
# Instalar NSSM (Non-Sucking Service Manager)
choco install nssm

# Crear servicio
nssm install ExcelScheduler "C:\Users\TuUsuario\AppData\Local\pypoetry\venv\Scripts\python.exe" "-m app.main"
nssm set ExcelScheduler AppDirectory "C:\ruta\al\proyecto"
nssm start ExcelScheduler
```

O usando la Task Scheduler de Windows:
```
Acción: Ejecutar programa
Programa: poetry
Argumentos: run scheduler
Directorio: C:\ruta\al\proyecto
Trigger: Al iniciar sesión / Al arrancar
```

## Variables de entorno (`.env`)

```env
HOST=0.0.0.0
PORT=8000
DEBUG=false
DB_PATH=scheduler.db
LOGS_DIR=logs
LOCK_TIMEOUT_S=120
REFRESH_TIMEOUT_S=300
MAX_LOG_SIZE_MB=10
```

## Estructura del proyecto

```
excel-onedrive-scheduler/
├── app/
│   ├── __init__.py
│   ├── main.py          # FastAPI + lifespan
│   ├── config.py        # Settings (pydantic-settings)
│   ├── database.py      # SQLAlchemy async + modelos
│   ├── excel_engine.py  # Motor COM / openpyxl
│   ├── scheduler.py     # APScheduler + lógica de triggers
│   ├── routes.py        # Endpoints REST
│   └── static/
│       └── index.html   # Dashboard SPA
├── logs/                # Archivos .log por ejecución
├── pyproject.toml
├── .env.example
└── README.md
```

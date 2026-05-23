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
cd Excelater

# 2. Instalar dependencias con Poetry
poetry install

# 3. (Opcional) Crear archivo .env
cp .env.example .env
# Editar .env con tus valores
```

## Uso

```bash
# Instalar el paquete (primera vez o tras clonar)
poetry install

# Iniciar el servidor (dashboard en http://localhost:8000)
poetry run excelater

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

## Tarea automática de Reposición (Access)

Se agregó un endpoint especializado para crear una tarea con el flujo de Reposición:

1. Actualizar cubos Excel (Maestro, SKU_SUC, Transferencias).
2. Ejecutar macro Access: Ejecutar Elimina Cubos.
3. Compactar y reparar la BD Access.
4. Ejecutar importaciones guardadas.
5. Ejecutar macro Access: Ejecutar ETL Procesos.

Endpoint:

POST /api/tasks/pipeline/reposicion

Ejemplo de payload:

{
	"name": "Actualizacion Reposicion",
	"description": "Actualizacion automatica de herramienta de reposicion",
	"schedule_type": "cron",
	"schedule_config": { "cron": "0 6 * * 1-6" },
	"access_db": "C:/Tableros/6 DWH Access/DWH Grupo Vega.accdb",
	"cubo_sku_suc_maestro": "C:/Tableros/5 Cubos RMS/Reposicion de Mercaderia/Cubo_SKU_SUC_Maestro.xlsm",
	"cubo_sku_suc": "C:/Tableros/5 Cubos RMS/Reposicion de Mercaderia/Cubo_SKU_SUC.xlsm",
	"cubo_sku_suc_transferencias": "C:/Tableros/5 Cubos RMS/Reposicion de Mercaderia/Cubo_SKU_SUC_Transferencias.xlsm",
	"compact_before_import": true,
	"max_retries": 1,
	"retry_delay_s": 300
}

## Instalar como servicio de Windows (opcional)

```bash
# Instalar NSSM (Non-Sucking Service Manager)
choco install nssm

# Crear servicio
nssm install Excelater "C:\Users\TuUsuario\AppData\Local\pypoetry\venv\Scripts\python.exe" "-m app.main"
nssm set Excelater AppDirectory "C:\ruta\al\proyecto"
nssm start Excelater
```

O usando la Task Scheduler de Windows:
```
Acción: Ejecutar programa
Programa: poetry
Argumentos: run excelater
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
Excelater/
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI + lifespan + entry point
│   ├── config.py            # Settings (pydantic-settings)
│   ├── database.py          # SQLAlchemy async + modelos
│   ├── excel_engine.py      # Motor COM / openpyxl
│   ├── access_engine.py     # Motor COM para Access (macro, compact & repair)
│   ├── scheduler.py         # APScheduler + lógica de triggers
│   ├── routes.py            # Endpoints REST
│   ├── notifications.py     # Notificaciones (email / webhook)
│   └── static/
│       └── index.html       # Dashboard SPA
├── logs/                    # Archivos .log por ejecución
├── tests/
│   ├── conftest.py
│   ├── test_routes.py
│   └── test_scheduler.py
├── pyproject.toml
├── .env.example
└── README.md
```

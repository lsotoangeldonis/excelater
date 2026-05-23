# Excelater: Excel OneDrive Scheduler

Servicio web para programar y monitorear la actualización automática de archivos excel o access locales o sincronizados desde SharePoint a través de OneDrive.

---

## Requisitos

| Requisito | Versión mínima |
|-----------|----------------|
| Windows | 10 / Server 2019 |
| Python | 3.10 |
| [Poetry](https://python-poetry.org/docs/#installation) | 1.8+ |
| Microsoft Excel | Cualquier versión de escritorio |

> Excel de escritorio es obligatorio para el refresco de conexiones y tablas dinámicas vía COM.

---

## Instalación

```powershell
# 1. Clonar / descomprimir el proyecto
cd Excelater

# 2. Instalar dependencias
poetry install

# 3. Crear archivo de configuración
copy .env.example .env
# Editar .env con tus valores (ver sección Variables de entorno)
```

---

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

---

## Ejecución manual (desarrollo / pruebas)

```powershell
poetry run excelater
```

El dashboard estará disponible en **http://localhost:8000**

---

## Instalar como servicio de Windows

Para que Excelater arranque automáticamente sin necesidad de abrir una consola,
usa el script de instalación incluido.

> **¿Por qué Task Scheduler y no un servicio de Windows puro?**
> Excel COM y OneDrive Files On-Demand requieren una sesión de usuario activa.
> Un servicio en sesión 0 no tiene acceso al escritorio ni a los archivos sincronizados por OneDrive.

### Pasos

**1. Abrir PowerShell como Administrador**

**2. Ejecutar el script desde la carpeta del proyecto:**

```powershell
cd D:\ruta\al\proyecto\Excelater
.\install-service.ps1
```

El script realiza automáticamente:

- Detecta el `python.exe` del venv de Poetry
- Ejecuta `poetry install`
- Limpia instalaciones anteriores (NSSM o tarea previa)
- Registra una tarea en el **Task Scheduler** que arranca al iniciar sesión
- Inicia el servicio inmediatamente

Al finalizar, muestra la URL del dashboard y el path del archivo de log.

### Comandos de gestión

```powershell
# Ver estado
Get-ScheduledTask        -TaskName Excelater

# Iniciar / detener manualmente
Start-ScheduledTask      -TaskName Excelater
Stop-ScheduledTask       -TaskName Excelater

# Desinstalar
Unregister-ScheduledTask -TaskName Excelater -Confirm:$false
```

### Log del servicio

```
logs\excelater.log
```

---

## Configuración de tareas

Desde el dashboard → **Tareas** → **+ Nueva**:

| Campo | Descripción |
|-------|-------------|
| Nombre | Nombre descriptivo de la tarea |
| Ruta del archivo | Ruta completa al `.xlsx`/`.xlsm`, soporta `%USERPROFILE%`, `%OneDrive%` |
| Programación | Ver tipos abajo |

### Tipos de programación

**Una vez al día**
> Ejecuta a una hora fija todos los días.
> Ej: `06:00` → actualiza cada día a las 6 am.

**Repetir cada…**
> Define hora de inicio y frecuencia.
> Ej: Desde `06:00` cada `1h` → ejecuta a las 6, 7, 8, 9… hasta las 23 h.

**Expresión Cron**
> Control total con sintaxis cron estándar de 5 campos.
> ```
> minuto  hora  día  mes  día_semana
>    0      6    *    *      1-5      → lunes a viernes a las 6:00
>    0    6,12   *    *       *       → 6 am y 12 pm todos los días
>   30     8     *    *       1       → lunes 8:30 am
> ```

---

## Tarea automática de Reposición (Access)

Endpoint especializado que crea en un solo paso el flujo completo de Reposición:

1. Actualizar cubos Excel (Maestro, SKU_SUC, Transferencias)
2. Ejecutar macro Access: **Ejecutar Elimina Cubos**
3. Compactar y reparar la BD Access
4. Ejecutar importaciones guardadas
5. Ejecutar macro Access: **Ejecutar ETL Procesos**

**Endpoint:**

```
POST /api/tasks/pipeline/reposicion
```

**Payload de ejemplo:**

```json
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
```

---

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
├── install-service.ps1      # Script de instalación como servicio
├── pyproject.toml
├── .env.example
└── README.md
```

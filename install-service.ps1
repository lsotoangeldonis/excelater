#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Instala Excelater como tarea programada de Windows (Task Scheduler).
.DESCRIPTION
    - Detecta automáticamente el python.exe del venv de Poetry
    - Instala las dependencias del proyecto
    - Elimina el servicio NSSM si existe
    - Registra la tarea para que arranque al iniciar sesión (Sesión 1)
      → Necesario para Excel COM y OneDrive Files On-Demand
.NOTES
    Debe ejecutarse como Administrador.
    Ejecutar desde la carpeta raíz del proyecto.
    IMPORTANTE: El usuario debe estar logueado para que la tarea funcione.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ServiceName = "Excelater"
$ProjectDir   = $PSScriptRoot   # carpeta donde está este script

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
function Write-Step([string]$msg) {
    Write-Host "`n>> $msg" -ForegroundColor Cyan
}

function Write-Ok([string]$msg) {
    Write-Host "   OK  $msg" -ForegroundColor Green
}

function Write-Fail([string]$msg) {
    Write-Host "   ERR $msg" -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────
# 1. Detectar poetry.exe
# ─────────────────────────────────────────────
Write-Step "Buscando poetry.exe..."

$poetryPath = (Get-Command poetry -ErrorAction SilentlyContinue)?.Source

if (-not $poetryPath) {
    # Ubicaciones comunes cuando no está en PATH
    $candidates = @(
        "$env:APPDATA\Python\Scripts\poetry.exe",
        "$env:APPDATA\pypoetry\venv\Scripts\poetry.exe",
        "$env:LOCALAPPDATA\Programs\Python\Scripts\poetry.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $poetryPath = $c; break }
    }
}

if (-not $poetryPath) {
    Write-Fail "poetry.exe no encontrado. Asegúrate de tener Poetry instalado (pip install poetry)."
}

Write-Ok "poetry: $poetryPath"

# ─────────────────────────────────────────────
# 2. Instalar dependencias y obtener path del venv
# ─────────────────────────────────────────────
Write-Step "Instalando dependencias con Poetry..."
Push-Location $ProjectDir
& $poetryPath install --without dev
if ($LASTEXITCODE -ne 0) { Write-Fail "poetry install falló. Revisa los errores anteriores." }

$venvPath = (& $poetryPath env info --path).Trim()
if (-not $venvPath -or -not (Test-Path $venvPath)) {
    Write-Fail "No se pudo obtener el path del venv de Poetry."
}
$pythonExe = Join-Path $venvPath "Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Fail "No se encontró python.exe en el venv: $pythonExe"
}
Pop-Location

Write-Ok "venv Python: $pythonExe"

# ─────────────────────────────────────────────
# 3. Validar estructura del proyecto
# ─────────────────────────────────────────────
Write-Step "Validando proyecto en: $ProjectDir"

if (-not (Test-Path "$ProjectDir\pyproject.toml")) {
    Write-Fail "No se encontró pyproject.toml en $ProjectDir. Ejecuta este script desde la carpeta raíz del proyecto."
}

if (-not (Test-Path "$ProjectDir\.env")) {
    Write-Host "   WARN .env no encontrado. Copiando desde .env.example..." -ForegroundColor Yellow
    if (Test-Path "$ProjectDir\.env.example") {
        Copy-Item "$ProjectDir\.env.example" "$ProjectDir\.env"
        Write-Host "   Edita $ProjectDir\.env antes de continuar." -ForegroundColor Yellow
    } else {
        Write-Fail ".env y .env.example ausentes. Crea el archivo .env antes de continuar."
    }
}

# Crear carpetas necesarias
@("logs", "data") | ForEach-Object {
    $folder = Join-Path $ProjectDir $_
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
        Write-Ok "Carpeta creada: $folder"
    }
}

Write-Ok "Estructura válida"

# ─────────────────────────────────────────────
# 4. Eliminar tarea/servicio previo si existen
# ─────────────────────────────────────────────
Write-Step "Verificando tarea/servicio existente..."

# Si quedó un servicio NSSM del pasado, eliminarlo primero
$legacyService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($legacyService) {
    Write-Host "   Detectado servicio NSSM '$ServiceName'. Eliminando..." -ForegroundColor Yellow
    $nssmExe = (Get-Command nssm -ErrorAction SilentlyContinue)?.Source
    if ($nssmExe) {
        & $nssmExe stop $ServiceName 2>$null
        Start-Sleep -Seconds 2
        & $nssmExe remove $ServiceName confirm
    } else {
        & sc.exe stop   $ServiceName 2>$null
        & sc.exe delete $ServiceName
    }
    Write-Ok "Servicio NSSM anterior eliminado"
}

# Si ya existe la tarea programada, eliminarla para recrearla limpiamente
$existingTask = Get-ScheduledTask -TaskName $ServiceName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "   Tarea programada '$ServiceName' ya existe. Reemplazando..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $ServiceName -Confirm:$false
    Write-Ok "Tarea anterior eliminada"
}

# ─────────────────────────────────────────────
# 5. Registrar tarea programada (Task Scheduler)
# ─────────────────────────────────────────────
Write-Step "Registrando tarea programada '$ServiceName'..."

$logsDir = Join-Path $ProjectDir "logs"
$logFile = Join-Path $logsDir "excelater.log"

# Wrappear con cmd /c para redirigir stdout+stderr al log
$argument = "/c `"$pythonExe`" -m uvicorn app.main:app --host 0.0.0.0 --port 8000 >> `"$logFile`" 2>&1"

$action = New-ScheduledTaskAction `
    -Execute          "cmd.exe" `
    -Argument         $argument `
    -WorkingDirectory $ProjectDir

# AtLogOn sin -Password → corre en la sesión interactiva (Sesión 1)
# Necesario para OneDrive Files On-Demand y Excel COM
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 5 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName $ServiceName `
    -Action   $action `
    -Trigger  $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -User     $env:USERNAME `
    -Force | Out-Null

Write-Ok "Tarea registrada para usuario: $env:USERNAME"

# ─────────────────────────────────────────────
# 6. Iniciar tarea
# ─────────────────────────────────────────────
Write-Step "Iniciando tarea..."

Start-ScheduledTask -TaskName $ServiceName
Start-Sleep -Seconds 4

$taskState = (Get-ScheduledTask -TaskName $ServiceName).State
if ($taskState -eq "Running") {
    Write-Ok "Tarea corriendo"
} else {
    Write-Host "   WARN Estado: $taskState — revisa: $logFile" -ForegroundColor Yellow
}

# ─────────────────────────────────────────────
# 7. Resumen
# ─────────────────────────────────────────────
Write-Host ""
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Excelater registrado en Task Scheduler" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  URL:      http://localhost:8000"
Write-Host "  Health:   http://localhost:8000/health"
Write-Host "  Log:      $logFile"
Write-Host "  Usuario:  $env:USERNAME  (Sesión 1 — OneDrive + Excel COM OK)"
Write-Host ""
Write-Host "  Comandos útiles:"
Write-Host "    Get-ScheduledTask        -TaskName $ServiceName"
Write-Host "    Start-ScheduledTask      -TaskName $ServiceName"
Write-Host "    Stop-ScheduledTask       -TaskName $ServiceName"
Write-Host "    Unregister-ScheduledTask -TaskName $ServiceName -Confirm:`$false"
Write-Host ""
Write-Host "  NOTA: La tarea arranca automáticamente cuando $env:USERNAME inicia sesión."
Write-Host "  Si necesitas ejecutarla sin sesión activa, considera mover los archivos"
Write-Host "  Excel a una ruta local fuera de OneDrive Files On-Demand."
Write-Host ""

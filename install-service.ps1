#Requires -RunAsAdministrator
param(
    # Puerto en que escucha uvicorn
    [int]    $Port            = 8000,

    # Abre el puerto en el Firewall de Windows sin preguntar
    [switch] $OpenFirewall,

    # Instala sin preguntas interactivas (usa valores por defecto)
    [switch] $NonInteractive,

    # No inicia la tarea al terminar la instalación
    [switch] $SkipStart
)
<#
.SYNOPSIS
    Instala Excelater como tarea programada de Windows (Task Scheduler).
.DESCRIPTION
    - Detecta automáticamente el python.exe del venv de Poetry
    - Instala las dependencias del proyecto
    - Elimina el servicio NSSM si existe
    - Registra la tarea para que arranque al iniciar sesión (Sesión 1)
      → Necesario para Excel COM y OneDrive Files On-Demand
.PARAMETER Port
    Puerto en que escucha uvicorn. Por defecto: 8000.
.PARAMETER OpenFirewall
    Abre el puerto en el Firewall de Windows sin preguntar.
.PARAMETER NonInteractive
    Instala con todos los valores por defecto sin hacer preguntas.
.PARAMETER SkipStart
    No inicia la tarea al finalizar la instalacion.
.EXAMPLE
    # Instalacion interactiva (pregunta puerto, firewall y arranque)
    .\install-service.ps1
.EXAMPLE
    # No interactivo con firewall abierto en puerto 9000
    .\install-service.ps1 -NonInteractive -OpenFirewall -Port 9000
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
# Configuración interactiva
# ─────────────────────────────────────────────
$doFirewall = $OpenFirewall.IsPresent
$doStart    = -not $SkipStart.IsPresent

if (-not $NonInteractive) {
    Write-Host ""
    Write-Host "  Configuracion de instalacion  (Enter = valor por defecto)" -ForegroundColor Cyan
    Write-Host ""

    # Puerto
    $inputPort = Read-Host "  Puerto de escucha [default: $Port]"
    if ($inputPort -match '^\d+$' -and [int]$inputPort -ge 1 -and [int]$inputPort -le 65535) {
        $Port = [int]$inputPort
    }

    # Abrir firewall (solo si no fue forzado por -OpenFirewall)
    if (-not $OpenFirewall) {
        $ans = Read-Host "  Abrir puerto $Port en el Firewall para acceso en red local? [s/N]"
        $doFirewall = $ans -match '^[sS]$'
    }

    # Inicio inmediato (solo si no fue inhibido por -SkipStart)
    if (-not $SkipStart) {
        $ans = Read-Host "  Iniciar el servicio al terminar la instalacion? [S/n]"
        if ($ans -match '^[nN]$') { $doStart = $false }
    }

    Write-Host ""
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

# Generar JWT_SECRET si no está configurado en .env
$envContent = Get-Content "$ProjectDir\.env" -Raw -ErrorAction SilentlyContinue
if ($envContent -notmatch 'JWT_SECRET\s*=\s*\S') {
    $jwtSecret = -join ((48..57 + 97..122) | Get-Random -Count 64 | ForEach-Object { [char]$_ })
    Add-Content -Path "$ProjectDir\.env" -Value "`nJWT_SECRET=$jwtSecret"
    Write-Ok "JWT_SECRET generado y guardado en .env"
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
# 4. Crear superusuario (primera instalación)
# ─────────────────────────────────────────────
Write-Step "Configurando usuario administrador..."

$dbFile = Join-Path $ProjectDir "scheduler.db"
$superadminScript = Join-Path $ProjectDir "scripts\create_superadmin.py"

if (-not (Test-Path $superadminScript)) {
    Write-Host "   WARN script create_superadmin.py no encontrado. Omitiendo paso." -ForegroundColor Yellow
} else {
    # Verificar si ya existe la tabla users con al menos un superusuario
    $hasSuperuser = $false
    if (Test-Path $dbFile) {
        try {
            $checkResult = & $pythonExe -c @"
import sqlite3, sys
try:
    c = sqlite3.connect(r'$dbFile')
    r = c.execute("SELECT COUNT(*) FROM users WHERE role='superuser'").fetchone()
    sys.exit(0 if r[0] > 0 else 1)
except: sys.exit(1)
"@
            $hasSuperuser = ($LASTEXITCODE -eq 0)
        } catch { $hasSuperuser = $false }
    }

    if ($hasSuperuser) {
        Write-Host "   Ya existe un superusuario. Omitiendo creación." -ForegroundColor Yellow
        Write-Host "   Para crear otro, ejecuta manualmente:" -ForegroundColor Yellow
        Write-Host "   poetry run python scripts/create_superadmin.py" -ForegroundColor Yellow
    } else {
        if (-not $NonInteractive) {
            Write-Host ""
            Write-Host "   No existe ningún superusuario. Se ejecutará el asistente de creación." -ForegroundColor Cyan
            Write-Host "   Completa los datos para acceder al dashboard de Excelater." -ForegroundColor Cyan
            Write-Host ""
            Push-Location $ProjectDir
            & $pythonExe scripts/create_superadmin.py
            if ($LASTEXITCODE -ne 0) {
                Write-Host "   WARN La creación del superusuario no se completó. Puedes ejecutarla manualmente:" -ForegroundColor Yellow
                Write-Host "   poetry run python scripts/create_superadmin.py" -ForegroundColor Yellow
            } else {
                Write-Ok "Superusuario creado"
            }
            Pop-Location
        } else {
            Write-Host "   Modo no-interactivo: omitiendo creación de superusuario." -ForegroundColor Yellow
            Write-Host "   Ejecuta manualmente: poetry run python scripts/create_superadmin.py" -ForegroundColor Yellow
        }
    }
}

# ─────────────────────────────────────────────
# 5. Eliminar tarea/servicio previo si existen
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
# 6. Registrar tarea programada (Task Scheduler)
# ─────────────────────────────────────────────
Write-Step "Registrando tarea programada '$ServiceName'..."

$logsDir = Join-Path $ProjectDir "logs"
$logFile = Join-Path $logsDir "excelater.log"

# Wrappear con cmd /c para redirigir stdout+stderr al log
$argument = "/c `"$pythonExe`" -m uvicorn app.main:app --host 0.0.0.0 --port $Port >> `"$logFile`" 2>&1"

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
# 7. Regla de Firewall
# ─────────────────────────────────────────────
Write-Step "Configurando regla de firewall..."

$firewallRuleName = "Excelater API"
if ($doFirewall) {
    $existing = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue
    if ($existing) {
        Set-NetFirewallRule -DisplayName $firewallRuleName -LocalPort $Port | Out-Null
        Write-Host "   Regla '$firewallRuleName' actualizada -> TCP/$Port" -ForegroundColor Yellow
    } else {
        New-NetFirewallRule -DisplayName $firewallRuleName -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow | Out-Null
        Write-Ok "Regla '$firewallRuleName' creada -> TCP/$Port"
    }
    $localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
        ($_.PrefixOrigin -eq "Dhcp" -or $_.PrefixOrigin -eq "Manual") -and
        $_.IPAddress -notmatch "^169\." -and $_.IPAddress -ne "127.0.0.1"
    } | Select-Object -First 1).IPAddress
    if ($localIP) {
        Write-Host "   Acceso en red: http://${localIP}:$Port" -ForegroundColor Green
    }
} else {
    Write-Host "   Sin cambios en firewall. Acceso solo desde localhost." -ForegroundColor Yellow
}

# ─────────────────────────────────────────────
# 8. Iniciar tarea
# ─────────────────────────────────────────────
Write-Step "Iniciando tarea..."

if (-not $doStart) {
    Write-Host "   Omitido (-SkipStart). Inicia manualmente:" -ForegroundColor Yellow
    Write-Host "   Start-ScheduledTask -TaskName $ServiceName" -ForegroundColor Yellow
} else {
    Start-ScheduledTask -TaskName $ServiceName
    Start-Sleep -Seconds 4

    $taskState = (Get-ScheduledTask -TaskName $ServiceName).State
    if ($taskState -eq "Running") {
        Write-Ok "Tarea corriendo"
    } else {
        Write-Host "   WARN Estado: $taskState — revisa: $logFile" -ForegroundColor Yellow
    }
}

# ─────────────────────────────────────────────
# 9. Resumen
# ─────────────────────────────────────────────
Write-Host ""
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Excelater registrado en Task Scheduler" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  URL:      http://localhost:$Port"
Write-Host "  Health:   http://localhost:$Port/health"
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

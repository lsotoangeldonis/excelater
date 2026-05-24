#Requires -RunAsAdministrator
param(
    # Puerto en que escucha uvicorn
    [int]    $Port            = 8000,

    # Abre el puerto en el Firewall de Windows sin preguntar
    [switch] $OpenFirewall,

    # Instala sin preguntas interactivas (usa valores por defecto)
    [switch] $NonInteractive,

    # No inicia la tarea al terminar la instalacion
    [switch] $SkipStart
)
<#
.SYNOPSIS
    Instala Excelater como tarea programada de Windows (Task Scheduler).
.DESCRIPTION
    - Detecta automaticamente el python.exe del venv de Poetry
    - Instala las dependencias del proyecto
    - Elimina el servicio NSSM si existe
    - Registra la tarea para que arranque al iniciar sesion (Sesion 1)
      -> Necesario para Excel COM y OneDrive Files On-Demand
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
    Ejecutar desde la carpeta raiz del proyecto.
    IMPORTANTE: El usuario debe estar logueado para que la tarea funcione.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ServiceName = "Excelater"
$ProjectDir   = $PSScriptRoot   # carpeta donde esta este script

# ---------------------------------------------
# Helpers
# ---------------------------------------------
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

# ---------------------------------------------
# Configuracion interactiva
# ---------------------------------------------
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

# ---------------------------------------------
# 1. Detectar poetry.exe
# ---------------------------------------------
Write-Step "Buscando poetry.exe..."

$_poetryCmd = Get-Command poetry -ErrorAction SilentlyContinue
$poetryPath = if ($_poetryCmd) { $_poetryCmd.Source } else { $null }

if (-not $poetryPath) {
    # Ubicaciones comunes cuando no esta en PATH
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
    Write-Fail "poetry.exe no encontrado. Asegurate de tener Poetry instalado (pip install poetry)."
}

Write-Ok "poetry: $poetryPath"

# ---------------------------------------------
# 2. Instalar dependencias y obtener path del venv
# ---------------------------------------------
Write-Step "Instalando dependencias con Poetry..."
Push-Location $ProjectDir
& $poetryPath install --without dev
if ($LASTEXITCODE -ne 0) { Write-Fail "poetry install fallo. Revisa los errores anteriores." }

$venvPath = (& $poetryPath env info --path).Trim()
if (-not $venvPath -or -not (Test-Path $venvPath)) {
    Write-Fail "No se pudo obtener el path del venv de Poetry."
}
$pythonExe = Join-Path $venvPath "Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    Write-Fail "No se encontro python.exe en el venv: $pythonExe"
}
Pop-Location

Write-Ok "venv Python: $pythonExe"

# ---------------------------------------------
# 3. Validar estructura del proyecto
# ---------------------------------------------
Write-Step "Validando proyecto en: $ProjectDir"

if (-not (Test-Path "$ProjectDir\pyproject.toml")) {
    Write-Fail "No se encontro pyproject.toml en $ProjectDir. Ejecuta este script desde la carpeta raiz del proyecto."
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

# Generar JWT_SECRET si no esta configurado en .env
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

Write-Ok "Estructura valida"

# ---------------------------------------------
# 4. Crear superusuario (primera instalacion)
# ---------------------------------------------
Write-Step "Configurando usuario administrador..."

$needsSuperadmin = $false  # flag para aviso en resumen final

$dbFile = Join-Path $ProjectDir "scheduler.db"
$superadminScript = Join-Path $ProjectDir "scripts\create_superadmin.py"

if (-not (Test-Path $superadminScript)) {
    Write-Host "   WARN script create_superadmin.py no encontrado. Omitiendo paso." -ForegroundColor Yellow
} else {
    # Verificar si ya existe la tabla users con al menos un superusuario
    $hasSuperuser = $false
    if (Test-Path $dbFile) {
        try {
            $tmpPy = [System.IO.Path]::GetTempFileName() + ".py"
            @'
import sqlite3, sys
found = False
try:
    db = sys.argv[1]
    c = sqlite3.connect(db)
    r = c.execute("SELECT COUNT(*) FROM users WHERE role='superuser'").fetchone()
    found = r is not None and r[0] > 0
except Exception:
    found = False
sys.exit(0 if found else 1)
'@ | Set-Content -Path $tmpPy -Encoding ASCII
            & $pythonExe $tmpPy $dbFile 2>$null
            $hasSuperuser = ($LASTEXITCODE -eq 0)
            Remove-Item $tmpPy -ErrorAction SilentlyContinue
        } catch { $hasSuperuser = $false }
    }

    if ($hasSuperuser) {
        Write-Ok "Ya existe un superusuario en la base de datos"
    } else {
        if (-not $NonInteractive) {
            Write-Host ""
            Write-Host "   No existe ningun superusuario." -ForegroundColor Yellow
            $ans = Read-Host "   Crear superusuario ahora? [S/n]"
            if ($ans -notmatch '^[nN]$') {
                Write-Host ""
                Push-Location $ProjectDir
                & $pythonExe scripts/create_superadmin.py
                if ($LASTEXITCODE -ne 0) {
                    Write-Host "   WARN La creacion no se completo. Ejecuta manualmente:" -ForegroundColor Yellow
                    Write-Host "   poetry run python scripts/create_superadmin.py" -ForegroundColor Yellow
                    $needsSuperadmin = $true
                } else {
                    Write-Ok "Superusuario creado"
                }
                Pop-Location
            } else {
                $needsSuperadmin = $true
                Write-Warn "Omitido. Recuerda crear el superusuario antes de usar el dashboard."
            }
        } else {
            $needsSuperadmin = $true
            Write-Warn "Modo no-interactivo: sin superusuario. Se recomienda crear uno tras iniciar el servicio."
        }
    }
}

# ---------------------------------------------
# 5. Eliminar tarea/servicio previo si existen
# ---------------------------------------------
Write-Step "Verificando tarea/servicio existente..."

# Si quedo un servicio NSSM del pasado, eliminarlo primero
$legacyService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($legacyService) {
    Write-Host "   Detectado servicio NSSM '$ServiceName'. Eliminando..." -ForegroundColor Yellow
    $_nssmCmd = Get-Command nssm -ErrorAction SilentlyContinue
    $nssmExe = if ($_nssmCmd) { $_nssmCmd.Source } else { $null }
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

# Verificar si la tarea ya existe con el mismo puerto (evitar recrear innecesariamente)
$existingTask = Get-ScheduledTask -TaskName $ServiceName -ErrorAction SilentlyContinue
$needsRecreate = $false
if ($existingTask) {
    $existingArg = $existingTask.Actions[0].Arguments
    if ($existingArg -match "--port\s+$Port\b") {
        Write-Ok "Tarea '$ServiceName' ya existe con puerto $Port - se actualizara en el registro"
    } else {
        Write-Host "   Puerto cambio o tarea desactualizada. Recreando..." -ForegroundColor Yellow
        $needsRecreate = $true
    }
    # Detener si esta corriendo antes de modificarla
    if ($existingTask.State -eq "Running") {
        Stop-ScheduledTask -TaskName $ServiceName
        Start-Sleep -Seconds 2
        Write-Ok "Tarea detenida para actualizacion"
    }
    if ($needsRecreate) {
        Unregister-ScheduledTask -TaskName $ServiceName -Confirm:$false
        Write-Ok "Tarea anterior eliminada"
    }
}

# ---------------------------------------------
# 6. Registrar tarea programada (Task Scheduler)
# ---------------------------------------------
Write-Step "Registrando tarea programada '$ServiceName'..."

# Si ya existe y no necesita recrearse, saltamos el registro
if ($existingTask -and -not $needsRecreate) {
    Write-Ok "Tarea ya registrada correctamente. Sin cambios."
} else {

$logsDir = Join-Path $ProjectDir "logs"
$logFile = Join-Path $logsDir "excelater.log"

# Crear directorio de logs antes de registrar la tarea
if (-not (Test-Path $logsDir)) {
    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    Write-Ok "Directorio de logs creado: $logsDir"
}

# cmd /c con comillas externas para que >> y 2>&1 sean operadores shell
$argument = "/c `"`"$pythonExe`" -m uvicorn app.main:app --host 0.0.0.0 --port $Port >> `"$logFile`" 2>&1`""

$action = New-ScheduledTaskAction `
    -Execute          "cmd.exe" `
    -Argument         $argument `
    -WorkingDirectory $ProjectDir

# AtLogOn sin -Password -> corre en la sesion interactiva (Sesion 1)
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
}

# ---------------------------------------------
# 7. Regla de Firewall
# ---------------------------------------------
Write-Step "Configurando regla de firewall..."

$firewallRuleName = "Excelater API"
if ($doFirewall) {
    $existing = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue
    if ($existing) {
        # Verificar si ya tiene el puerto correcto
        $existingPort = ($existing | Get-NetFirewallPortFilter).LocalPort
        if ($existingPort -eq $Port.ToString()) {
            Write-Ok "Regla '$firewallRuleName' ya existe con TCP/$Port - sin cambios"
        } else {
            Set-NetFirewallRule -DisplayName $firewallRuleName -LocalPort $Port | Out-Null
            Write-Ok "Regla '$firewallRuleName' actualizada: TCP/$existingPort -> TCP/$Port"
        }
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

# ---------------------------------------------
# 8. Iniciar tarea
# ---------------------------------------------
Write-Step "Iniciando tarea..."

if (-not $doStart) {
    Write-Host "   Omitido (-SkipStart). Inicia manualmente:" -ForegroundColor Yellow
    Write-Host "   Start-ScheduledTask -TaskName $ServiceName" -ForegroundColor Yellow
} else {
    $currentState = (Get-ScheduledTask -TaskName $ServiceName).State
    if ($currentState -eq "Running") {
        Write-Ok "Tarea ya esta corriendo"
    } else {
        Start-ScheduledTask -TaskName $ServiceName
        Start-Sleep -Seconds 4

        $taskState = (Get-ScheduledTask -TaskName $ServiceName).State
        if ($taskState -eq "Running") {
            Write-Ok "Tarea corriendo"
        } else {
            Write-Host "   WARN Estado: $taskState - revisa: $logFile" -ForegroundColor Yellow
        }
    }
}

# ---------------------------------------------
# 9. Resumen
# ---------------------------------------------
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Excelater registrado en Task Scheduler" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  URL:      http://localhost:$Port"
Write-Host "  Health:   http://localhost:$Port/health"
Write-Host "  Log:      $logFile"
Write-Host "  Usuario:  $env:USERNAME  (Sesion 1 - OneDrive + Excel COM OK)"
Write-Host ""
Write-Host "  Comandos utiles:"
Write-Host "    Get-ScheduledTask        -TaskName $ServiceName"
Write-Host "    Start-ScheduledTask      -TaskName $ServiceName"
Write-Host "    Stop-ScheduledTask       -TaskName $ServiceName"
Write-Host "    Unregister-ScheduledTask -TaskName $ServiceName -Confirm:`$false"
Write-Host ""
Write-Host "  NOTA: La tarea arranca automaticamente cuando $env:USERNAME inicia sesion."
Write-Host "  Si necesitas ejecutarla sin sesion activa, considera mover los archivos"
Write-Host "  Excel a una ruta local fuera de OneDrive Files On-Demand."
Write-Host ""

if ($needsSuperadmin) {
    Write-Host "  ================================================================" -ForegroundColor Red
    Write-Host "  PENDIENTE: No existe ningun superusuario en la base de datos." -ForegroundColor Red
    Write-Host "  El dashboard no sera accesible hasta que crees uno." -ForegroundColor Red
    Write-Host "  Ejecuta:" -ForegroundColor Yellow
    Write-Host "    poetry run python scripts/create_superadmin.py" -ForegroundColor Yellow
    Write-Host "  ================================================================" -ForegroundColor Red
    Write-Host ""
}

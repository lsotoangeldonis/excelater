#Requires -RunAsAdministrator
param(
    # No reinicia el servicio al terminar (solo actualiza código y deps)
    [switch] $NoRestart,

    # No ejecuta poetry install (más rápido si no cambiaron dependencias)
    [switch] $SkipInstall,

    # Rama a usar en git pull
    [string] $Branch = "main"
)
<#
.SYNOPSIS
    Actualiza Excelater desde git y reinicia el servicio.
.DESCRIPTION
    1. Detiene la tarea programada si está corriendo
    2. git pull
    3. poetry install (sincroniza dependencias)
    4. Reinicia la tarea programada
.EXAMPLE
    .\deploy.ps1
.EXAMPLE
    .\deploy.ps1 -SkipInstall     # solo git pull + restart
    .\deploy.ps1 -NoRestart       # solo actualiza, sin reiniciar
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ServiceName = "Excelater"
$ProjectDir  = $PSScriptRoot

# ── Helpers ───────────────────────────────────────────────────────────────────
function Write-Step([string]$msg) { Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-Ok([string]$msg)   { Write-Host "   OK  $msg" -ForegroundColor Green }
function Write-Warn([string]$msg) { Write-Host "   WARN $msg" -ForegroundColor Yellow }
function Write-Fail([string]$msg) { Write-Host "   ERR $msg" -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Excelater — Deploy                    " -ForegroundColor Cyan
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan

# ── 1. Detener tarea ──────────────────────────────────────────────────────────
Write-Step "Verificando tarea programada '$ServiceName'..."

$task = Get-ScheduledTask -TaskName $ServiceName -ErrorAction SilentlyContinue
$wasRunning = $false

if (-not $task) {
    Write-Warn "Tarea '$ServiceName' no registrada. Ejecuta install-service.ps1 primero."
} else {
    if ($task.State -eq "Running") {
        $wasRunning = $true
        Write-Host "   Deteniendo tarea..." -ForegroundColor Yellow
        Stop-ScheduledTask -TaskName $ServiceName
        # Esperar hasta 10 s a que pare
        $waited = 0
        while ((Get-ScheduledTask -TaskName $ServiceName).State -eq "Running" -and $waited -lt 10) {
            Start-Sleep -Seconds 1
            $waited++
        }
        Write-Ok "Tarea detenida"
    } else {
        Write-Ok "Tarea en estado: $($task.State)"
    }
}

# ── 2. Git pull ────────────────────────────────────────────────────────────────
Write-Step "Actualizando código (git pull origin $Branch)..."

Push-Location $ProjectDir

$_gitCmd = Get-Command git -ErrorAction SilentlyContinue
$gitExe = if ($_gitCmd) { $_gitCmd.Source } else { $null }
if (-not $gitExe) { Write-Fail "git no encontrado en PATH." }

# Verificar que no hay cambios locales que bloqueen el pull
$status = & git status --porcelain 2>&1
if ($status) {
    Write-Warn "Hay cambios locales sin commitear:"
    $status | ForEach-Object { Write-Host "   $_" -ForegroundColor Yellow }
    Write-Warn "Haciendo stash automático para permitir el pull..."
    & git stash push -m "deploy-auto-stash-$(Get-Date -Format 'yyyyMMdd-HHmmss')" | Out-Null
}

& git pull origin $Branch
if ($LASTEXITCODE -ne 0) { Pop-Location; Write-Fail "git pull falló. Revisa la conexión y los permisos del repositorio." }

$commitHash = (& git rev-parse --short HEAD).Trim()
$commitMsg  = (& git log -1 --pretty="%s").Trim()
Write-Ok "En commit: $commitHash - $commitMsg"

Pop-Location

# ── 3. Dependencias ────────────────────────────────────────────────────────────
if (-not $SkipInstall) {
    Write-Step "Actualizando dependencias (poetry install)..."

    $_poetryCmd = Get-Command poetry -ErrorAction SilentlyContinue
    $poetryPath = if ($_poetryCmd) { $_poetryCmd.Source } else { $null }
    if (-not $poetryPath) {
        $candidates = @(
            "$env:APPDATA\Python\Scripts\poetry.exe",
            "$env:APPDATA\pypoetry\venv\Scripts\poetry.exe",
            "$env:LOCALAPPDATA\Programs\Python\Scripts\poetry.exe"
        )
        foreach ($c in $candidates) { if (Test-Path $c) { $poetryPath = $c; break } }
    }
    if (-not $poetryPath) { Write-Fail "poetry.exe no encontrado." }

    Push-Location $ProjectDir
    & $poetryPath install --without dev
    if ($LASTEXITCODE -ne 0) { Pop-Location; Write-Fail "poetry install falló." }
    Pop-Location

    Write-Ok "Dependencias sincronizadas"
} else {
    Write-Warn "poetry install omitido (-SkipInstall)"
}

# ── 4. Reiniciar tarea ────────────────────────────────────────────────────────
if ($NoRestart) {
    Write-Warn "Reinicio omitido (-NoRestart). Inicia manualmente:"
    Write-Host "   Start-ScheduledTask -TaskName $ServiceName" -ForegroundColor Yellow
} else {
    Write-Step "Reiniciando tarea '$ServiceName'..."

    $task = Get-ScheduledTask -TaskName $ServiceName -ErrorAction SilentlyContinue
    if (-not $task) {
        Write-Warn "Tarea no registrada. No se puede iniciar. Ejecuta install-service.ps1."
    } else {
        Start-ScheduledTask -TaskName $ServiceName
        Start-Sleep -Seconds 4

        $state = (Get-ScheduledTask -TaskName $ServiceName).State
        if ($state -eq "Running") {
            Write-Ok "Tarea corriendo"
        } else {
            $logFile = Join-Path $ProjectDir "logs\excelater.log"
            Write-Warn "Estado: $state — revisa: $logFile"
        }
    }
}

# ── Resumen ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Deploy completado" -ForegroundColor Green
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Commit : $commitHash - $commitMsg"
$logFile = Join-Path $ProjectDir "logs\excelater.log"
Write-Host "  Log    : $logFile"
Write-Host ""
Write-Host "  Para ver el log en vivo:"
Write-Host ("    Get-Content `"" + $logFile + "`" -Wait -Tail 30") -ForegroundColor Yellow
Write-Host ""

<#
.SYNOPSIS
  Diagnostico SMTP de Excelater desde el lado de M365/Exchange Online.

.DESCRIPTION
  Comprueba en orden:
    1) Conectividad TCP al servidor SMTP (smtp.office365.com:587 por defecto).
    2) Estado de SMTP AUTH en el buzon (Set-CASMailbox).
    3) Estado de SMTP AUTH a nivel tenant (Get-TransportConfig).
    4) Envio real con las credenciales del buzon.

  Requiere modulo ExchangeOnlineManagement (lo instala si falta y lo autorizas).
  Debes correrlo como Global Admin (o con rol Exchange Admin) para los pasos 2 y 3.

.PARAMETER Mailbox
  UPN del buzon de servicio. Ejemplo: serviciosti@thebox.com.pe

.PARAMETER To
  Direccion a la que se envia el correo de prueba.

.PARAMETER Password
  Password del buzon. Si se omite, se pide interactivamente.

.PARAMETER SmtpServer
  Por defecto smtp.office365.com.

.PARAMETER SkipExoChecks
  Salta los pasos 2 y 3. Util si solo quieres probar el envio desde la maquina
  donde corre Excelater (sin necesidad de conectar a Exchange Online).

.EXAMPLE
  .\scripts\check_smtp.ps1 -Mailbox serviciosti@thebox.com.pe -To lsoto@thebox.com.pe

.EXAMPLE
  .\scripts\check_smtp.ps1 -Mailbox serviciosti@thebox.com.pe -To x@y.com -SkipExoChecks
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Mailbox,
    [Parameter(Mandatory = $true)][string]$To,
    [SecureString]$Password,
    [string]$SmtpServer = "smtp.office365.com",
    [int]$Port = 587,
    [switch]$SkipExoChecks
)

$ErrorActionPreference = 'Continue'

function Write-Section($title) {
    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host "  $title" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
}
function Write-Ok($m)   { Write-Host "[ OK   ] $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "[ WARN ] $m" -ForegroundColor Yellow }
function Write-Err($m)  { Write-Host "[ FAIL ] $m" -ForegroundColor Red }
function Write-Info($m) { Write-Host "[ INFO ] $m" -ForegroundColor Gray }

# ---------------------------------------------------------------------------
# 1) Conectividad TCP
# ---------------------------------------------------------------------------
Write-Section "1) Conectividad TCP a $SmtpServer`:$Port"
$net = Test-NetConnection -ComputerName $SmtpServer -Port $Port -InformationLevel Quiet -WarningAction SilentlyContinue
if ($net) {
    Write-Ok "Puerto $Port alcanzable en $SmtpServer."
} else {
    Write-Err "No se puede conectar a $SmtpServer`:$Port."
    Write-Info "Causas: firewall corporativo, proxy, DNS, o salida bloqueada al puerto 587."
    exit 1
}

# ---------------------------------------------------------------------------
# 2) y 3) Checks contra Exchange Online
# ---------------------------------------------------------------------------
if (-not $SkipExoChecks) {
    Write-Section "2) Modulo ExchangeOnlineManagement"
    $mod = Get-Module -ListAvailable -Name ExchangeOnlineManagement
    if (-not $mod) {
        Write-Warn "Modulo ExchangeOnlineManagement no instalado."
        $r = Read-Host "Instalarlo ahora (Install-Module ... -Scope CurrentUser)? (s/N)"
        if ($r -match '^[sSyY]') {
            try {
                Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Ok "Modulo instalado."
            } catch {
                Write-Err ("No se pudo instalar el modulo: " + $_.Exception.Message)
                $SkipExoChecks = $true
            }
        } else {
            Write-Warn "Saltando checks de EXO. Usa -SkipExoChecks para silenciar este aviso."
            $SkipExoChecks = $true
        }
    } else {
        Write-Ok "Modulo ExchangeOnlineManagement presente (version $($mod[0].Version))."
    }
}

if (-not $SkipExoChecks) {
    Write-Section "3) Estado de SMTP AUTH (buzon + tenant)"
    Write-Info "Conectando a Exchange Online. Te pedira login de Global Admin..."
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Ok "Conectado a Exchange Online."
    } catch {
        Write-Err ("Fallo Connect-ExchangeOnline: " + $_.Exception.Message)
        exit 1
    }

    # Buzon
    try {
        $cas = Get-CASMailbox -Identity $Mailbox -ErrorAction Stop
        $val = $cas.SmtpClientAuthenticationDisabled
        if ($val -eq $false) {
            Write-Ok "Buzon $Mailbox -> SmtpClientAuthenticationDisabled = False (HABILITADO)."
        } elseif ($null -eq $val) {
            Write-Info "Buzon $Mailbox -> SmtpClientAuthenticationDisabled = (null). Hereda del tenant."
        } else {
            Write-Err "Buzon $Mailbox -> SmtpClientAuthenticationDisabled = True (DESHABILITADO)."
            Write-Info "Solucion:"
            Write-Info "  Set-CASMailbox -Identity $Mailbox -SmtpClientAuthenticationDisabled `$false"
        }
    } catch {
        Write-Err ("No se pudo leer el buzon $Mailbox -> " + $_.Exception.Message)
    }

    # Tenant
    try {
        $tc = Get-TransportConfig -ErrorAction Stop
        if ($tc.SmtpClientAuthenticationDisabled -eq $true) {
            Write-Warn "Tenant: SmtpClientAuthenticationDisabled = True (DESHABILITADO globalmente)."
            Write-Info "Esto no impide enviar si el buzon tiene un override en False, pero es indicio de tenant moderno."
        } elseif ($tc.SmtpClientAuthenticationDisabled -eq $false) {
            Write-Ok "Tenant: SMTP AUTH habilitado globalmente."
        } else {
            Write-Info "Tenant: SmtpClientAuthenticationDisabled no esta definido (default permite SMTP AUTH si el buzon lo permite)."
        }
    } catch {
        Write-Warn ("No se pudo leer Get-TransportConfig: " + $_.Exception.Message)
    }

    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Write-Info "Desconectado de Exchange Online."
}

# ---------------------------------------------------------------------------
# 4) Envio real
# ---------------------------------------------------------------------------
Write-Section "4) Envio SMTP real ($SmtpServer`:$Port, STARTTLS)"
if (-not $Password) {
    $Password = Read-Host "Ingresa la password del buzon $Mailbox" -AsSecureString
}
$cred = New-Object System.Management.Automation.PSCredential($Mailbox, $Password)

$subject = "[Excelater] Test SMTP " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$body    = "Test de envio SMTP desde el host que corre Excelater. Si lo recibes, SMTP AUTH funciona."

try {
    Send-MailMessage `
        -From $Mailbox `
        -To $To `
        -Subject $subject `
        -Body $body `
        -SmtpServer $SmtpServer `
        -Port $Port `
        -UseSsl `
        -Credential $cred `
        -ErrorAction Stop
    Write-Ok "Correo enviado a $To. Revisa la bandeja (puede tardar unos segundos)."
    exit 0
} catch {
    $errMsg = $_.Exception.Message
    Write-Err "Fallo envio: $errMsg"

    if ($errMsg -match "5\.7\.139|SmtpClientAuthentication is disabled") {
        Write-Info "Diagnostico: SMTP AUTH deshabilitado en el buzon. Corre:"
        Write-Info "  Set-CASMailbox -Identity $Mailbox -SmtpClientAuthenticationDisabled `$false"
    } elseif ($errMsg -match "5\.7\.57|5\.7\.3|535|not authenticated|InvalidCredentials") {
        Write-Info "Diagnostico: credenciales rechazadas. Causas comunes:"
        Write-Info "  (a) Password incorrecto. Si tiene `$` en .env, debe ir entre comillas simples."
        Write-Info "  (b) MFA activado en la cuenta -> necesitas App Password o quitar MFA."
        Write-Info "  (c) Security Defaults ON en Entra ID -> bloquea legacy auth."
        Write-Info "  (d) Conditional Access bloqueando esta IP / app."
    } elseif ($errMsg -match "5\.4\.1|access denied|relay") {
        Write-Info "Diagnostico: relay denegado. Casi siempre el From no coincide con el usuario autenticado."
    } else {
        Write-Info "Codigo de error sin patron conocido. Busca el codigo SMTP en docs.microsoft.com."
    }
    exit 1
}

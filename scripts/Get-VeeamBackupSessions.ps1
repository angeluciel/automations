##################################################################
####  Script para automação das rotinas de Backup  ###############
##################################################################


# ──  Instala Módulos de Segurança ───────────────────────────── #

if (-not (Get-Module -ListAvailable -Name Microsoft.Powershell.SecretManagement )) {
    Install-Module -Name Microsoft.Powershell.SecretManagement -AllowClobber -Force
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Powershell.SecretStore)) {
    Install-Module -Name Microsoft.Powershell.SecretStore
}

# ──  Instala Módulos de Segurança ───────────────────────────── #

# ──  Configuração  ──────────────────────────────────────────── #

$pathToCredential = "C:\path\secret.xml"
$loadedCredential = Import-CliXML -Path $pathToCredential

Unlock-SecretStore -Password $loadedCredential

$BaseUrl     = "https://server:9419"
$SecretToken = (Invoke-RestMethod `
    -Uri     "$BaseUrl/api/oauth2/token" `
    -Method  Post `
    -Headers @{
        "accept"        = "application/json"
        "x-api-version" = "1.3-rev1"
        "Content-Type"  = "application/x-www-form-urlencoded"
    } `
    -Body @{
        grant_type      = "Password"
        username        = (Get-Secret -Name "backup-user" -AsPlainText)
        password        = (Get-Secret -Name "backup-user_password" -AsPlainText)
    } `
    -SkipCertificateCheck).access_token

# ────────────────────────────────────────────────────────────── #


# ──  Derivar a data do mês anterior  ────────────────────────── #

$Today              = Get-Date
$FirstOfThisMonth   = Get-Date -Year $Today.Year -Month $Today.Month -Day 1 `
    -Hour 0 -Minute 0 -Millisecond 0
$FirstOfLastMonth   = $FirstOfThisMonth.AddMonths(-1)
$LastOfLastMonth    = $FirstOfThisMonth.AddSeconds(-1)

# ──  Formatar para ISO 8601 UTC (api só aceita assim) ───────── #
$CreatedAfter       = $FirstOfLastMonth.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$EndedBefore        = $LastOfLastMonth.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host "Consultando sessões de: $CreatedAfter -> $EndedBefore" -ForegroundColor Cyan

# ────────────────────────────────────────────────────────────── #


# ──  Requisição HTTP  ───────────────────────────────────────── #

$Headers = @{
    "Authorization" = "Bearer $SecretToken"
    "x-api-version" = "1.3-rev1"
    "Accept"        = "application/json"
}

$QueryParams = @{
    typeFilter          = "BackupJob"
    createdAfterFilter  = $CreatedAfter
    endedBeforeFilter   = $EndedBefore
}

$QueryString = ($QueryParams.GetEnumerator() |
    ForEach-Object { "$($_.Key)=$([Uri]::EscapeDataString($_.Value))"}) -join "&"

$Uri = "$BaseUrl/api/v1/sessions?$QueryString"

# ────────────────────────────────────────────────────────────── #


# ──  Roda a Request  ────────────────────────────────────────── #

try {
    $Response = Invoke-RestMethod `
        -Uri            $Uri `
        -Method         GET `
        -Headers        $Headers `
        -SkipCertificateCheck

    Write-Host "Foram retornadas: $($Response.data.Count) sessões." -ForegroundColor Green
    $Response.data | ConvertTo-Json -Depth 10
} catch {
    Write-Error "Request failed: $_"
}
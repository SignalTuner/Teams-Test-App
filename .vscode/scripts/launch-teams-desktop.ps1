param(
    [string] $EnvFile = ".\env\.env.local"
)

$ErrorActionPreference = "Stop"

$envFile = $EnvFile
if (-not (Test-Path -LiteralPath $envFile)) {
    throw "Missing $envFile."
}

$line = Get-Content -LiteralPath $envFile | Where-Object { $_ -match "^TEAMS_APP_ID=" } | Select-Object -First 1
if (-not $line) {
    throw "Missing TEAMS_APP_ID in $envFile."
}

$teamsAppId = ($line -split "=", 2)[1].Trim()
if (-not $teamsAppId) {
    throw "TEAMS_APP_ID is empty in $envFile."
}

Start-Process "msteams://teams.microsoft.com/l/app/$teamsAppId`?installAppPackage=true"
Write-Host "Opened Teams desktop for app $teamsAppId."

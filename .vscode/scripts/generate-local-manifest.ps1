$ErrorActionPreference = "Stop"

$envFile = ".\env\.env.local"
if (-not (Test-Path -LiteralPath $envFile)) {
    throw "Missing env\.env.local. Run Microsoft 365 Agents Toolkit provision for local once."
}

$envValues = @{}
Get-Content -LiteralPath $envFile |
    Where-Object { $_ -match "^[A-Za-z_][A-Za-z0-9_]*=" } |
    ForEach-Object {
        $parts = $_ -split "=", 2
        $envValues[$parts[0]] = $parts[1]
    }

foreach ($required in @("TEAMS_APP_ID", "APP_NAME_SUFFIX", "TAB_ENDPOINT")) {
    if (-not $envValues.ContainsKey($required) -or [string]::IsNullOrWhiteSpace($envValues[$required])) {
        throw "Missing $required in env\.env.local."
    }
}

New-Item -ItemType Directory -Path ".\appPackage\build" -Force | Out-Null

$localManifestPath = ".\appPackage\build\manifest.local.json"
if (Test-Path -LiteralPath $localManifestPath) {
    Set-ItemProperty -LiteralPath $localManifestPath -Name IsReadOnly -Value $false
}

$manifest = Get-Content -LiteralPath ".\appPackage\manifest.json" -Raw
foreach ($key in $envValues.Keys) {
    $manifest = $manifest.Replace('${{' + $key + '}}', $envValues[$key])
}

Set-Content -LiteralPath $localManifestPath -Value $manifest -Encoding UTF8
Write-Host "Generated appPackage\build\manifest.local.json from env\.env.local."

param([switch]$ValidateOnly)

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

$manifest = Get-Content -LiteralPath ".\appPackage\manifest.json" -Raw
$manifestSettings = $manifest | ConvertFrom-Json
$tabEndpoint = [Uri]$envValues["TAB_ENDPOINT"]
if (-not $tabEndpoint.IsAbsoluteUri -or $tabEndpoint.Scheme -ne "https") {
    throw "TAB_ENDPOINT must be an absolute HTTPS URL."
}
if ($manifestSettings.webApplicationInfo.id -ne $manifestSettings.bots[0].botId) {
    throw "Teams SSO/RSC identity must match the meeting bot identity."
}
$expectedResource = "api://$($tabEndpoint.Authority)/botid-$($manifestSettings.bots[0].botId)"
if ($envValues.ContainsKey("TEAMS_SSO_RESOURCE") -and $envValues["TEAMS_SSO_RESOURCE"] -ne $expectedResource) {
    throw "TEAMS_SSO_RESOURCE must match TAB_ENDPOINT and the meeting bot identity."
}
$envValues["TEAMS_SSO_RESOURCE"] = $expectedResource
$envValues["TAB_DOMAIN"] = $tabEndpoint.DnsSafeHost

foreach ($key in $envValues.Keys) {
    $manifest = $manifest.Replace('${{' + $key + '}}', $envValues[$key])
}
if ($manifest -match '\$\{\{[^}]+\}\}') {
    throw "The local manifest still contains unresolved environment placeholders."
}
$renderedManifest = $manifest | ConvertFrom-Json
if ($ValidateOnly) {
    Write-Host "Validated local manifest $($renderedManifest.version): bot/RSC identity, HTTPS origin, SSO resource, and environment placeholders. No package generated."
    return
}

New-Item -ItemType Directory -Path ".\appPackage\build" -Force | Out-Null
$localManifestPath = ".\appPackage\build\manifest.local.json"
if (Test-Path -LiteralPath $localManifestPath) {
    Set-ItemProperty -LiteralPath $localManifestPath -Name IsReadOnly -Value $false
}

Set-Content -LiteralPath $localManifestPath -Value $manifest -Encoding UTF8
Write-Host "Generated appPackage\build\manifest.local.json from env\.env.local."

param([string]$EnvironmentFile, [switch]$RequireProduction)

# Read-only preflight, not Microsoft's schema/package validator or certification.
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$problems = [System.Collections.Generic.List[string]]::new()
function Check([bool]$condition, [string]$message) {
    if (-not $condition) { $problems.Add($message) }
}
$source = Get-Content -LiteralPath (Join-Path $root "appPackage/manifest.json") -Raw -Encoding UTF8
$manifest = $source | ConvertFrom-Json
Check ($manifest.name.short -eq "SignalTuner") "Use SignalTuner without a local/dev/custom suffix."
Check ($manifest.developer.name -eq "SignalTuner Inc.") "Confirm the SignalTuner publisher name."
Check ($manifest.developer.websiteUrl -eq "https://www.signaltuner.com") "Replace the placeholder publisher website."
Check ($manifest.description.short.Length -le 80) "Short description exceeds 80 characters."
Check ($manifest.description.full.Length -le 4000) "Full description exceeds 4000 characters."
Check ($manifest.webApplicationInfo.id -eq $manifest.bots[0].botId) "Bot and SSO/RSC identities differ."
Check ($manifest.version -match '^\d+\.\d+\.\d+$') "App version must have three numeric components."
Check ($source -notmatch 'example\.com|My App, Inc\.|Teams Test App') "Manifest contains template branding."

Add-Type -AssemblyName System.Drawing
foreach ($variant in @(@{ Key = "color"; Size = 192 }, @{ Key = "outline"; Size = 32 })) {
    $file = $manifest.icons.($variant.Key)
    Check ($file -match '^[\w-]+\.png$') "Icon must be a PNG filename at the package root."
    $icon = [System.Drawing.Bitmap]::FromFile((Join-Path $root "appPackage/$file"))
    try {
        Check ($icon.RawFormat.Guid -eq [System.Drawing.Imaging.ImageFormat]::Png.Guid) "$file must contain PNG image data."
        Check ($icon.Width -eq $variant.Size -and $icon.Height -eq $variant.Size) "$file must be $($variant.Size) x $($variant.Size)."
        $hasVisible = $false; $hasTransparent = $false; $validPixels = $true
        for ($y = 0; $y -lt $icon.Height; $y++) {
            for ($x = 0; $x -lt $icon.Width; $x++) {
                $pixel = $icon.GetPixel($x, $y)
                if ($pixel.A -gt 0) { $hasVisible = $true }
                if ($pixel.A -eq 0) { $hasTransparent = $true }
                if ($variant.Key -eq "outline") {
                    if ($pixel.A -gt 0 -and ($pixel.R -ne 255 -or $pixel.G -ne 255 -or $pixel.B -ne 255)) { $validPixels = $false }
                } elseif ($pixel.A -ne 255) { $validPixels = $false }
            }
        }
        Check $hasVisible "$file is empty."
        Check $validPixels "$file must be opaque for color, or white with transparency for outline."
        if ($variant.Key -eq "outline") { Check $hasTransparent "Outline must have a transparent background." }
    } finally { $icon.Dispose() }
}

if ($EnvironmentFile) {
    # Only read non-secret keys used by the manifest; never print env-file values.
    $values = @{}
    foreach ($line in Get-Content -LiteralPath $EnvironmentFile) {
        if ($line -match '^\s*(TEAMSFX_ENV|TEAMS_APP_ID|TAB_DOMAIN|TAB_ENDPOINT|TEAMS_SSO_RESOURCE)=(.*)$') {
            $values[$Matches[1]] = $Matches[2].Trim().Trim('"').Trim("'")
        }
    }
    foreach ($key in @("TEAMSFX_ENV", "TEAMS_APP_ID", "TAB_DOMAIN", "TAB_ENDPOINT", "TEAMS_SSO_RESOURCE")) {
        Check (-not [string]::IsNullOrWhiteSpace($values[$key])) "Missing production $key."
    }
    Check ($values["TEAMSFX_ENV"] -eq "prod") "Use an explicitly production environment, not local/dev."
    $appId = [Guid]::Empty
    Check ([Guid]::TryParse($values["TEAMS_APP_ID"], [ref]$appId) -and $appId -ne [Guid]::Empty) "Production TEAMS_APP_ID must be a nonempty GUID."
    $endpoint = $null
    $validEndpoint = [Uri]::TryCreate($values["TAB_ENDPOINT"], [UriKind]::Absolute, [ref]$endpoint)
    Check ($validEndpoint -and $endpoint.Scheme -eq "https") "TAB_ENDPOINT must be an absolute HTTPS origin."
    if ($validEndpoint) {
        Check ($endpoint.AbsolutePath -eq "/" -and -not $endpoint.Query -and -not $endpoint.Fragment -and -not $endpoint.UserInfo) "TAB_ENDPOINT must be an origin without a path, query, fragment, or credentials."
        Check ($values["TAB_ENDPOINT"] -notmatch '/$') "TAB_ENDPOINT must omit the trailing slash for template routes."
        Check ($values["TAB_DOMAIN"] -eq $endpoint.DnsSafeHost) "TAB_DOMAIN must match TAB_ENDPOINT."
        Check ($values["TEAMS_SSO_RESOURCE"] -eq "api://$($endpoint.Authority)/botid-$($manifest.bots[0].botId)") "Production SSO resource must match the tab origin and bot."
        Check (-not $endpoint.IsLoopback -and $endpoint.HostNameType -eq [UriHostNameType]::Dns -and $endpoint.Host.Contains('.')) "A public production DNS hostname is required."
    }
    $rendered = $source
    foreach ($key in $values.Keys) {
        # JSON-escape substitution values before rendering, without writing a file.
        $escaped = ConvertTo-Json -InputObject ([string]$values[$key]) -Compress
        $rendered = $rendered.Replace(('${{' + $key + '}}'), $escaped.Substring(1, $escaped.Length - 2))
    }
    Check ($rendered -notmatch '\$\{\{') "Manifest has unresolved environment placeholders."
    Check ($rendered -notmatch '(?i)localhost|127\.0\.0\.1|0\.0\.0\.0|\.devtunnels\.ms|\.ngrok\.|\.local\b|\.test\b|\.invalid\b|example\.(com|org|net)') "Manifest contains development, tunnel, or placeholder hosts."
    $null = $rendered | ConvertFrom-Json
} elseif ($RequireProduction) {
    $problems.Add("Supply -EnvironmentFile env/.env.prod with confirmed production values.")
}

if ($problems.Count) {
    foreach ($problem in $problems) { Write-Host "FAIL: $problem" }
    throw "Store preflight failed ($($problems.Count) checks). No files written or packages created."
}
if ($EnvironmentFile) {
    Write-Host "Production manifest inputs and icon checks passed. Hosting, Entra consent, cross-tenant access, mobile policy, live workflows and Microsoft package validation still require verification."
} else {
    Write-Host "Source branding and icon checks passed. Production readiness NOT checked; use -RequireProduction -EnvironmentFile env/.env.prod before packaging."
}
Write-Host "No files written, package created, upload, or deployment performed. See docs/store-submission.md."

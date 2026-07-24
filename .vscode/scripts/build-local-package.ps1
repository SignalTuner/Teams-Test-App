$ErrorActionPreference = "Stop"

npm install
npm run build

$stage = Join-Path $env:TEMP "signaltuner-appPackage-local"
if (Test-Path -LiteralPath $stage) {
    Remove-Item -LiteralPath $stage -Recurse -Force
}

New-Item -ItemType Directory -Path $stage | Out-Null
Copy-Item -LiteralPath ".\appPackage\build\manifest.local.json" -Destination (Join-Path $stage "manifest.json")
Copy-Item -LiteralPath ".\appPackage\color.png" -Destination (Join-Path $stage "color.png")
Copy-Item -LiteralPath ".\appPackage\outline.png" -Destination (Join-Path $stage "outline.png")
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath ".\appPackage\build\appPackage.local.zip" -Force
Remove-Item -LiteralPath $stage -Recurse -Force

Write-Host "Built appPackage\build\appPackage.local.zip."

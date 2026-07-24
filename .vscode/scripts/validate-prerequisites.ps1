$ErrorActionPreference = "Stop"

$nodeVersion = node --version
if (-not $nodeVersion) {
    throw "Node.js is required."
}

$ports = @(3978, 3979, 9239)
foreach ($port in $ports) {
    $listener = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
    if ($listener) {
        throw "Port $port is already in use by process $($listener.OwningProcess). Close that process and try Teams debug again."
    }
}

if (-not (Test-Path -LiteralPath ".\env\.env.local")) {
    throw "Missing env\.env.local. Run Microsoft 365 Agents Toolkit provision for local once, or restore the local environment file."
}

Write-Host "Prerequisites OK. Node $nodeVersion; ports 3978, 3979, and 9239 are available."

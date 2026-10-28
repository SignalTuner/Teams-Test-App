param([switch]$ValidateOnly)
$ErrorActionPreference = 'Stop'

# Run after closing VS Code/terminals/debuggers using the old folder.
# Native PowerShell only; never merge/copy two checkouts or change Git remotes.
$checkout = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$repoParent = Split-Path $checkout -Parent
$source = [IO.Path]::GetFullPath((Join-Path $repoParent 'Teams Test App'))
$target = [IO.Path]::GetFullPath((Join-Path $repoParent 'SignalTuner Teams App'))
if ((Split-Path $source -Parent) -ne $repoParent -or (Split-Path $target -Parent) -ne $repoParent -or
    ($checkout -ne $source -and $checkout -ne $target)) {
    throw 'Unexpected checkout/rename paths. Inspect before proceeding.'
}
$sourceExists = Test-Path -LiteralPath $source -PathType Container
$targetExists = Test-Path -LiteralPath $target -PathType Container
if ($sourceExists -eq $targetExists) { throw 'Expected exactly one checkout. Refusing to merge directories or guess a source.' }
$active = if ($sourceExists) { $source } else { $target }
if ((Get-Item -LiteralPath $active).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Refusing to rename a redirected checkout.' }

$projects = @('SignalTuner','signaltuner-back','signaltuner-front','is-service-down','SignalTuner Teams App')
$contents = @{}
foreach ($project in $projects) {
    $folder = if ($project -eq 'SignalTuner Teams App') { $active } else { Join-Path $repoParent $project }
    $file = Join-Path $folder 'AGENTS.md'
    if ((Get-Item -LiteralPath $file).IsReadOnly) { throw "Read-only AGENTS.md in $project." }
    $contents[$project] = [IO.File]::ReadAllText($file)
    if (-not ($contents[$project].Contains($source) -or $contents[$project].Contains($target))) {
        throw "Missing expected project path in $project AGENTS.md."
    }
}
$workspaceFile = Join-Path $repoParent 'SignalTuner.code-workspace'
if ((Get-Item -LiteralPath $workspaceFile).IsReadOnly) { throw 'Shared workspace is read-only.' }
$workspaceText = [IO.File]::ReadAllText($workspaceFile)
if (-not ($workspaceText.Contains('"path": "Teams Test App"') -or $workspaceText.Contains('"path": "SignalTuner Teams App"'))) {
    throw 'Unexpected workspace entry. Inspect before proceeding.'
}

$verifyFiles = @('.git/config','appPackage/manifest.json','m365agents.local.yml','src/index.ts',
    'env/.env.local','.env.local','.localConfigs','.vscode/launch.json','.vscode/tasks.json')
$hashes = @{}
foreach ($file in $verifyFiles) {
    $full = Join-Path $active $file
    if (Test-Path -LiteralPath $full -PathType Leaf) { $hashes[$file] = (Get-FileHash -LiteralPath $full).Hash }
}
if ($ValidateOnly) {
    Write-Output "Rename preflight passed. Current checkout: $active"
    Write-Output "Target: $target. Five AGENTS maps and workspace are accessible. No files changed."
    exit 0
}

# All absolute paths have been checked before this directory operation.
Set-Location -LiteralPath $repoParent
if ($sourceExists) {
    try { Rename-Item -LiteralPath $source -NewName 'SignalTuner Teams App' }
    catch { throw 'Rename blocked. Close VS Code, terminals and debuggers using the old folder, then rerun from a separate PowerShell window. No path metadata was changed.' }
}
$utf8 = New-Object System.Text.UTF8Encoding($false)
foreach ($project in $projects) {
    $updated = $contents[$project].Replace($source, $target).Replace(
        'Microsoft 365/Teams Agents Toolkit project; inspect `package.json` and manifests for the exact runtime',
        'React/Vite static frontend on DreamHost Apache; Microsoft 365/Teams Agents Toolkit and Node.js for local debugging')
    [IO.File]::WriteAllText((Join-Path (Join-Path $repoParent $project) 'AGENTS.md'), $updated, $utf8)
}
$workspaceText = $workspaceText.Replace('"path": "Teams Test App"', '"path": "SignalTuner Teams App"')
# Use the folder entry's path to identify it without depending on punctuation in its old display name.
$workspaceText = [regex]::Replace($workspaceText,
    '"name"\s*:\s*"[^"]+"(\s*,\s*"path"\s*:\s*"SignalTuner Teams App")',
    '"name": "SignalTuner Teams App"$1')
[IO.File]::WriteAllText($workspaceFile, $workspaceText, $utf8)
foreach ($file in $hashes.Keys) {
    if ((Get-FileHash -LiteralPath (Join-Path $target $file)).Hash -ne $hashes[$file]) { throw "Protected file changed during rename: $file" }
}
foreach ($project in $projects) {
    if ([IO.File]::ReadAllText((Join-Path (Join-Path $repoParent $project) 'AGENTS.md')).Contains($source)) {
        throw "Stale path in $project AGENTS.md. Rerun this script from the new directory."
    }
}
Write-Output "Renamed project: $target"
Write-Output 'Updated all five AGENTS.md maps and SignalTuner.code-workspace. Git remote, manifest IDs and local config preserved.'
Write-Output "Reopen: $workspaceFile"

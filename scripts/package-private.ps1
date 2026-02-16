param(
    [string]$SourceDir = "",
    [string]$OutDir = "",
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    $SourceDir = Join-Path $repoRoot "GoogleContactsSync\bin\Release"
}
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path $repoRoot "dist"
}

$assemblyInfo = Join-Path $repoRoot "GoogleContactsSync\Properties\AssemblyInfo.cs"
if ([string]::IsNullOrWhiteSpace($Version)) {
    $m = Select-String -Path $assemblyInfo -Pattern 'AssemblyInformationalVersion\("([^"]+)"\)' | Select-Object -First 1
    if ($m -and $m.Matches.Count -gt 0) {
        $Version = $m.Matches[0].Groups[1].Value
    } else {
        $Version = "private"
    }
}

$stageDir = Join-Path $OutDir ("GOContactSyncMod-{0}" -f $Version)
$zipPath = "$stageDir.zip"

if (-not (Test-Path (Join-Path $SourceDir "GOContactSync.exe"))) {
    throw "Built EXE not found in $SourceDir. Run scripts\build-private.ps1 first."
}

if (Test-Path $stageDir) {
    Remove-Item $stageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

Get-ChildItem -Path $SourceDir -File | Where-Object { $_.Extension -ne ".pdb" } | ForEach-Object {
    Copy-Item $_.FullName -Destination (Join-Path $stageDir $_.Name) -Force
}

if (Test-Path (Join-Path $SourceDir "Resources")) {
    Copy-Item (Join-Path $SourceDir "Resources") -Destination (Join-Path $stageDir "Resources") -Recurse -Force
}

"Version: $Version" | Set-Content -Path (Join-Path $stageDir "VERSION.txt") -Encoding UTF8

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}
Compress-Archive -Path (Join-Path $stageDir "*") -DestinationPath $zipPath -CompressionLevel Optimal -Force

Write-Host "Package created:"
Write-Host "  Stage: $stageDir"
Write-Host "  Zip:   $zipPath"

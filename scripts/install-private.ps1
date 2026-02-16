param(
    [ValidateSet("CurrentUser", "AllUsers")]
    [string]$Scope = "CurrentUser",
    [string]$SourceDir = "",
    [string]$InstallDir = "",
    [switch]$NoDesktopShortcut,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    $SourceDir = Join-Path $repoRoot "GoogleContactsSync\bin\Release"
}

$exeName = "GOContactSync.exe"
$sourceExe = Join-Path $SourceDir $exeName
if (-not $Uninstall -and -not (Test-Path $sourceExe)) {
    throw "Built EXE not found: $sourceExe. Run scripts\build-private.ps1 first."
}

$isAllUsers = $Scope -eq "AllUsers"
if ([string]::IsNullOrWhiteSpace($InstallDir)) {
    if ($isAllUsers) {
        $InstallDir = Join-Path $env:ProgramFiles "GO Contact Sync Mod JKFix"
    } else {
        $InstallDir = Join-Path $env:LOCALAPPDATA "Programs\GO Contact Sync Mod JKFix"
    }
}

if ($isAllUsers) {
    $startMenuRoot = Join-Path $env:ProgramData "Microsoft\Windows\Start Menu\Programs"
    $desktopRoot = Join-Path $env:Public "Desktop"
} else {
    $startMenuRoot = [Environment]::GetFolderPath("Programs")
    $desktopRoot = [Environment]::GetFolderPath("Desktop")
}

$startMenuFolder = Join-Path $startMenuRoot "GO Contact Sync Mod JKFix"
$startMenuLnk = Join-Path $startMenuFolder "GO Contact Sync Mod JKFix.lnk"
$desktopLnk = Join-Path $desktopRoot "GO Contact Sync Mod JKFix.lnk"
$installedExe = Join-Path $InstallDir $exeName

function New-Shortcut {
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [string]$WorkingDirectory
    )

    $wsh = New-Object -ComObject WScript.Shell
    $shortcut = $wsh.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.IconLocation = "$TargetPath,0"
    $shortcut.Save()
}

if ($Uninstall) {
    Write-Host "Uninstalling private build from: $InstallDir"
    if (Test-Path $startMenuLnk) { Remove-Item $startMenuLnk -Force }
    if (Test-Path $desktopLnk) { Remove-Item $desktopLnk -Force }
    if (Test-Path $startMenuFolder) {
        Remove-Item $startMenuFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $InstallDir) {
        Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Uninstall complete."
    exit 0
}

Write-Host "Installing private build to: $InstallDir"
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

# Keep release payload clean by skipping pdb files.
Get-ChildItem -Path $SourceDir -File | Where-Object { $_.Extension -ne ".pdb" } | ForEach-Object {
    Copy-Item $_.FullName -Destination (Join-Path $InstallDir $_.Name) -Force
}

if (Test-Path (Join-Path $SourceDir "Resources")) {
    Copy-Item (Join-Path $SourceDir "Resources") -Destination (Join-Path $InstallDir "Resources") -Recurse -Force
}

New-Item -ItemType Directory -Path $startMenuFolder -Force | Out-Null
New-Shortcut -ShortcutPath $startMenuLnk -TargetPath $installedExe -WorkingDirectory $InstallDir

if (-not $NoDesktopShortcut) {
    New-Shortcut -ShortcutPath $desktopLnk -TargetPath $installedExe -WorkingDirectory $InstallDir
}

Write-Host "Install complete."
Write-Host "Start menu shortcut: $startMenuLnk"
if (-not $NoDesktopShortcut) {
    Write-Host "Desktop shortcut: $desktopLnk"
}
Write-Host "Executable: $installedExe"

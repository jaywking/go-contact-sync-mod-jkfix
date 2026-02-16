param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [ValidateSet("x86", "x64", "Any CPU")]
    [string]$Platform = "x86",
    [bool]$StopRunningApp = $true
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

if ($StopRunningApp) {
    $running = Get-Process | Where-Object { $_.ProcessName -eq "GOContactSync" -or $_.ProcessName -eq "Go Contact Sync Mod" }
    if ($running) {
        Write-Host "Stopping running GO Contact Sync Mod process(es)..."
        $running | Stop-Process -Force
        Start-Sleep -Seconds 1
    }
}

# Validate .NET Framework targeting pack before restore/build.
$netFxRefPath = "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8.1"
if (-not (Test-Path $netFxRefPath)) {
    throw @"
.NET Framework 4.8.1 targeting pack was not found.
Install '.NET Framework 4.8.1 SDK/Targeting Pack (Developer Pack)' and rerun.

Quick options:
1) Visual Studio Installer -> Modify Build Tools/VS -> Individual components -> search '4.8.1' -> install .NET Framework 4.8.1 targeting pack.
2) Download directly: https://aka.ms/msbuild/developerpacks
"@
}

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    throw "vswhere.exe not found. Install Visual Studio 2022 Build Tools or Visual Studio 2022."
}

$installPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
if (-not $installPath) {
    throw "No Visual Studio installation with MSBuild found."
}

$msbuild = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
    throw "MSBuild not found at: $msbuild"
}

$nugetExe = Join-Path $repoRoot ".nuget\NuGet.exe"
$nugetBootstrapUrl = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
if (-not (Test-Path $nugetExe)) {
    Write-Host "Downloading NuGet CLI..."
    Invoke-WebRequest -UseBasicParsing -Uri $nugetBootstrapUrl -OutFile $nugetExe
}
else {
    # Replace bundled legacy nuget.exe with a current build for modern TLS/protocol support.
    try {
        Write-Host "Refreshing NuGet CLI..."
        Invoke-WebRequest -UseBasicParsing -Uri $nugetBootstrapUrl -OutFile $nugetExe
    }
    catch {
        Write-Warning "Could not refresh NuGet CLI, continuing with existing .nuget\\NuGet.exe"
    }
}

$nugetConfig = Join-Path $repoRoot ".nuget\NuGet.local.config"
@'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
    <add key="nugetv2" value="https://www.nuget.org/api/v2/" />
  </packageSources>
</configuration>
'@ | Set-Content -Encoding UTF8 $nugetConfig

Write-Host "Restoring NuGet packages..."
& $nugetExe restore "$repoRoot\GoogleContactsSync.sln" -ConfigFile $nugetConfig -Source "https://api.nuget.org/v3/index.json;https://www.nuget.org/api/v2/" -NonInteractive -Verbosity detailed

Write-Host "Building GoogleContactsSync project..."
& $msbuild "$repoRoot\GoogleContactsSync\GoogleContactsSync.csproj" /t:Build /p:Configuration=$Configuration /p:Platform="$Platform" /p:RestorePackages=false /m

$candidateExePaths = @(
    (Join-Path $repoRoot "GoogleContactsSync\bin\$Configuration\GOContactSync.exe"),
    (Join-Path $repoRoot "GoogleContactsSync\bin\$Platform\$Configuration\GOContactSync.exe"),
    (Join-Path $repoRoot "GoogleContactsSync\bin\$Configuration\GoContactSyncMod.exe"),
    (Join-Path $repoRoot "GoogleContactsSync\bin\$Platform\$Configuration\GoContactSyncMod.exe")
)

$resolvedExe = $candidateExePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($resolvedExe) {
    Write-Host "Build complete: $resolvedExe"
}
else {
    Write-Warning "Build finished but EXE was not found in known output folders."
    Write-Warning ("Checked: " + ($candidateExePaths -join "; "))
}

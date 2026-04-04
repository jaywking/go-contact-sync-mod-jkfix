param(
    [Parameter(Mandatory = $true)]
    [string]$VersionDirectory,

    [Parameter(Mandatory = $true)]
    [string]$MsiPath,

    [Parameter(Mandatory = $true)]
    [string]$ZipPath,

    [string]$Version = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Version)) {
    $assemblyInfo = Join-Path $repoRoot "GoogleContactsSync\Properties\AssemblyInfo.cs"
    $match = Select-String -Path $assemblyInfo -Pattern 'AssemblyInformationalVersion\("([^"]+)"\)' | Select-Object -First 1
    if (-not $match -or $match.Matches.Count -eq 0) {
        throw "Could not determine AssemblyInformationalVersion from $assemblyInfo."
    }

    $Version = $match.Matches[0].Groups[1].Value
}

if (-not (Test-Path -LiteralPath $VersionDirectory)) {
    throw "Release version directory does not exist: $VersionDirectory"
}

if (-not (Test-Path -LiteralPath $MsiPath)) {
    throw "MSI artifact not found: $MsiPath"
}

if (-not (Test-Path -LiteralPath $ZipPath)) {
    throw "ZIP artifact not found: $ZipPath"
}

$cleanMsiPath = Join-Path $VersionDirectory ("SetupGCSM-{0}.msi" -f $Version)
$cleanZipPath = Join-Path $VersionDirectory ("GOContactSyncMod-{0}.zip" -f $Version)

Copy-Item -LiteralPath $MsiPath -Destination $cleanMsiPath -Force
Copy-Item -LiteralPath $ZipPath -Destination $cleanZipPath -Force

Write-Host "Prepared clean release assets:"
Write-Host "  MSI: $cleanMsiPath"
Write-Host "  ZIP: $cleanZipPath"

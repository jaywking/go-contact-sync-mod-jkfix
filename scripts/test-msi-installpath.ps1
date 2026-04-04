param(
    [Parameter(Mandatory = $true)]
    [string]$BaseMsiPath,

    [string]$UpgradeMsiPath = "",

    [string]$CustomInstallPath = "",

    [string]$LogDir = "",

    [switch]$AllowReplaceExistingInstall,

    [switch]$KeepInstalledVersion
)

$ErrorActionPreference = "Stop"

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Run this script from an elevated PowerShell session."
    }
}

function Normalize-PathValue {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return ""
    }

    return $PathValue.Trim().TrimEnd("\")
}

function Format-MsiDirectoryValue {
    param([string]$PathValue)

    $normalized = Normalize-PathValue $PathValue
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ""
    }

    return "$normalized\"
}

function Assert-InstallPathDriveExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue,

        [Parameter(Mandatory = $true)]
        [string]$ParameterName
    )

    $normalized = Normalize-PathValue $PathValue
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return
    }

    $driveRoot = [System.IO.Path]::GetPathRoot($normalized)
    if ([string]::IsNullOrWhiteSpace($driveRoot)) {
        throw "$ParameterName must include a rooted filesystem path. Received: $PathValue"
    }

    if (-not (Test-Path -LiteralPath $driveRoot)) {
        throw "$ParameterName uses a drive that does not exist on this machine: $driveRoot"
    }
}

function Assert-InstallPathParentWritable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue,

        [Parameter(Mandatory = $true)]
        [string]$ParameterName
    )

    $normalized = Normalize-PathValue $PathValue
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return
    }

    $probePath = $normalized
    while (-not [string]::IsNullOrWhiteSpace($probePath) -and -not (Test-Path -LiteralPath $probePath)) {
        $parent = Split-Path -Path $probePath -Parent
        if ($parent -eq $probePath) {
            break
        }

        $probePath = $parent
    }

    if ([string]::IsNullOrWhiteSpace($probePath) -or -not (Test-Path -LiteralPath $probePath)) {
        throw "$ParameterName could not resolve an existing parent directory for: $PathValue"
    }

    $testFile = Join-Path $probePath ([System.IO.Path]::GetRandomFileName())
    try {
        [System.IO.File]::WriteAllText($testFile, "write-test")
        Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
    }
    catch {
        throw "$ParameterName is not writable at or above '$probePath'. Choose a folder you can write to."
    }
}

function Get-DefaultInstallPath {
    $programFilesX86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
    if ($env:ProgramFiles -and -not [string]::IsNullOrWhiteSpace($programFilesX86)) {
        return (Join-Path $programFilesX86 "GO Contact Sync Mod")
    }

    return (Join-Path $env:ProgramFiles "GO Contact Sync Mod")
}

function Get-ProductEntries {
    $registryRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($root in $registryRoots) {
        if (-not (Test-Path $root)) {
            continue
        }

        foreach ($subKey in Get-ChildItem -Path $root) {
            $item = Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue
            if (-not $item -or [string]::IsNullOrWhiteSpace($item.DisplayName)) {
                continue
            }

            if ($item.DisplayName -ne "GO Contact Sync Mod") {
                continue
            }

            [pscustomobject]@{
                RegistryPath    = $subKey.PSPath
                ProductCode     = $subKey.PSChildName
                DisplayName     = $item.DisplayName
                DisplayVersion  = $item.DisplayVersion
                InstallLocation = $item.InstallLocation
                UninstallString = $item.UninstallString
                WindowsInstaller = $item.WindowsInstaller
                InstallDate     = $item.InstallDate
            }
        }
    }
}

function Get-RememberedInstallLocationRegistryPaths {
    @(
        "HKLM:\SOFTWARE\WOW6432Node\GOContactSyncModJKFix\Installer",
        "HKLM:\SOFTWARE\GOContactSyncModJKFix\Installer"
    )
}

function Get-RememberedInstallLocation {
    foreach ($registryPath in Get-RememberedInstallLocationRegistryPaths) {
        if (-not (Test-Path $registryPath)) {
            continue
        }

        $item = Get-ItemProperty -Path $registryPath -ErrorAction SilentlyContinue
        if ($item -and -not [string]::IsNullOrWhiteSpace($item.InstallLocation)) {
            return $item.InstallLocation
        }
    }

    return ""
}

function Write-DiagnosticSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScenarioName,

        [Parameter(Mandatory = $true)]
        [string]$LogDirPath
    )

    Write-Host ""
    Write-Host "Diagnostic snapshot for: $ScenarioName"

    $product = Get-InstalledProduct
    if ($product) {
        Write-Host "Installed product:"
        $product | Format-List | Out-String | Write-Host
    }
    else {
        Write-Host "Installed product: <none>"
    }

    $rememberedPath = Get-RememberedInstallLocation
    if ([string]::IsNullOrWhiteSpace($rememberedPath)) {
        Write-Host "Remembered install path: <none>"
    }
    else {
        Write-Host "Remembered install path: $rememberedPath"
    }

    $recentLogs = Get-ChildItem -Path $LogDirPath -Filter *.log -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 2

    foreach ($logFile in $recentLogs) {
        Write-Host ""
        Write-Host "Tail of $($logFile.Name):"
        Get-Content -Path $logFile.FullName -Tail 40 | Write-Host
    }
}

function Remove-RememberedInstallLocation {
    foreach ($registryPath in Get-RememberedInstallLocationRegistryPaths) {
        if (Test-Path $registryPath) {
            Remove-Item -LiteralPath $registryPath -Recurse -Force
        }
    }
}

function Get-InstalledProduct {
    $products = @(Get-ProductEntries)
    if ($products.Count -eq 0) {
        return $null
    }

    return $products |
        Sort-Object @{
            Expression = { $_.InstallDate }
            Descending = $true
        }, @{
            Expression = { $_.DisplayVersion }
            Descending = $true
        } |
        Select-Object -First 1
}

function Invoke-MsiInstall {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MsiPath,

        [string]$AdditionalArguments = "",

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    $arguments = "/i `"$MsiPath`" /qn /norestart /l*v `"$LogPath`" $AdditionalArguments"
    Write-Host "Installing MSI: $MsiPath"
    Write-Host "  Log: $LogPath"

    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru
    if ($process.ExitCode -notin @(0, 3010, 1641)) {
        throw "MSI install failed with exit code $($process.ExitCode). See $LogPath"
    }
}

function Invoke-MsiUninstall {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProductCode,

        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    $arguments = "/x $ProductCode /qn /norestart /l*v `"$LogPath`""
    Write-Host "Uninstalling product: $ProductCode"
    Write-Host "  Log: $LogPath"

    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru
    if ($process.ExitCode -notin @(0, 3010, 1605)) {
        throw "MSI uninstall failed with exit code $($process.ExitCode). See $LogPath"
    }
}

function Remove-TestInstallDirectory {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue) -or -not (Test-Path -LiteralPath $PathValue)) {
        return
    }

    $fullPath = [System.IO.Path]::GetFullPath($PathValue)
    $systemDriveRoot = [System.IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.Length -le $systemDriveRoot.Length + 3) {
        throw "Refusing to remove suspiciously short path: $fullPath"
    }

    Remove-Item -LiteralPath $fullPath -Recurse -Force
}

function Assert-InstalledState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedInstallLocation,

        [Parameter(Mandatory = $true)]
        [string]$ScenarioName
    )

    $product = Get-InstalledProduct
    if (-not $product) {
        throw "$ScenarioName failed: GO Contact Sync Mod is not installed."
    }

    $actualLocation = Normalize-PathValue $product.InstallLocation
    $expectedLocation = Normalize-PathValue $ExpectedInstallLocation
    if ($actualLocation -ne $expectedLocation) {
        throw "$ScenarioName failed: InstallLocation was '$actualLocation' instead of '$expectedLocation'."
    }

    $installedExe = Join-Path $expectedLocation "GOContactSync.exe"
    if (-not (Test-Path -LiteralPath $installedExe)) {
        throw "$ScenarioName failed: installed EXE not found at $installedExe"
    }

    [pscustomobject]@{
        Scenario       = $ScenarioName
        InstallLocation = $actualLocation
        DisplayVersion = $product.DisplayVersion
        ExePath        = $installedExe
    }
}

function Assert-RememberedInstallLocation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedInstallLocation,

        [Parameter(Mandatory = $true)]
        [string]$ScenarioName
    )

    $actualLocation = Normalize-PathValue (Get-RememberedInstallLocation)
    $expectedLocation = Normalize-PathValue $ExpectedInstallLocation

    if ($actualLocation -ne $expectedLocation) {
        throw "$ScenarioName failed: remembered InstallLocation was '$actualLocation' instead of '$expectedLocation'."
    }

    [pscustomobject]@{
        Scenario        = $ScenarioName
        InstallLocation = $actualLocation
        DisplayVersion  = "(remembered)"
        ExePath         = ""
    }
}

Assert-Administrator

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($LogDir)) {
    $LogDir = Join-Path $repoRoot "dist\installer-test-logs"
}

if ([string]::IsNullOrWhiteSpace($CustomInstallPath)) {
    $CustomInstallPath = Join-Path $env:SystemDrive "InstallerTest\GO Contact Sync Mod"
}

Assert-InstallPathDriveExists -PathValue $CustomInstallPath -ParameterName "CustomInstallPath"
Assert-InstallPathParentWritable -PathValue $CustomInstallPath -ParameterName "CustomInstallPath"

$BaseMsiPath = [System.IO.Path]::GetFullPath($BaseMsiPath)
if (-not (Test-Path -LiteralPath $BaseMsiPath)) {
    throw "BaseMsiPath not found: $BaseMsiPath"
}

if (-not [string]::IsNullOrWhiteSpace($UpgradeMsiPath)) {
    $UpgradeMsiPath = [System.IO.Path]::GetFullPath($UpgradeMsiPath)
    if (-not (Test-Path -LiteralPath $UpgradeMsiPath)) {
        throw "UpgradeMsiPath not found: $UpgradeMsiPath"
    }
}

$defaultInstallPath = Get-DefaultInstallPath
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$transcriptPath = Join-Path $LogDir ("script-run-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
Start-Transcript -Path $transcriptPath -Force | Out-Null

$preexistingProduct = Get-InstalledProduct
if ($preexistingProduct -and -not $AllowReplaceExistingInstall) {
    throw @"
GO Contact Sync Mod is already installed at:
$($preexistingProduct.InstallLocation)

This script uninstalls and reinstalls the product as part of the test flow.
Run it on a disposable machine or rerun with -AllowReplaceExistingInstall.
"@
}

$results = New-Object System.Collections.Generic.List[object]

try {
    if ($preexistingProduct) {
        Invoke-MsiUninstall -ProductCode $preexistingProduct.ProductCode -LogPath (Join-Path $LogDir "00-remove-existing.log")
    }

    Remove-RememberedInstallLocation
    Remove-TestInstallDirectory -PathValue $CustomInstallPath

    Invoke-MsiInstall -MsiPath $BaseMsiPath -LogPath (Join-Path $LogDir "01-fresh-default-install.log")
    $results.Add((Assert-InstalledState -ExpectedInstallLocation $defaultInstallPath -ScenarioName "Fresh install uses default path"))

    $installedProduct = Get-InstalledProduct
    if ($installedProduct) {
        Invoke-MsiUninstall -ProductCode $installedProduct.ProductCode -LogPath (Join-Path $LogDir "02-remove-default-install.log")
    }

    Remove-TestInstallDirectory -PathValue $CustomInstallPath

    $msiCustomInstallPath = Format-MsiDirectoryValue $CustomInstallPath
    Invoke-MsiInstall -MsiPath $BaseMsiPath -AdditionalArguments "INSTALLLOCATION=`"$msiCustomInstallPath`"" -LogPath (Join-Path $LogDir "03-custom-install.log")
    $results.Add((Assert-InstalledState -ExpectedInstallLocation $CustomInstallPath -ScenarioName "Custom install path is stored"))
    $results.Add((Assert-RememberedInstallLocation -ExpectedInstallLocation $CustomInstallPath -ScenarioName "Custom install path is remembered"))

    if (-not [string]::IsNullOrWhiteSpace($UpgradeMsiPath)) {
        Invoke-MsiInstall -MsiPath $UpgradeMsiPath -LogPath (Join-Path $LogDir "04-upgrade-install.log")
        $results.Add((Assert-InstalledState -ExpectedInstallLocation $CustomInstallPath -ScenarioName "Upgrade preserves custom install path"))
        $results.Add((Assert-RememberedInstallLocation -ExpectedInstallLocation $CustomInstallPath -ScenarioName "Upgrade keeps remembered custom path"))
    }
    else {
        Write-Warning "UpgradeMsiPath not provided; skipping upgrade-preserves-path scenario."
    }

    $installedProduct = Get-InstalledProduct
    if ($installedProduct) {
        Invoke-MsiUninstall -ProductCode $installedProduct.ProductCode -LogPath (Join-Path $LogDir "05-remove-custom-install.log")
    }

    Remove-TestInstallDirectory -PathValue $CustomInstallPath

    $reinstallMsiPath = if ([string]::IsNullOrWhiteSpace($UpgradeMsiPath)) { $BaseMsiPath } else { $UpgradeMsiPath }
    Invoke-MsiInstall -MsiPath $reinstallMsiPath -LogPath (Join-Path $LogDir "06-reinstall-remembered-path.log")
    $results.Add((Assert-InstalledState -ExpectedInstallLocation $CustomInstallPath -ScenarioName "Reinstall reuses remembered custom path"))
    $results.Add((Assert-RememberedInstallLocation -ExpectedInstallLocation $CustomInstallPath -ScenarioName "Reinstall preserves remembered custom path"))

    $installedProduct = Get-InstalledProduct
    if ($installedProduct) {
        Invoke-MsiUninstall -ProductCode $installedProduct.ProductCode -LogPath (Join-Path $LogDir "07-remove-reinstalled-product.log")
    }

    Remove-TestInstallDirectory -PathValue $CustomInstallPath
    Remove-RememberedInstallLocation

    $finalMsiPath = if ([string]::IsNullOrWhiteSpace($UpgradeMsiPath)) { $BaseMsiPath } else { $UpgradeMsiPath }
    Invoke-MsiInstall -MsiPath $finalMsiPath -LogPath (Join-Path $LogDir "08-fallback-default-install.log")
    $results.Add((Assert-InstalledState -ExpectedInstallLocation $defaultInstallPath -ScenarioName "No previous install falls back to default path"))
}
catch {
    Write-DiagnosticSnapshot -ScenarioName "Installer path regression failure" -LogDirPath $LogDir
    throw
}
finally {
    if (-not $KeepInstalledVersion) {
        $installedProduct = Get-InstalledProduct
        if ($installedProduct) {
            Invoke-MsiUninstall -ProductCode $installedProduct.ProductCode -LogPath (Join-Path $LogDir "99-cleanup.log")
        }

        Remove-RememberedInstallLocation
        Remove-TestInstallDirectory -PathValue $CustomInstallPath
    }

    Stop-Transcript | Out-Null
}

Write-Host ""
Write-Host "Installer path test summary"
$results | Format-Table -AutoSize
Write-Host ""
Write-Host "Logs written to: $LogDir"

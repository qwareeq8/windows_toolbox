[CmdletBinding()]
param(
    [switch]$AllowDirty,
    [switch]$SkipAppBuild
)

$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot
. "$PSScriptRoot\release-common.ps1"

if (-not $SkipAppBuild) {
    Write-Host "[build-installer] Building the application bundle."
    & "$PSScriptRoot\build-app.ps1" -AllowDirty:$AllowDirty
    if ($LASTEXITCODE -ne 0) {
        throw "build-app.ps1 failed."
    }
} else {
    Write-Host "[build-installer] Verifying the existing application bundle."
    & "$PSScriptRoot\verify-release.ps1" `
        -AllowDirty:$AllowDirty `
        -SkipInstaller `
        -SkipSmokeTest
    if ($LASTEXITCODE -ne 0) {
        throw "The existing application bundle failed static release verification."
    }
}

$appVersion = Get-VireloAppVersion -ProjectRoot $projectRoot
$isccPath = $env:ISCC_PATH
if (-not $isccPath) {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )
    $isccPath = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1
}
if (-not $isccPath) {
    throw "ISCC.exe was not found. Install Inno Setup 6 or set ISCC_PATH."
}

$installerDirectory = "installer\dist"
if (Test-Path -LiteralPath $installerDirectory -PathType Container) {
    Remove-Item -LiteralPath $installerDirectory -Recurse -Force
}

Write-Host "[build-installer] Building Virelo $appVersion with $isccPath."
& $isccPath "/DMyAppVersion=$appVersion" "installer\virelo.iss"
if ($LASTEXITCODE -ne 0) {
    throw "ISCC.exe failed with exit code $LASTEXITCODE."
}

$installerPath = "installer\dist\VireloSetup.exe"
if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    throw "The installer build did not create $installerPath."
}
$installerVersion = (Get-Item -LiteralPath $installerPath).VersionInfo.ProductVersion.Trim()
if ($installerVersion -ne $appVersion) {
    throw "The installer product version '$installerVersion' does not match '$appVersion'."
}

& "$PSScriptRoot\write-release-checksums.ps1"
if ($LASTEXITCODE -ne 0) {
    throw "Writing the release checksums failed."
}

Write-Host "[build-installer] OK: The installer and release checksums were created."

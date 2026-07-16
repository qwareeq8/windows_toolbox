[CmdletBinding()]
param(
    [switch]$AllowDirty,
    [switch]$SkipSmokeTest
)

$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot
. "$PSScriptRoot\release-common.ps1"

$venvPython = ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
    throw "The virtual environment is missing. Run scripts\bootstrap.ps1 first."
}
if (-not (Test-Path -LiteralPath "Virelo.spec" -PathType Leaf)) {
    throw "Virelo.spec is missing from the project root."
}

$gitCommit = (& git rev-parse HEAD 2>&1 | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $gitCommit -notmatch '^[0-9a-f]{40}$') {
    throw "The source commit could not be resolved. Build from a Git checkout."
}
$gitStatus = @(& git status --porcelain --untracked-files=normal)
if ($LASTEXITCODE -ne 0) {
    throw "The current Git working-tree status could not be read."
}
$sourceDirty = $gitStatus.Count -gt 0
if ($sourceDirty -and -not $AllowDirty) {
    throw "The working tree is not clean. Commit the intended source or pass -AllowDirty for a non-release build."
}

& $venvPython -m pip check
if ($LASTEXITCODE -ne 0) {
    throw "The Python environment has inconsistent requirements."
}
$pythonRuntimeJson = (& $venvPython -c @'
import json
import platform
import struct

print(json.dumps({
    "machine": platform.machine(),
    "pointerBits": struct.calcsize("P") * 8,
    "version": platform.python_version(),
}))
'@ 2>&1 | Out-String).Trim()
if ($LASTEXITCODE -ne 0) {
    throw "The Python runtime architecture could not be determined."
}
$pythonRuntime = $pythonRuntimeJson | ConvertFrom-Json
$pythonVersion = [version]$pythonRuntime.version
if ($pythonVersion.Major -ne 3 -or $pythonVersion.Minor -lt 12 -or $pythonVersion.Minor -gt 14) {
    throw "Python 3.12 to 3.14 is required, but $($pythonRuntime.version) is active."
}
if ($pythonRuntime.pointerBits -ne 64) {
    throw "A 64-bit Python environment is required, but the virtual environment is $($pythonRuntime.pointerBits)-bit."
}

Write-Host "[build-app] Running Python static checks and tests."
foreach ($arguments in @(
    @("-m", "ruff", "check", "."),
    @("-m", "ruff", "format", "--check", "."),
    @("-m", "mypy", "."),
    @("-m", "pytest", "-q")
)) {
    & $venvPython @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Python verification failed: $($arguments -join ' ')."
    }
}

Write-Host "[build-app] Building and checking the frontend."
& "$PSScriptRoot\build-frontend.ps1"
if ($LASTEXITCODE -ne 0) {
    throw "build-frontend.ps1 failed."
}

Write-Host "[build-app] Running PyInstaller."
& $venvPython -m PyInstaller --clean --noconfirm "Virelo.spec"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE."
}

$bundleRoot = "dist\Virelo"
$exePath = Join-Path $bundleRoot "Virelo.exe"
if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
    throw "The application build did not create $exePath."
}

$releaseInputs = Get-VireloReleaseInputPaths -ProjectRoot $projectRoot
$frontendRoot = (Resolve-Path "frontend\dist").Path
$frontendPaths = Get-ChildItem -LiteralPath $frontendRoot -Recurse -File -Force |
    ForEach-Object { Get-VireloRelativePath -Root $frontendRoot -Path $_.FullName }
$manifest = [ordered]@{
    schemaVersion = 2
    appName = "Virelo"
    appVersion = Get-VireloAppVersion -ProjectRoot $projectRoot
    sourceCommit = $gitCommit
    sourceDirty = $sourceDirty
    sourceStatus = @($gitStatus)
    builtAtUtc = [DateTime]::UtcNow.ToString("o")
    toolchain = [ordered]@{
        python = (& $venvPython --version 2>&1 | Out-String).Trim()
        pythonMachine = $pythonRuntime.machine
        pythonPointerBits = $pythonRuntime.pointerBits
        pip = (& $venvPython -m pip --version 2>&1 | Out-String).Trim()
        pyinstaller = (& $venvPython -m PyInstaller --version 2>&1 | Out-String).Trim()
        node = (& node --version 2>&1 | Out-String).Trim()
        npm = (& npm --version 2>&1 | Out-String).Trim()
        powershell = $PSVersionTable.PSVersion.ToString()
    }
    platform = [ordered]@{
        osDescription = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
        osArchitecture = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
        processArchitecture = [System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString()
        windowsVersion = [Environment]::OSVersion.Version.ToString()
        is64BitOperatingSystem = [Environment]::Is64BitOperatingSystem
        is64BitProcess = [Environment]::Is64BitProcess
    }
    inputs = Get-VireloFileHashMap -ProjectRoot $projectRoot -RelativePaths $releaseInputs
    frontend = Get-VireloFileHashMap -ProjectRoot $frontendRoot -RelativePaths $frontendPaths
}
$manifestPath = Join-Path $bundleRoot ".release.json"
$manifestJson = $manifest | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($manifestPath),
    $manifestJson + [Environment]::NewLine,
    (New-Object System.Text.UTF8Encoding($false))
)

& "$PSScriptRoot\write-bundle-checksums.ps1"
if ($LASTEXITCODE -ne 0) {
    throw "Writing the bundle checksums failed."
}

$productVersion = (Get-Item -LiteralPath $exePath).VersionInfo.ProductVersion
if ($productVersion -ne $manifest.appVersion) {
    throw "The executable product version '$productVersion' does not match '$($manifest.appVersion)'."
}

if (-not $SkipSmokeTest) {
    Write-Host "[build-app] Running the packaged smoke test."
    & ".\$exePath" --smoke-test
    if ($LASTEXITCODE -ne 0) {
        throw "The packaged smoke test failed with exit code $LASTEXITCODE."
    }
}

Write-Host "[build-app] OK: The application bundle and provenance records were created."

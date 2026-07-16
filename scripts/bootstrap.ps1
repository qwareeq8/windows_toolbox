[CmdletBinding()]
param(
    [switch]$Recreate
)

$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot
$constraints = "requirements\build-constraints.txt"

if (-not (Test-Path -LiteralPath $constraints -PathType Leaf)) {
    throw "The pinned build constraints file is missing: $constraints."
}

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    throw "Python was not found. Install Python 3.12 to 3.14 from https://www.python.org/."
}

$pythonVersionOutput = (& python --version 2>&1 | Out-String).Trim()
if ($LASTEXITCODE -ne 0) {
    throw "The python --version command failed."
}
$versionMatch = [regex]::Match($pythonVersionOutput, '^Python (\d+)\.(\d+)')
if (-not $versionMatch.Success) {
    throw "Could not parse the Python version from '$pythonVersionOutput'."
}
$major = [int]$versionMatch.Groups[1].Value
$minor = [int]$versionMatch.Groups[2].Value
if ($major -ne 3 -or $minor -lt 12 -or $minor -gt 14) {
    throw "Python 3.12 to 3.14 is required, but $pythonVersionOutput is active."
}
$pythonPointerBits = (& python -c "import struct; print(struct.calcsize('P') * 8)" 2>&1 |
    Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or $pythonPointerBits -notmatch '^\d+$') {
    throw "The active Python architecture could not be determined."
}
if ([int]$pythonPointerBits -ne 64) {
    throw "A 64-bit Python interpreter is required, but the active interpreter is $pythonPointerBits-bit."
}
Write-Host "[bootstrap] $pythonVersionOutput ($pythonPointerBits-bit)."

if ($Recreate -and (Test-Path -LiteralPath ".venv" -PathType Container)) {
    Write-Host "[bootstrap] Recreating the virtual environment."
    Remove-Item -LiteralPath ".venv" -Recurse -Force
}
if (-not (Test-Path -LiteralPath ".venv\Scripts\python.exe" -PathType Leaf)) {
    Write-Host "[bootstrap] Creating the virtual environment."
    & python -m venv .venv
    if ($LASTEXITCODE -ne 0) {
        throw "python -m venv failed."
    }
}

$venvPython = ".venv\Scripts\python.exe"
Write-Host "[bootstrap] Installing the pinned packaging tools."
& $venvPython -m pip install --upgrade --constraint $constraints pip setuptools
if ($LASTEXITCODE -ne 0) {
    throw "Installing the pinned packaging tools failed."
}

Write-Host "[bootstrap] Installing the constrained application environment."
& $venvPython -m pip install --constraint $constraints --editable ".[dev,build]"
if ($LASTEXITCODE -ne 0) {
    throw "Installing the constrained Python environment failed."
}
& $venvPython -m pip check
if ($LASTEXITCODE -ne 0) {
    throw "The installed Python environment has inconsistent requirements."
}

$node = Get-Command node -ErrorAction SilentlyContinue
$npm = Get-Command npm -ErrorAction SilentlyContinue
if (-not $node -or -not $npm) {
    throw "Node.js and npm are required. Install a supported Node.js version before continuing."
}
$nodeVersion = (& node --version 2>&1 | Out-String).Trim()
if ($nodeVersion -notmatch '^v(22|23|24)\.') {
    throw "Node.js 22 to 24 is required, but $nodeVersion is active."
}

Write-Host "[bootstrap] Installing the locked frontend dependencies."
Push-Location frontend
try {
    & npm ci
    if ($LASTEXITCODE -ne 0) {
        throw "npm ci failed."
    }
} finally {
    Pop-Location
}

Write-Host "[bootstrap] OK: The constrained environment is ready."

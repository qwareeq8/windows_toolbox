$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

if (-not (Test-Path "Virelo.spec")) {
    throw "PyInstaller spec not found: Virelo.spec"
}

Write-Host "Building EXE with PyInstaller..."
& "$projectRoot\.venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm "Virelo.spec"

$IsccPath = $env:ISCC_PATH
if (-not $IsccPath) {
    $candidates = @(
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    $IsccPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

if (-not $IsccPath) {
    throw "ISCC.exe not found. Install Inno Setup 6 or set ISCC_PATH."
}

Write-Host "Building installer with Inno Setup..."
& $IsccPath "installer\virelo.iss"

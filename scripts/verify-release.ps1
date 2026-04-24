$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

Write-Host "[verify-release] Checking build artifacts..."

$errors = @()

# --- Read expected version ---
$versionMatch = Select-String -Path "virelo\app\config.py" -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $versionMatch) {
    $errors += "APP_VERSION not found in virelo/app/config.py"
} else {
    $AppVersion = $versionMatch.Matches.Groups[1].Value
    Write-Host "[verify-release] Expected version: $AppVersion"
}

# --- Check frontend build ---
if (-not (Test-Path "frontend\dist\index.html")) {
    $errors += "Missing: frontend\dist\index.html"
} else {
    Write-Host "[verify-release] OK: frontend/dist/index.html"
}

# --- Check PyInstaller output ---
if (-not (Test-Path "dist\Virelo\Virelo.exe")) {
    $errors += "Missing: dist\Virelo\Virelo.exe"
} else {
    Write-Host "[verify-release] OK: dist/Virelo/Virelo.exe"
}

# --- Check installer output ---
if (-not (Test-Path "installer\dist\VireloSetup.exe")) {
    $errors += "Missing: installer\dist\VireloSetup.exe"
} else {
    Write-Host "[verify-release] OK: installer/dist/VireloSetup.exe"
}

# --- Check Virelo.spec exists ---
if (-not (Test-Path "Virelo.spec")) {
    $errors += "Missing: Virelo.spec"
} else {
    Write-Host "[verify-release] OK: Virelo.spec"
}

# --- Check no stale name ---
if (Test-Path "Windows Toolbox.spec") {
    $errors += "Stale file found: Windows Toolbox.spec (should have been renamed to Virelo.spec)"
}

# --- Report ---
if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "[verify-release] FAILED: $($errors.Count) issue(s) found:" -ForegroundColor Red
    foreach ($err in $errors) {
        Write-Host "  - $err" -ForegroundColor Red
    }
    throw "Release verification failed"
}

Write-Host ""
Write-Host "[verify-release] PASSED: All artifacts verified" -ForegroundColor Green

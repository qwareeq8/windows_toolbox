$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

$artifactTargets = @(
    "build",
    "dist",
    "frontend\dist",
    "installer\dist"
)

$cacheTargets = @(
    ".pytest_cache",
    ".ruff_cache",
    "htmlcov",
    "virelo.egg-info"
)

foreach ($target in $artifactTargets) {
    if (Test-Path $target) {
        Write-Host "[clean] Removing $target"
        Remove-Item -Recurse -Force $target
    }
}

foreach ($target in $cacheTargets) {
    if (-not (Test-Path $target)) {
        continue
    }
    Write-Host "[clean] Removing optional cache $target"
    try {
        Remove-Item -Recurse -Force $target -ErrorAction Stop
    } catch {
        Write-Warning "Could not remove optional cache '$target': $($_.Exception.Message)"
    }
}

# Clean *.spec.bak files
Get-ChildItem -Filter "*.spec.bak" -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "[clean] Removing $($_.Name)"
    Remove-Item -Force $_.FullName
}

# Recursive __pycache__ and *.pyc removal (D-03)
Get-ChildItem -Recurse -Filter "__pycache__" -Directory -ErrorAction SilentlyContinue |
    ForEach-Object {
        Write-Host "[clean] Removing $($_.FullName)"
        Remove-Item -Recurse -Force $_.FullName
    }

Get-ChildItem -Recurse -Filter "*.pyc" -ErrorAction SilentlyContinue |
    ForEach-Object {
        Remove-Item -Force $_.FullName
    }

if (Test-Path -LiteralPath ".coverage" -PathType Leaf) {
    Write-Host "[clean] Removing .coverage"
    Remove-Item -LiteralPath ".coverage" -Force
}

Write-Host "[clean] OK: Build artifacts were removed."

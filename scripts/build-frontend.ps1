$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot
. "$PSScriptRoot\release-common.ps1"

$node = Get-Command node -ErrorAction SilentlyContinue
$npm = Get-Command npm -ErrorAction SilentlyContinue
if (-not $node -or -not $npm) {
    throw "Node.js and npm are required to build the frontend."
}
$nodeVersion = (& node --version 2>&1 | Out-String).Trim()
if ($nodeVersion -notmatch '^v(22|23|24)\.') {
    throw "Node.js 22 to 24 is required, but $nodeVersion is active."
}

$appVersion = Get-VireloAppVersion -ProjectRoot $projectRoot
Write-Host "[build-frontend] Node $nodeVersion, npm $(& npm --version), version $appVersion."

$previousVersion = $env:VITE_APP_VERSION
$hadPreviousVersion = Test-Path Env:VITE_APP_VERSION
$env:VITE_APP_VERSION = $appVersion
Push-Location frontend
try {
    Write-Host "[build-frontend] Installing locked dependencies."
    & npm ci
    if ($LASTEXITCODE -ne 0) {
        throw "npm ci failed."
    }

    Write-Host "[build-frontend] Running static checks and tests."
    foreach ($script in @("lint", "format:check", "test", "build")) {
        & npm run $script
        if ($LASTEXITCODE -ne 0) {
            throw "npm run $script failed."
        }
    }
} finally {
    Pop-Location
    if ($hadPreviousVersion) {
        $env:VITE_APP_VERSION = $previousVersion
    } else {
        Remove-Item Env:VITE_APP_VERSION -ErrorAction SilentlyContinue
    }
}

$indexPath = "frontend\dist\index.html"
if (-not (Test-Path -LiteralPath $indexPath -PathType Leaf)) {
    throw "The frontend build did not create $indexPath."
}
$assetPaths = [regex]::Matches(
    [System.IO.File]::ReadAllText((Resolve-Path $indexPath)),
    '(?:src|href)="\.\/([^"#?]+)"'
) | ForEach-Object { $_.Groups[1].Value }
if (-not $assetPaths) {
    throw "The frontend index does not reference any built assets."
}
foreach ($assetPath in $assetPaths) {
    if (-not (Test-Path -LiteralPath (Join-Path "frontend\dist" $assetPath) -PathType Leaf)) {
        throw "The frontend index references a missing asset: $assetPath."
    }
}
$javascript = Get-ChildItem -LiteralPath "frontend\dist" -Recurse -File -Filter "*.js"
$versionFound = $javascript | Select-String -SimpleMatch $appVersion -Quiet
if (-not $versionFound) {
    throw "The release version was not embedded in the frontend JavaScript."
}

Write-Host "[build-frontend] OK: The versioned frontend passed checks and was built."

$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot
. "$PSScriptRoot\release-common.ps1"

$bundleRoot = "dist\Virelo"
$manifestPath = Join-Path $bundleRoot ".release.json"
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "The release manifest is missing: $manifestPath."
}
$checksumPath = Join-Path $bundleRoot "bundle-files.sha256"
Write-VireloSha256Manifest -Root $bundleRoot -OutputPath $checksumPath
Write-Host "[write-bundle-checksums] OK: $checksumPath was updated."

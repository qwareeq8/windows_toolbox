[CmdletBinding()]
param(
    [string]$OutputRoot
)

$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot

if (-not $OutputRoot) {
    $OutputRoot = "installer\dist"
}
if (-not (Test-Path -LiteralPath $OutputRoot -PathType Container)) {
    throw "The release output directory is missing: $OutputRoot."
}
$artifactNames = @("VireloSetup.exe")
$lines = foreach ($artifactName in $artifactNames) {
    $artifactPath = Join-Path $OutputRoot $artifactName
    if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
        throw "A release artifact is missing: $artifactPath."
    }
    $hash = (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash
    "$hash *$artifactName"
}
$outputPath = Join-Path $OutputRoot "CHECKSUMS.sha256"
[System.IO.File]::WriteAllLines(
    [System.IO.Path]::GetFullPath($outputPath),
    $lines,
    (New-Object System.Text.UTF8Encoding($false))
)
Write-Host "[write-release-checksums] OK: $outputPath was updated."

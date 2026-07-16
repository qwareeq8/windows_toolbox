[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot
. "$PSScriptRoot\release-common.ps1"

$systemTempRoot = [System.IO.Path]::GetFullPath(
    [System.IO.Path]::GetTempPath()
).TrimEnd("\", "/")
$tempRoot = Join-Path $systemTempRoot "Virelo-release-test-$([guid]::NewGuid())"
$resolvedTempRoot = [System.IO.Path]::GetFullPath($tempRoot)
$expectedPrefix = $systemTempRoot + [System.IO.Path]::DirectorySeparatorChar
if (-not $resolvedTempRoot.StartsWith(
        $expectedPrefix,
        [System.StringComparison]::OrdinalIgnoreCase
    )) {
    throw "The temporary release-test path is outside the system temporary directory."
}

$null = New-Item -ItemType Directory -Path $resolvedTempRoot
$payloadPath = Join-Path $resolvedTempRoot "payload.txt"
$manifestPath = Join-Path $resolvedTempRoot "bundle-files.sha256"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

try {
    [System.IO.File]::WriteAllText($payloadPath, "original", $utf8NoBom)
    Write-VireloSha256Manifest -Root $resolvedTempRoot -OutputPath $manifestPath

    $failures = @(Test-VireloSha256Manifest -Root $resolvedTempRoot -ManifestPath $manifestPath)
    if ($failures.Count -ne 0) {
        throw "A freshly written checksum manifest failed verification: $($failures -join '; ')."
    }

    [System.IO.File]::WriteAllText($payloadPath, "tampered", $utf8NoBom)
    $failures = @(Test-VireloSha256Manifest -Root $resolvedTempRoot -ManifestPath $manifestPath)
    if ($failures -notcontains "Checksum mismatch: payload.txt") {
        throw "Checksum verification did not reject a modified payload."
    }

    $fakeHash = "0" * 64
    [System.IO.File]::WriteAllText(
        $manifestPath,
        "$fakeHash *../outside.txt$([Environment]::NewLine)",
        $utf8NoBom
    )
    $failures = @(Test-VireloSha256Manifest -Root $resolvedTempRoot -ManifestPath $manifestPath)
    if (-not ($failures -like "Checksum path escapes the bundle root: *")) {
        throw "Checksum verification did not reject a path outside the bundle root."
    }

    [System.IO.File]::WriteAllText(
        $manifestPath,
        "$fakeHash *payload.txt$([Environment]::NewLine)$fakeHash *payload.txt$([Environment]::NewLine)",
        $utf8NoBom
    )
    $failures = @(Test-VireloSha256Manifest -Root $resolvedTempRoot -ManifestPath $manifestPath)
    if ($failures -notcontains "Duplicate checksum path: payload.txt") {
        throw "Checksum verification did not reject a duplicate manifest path."
    }

    $releaseInputs = @(Get-VireloReleaseInputPaths -ProjectRoot $projectRoot)
    if ($releaseInputs -notcontains "tests/unit/test_release_metadata.py") {
        throw "Python tests are missing from the release input inventory."
    }

    $releaseOutput = Join-Path $resolvedTempRoot "release-output"
    $null = New-Item -ItemType Directory -Path $releaseOutput
    [System.IO.File]::WriteAllText(
        (Join-Path $releaseOutput "VireloSetup.exe"),
        "installer fixture",
        $utf8NoBom
    )
    & "$PSScriptRoot\write-release-checksums.ps1" -OutputRoot $releaseOutput
    $releaseChecksumLines = [System.IO.File]::ReadAllLines(
        (Join-Path $releaseOutput "CHECKSUMS.sha256")
    )
    if (
        $releaseChecksumLines.Count -ne 1 -or
        $releaseChecksumLines[0] -notmatch '^[A-F0-9]{64} \*VireloSetup[.]exe$'
    ) {
        throw "The public release checksum manifest does not use a portable artifact name."
    }
} finally {
    if (Test-Path -LiteralPath $resolvedTempRoot -PathType Container) {
        Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
    }
}

Write-Host "[test-release-tools] PASSED: Release helper behavior is correct."

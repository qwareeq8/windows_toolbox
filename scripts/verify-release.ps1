[CmdletBinding()]
param(
    [switch]$AllowDirty,
    [switch]$RequireSignature,
    [switch]$SkipInstaller,
    [switch]$SkipSmokeTest
)

$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $projectRoot
. "$PSScriptRoot\release-common.ps1"

$errors = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]

function Add-ReleaseError([string]$Message) {
    $null = $errors.Add($Message)
}

function Add-ReleaseWarning([string]$Message) {
    $null = $warnings.Add($Message)
}

function Test-ArtifactSignature([string]$Path) {
    $signatureCommand = Get-Command Get-AuthenticodeSignature -ErrorAction SilentlyContinue
    if (-not $signatureCommand) {
        if ($RequireSignature) {
            Add-ReleaseError "Authenticode verification is required, but Get-AuthenticodeSignature is unavailable."
        } else {
            Add-ReleaseWarning "Authenticode status was not checked because Get-AuthenticodeSignature is unavailable."
        }
        return
    }
    $signature = Get-AuthenticodeSignature -FilePath $Path
    Write-Host "[verify-release] Signature: $Path = $($signature.Status)."
    if ($signature.Status -ne "Valid") {
        if ($RequireSignature) {
            Add-ReleaseError "The required Authenticode signature is not valid for ${Path}: $($signature.Status)."
        } else {
            Add-ReleaseWarning "$Path is not Authenticode-signed with a valid signature."
        }
    }
}

Write-Host "[verify-release] Verifying static release evidence."
try {
    $appVersion = Get-VireloAppVersion -ProjectRoot $projectRoot
} catch {
    Add-ReleaseError $_.Exception.Message
    $appVersion = $null
}

$currentGitStatus = @(& git status --porcelain --untracked-files=normal 2>&1)
if ($LASTEXITCODE -ne 0) {
    Add-ReleaseError "The current Git working-tree status could not be read."
} elseif ($currentGitStatus.Count -gt 0 -and -not $AllowDirty) {
    Add-ReleaseError "The current working tree is not clean. Commit the intended source or use -AllowDirty for diagnostic verification."
}

try {
    $metadataScript = @'
const fs = require("node:fs");
const packageJson = JSON.parse(fs.readFileSync("frontend/package.json", "utf8"));
const packageLock = JSON.parse(fs.readFileSync("frontend/package-lock.json", "utf8"));
process.stdout.write(JSON.stringify({
  packageVersion: packageJson.version,
  lockVersion: packageLock.version,
  lockRootVersion: packageLock.packages[""].version,
}));
'@
    $frontendVersions = (& node -e $metadataScript 2>&1 | Out-String).Trim() |
        ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) {
        throw "Node.js could not read the frontend metadata."
    }
    $versionRecords = [ordered]@{
        "frontend/package.json" = $frontendVersions.packageVersion
        "frontend/package-lock.json" = $frontendVersions.lockVersion
        "frontend/package-lock.json root package" = $frontendVersions.lockRootVersion
    }
    foreach ($entry in $versionRecords.GetEnumerator()) {
        if ($entry.Value -ne $appVersion) {
            Add-ReleaseError "$($entry.Key) reports version '$($entry.Value)' instead of '$appVersion'."
        }
    }
} catch {
    Add-ReleaseError "Frontend version metadata could not be read: $($_.Exception.Message)"
}

$venvPython = ".venv\Scripts\python.exe"
if (Test-Path -LiteralPath $venvPython -PathType Leaf) {
    $installedVersion = (& $venvPython -c "import importlib.metadata; print(importlib.metadata.version('virelo'))" 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
        Add-ReleaseError "The installed Virelo package version could not be read."
    } elseif ($installedVersion -ne $appVersion) {
        Add-ReleaseError "The installed package version '$installedVersion' does not match '$appVersion'."
    }
    & $venvPython -m pip check
    if ($LASTEXITCODE -ne 0) {
        Add-ReleaseError "The Python environment has inconsistent requirements."
    }
} else {
    Add-ReleaseWarning "The virtual environment is absent, so installed Python metadata was not checked."
}

$bundleRoot = "dist\Virelo"
$exePath = Join-Path $bundleRoot "Virelo.exe"
$manifestPath = Join-Path $bundleRoot ".release.json"
$bundleChecksums = Join-Path $bundleRoot "bundle-files.sha256"
foreach ($requiredPath in @($exePath, $manifestPath, $bundleChecksums)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        Add-ReleaseError "A required bundle artifact is missing: $requiredPath."
    }
}

$manifest = $null
if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        if ($manifest.schemaVersion -ne 2) {
            Add-ReleaseError "The release manifest schema '$($manifest.schemaVersion)' is unsupported."
        }
        if ($manifest.appVersion -ne $appVersion) {
            Add-ReleaseError "The release manifest version '$($manifest.appVersion)' does not match '$appVersion'."
        }
        if ($manifest.sourceDirty -and -not $AllowDirty) {
            Add-ReleaseError "The application bundle was built from a dirty working tree."
        }
        if ($manifest.toolchain.pythonPointerBits -ne 64) {
            Add-ReleaseError "The bundle was not built with a 64-bit Python interpreter."
        }
        if (
            $manifest.platform.osArchitecture -ne "X64" -or
            $manifest.platform.processArchitecture -ne "X64" -or
            -not $manifest.platform.is64BitOperatingSystem -or
            -not $manifest.platform.is64BitProcess
        ) {
            Add-ReleaseError "The recorded build platform is not a 64-bit x64 Windows environment."
        }

        $currentCommit = (& git rev-parse HEAD 2>&1 | Out-String).Trim()
        if ($LASTEXITCODE -ne 0 -or $currentCommit -ne $manifest.sourceCommit) {
            Add-ReleaseError "The bundle source commit '$($manifest.sourceCommit)' does not match the current commit '$currentCommit'."
        }

        $currentInputPaths = Get-VireloReleaseInputPaths -ProjectRoot $projectRoot
        $currentInputSet = @($currentInputPaths | Sort-Object)
        $manifestInputSet = @($manifest.inputs.PSObject.Properties.Name | Sort-Object)
        $pathDifference = @(Compare-Object $manifestInputSet $currentInputSet)
        if ($pathDifference.Count -gt 0) {
            Add-ReleaseError "The release input inventory has changed since the bundle was built."
        }
        foreach ($property in $manifest.inputs.PSObject.Properties) {
            $inputPath = Join-Path $projectRoot $property.Name
            if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
                Add-ReleaseError "A recorded release input is missing: $($property.Name)."
                continue
            }
            $actualHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
            if ($actualHash -ne $property.Value) {
                Add-ReleaseError "A release input changed after the bundle was built: $($property.Name)."
            }
        }

        $currentFrontendRoot = "frontend\dist"
        $bundledFrontendRoot = Join-Path $bundleRoot "_internal\frontend\dist"
        foreach ($frontendRoot in @($currentFrontendRoot, $bundledFrontendRoot)) {
            if (-not (Test-Path -LiteralPath $frontendRoot -PathType Container)) {
                Add-ReleaseError "A frontend output directory is missing: $frontendRoot."
            }
        }
        if (
            (Test-Path -LiteralPath $currentFrontendRoot -PathType Container) -and
            (Test-Path -LiteralPath $bundledFrontendRoot -PathType Container)
        ) {
            $currentFrontendPaths = @(Get-ChildItem -LiteralPath $currentFrontendRoot -Recurse -File |
                ForEach-Object {
                    Get-VireloRelativePath -Root $currentFrontendRoot -Path $_.FullName
                } | Sort-Object)
            $bundledFrontendPaths = @(Get-ChildItem -LiteralPath $bundledFrontendRoot -Recurse -File |
                ForEach-Object {
                    Get-VireloRelativePath -Root $bundledFrontendRoot -Path $_.FullName
                } | Sort-Object)
            $manifestFrontendPaths = @($manifest.frontend.PSObject.Properties.Name | Sort-Object)
            if (@(Compare-Object $manifestFrontendPaths $currentFrontendPaths).Count -gt 0) {
                Add-ReleaseError "The current frontend output does not match the recorded frontend inventory."
            }
            if (@(Compare-Object $manifestFrontendPaths $bundledFrontendPaths).Count -gt 0) {
                Add-ReleaseError "The bundled frontend does not match the recorded frontend inventory."
            }
            foreach ($property in $manifest.frontend.PSObject.Properties) {
                $currentPath = Join-Path $currentFrontendRoot $property.Name
                $bundledPath = Join-Path $bundledFrontendRoot $property.Name
                foreach ($candidate in @($currentPath, $bundledPath)) {
                    if ((Test-Path -LiteralPath $candidate -PathType Leaf) -and
                        (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash -ne $property.Value) {
                        Add-ReleaseError "Frontend hash mismatch: $candidate."
                    }
                }
            }
        }
    } catch {
        Add-ReleaseError "The release manifest could not be validated: $($_.Exception.Message)"
    }
}

if (Test-Path -LiteralPath $bundleChecksums -PathType Leaf) {
    $bundleFailures = @(Test-VireloSha256Manifest -Root $bundleRoot -ManifestPath $bundleChecksums)
    foreach ($failure in $bundleFailures) {
        Add-ReleaseError $failure
    }
}

if (Test-Path -LiteralPath $exePath -PathType Leaf) {
    $exeVersion = (Get-Item -LiteralPath $exePath).VersionInfo.ProductVersion
    if ($exeVersion -ne $appVersion) {
        Add-ReleaseError "The executable product version '$exeVersion' does not match '$appVersion'."
    }
    Test-ArtifactSignature -Path $exePath
    if (-not $SkipSmokeTest) {
        & ".\$exePath" --smoke-test
        if ($LASTEXITCODE -ne 0) {
            Add-ReleaseError "The packaged smoke test failed with exit code $LASTEXITCODE."
        }
    }
}

if (-not $SkipInstaller) {
    $installerPath = "installer\dist\VireloSetup.exe"
    $releaseChecksums = "installer\dist\CHECKSUMS.sha256"
    foreach ($requiredPath in @($installerPath, $releaseChecksums)) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            Add-ReleaseError "A required installer artifact is missing: $requiredPath."
        }
    }
    if (Test-Path -LiteralPath $installerPath -PathType Leaf) {
        $installerVersion = (Get-Item -LiteralPath $installerPath).VersionInfo.ProductVersion.Trim()
        if ($installerVersion -ne $appVersion) {
            Add-ReleaseError "The installer product version '$installerVersion' does not match '$appVersion'."
        }
        Test-ArtifactSignature -Path $installerPath
    }
    if (Test-Path -LiteralPath $releaseChecksums -PathType Leaf) {
        $releaseRoot = Split-Path -Parent $releaseChecksums
        $expectedReleaseArtifacts = @("VireloSetup.exe")
        $expectedReleaseSet = New-Object 'System.Collections.Generic.HashSet[string]' (
            [System.StringComparer]::OrdinalIgnoreCase
        )
        $seenReleaseSet = New-Object 'System.Collections.Generic.HashSet[string]' (
            [System.StringComparer]::OrdinalIgnoreCase
        )
        foreach ($expectedArtifact in $expectedReleaseArtifacts) {
            $null = $expectedReleaseSet.Add($expectedArtifact)
        }
        foreach ($line in [System.IO.File]::ReadAllLines((Resolve-Path $releaseChecksums))) {
            if (-not $line.Trim()) {
                continue
            }
            $match = [regex]::Match($line, '^([A-Fa-f0-9]{64}) \*(.+)$')
            if (-not $match.Success) {
                Add-ReleaseError "Malformed release checksum line: $line"
                continue
            }
            $artifactName = $match.Groups[2].Value.Replace("/", "\")
            if (-not $expectedReleaseSet.Contains($artifactName)) {
                Add-ReleaseError "Unexpected release checksum path: $artifactName."
                continue
            }
            if (-not $seenReleaseSet.Add($artifactName)) {
                Add-ReleaseError "Duplicate release checksum path: $artifactName."
                continue
            }
            $artifactPath = Join-Path $releaseRoot $artifactName
            if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
                Add-ReleaseError "A checksummed release artifact is missing: $artifactName."
                continue
            }
            $actualHash = (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash
            if ($actualHash -ne $match.Groups[1].Value.ToUpperInvariant()) {
                Add-ReleaseError "Release checksum mismatch: $artifactName."
            }
        }
        foreach ($expectedArtifact in $expectedReleaseArtifacts) {
            if (-not $seenReleaseSet.Contains($expectedArtifact)) {
                Add-ReleaseError "A release artifact is missing from CHECKSUMS.sha256: $expectedArtifact."
            }
        }
    }
}

Add-ReleaseWarning "Installation, upgrade, recovery, and uninstall lifecycle tests are separate acceptance gates."
foreach ($warning in $warnings) {
    Write-Host "[verify-release] WARNING: $warning" -ForegroundColor Yellow
}
if ($errors.Count -gt 0) {
    Write-Host "[verify-release] FAILED: $($errors.Count) issue(s) were found." -ForegroundColor Red
    foreach ($releaseError in $errors) {
        Write-Host "  - $releaseError" -ForegroundColor Red
    }
    throw "Static release verification failed."
}

Write-Host "[verify-release] PASSED: Static artifact, provenance, checksum, and smoke checks passed." -ForegroundColor Green

function Get-VireloRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $rootPath = [System.IO.Path]::GetFullPath($Root).TrimEnd("\", "/")
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $prefix = $rootPath + [System.IO.Path]::DirectorySeparatorChar
    if (-not $fullPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "The path '$fullPath' is outside the release root '$rootPath'."
    }
    return $fullPath.Substring($prefix.Length).Replace("\", "/")
}

function Get-VireloAppVersion {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)

    $configPath = Join-Path $ProjectRoot "virelo\app\config.py"
    $configText = [System.IO.File]::ReadAllText($configPath)
    $match = [regex]::Match(
        $configText,
        '(?m)^APP_VERSION\s*=\s*"(\d+\.\d+\.\d+)"\s*$'
    )
    if (-not $match.Success) {
        throw "APP_VERSION was not found or is not a three-component numeric version."
    }
    return $match.Groups[1].Value
}

function Get-VireloReleaseInputPaths {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)

    $explicitFiles = @(
        "Virelo.spec",
        "icon.ico",
        "installer\virelo.iss",
        "main.py",
        "pyproject.toml",
        "frontend\.prettierignore",
        "frontend\.prettierrc.json",
        "frontend\eslint.config.js",
        "frontend\index.html",
        "frontend\package-lock.json",
        "frontend\package.json",
        "frontend\vite.config.js"
    )
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($relativePath in $explicitFiles) {
        $fullPath = Join-Path $ProjectRoot $relativePath
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            throw "A required release input is missing: $relativePath."
        }
        $null = $paths.Add($relativePath.Replace("\", "/"))
    }

    $directoryExtensions = [ordered]@{
        "branding" = @(".bmp", ".svg")
        "frontend\src" = @(".css", ".js", ".jsx")
        "requirements" = @(".txt")
        "scripts" = @(".ps1", ".py")
        "tests" = @(".py")
        "virelo" = @(".py")
    }
    foreach ($entry in $directoryExtensions.GetEnumerator()) {
        $directory = Join-Path $ProjectRoot $entry.Key
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
            throw "A required release input directory is missing: $($entry.Key)."
        }
        Get-ChildItem -LiteralPath $directory -Recurse -File |
            Where-Object { $entry.Value -contains $_.Extension } |
            ForEach-Object {
                $null = $paths.Add(
                    (Get-VireloRelativePath -Root $ProjectRoot -Path $_.FullName)
                )
            }
    }

    return @($paths | Sort-Object -Unique)
}

function Get-VireloFileHashMap {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [Parameter(Mandatory = $true)][string[]]$RelativePaths
    )

    $hashes = [ordered]@{}
    foreach ($relativePath in ($RelativePaths | Sort-Object)) {
        $fullPath = Join-Path $ProjectRoot $relativePath
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            throw "A release input disappeared while hashing: $relativePath."
        }
        $hashes[$relativePath.Replace("\", "/")] = (
            Get-FileHash -LiteralPath $fullPath -Algorithm SHA256
        ).Hash
    }
    return $hashes
}

function Write-VireloSha256Manifest {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $resolvedRoot = (Resolve-Path -LiteralPath $Root).Path
    $outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
    $lines = Get-ChildItem -LiteralPath $resolvedRoot -Recurse -File |
        Where-Object {
            -not $_.FullName.Equals(
                $outputFullPath,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        } |
        Sort-Object FullName |
        ForEach-Object {
            $relativePath = Get-VireloRelativePath -Root $resolvedRoot -Path $_.FullName
            $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            "$hash *$relativePath"
        }
    [System.IO.File]::WriteAllLines(
        $outputFullPath,
        $lines,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

function Test-VireloSha256Manifest {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$ManifestPath
    )

    $failures = New-Object System.Collections.Generic.List[string]
    $resolvedRoot = (Resolve-Path -LiteralPath $Root).Path.TrimEnd("\", "/")
    $rootPrefix = $resolvedRoot + [System.IO.Path]::DirectorySeparatorChar
    $expectedPaths = New-Object 'System.Collections.Generic.HashSet[string]' (
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($line in [System.IO.File]::ReadAllLines($ManifestPath)) {
        if (-not $line.Trim()) {
            continue
        }
        $match = [regex]::Match($line, '^([A-Fa-f0-9]{64}) \*(.+)$')
        if (-not $match.Success) {
            $null = $failures.Add("Malformed checksum line: $line")
            continue
        }
        $expectedHash = $match.Groups[1].Value.ToUpperInvariant()
        $relativePath = $match.Groups[2].Value.Replace("/", "\")
        if (-not $expectedPaths.Add($relativePath)) {
            $null = $failures.Add("Duplicate checksum path: $relativePath")
            continue
        }
        $fullPath = [System.IO.Path]::GetFullPath((Join-Path $resolvedRoot $relativePath))
        if (-not $fullPath.StartsWith(
                $rootPrefix,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
            $null = $failures.Add("Checksum path escapes the bundle root: $relativePath")
            continue
        }
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            $null = $failures.Add("Missing checksummed file: $relativePath")
            continue
        }
        $actualHash = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash
        if ($actualHash -ne $expectedHash) {
            $null = $failures.Add("Checksum mismatch: $relativePath")
        }
    }

    $manifestFullPath = [System.IO.Path]::GetFullPath($ManifestPath)
    Get-ChildItem -LiteralPath $Root -Recurse -File | ForEach-Object {
        if (-not $_.FullName.Equals(
                $manifestFullPath,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
            $relativePath = (
                Get-VireloRelativePath -Root $Root -Path $_.FullName
            ).Replace("/", "\")
            if (-not $expectedPaths.Contains($relativePath)) {
                $null = $failures.Add("Unlisted bundle file: $relativePath")
            }
        }
    }
    return @($failures)
}

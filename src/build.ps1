# Build-Skript fuer LEK-Bastler
# Erstellt die EXE via PyInstaller und stellt ein fertiges Deploy-Paket zusammen

param(
    [switch]$SkipBuild,
    [switch]$Help
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $OutputEncoding = [System.Text.Encoding]::UTF8
}

if ($Help) {
    Write-Host "LEK-Bastler Build-Skript - Optionen:" -ForegroundColor DarkCyan
    Write-Host "  -Help       : Nur diese Hilfe anzeigen" -ForegroundColor DarkGray
    Write-Host "  -SkipBuild  : PyInstaller-Build ueberspringen (nur Deploy-Paket neu erstellen)" -ForegroundColor DarkGray
    exit 0
}

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $projectRoot

$BaseExeName = "LEK-Bastler"
$SpecFile    = Join-Path $PSScriptRoot "LEK-Bastler.spec"
$PyExePath   = "dist\LEK-Bastler.exe"
$VenvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"

function Get-BuildVersion {
    $versionFile = Join-Path $PSScriptRoot 'build_version_info.txt'
    if (-not (Test-Path $versionFile)) {
        return '0.0.0.0'
    }

    $content = Get-Content $versionFile -Raw

    if ($content -match "StringStruct\('FileVersion',\s*'([^']+)'\)") {
        return $matches[1]
    }

    if ($content -match 'filevers=\((\d+),\s*(\d+),\s*(\d+),\s*(\d+)\)') {
        return "$($matches[1]).$($matches[2]).$($matches[3]).$($matches[4])"
    }

    return '0.0.0.0'
}

function Convert-ToThreePartVersion {
    param(
        [string]$Version
    )

    if ([string]::IsNullOrWhiteSpace($Version)) {
        return '0.0.0'
    }

    $parts = $Version.Split('.') | Where-Object { $_ -ne '' }
    while ($parts.Count -lt 3) {
        $parts += '0'
    }

    return "$($parts[0]).$($parts[1]).$($parts[2])"
}

$Version = Convert-ToThreePartVersion (Get-BuildVersion)
$VersionedExeName = "${BaseExeName}_$Version.exe"
$DeployFolderName = "${BaseExeName}_$Version"
$DeployDir = Join-Path 'dist' $DeployFolderName
$VersionedExePath = Join-Path $DeployDir $VersionedExeName

$ReleaseDir = 'release'
$ReleaseZipName = "$DeployFolderName.zip"
$ReleaseZipPath = Join-Path $ReleaseDir $ReleaseZipName

$DocsFilesToExcludeFromRelease = @(
    'CHANGELOG.md',
    'DOKUMENTATION_DIAGRAMME.md',
    'DOKUMENTATION_PROJEKT.md',
    'DOKUMENTATION_RELEASES.md',
    'RELEASE_QA_CHECKLISTE.md',
    'RELEASE_SMOKETEST_PROTOKOLL.md'
)

function Set-FileContentIfChanged {
    param(
        [string]$Path,
        [string]$Content
    )

    $existing = ''
    if (Test-Path $Path) {
        $existing = Get-Content -Path $Path -Raw -Encoding UTF8
    }

    if ($existing -ne $Content) {
        Set-Content -Path $Path -Value $Content -Encoding UTF8
        return $true
    }
    return $false
}

function Update-VersionReferences {
    param(
        [string]$Version,
        [string]$ProjectRoot,
        [string]$ReleaseNotesPath,
        [string]$ReleaseZipName,
        [string]$DeployFolderName,
        [string]$VersionedExeName
    )

    Write-Host "0. Versionsreferenzen werden aktualisiert..." -ForegroundColor Yellow

    $targets = @(
        (Join-Path $ProjectRoot 'docs\DOKUMENTATION_ANWENDER.md'),
        (Join-Path $ProjectRoot 'docs\DOKUMENTATION_TECHNIK.md'),
        (Join-Path $ProjectRoot 'docs\DOKUMENTATION_PROJEKT.md'),
        $ReleaseNotesPath
    )

    $updated = 0
    foreach ($path in $targets) {
        if (-not (Test-Path $path)) {
            continue
        }

        $content = Get-Content -Path $path -Raw -Encoding UTF8
        $newContent = $content

        if ($path -like '*DOKUMENTATION_ANWENDER.md') {
            $newContent = [regex]::Replace(
                $newContent,
                'Hinweis zum aktuellen Stand:\s*Version\s*\*\*[0-9]+\.[0-9]+\.[0-9]+\*\*\.',
                "Hinweis zum aktuellen Stand: Version **$Version**."
            )
        }

        if ($path -like '*DOKUMENTATION_TECHNIK.md') {
            $newContent = [regex]::Replace(
                $newContent,
                'Aktueller Versionsstand:\s*\*\*[0-9]+\.[0-9]+\.[0-9]+\*\*\.',
                "Aktueller Versionsstand: **$Version**."
            )
        }

        if ($path -like '*DOKUMENTATION_PROJEKT.md') {
            $newContent = [regex]::Replace(
                $newContent,
                'Aktueller Versionsstand:\s*\*\*[0-9]+\.[0-9]+\.[0-9]+\*\*\s*\(aus\s+`?src/build_version_info\.txt`?\)\.',
                ('Aktueller Versionsstand: **{0}** (aus `src/build_version_info.txt`).' -f $Version)
            )

            $newContent = [regex]::Replace(
                $newContent,
                'dist/LEK-Bastler_[0-9]+\.[0-9]+\.[0-9]+/LEK-Bastler_[0-9]+\.[0-9]+\.[0-9]+\.exe',
                "dist/$DeployFolderName/$VersionedExeName"
            )

            $newContent = [regex]::Replace(
                $newContent,
                'release/LEK-Bastler_[0-9]+\.[0-9]+\.[0-9]+\.zip',
                "release/$ReleaseZipName"
            )
        }

        if ($path -like '*RELEASE_NOTES_v*.md') {
            $newContent = [regex]::Replace(
                $newContent,
                '^#\s+Release Notes\s+v[0-9]+\.[0-9]+\.[0-9]+\s*$',
                "# Release Notes v$Version",
                [System.Text.RegularExpressions.RegexOptions]::Multiline
            )

            $newContent = [regex]::Replace(
                $newContent,
                'Release-Build für\s*`[0-9]+\.[0-9]+\.[0-9]+`',
                ('Release-Build für `{0}`' -f $Version)
            )

            $newContent = [regex]::Replace(
                $newContent,
                'dist/LEK-Bastler_[0-9]+\.[0-9]+\.[0-9]+/LEK-Bastler_[0-9]+\.[0-9]+\.[0-9]+\.exe',
                "dist/$DeployFolderName/$VersionedExeName"
            )

            $newContent = [regex]::Replace(
                $newContent,
                'release/LEK-Bastler_[0-9]+\.[0-9]+\.[0-9]+\.zip',
                "release/$ReleaseZipName"
            )
        }

        if (Set-FileContentIfChanged -Path $path -Content $newContent) {
            $updated += 1
            Write-Host "   aktualisiert: $([System.IO.Path]::GetFileName($path))" -ForegroundColor DarkGray
        }
    }

    if ($updated -eq 0) {
        Write-Host "   keine Versionsanpassungen erforderlich" -ForegroundColor DarkGray
    }
}

Write-Host "`nLEK-Bastler Build-Prozess" -ForegroundColor Green
Write-Host "=========================`n" -ForegroundColor Green

Update-VersionReferences `
    -Version $Version `
    -ProjectRoot $projectRoot `
    -ReleaseNotesPath (Join-Path $projectRoot "release\RELEASE_NOTES_v$Version.md") `
    -ReleaseZipName $ReleaseZipName `
    -DeployFolderName $DeployFolderName `
    -VersionedExeName $VersionedExeName

# ─── Schritt 1: PyInstaller-Build ───────────────────────────────────────────
if (-not $SkipBuild) {
    Write-Host "1. PyInstaller-Build wird erstellt..." -ForegroundColor Yellow

    if (-not (Test-Path $VenvPython)) {
        Write-Host "   venv nicht gefunden ($VenvPython)." -ForegroundColor Red
        Write-Host "   Bitte zuerst: python -m venv .venv && .venv\Scripts\pip install -r src\REQUIREMENTS.txt" -ForegroundColor Yellow
        exit 1
    }

    & $VenvPython -m PyInstaller $SpecFile --noconfirm
    if ($LASTEXITCODE -ne 0) {
        Write-Host "   PyInstaller fehlgeschlagen!" -ForegroundColor Red
        exit 1
    }

    if (-not (Test-Path $PyExePath)) {
        Write-Host "   EXE nicht gefunden nach Build: $PyExePath" -ForegroundColor Red
        exit 1
    }
    Write-Host "   EXE erstellt: $PyExePath" -ForegroundColor Green
} else {
    Write-Host "1. PyInstaller-Build uebersprungen (-SkipBuild)." -ForegroundColor DarkGray
    if (-not (Test-Path $PyExePath)) {
        Write-Host "   EXE nicht gefunden: $PyExePath" -ForegroundColor Red
        exit 1
    }
}

# ─── Schritt 2: Deploy-Verzeichnis vorbereiten ──────────────────────────────
Write-Host "`n2. Deploy-Verzeichnis wird erstellt: $DeployDir" -ForegroundColor Yellow

if (Test-Path $DeployDir) {
    Remove-Item $DeployDir -Recurse -Force
}
New-Item -ItemType Directory -Path $DeployDir | Out-Null

# ─── Schritt 3: EXE kopieren ─────────────────────────────────────────────────
Write-Host "`n3. Dateien und Ordner werden kopiert..." -ForegroundColor Yellow

Copy-Item $PyExePath $VersionedExePath
Write-Host "   $VersionedExeName" -ForegroundColor DarkGray

# ─── Schritt 4: Pflicht-Ordner kopieren ──────────────────────────────────────
$foldersToInclude = @("data", "docs")
foreach ($folder in $foldersToInclude) {
    if (Test-Path $folder) {
        Copy-Item $folder "$DeployDir\" -Recurse

        if ($folder -eq 'docs') {
            foreach ($excludedDoc in $DocsFilesToExcludeFromRelease) {
                $excludedPath = Join-Path $DeployDir "docs\$excludedDoc"
                if (Test-Path $excludedPath) {
                    Remove-Item $excludedPath -Force
                }
            }
        }

        $count = (Get-ChildItem "$DeployDir\$folder" -File -Recurse).Count
        Write-Host "   $folder\ ($count Datei(en))" -ForegroundColor DarkGray

        if ($folder -eq 'docs' -and $DocsFilesToExcludeFromRelease.Count -gt 0) {
            Write-Host ("   docs\ (ohne: {0})" -f (($DocsFilesToExcludeFromRelease | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join ', ')) -ForegroundColor DarkGray
        }
    } else {
        Write-Host "   WARNUNG: Ordner nicht gefunden – $folder\" -ForegroundColor Yellow
    }
}

# LEKs-Ordner: nur Platzhalterdatei(en), keine erzeugten LEK-Dokumente
New-Item -ItemType Directory -Path "$DeployDir\data\LEKs" -Force | Out-Null
$leksReadme = "data\LEKs\README.md"
if (Test-Path $leksReadme) {
    Copy-Item $leksReadme "$DeployDir\data\LEKs\"
    Write-Host "   data\LEKs\ (nur README.md)" -ForegroundColor DarkGray
} else {
    $fallbackReadmePath = Join-Path $DeployDir "data\LEKs\README.md"
    @(
        "# LEKs"
        ""
        "Dieser Ordner wird für erzeugte LEK-Dokumente verwendet."
    ) | Set-Content -Path $fallbackReadmePath -Encoding UTF8
    Write-Host "   data\LEKs\ (README.md automatisch erstellt)" -ForegroundColor DarkGray
}

# ─── Schritt 5: Begleitdateien kopieren ──────────────────────────────────────
$filesToInclude = @("README.md", "src\LIZENZ.txt")
foreach ($file in $filesToInclude) {
    if (Test-Path $file) {
        $targetName = Split-Path $file -Leaf
        Copy-Item $file (Join-Path $DeployDir $targetName)
        Write-Host "   $targetName" -ForegroundColor DarkGray
    } else {
        Write-Host "   WARNUNG: Datei nicht gefunden – $file" -ForegroundColor Yellow
    }
}

# ─── Ergebnis ─────────────────────────────────────────────────────────────────
# ─── Schritt 6: Release-ZIP erzeugen ─────────────────────────────────────────
Write-Host "`n4. Release-ZIP wird erstellt..." -ForegroundColor Yellow

if (-not (Test-Path $ReleaseDir)) {
    New-Item -ItemType Directory -Path $ReleaseDir | Out-Null
}

if (Test-Path $ReleaseZipPath) {
    Remove-Item $ReleaseZipPath -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory(
    (Resolve-Path $DeployDir).Path,
    (Join-Path (Get-Location) $ReleaseZipPath),
    [System.IO.Compression.CompressionLevel]::Optimal,
    $false
)
Write-Host "   $ReleaseZipName" -ForegroundColor DarkGray

# ─── Schritt 7: Release Notes erzeugen ───────────────────────────────────────
Write-Host "`n5. Release Notes werden erstellt..." -ForegroundColor Yellow
$ReleaseNotesName = "RELEASE_NOTES_v$Version.md"
$ReleaseNotesPath = Join-Path $ReleaseDir $ReleaseNotesName
if (-not (Test-Path $ReleaseNotesPath)) {
    @(
        "# Release Notes v$Version"
        ""
        "## Enthalten"
        ""
        "- LEK-Bastler im Versionsstand $Version."
        "- Versioniertes Deploy-Verzeichnis und Release-ZIP wurden erzeugt."
        "- Technische Dokumentation und Anwenderdokumentation sind enthalten."
        ""
        "## Artefakte"
        ""
        "- dist/$DeployFolderName/$VersionedExeName"
        "- release/$ReleaseZipName"
        "- release/$ReleaseNotesName"
    ) | Set-Content -Path $ReleaseNotesPath -Encoding UTF8
    Write-Host "   $ReleaseNotesName" -ForegroundColor DarkGray
} else {
    Write-Host "   $ReleaseNotesName (bereits vorhanden, bleibt erhalten)" -ForegroundColor DarkGray
}

# ─── Ergebnis ─────────────────────────────────────────────────────────────────
Write-Host "`nBuild erfolgreich abgeschlossen!" -ForegroundColor Green
$exeSize = [math]::Round((Get-Item $VersionedExePath).Length / 1MB, 1)
Write-Host "Version            : $Version" -ForegroundColor Cyan
Write-Host "Deploy-Verzeichnis : $DeployDir" -ForegroundColor Cyan
Write-Host "EXE-Datei          : $VersionedExeName ($exeSize MB)" -ForegroundColor Cyan
$zipSize = [math]::Round((Get-Item $ReleaseZipPath).Length / 1MB, 1)
Write-Host "Release-ZIP        : $ReleaseZipPath ($zipSize MB)" -ForegroundColor Cyan
Write-Host "Release Notes      : $ReleaseNotesPath" -ForegroundColor Cyan
$deployFiles = (Get-ChildItem $DeployDir -Recurse -File).Count
Write-Host "Dateien gesamt     : $deployFiles" -ForegroundColor Cyan

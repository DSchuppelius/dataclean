<#
.SYNOPSIS
    Erstellt ein portables ZIP-Paket von DataClean aus einem Release-Build.

.DESCRIPTION
    Baut das Projekt DataClean.csproj in der Konfiguration Release mit MSBuild
    und packt das Ausgabeverzeichnis in ein versioniertes ZIP-Archiv unter
    .\artifacts.

.PARAMETER Version
    Versionskennung für den Archivnamen. Standard: aktuelles Datum (yyyy.MM.dd).

.PARAMETER Configuration
    Build-Konfiguration. Standard: Release.

.EXAMPLE
    ./packaging/build-package.ps1 -Version 1.2.0
#>
[CmdletBinding()]
param(
    [string]$Version = (Get-Date -Format 'yyyy.MM.dd'),
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'

# Projektstamm = übergeordnetes Verzeichnis dieses Skripts.
$root = Split-Path -Parent $PSScriptRoot
$projectFile = Join-Path $root 'DataClean.csproj'
if (-not (Test-Path $projectFile)) {
    throw "DataClean.csproj wurde unter '$projectFile' nicht gefunden."
}

# MSBuild lokalisieren: zuerst über vswhere, dann gängige Fallback-Pfade.
function Resolve-MSBuild {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $path = & $vswhere -latest -requires Microsoft.Component.MSBuild `
            -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
        if ($path) { return $path }
    }

    $candidates = @(
        'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe'
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }

    throw 'MSBuild.exe konnte nicht gefunden werden. Bitte Visual Studio oder die Build Tools installieren.'
}

$msbuild = Resolve-MSBuild
Write-Host "Verwende MSBuild: $msbuild"

# Release-Build ausführen.
& $msbuild $projectFile /t:Rebuild /p:Configuration=$Configuration /nologo /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Build fehlgeschlagen (ExitCode $LASTEXITCODE)."
}

$outputDir = Join-Path $root "bin\$Configuration"
if (-not (Test-Path $outputDir)) {
    throw "Ausgabeverzeichnis '$outputDir' wurde nicht gefunden."
}

# Artefaktverzeichnis vorbereiten.
$artifacts = Join-Path $root 'artifacts'
New-Item -ItemType Directory -Force -Path $artifacts | Out-Null

$zipPath = Join-Path $artifacts "DataClean-$Version.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

# Nur Laufzeitdateien einpacken (keine .pdb/.xml-Dokumentation).
$include = Get-ChildItem -Path $outputDir -File |
    Where-Object { $_.Extension -notin '.pdb', '.xml' }

Compress-Archive -Path $include.FullName -DestinationPath $zipPath -CompressionLevel Optimal

Write-Host ""
Write-Host "Paket erstellt: $zipPath" -ForegroundColor Green

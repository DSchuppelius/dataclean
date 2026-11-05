<#
.SYNOPSIS
    Baut DataClean und erzeugt daraus die verteilbaren Pakete.

.DESCRIPTION
    Baut das Projekt in der angegebenen Konfiguration mit MSBuild und legt unter .\artifacts ein
    portables ZIP-Archiv, die einzelne DataClean.exe und auf Wunsch einen Installer ab.

    ClickOnce-Manifeste entstehen dabei nicht: Ihr Signaturzertifikat liegt nur auf dem
    Entwicklungsrechner, auf einem Build-Server würde der Build daran scheitern.

.PARAMETER Version
    Versionsnummer in der Form 2.0.0. Sie steht in den Dateinamen und wird für diesen Build in die
    Exe geschrieben; Properties\AssemblyInfo.cs ist danach wieder unverändert.
    Standard: die Version aus Properties\AssemblyInfo.cs.

.PARAMETER Configuration
    Build-Konfiguration. Standard: Release.

.PARAMETER RunTests
    Baut auch die Tests und führt sie aus. Schlagen sie fehl, entsteht kein Paket.

.PARAMETER Installer
    Erzeugt zusätzlich den Installer. Benötigt Inno Setup.

.PARAMETER IsccPath
    Pfad zur ISCC.exe, falls Inno Setup nicht am üblichen Ort installiert ist.

.EXAMPLE
    ./packaging/build-package.ps1 -Version 2.0.0 -RunTests -Installer
#>
[CmdletBinding()]
param(
    [string]$Version,
    [string]$Configuration = 'Release',
    [switch]$RunTests,
    [switch]$Installer,
    [string]$IsccPath
)

$ErrorActionPreference = 'Stop'

# Projektstamm = übergeordnetes Verzeichnis dieses Skripts.
$root = Split-Path -Parent $PSScriptRoot
$projectFile = Join-Path $root 'DataClean.csproj'
if (-not (Test-Path $projectFile)) {
    throw "DataClean.csproj wurde unter '$projectFile' nicht gefunden."
}
$solutionFile = Join-Path $root 'DataClean.sln'
$assemblyInfo = Join-Path $root 'Properties\AssemblyInfo.cs'

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

function Resolve-VsTest {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $path = & $vswhere -latest -find '**\vstest.console.exe' | Select-Object -First 1
        if ($path) { return $path }
    }
    throw 'vstest.console.exe konnte nicht gefunden werden. Bitte Visual Studio installieren oder ohne -RunTests bauen.'
}

function Resolve-Iscc {
    if ($IsccPath) {
        if (Test-Path $IsccPath) { return $IsccPath }
        throw "ISCC.exe wurde unter '$IsccPath' nicht gefunden."
    }

    $command = Get-Command 'ISCC.exe' -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $candidates = @(
        (Join-Path $env:ProgramFiles 'Inno Setup 7\ISCC.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe')
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }

    throw 'ISCC.exe konnte nicht gefunden werden. Bitte Inno Setup installieren oder -IsccPath angeben.'
}

# Version bestimmen und prüfen, bevor etwas gebaut wird.
if (-not $Version) {
    $match = [regex]::Match([IO.File]::ReadAllText($assemblyInfo), '(?m)^\[assembly: AssemblyVersion\("([^"]+)"\)\]')
    if (-not $match.Success) {
        throw "Keine AssemblyVersion in '$assemblyInfo' gefunden."
    }
    $Version = $match.Groups[1].Value -replace '\.0$', ''
}
if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Ungültige Version '$Version'. Erwartet wird die Form 2.0.0."
}

# Fehlende Werkzeuge sollen auffallen, bevor der Build läuft.
$msbuild = Resolve-MSBuild
Write-Host "Verwende MSBuild: $msbuild"
$vstest = if ($RunTests) { Resolve-VsTest }
$iscc = if ($Installer) { Resolve-Iscc }

# Die Version gilt nur für diesen Build, die Datei wird danach wiederhergestellt.
$originalAssemblyInfo = [IO.File]::ReadAllBytes($assemblyInfo)
try {
    $stamped = [regex]::Replace(
        [IO.File]::ReadAllText($assemblyInfo),
        '(?m)^(\[assembly: Assembly(?:File)?Version\(")[^"]+("\)\])',
        "`${1}$Version.0`${2}")
    [IO.File]::WriteAllText($assemblyInfo, $stamped, (New-Object Text.UTF8Encoding($true)))

    # Die Tests liegen in der Solution; ohne sie reicht das Projekt.
    $buildTarget = if ($RunTests) { $solutionFile } else { $projectFile }
    & $msbuild $buildTarget /restore /t:Rebuild /p:Configuration=$Configuration `
        /p:SignManifests=false /p:GenerateManifests=false /nologo /v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "Build fehlgeschlagen (ExitCode $LASTEXITCODE)."
    }
} finally {
    [IO.File]::WriteAllBytes($assemblyInfo, $originalAssemblyInfo)
}

if ($RunTests) {
    & $vstest (Join-Path $root "DataClean.Tests\bin\$Configuration\DataClean.Tests.dll")
    if ($LASTEXITCODE -ne 0) {
        throw "Tests fehlgeschlagen (ExitCode $LASTEXITCODE)."
    }
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

# Nur Laufzeitdateien einpacken. Manifeste können noch von einem früheren Build dort liegen.
$include = @(Get-ChildItem -Path $outputDir -File |
    Where-Object { $_.Extension -notin '.pdb', '.xml', '.application', '.manifest' } |
    ForEach-Object { $_.FullName })
$include += Join-Path $root 'LICENSE'

Compress-Archive -Path $include -DestinationPath $zipPath -CompressionLevel Optimal

Write-Host ""
Write-Host "Paket erstellt: $zipPath" -ForegroundColor Green

# Die Exe läuft auch allein: Die Standardwerte der Konfiguration sind in ihr enthalten.
$exePath = Join-Path $artifacts 'DataClean.exe'
Copy-Item -LiteralPath (Join-Path $outputDir 'DataClean.exe') -Destination $exePath -Force
Write-Host "Exe abgelegt: $exePath" -ForegroundColor Green

if ($Installer) {
    & $iscc "/DAppVersion=$Version" "/DBuildDir=$outputDir" '/Qp' (Join-Path $PSScriptRoot 'DataClean.iss')
    if ($LASTEXITCODE -ne 0) {
        throw "Installer fehlgeschlagen (ExitCode $LASTEXITCODE)."
    }
    Write-Host "Installer erstellt: $(Join-Path $artifacts "DataClean-Setup-$Version.exe")" -ForegroundColor Green
}

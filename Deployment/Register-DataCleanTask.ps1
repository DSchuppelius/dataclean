<#
.SYNOPSIS
    Registriert DataClean als geplante Aufgabe, die mit höchsten Privilegien
    bei der Benutzeranmeldung ausgeführt wird.

.DESCRIPTION
    Dieses Skript muss in einer PowerShell-Sitzung mit Administratorrechten
    ausgeführt werden. Es importiert die Aufgabe aus DataClean.Task.xml und
    setzt dabei den Pfad zur DataClean.exe sowie den Zielbenutzer.

.PARAMETER ExePath
    Vollständiger Pfad zur DataClean.exe.

.PARAMETER TaskName
    Name der geplanten Aufgabe (Standard: "DataClean").

.PARAMETER UserId
    Benutzer, unter dem die Aufgabe läuft (Standard: aktueller Benutzer).

.EXAMPLE
    .\Register-DataCleanTask.ps1 -ExePath "C:\Program Files\DataClean\DataClean.exe"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExePath,

    [string]$TaskName = "DataClean",

    [string]$UserId = "$env:USERDOMAIN\$env:USERNAME"
)

$ErrorActionPreference = "Stop"

# Administratorrechte prüfen
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Dieses Skript muss mit Administratorrechten ausgeführt werden."
}

if (-not (Test-Path -LiteralPath $ExePath)) {
    throw "Die angegebene DataClean.exe wurde nicht gefunden: $ExePath"
}

# Aktion, Auslöser und Principal (höchste Privilegien, nur bei Anmeldung)
$action    = New-ScheduledTaskAction -Execute $ExePath -Argument "/autostart"
$trigger   = New-ScheduledTaskTrigger -AtLogOn
$principalObj = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Highest
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1) -MultipleInstances IgnoreNew

# Bestehende Aufgabe ersetzen, falls vorhanden
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Bestehende Aufgabe '$TaskName' entfernt."
}

Register-ScheduledTask -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principalObj `
    -Settings $settings `
    -Description "Bereinigt temporäre Dateien, Caches, Papierkorb und optional Windows.old, DNS-Cache und Event-Logs." | Out-Null

Write-Host "Aufgabe '$TaskName' erfolgreich registriert."
Write-Host "Ausführung: bei Anmeldung von '$UserId' mit höchsten Privilegien."

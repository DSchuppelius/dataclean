; Installer für DataClean (Inno Setup). Aufruf über packaging\build-package.ps1 -Installer,
; das Skript übergibt Version und Build-Ausgabe.
;
; Der Installer legt bewusst keine geplante Aufgabe an: Die tägliche Ausführung lässt sich in den
; Einstellungen der Anwendung aktivieren, und eine weitere Aufgabe könnte zu doppelten Läufen führen.
; Auch das Starten nach der Installation fehlt bewusst: DataClean ist ein Löschwerkzeug und soll
; nur laufen, wenn jemand es ausdrücklich startet.

#ifndef AppVersion
  #error AppVersion fehlt. Bitte über packaging\build-package.ps1 -Installer bauen.
#endif
#ifndef BuildDir
  #error BuildDir fehlt. Bitte über packaging\build-package.ps1 -Installer bauen.
#endif

[Setup]
; Die AppId darf sich nie ändern, sonst gilt ein Update als zweite, getrennte Installation.
AppId={{3A437092-1282-4E61-82C1-F6A0809DFD89}
AppName=DataClean
AppVersion={#AppVersion}
AppPublisher=Daniel Jörg Schuppelius
AppPublisherURL=https://schuppelius.org
VersionInfoVersion={#AppVersion}
DefaultDirName={autopf}\DataClean
DisableProgramGroupPage=yes
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=commandline
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\artifacts
OutputBaseFilename=DataClean-Setup-{#AppVersion}
SetupIconFile=..\djs-ani.ico
UninstallDisplayIcon={app}\DataClean.exe
UninstallDisplayName=DataClean
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; Eine geschlossene Instanz wird nicht von selbst neu gestartet.
RestartApplications=no

[Languages]
Name: "de"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; Flags: unchecked

[Files]
Source: "{#BuildDir}\DataClean.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildDir}\DataClean.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\DataClean"; Filename: "{app}\DataClean.exe"
Name: "{autodesktop}\DataClean"; Filename: "{app}\DataClean.exe"; Tasks: desktopicon

[Code]
function InitializeSetup: Boolean;
begin
  Result := IsDotNetInstalled(net48, 0);
  if not Result then
    SuppressibleMsgBox('DataClean benötigt das .NET Framework 4.8. Bitte installieren Sie es zuerst und starten Sie die Einrichtung dann erneut.', mbCriticalError, MB_OK, IDOK);
end;

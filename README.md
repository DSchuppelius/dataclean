# DataClean

Windows-Programm, das temporäre Dateien und Caches löscht, von Hand oder automatisch über die
Aufgabenplanung. Ein eigener Menüpunkt findet verwaiste Registry-Einträge.

## Installation

Die Dateien liegen beim [aktuellen Release](https://github.com/DSchuppelius/dataclean/releases/latest):

| Datei                           | Zweck                                                         |
|---------------------------------|---------------------------------------------------------------|
| `DataClean-Setup-<Version>.exe` | Installer, legt das Programm in `%ProgramFiles%\DataClean` ab |
| `DataClean.exe`                 | Programm ohne Installation, läuft ohne weitere Dateien        |
| `DataClean-<Version>.zip`       | Programm mit `DataClean.exe.config` und Lizenz                |

Voraussetzung ist Windows mit .NET Framework 4.8. Die Dateien sind nicht signiert, SmartScreen kann
deshalb beim ersten Start warnen.

## Benutzung

Von Hand gestartet öffnet sich das Fenster im Zustand „Bereit“ und löscht noch nichts. Erst
**Bereinigung starten** beginnt den Lauf. Vor dem Löschen zeigt standardmäßig eine Rückfrage die
erwartete Größe. Mitglieder der Administratorengruppe bekommen beim Programmstart eine UAC-Abfrage,
weil einige Systemordner sonst nicht erreichbar sind.

### Was gelöscht wird

Ohne weitere Einstellungen bereinigt DataClean:

- die Temp-Ordner des Benutzers und von Windows sowie `C:\Temp`
- Internet-Cache und Verlauf (`%LOCALAPPDATA%\Microsoft\Windows\History`)
- Prefetch und die Downloads von Windows Update
- Absturzberichte (`%LOCALAPPDATA%\CrashDumps`) und den Miniaturansichten-Cache des Explorers
- die Caches von Edge, Chrome und Firefox
- den Papierkorb und den DNS-Cache

Windows-, System- und Programmordner sowie das Benutzerprofil sind geschützt. Nur die genannten
Unterordner sind davon ausgenommen.

### Einstellungen

- **Sicherheit:** Vorschau-Modus ohne Löschen, Wiederherstellungspunkt, Verschieben in den Papierkorb
  statt endgültigem Löschen, Ausschlussmuster wie `*.log` oder `**\Cache\**`
- **Mehr bereinigen:** weitere Browser-Caches, DISM-Komponentenbereinigung, Absturz-Dumps,
  Delivery Optimization, Papierkörbe aller Laufwerke, alte Protokolle
- **Ablauf:** nur bereinigen, wenn der Speicher knapp wird, gesperrte Dateien beim nächsten Neustart
  löschen, benannte Profile
- **Report:** als HTML oder CSV, auf Wunsch per E-Mail
- **Design:** folgt Windows oder fest hell oder dunkel

Die Ordnerliste, die geschützten Ordner, Mindestalter, Systemprüfung und SMTP-Server stehen nicht in
der Oberfläche, sondern in der `DataClean.exe.config`. Dort lassen sich auch `removeWindowsOld` und
`clearEventLogs` einschalten. Der Installer ersetzt diese Datei bei jedem Update.

### Automatisch bereinigen

> [!WARNING]
> Ein Start über die Aufgabenplanung oder mit `/autostart` löscht sofort und ohne Rückfrage.

Es gibt zwei Wege:

- **Täglich zu einer festen Uhrzeit:** in den Einstellungen einschalten.
- **Bei jeder Anmeldung:** das Skript `Deployment/Register-DataCleanTask.ps1` aus diesem Repository
  in einer PowerShell mit Administratorrechten ausführen.

  ```powershell
  .\Deployment\Register-DataCleanTask.ps1 -ExePath "C:\Program Files\DataClean\DataClean.exe"
  ```

Nach einem solchen Lauf schließt sich das Fenster nach 7,5 Sekunden. Offen bleibt es bei drohenden
Hardware-Ausfällen (SMART, Datenträgerstatus, Speicherfehler), nach einem unerwarteten Abbruch oder
solange jemand Details oder eine andere Seite geöffnet hat. Wenig freier Speicher und
Datenträgerfehler aus dem Ereignisprotokoll werden gemeldet, halten das Fenster aber nicht offen.

Der Installer legt keine geplanten Aufgaben an und entfernt auch keine. Die tägliche Ausführung
sollte vor dem Deinstallieren in den Einstellungen abgeschaltet werden.

### Registry

Die Registry-Prüfung läuft nur von Hand über den Menüpunkt **Registry**, nie im automatischen Lauf.
Als Fund gilt ein Eintrag nur, wenn seine Datei auf einem lokalen Laufwerk nachweislich fehlt.

- Vor dem Entfernen entsteht eine `.reg`-Sicherung, die sich über **Sicherung einspielen**
  zurückholen lässt.
- Funde unter „ActiveX und COM“ und „Typbibliotheken“ sind nicht vorausgewählt, weil andere Programme
  davon abhängen können.
- Windows-Dienste und Windows Installer werden bewusst nicht geprüft.
- Ohne Administratorrechte lassen sich nur Einträge des aktuellen Benutzers entfernen.

### Kommandozeile

| Option            | Wirkung                                 |
|-------------------|-----------------------------------------|
| `/autostart`      | Unbeaufsichtigter Lauf ohne UAC-Abfrage |
| `/dry-run`        | Vorschau-Modus, nichts wird gelöscht    |
| `/profile=<Name>` | Benanntes Profil für diesen Lauf        |
| `/report`         | Report nach dem Lauf erzeugen           |
| `/verbose`        | Ausführliches Protokoll                 |
| `/help`, `/?`     | Hilfe anzeigen                          |

### Wo die Daten liegen

| Ort                                        | Inhalt                    |
|--------------------------------------------|---------------------------|
| `%APPDATA%\DataClean`                      | Profile und Lauf-Historie |
| `%LOCALAPPDATA%\DataClean\Logs`            | Protokolle                |
| `%LOCALAPPDATA%\DataClean\RegistryBackups` | Registry-Sicherungen      |

## Entwicklung

Zum Bauen braucht es Visual Studio 2022 oder neuer (oder die Build Tools) mit dem .NET Framework 4.8
Targeting Pack.

```powershell
./packaging/build-package.ps1 -RunTests
```

Das baut das Programm, führt die Tests aus und legt ZIP und Exe unter `artifacts/` ab. Mit dem
Schalter `-Installer` entsteht auch der Installer, dafür wird Inno Setup gebraucht.

Beim Ausprobieren ist `/dry-run` die sichere Wahl: Ein echter Lauf löscht unter anderem den Inhalt
von `%TEMP%`.

Ein Release entsteht, sobald ein Versions-Tag wie `v2.0.0` gepusht wird. Details stehen in
[packaging/README.md](packaging/README.md).

## Lizenz

[MIT](LICENSE)

# DataClean – Paketierung / Veröffentlichung

Dieses Verzeichnis enthält das Skript und die Installer-Definition, aus denen die
verteilbaren Pakete von DataClean entstehen: ein Installer, die einzelne Exe und ein
portables ZIP.

## Release veröffentlichen

1. Alle Änderungen committen und pushen. Gebaut wird der Stand des Tags.
2. Die Version taggen und das Tag pushen:

   ```powershell
   git tag v2.0.0
   git push origin v2.0.0
   ```

Der Workflow `.github/workflows/release.yml` baut und testet dann und veröffentlicht
ein Release mit diesen Dateien. Die Version kommt aus dem Tag, die Release-Notizen
erzeugt GitHub aus den Commits.

| Datei                           | Inhalt                                       |
|---------------------------------|----------------------------------------------|
| `DataClean-Setup-<Version>.exe` | Installer                                    |
| `DataClean.exe`                 | Programm ohne Installation                   |
| `DataClean-<Version>.zip`       | Programm, `DataClean.exe.config` und Lizenz  |

Schlägt ein Schritt fehl, entsteht kein Release. Gibt es zum Tag schon ein Release,
etwa weil es auf GitHub zusammen mit dem Tag angelegt wurde, behält es seinen Text
und bekommt nur die Dateien. Ein Release zu einem schon vorhandenen Tag startet den
Workflow nicht.

Ohne Tag lässt er sich unter **Actions → Release → Run workflow** als Probelauf
starten. Die Dateien liegen danach als Build-Artefakt am Lauf.

Die Exe auf NETLOGON aktualisiert der Workflow nicht. Sie wird weiter von Hand
kopiert; dafür liegt die `DataClean.exe` am Release.

## Voraussetzungen zum lokalen Bauen

- Visual Studio 2022/2026 oder die Build Tools mit MSBuild
- Die Komponente ".NET Framework 4.8 SDK/Targeting Pack"
- Für den Installer: Inno Setup 6.3 oder neuer

## Pakete lokal erstellen

```powershell
# Aus dem Projektstamm ausführen:
./packaging/build-package.ps1 -RunTests -Installer
```

Ohne `-Version` gilt die Version aus `Properties/AssemblyInfo.cs`. Das Ergebnis liegt
unter `artifacts/`. Das ZIP enthält die `DataClean.exe`, die `DataClean.exe.config`
und die Lizenz, daneben liegt die `DataClean.exe` noch einmal einzeln.

## Installation beim Endanwender

Der Installer legt DataClean unter `%ProgramFiles%\DataClean` ab, trägt es ins
Startmenü ein und prüft, ob das .NET Framework 4.8 vorhanden ist. Beim ZIP genügt es,
den Inhalt in einen Ordner zu entpacken. Die einzelne `DataClean.exe` läuft ohne
weitere Dateien mit den eingebauten Standardwerten; um etwa die Ordnerliste zu ändern,
braucht sie die `DataClean.exe.config` aus dem ZIP.

Einstellungen und Profile werden benutzerbezogen unter `%APPDATA%\DataClean` bzw.
`%LOCALAPPDATA%\DataClean` abgelegt.

- **Update:** Ein neuer Installer ersetzt die alte Version und übernimmt die
  Benutzereinstellungen. Die `DataClean.exe.config` wird dabei ersetzt, Änderungen
  von Hand darin gehen also verloren.
- **Anderer Speicherort:** Windows trennt die Benutzereinstellungen nach dem Pfad der
  Exe. Wer von NETLOGON auf die installierte Exe wechselt, beginnt mit den
  Standardwerten.
- **Geplante Aufgaben:** Der Installer legt keine an und entfernt keine, auch nicht
  beim Deinstallieren. Die tägliche Ausführung wird in den Einstellungen aktiviert
  und sollte dort vor dem Deinstallieren wieder abgeschaltet werden.

## Kommandozeilenoptionen

| Option            | Wirkung                                            |
|-------------------|----------------------------------------------------|
| `/autostart`      | Unbeaufsichtigter Lauf ohne UAC-Abfrage            |
| `/dry-run`        | Vorschau-Modus (es wird nichts gelöscht)           |
| `/profile=<Name>` | Benanntes Profil für diesen Lauf anwenden          |
| `/report`         | Report nach dem Lauf erzeugen                      |
| `/verbose`        | Ausführliches Protokoll aktivieren                 |
| `/help`, `/?`     | Hilfe anzeigen                                     |

Eine geplante Aufgabe gilt auch ohne `/autostart` als unbeaufsichtigt: Die
Anwendung löscht dann ohne Rückfrage und schließt sich 7,5 Sekunden nach dem
Lauf von selbst. Offen bleibt sie nur bei drohenden Hardware-Ausfällen (SMART,
Datenträgerstatus, Speicherfehler), einem unerwarteten Abbruch oder solange
jemand Details oder eine andere Seite geöffnet hat. Wenig freier Speicher und
Datenträgerfehler aus dem Ereignisprotokoll werden gemeldet und protokolliert,
halten das Fenster aber nicht offen.

## Hinweis zu MSIX

Eine MSIX-Paketierung ist möglich, erfordert jedoch ein zusätzliches
Windows Application Packaging Project (`.wapproj`) sowie das Windows SDK. Da die
Zielumgebung dieses Repositories ohne .NET-SDK/MSIX-Tooling arbeitet, wird hier
das portable ZIP-Paket als Standard ausgeliefert. Das `.wapproj` kann bei Bedarf
später ergänzt werden, ohne die bestehende Build-Logik zu verändern.

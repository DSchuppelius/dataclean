# DataClean – Paketierung / Veröffentlichung

Dieses Verzeichnis enthält ein Skript, das eine verteilbare Version von DataClean
erzeugt. Es wird bewusst auf ein portables ZIP-Paket gesetzt, da das Projekt auf
.NET Framework 4.8 (WinForms, klassisches csproj) basiert und damit ohne
zusätzliche SDK-Abhängigkeiten gebaut werden kann.

## Voraussetzungen

- Visual Studio 2022/2026 oder die Build Tools mit MSBuild
- Die Komponente ".NET Framework 4.8 SDK/Targeting Pack"

## Paket erstellen

```powershell
# Aus dem Projektstamm ausführen:
./packaging/build-package.ps1
```

Optional kann eine Version mitgegeben werden:

```powershell
./packaging/build-package.ps1 -Version 1.2.0
```

Das Ergebnis liegt anschließend unter `artifacts/DataClean-<Version>.zip` und
enthält die `DataClean.exe`, die `DataClean.exe.config` sowie alle benötigten
Begleitdateien aus dem Release-Ausgabeverzeichnis.

## Installation beim Endanwender

1. ZIP in einen Ordner entpacken (z. B. `%ProgramFiles%\DataClean`).
2. `DataClean.exe` starten. Einstellungen und Profile werden benutzerbezogen
   unter `%APPDATA%\DataClean` bzw. `%LOCALAPPDATA%\DataClean` abgelegt.
3. Für den automatischen Betrieb kann in den Einstellungen die tägliche
   Ausführung aktiviert werden (legt eine geplante Aufgabe an).

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

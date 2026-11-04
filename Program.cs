using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataClean
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // Vor dem ersten Fenster setzen – auch die Hilfe (--help) ist ein App-Dialog.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Das Bereinigen von Systemordnern wie SoftwareDistribution erfordert Administratorrechte.
            // Eine UAC-Abfrage wird aber nur ausgelöst, wenn der Benutzer auch Mitglied der
            // Administratorengruppe ist. Ein echter Standardbenutzer (z. B. beim Autostart) wird
            // nicht mit einer zwecklosen Abfrage behelligt und startet direkt ohne erhöhte Rechte.
            //
            // Beim Autostart (Argument /autostart) wird grundsätzlich KEINE UAC-Abfrage angezeigt,
            // damit im Hintergrund kein störender Dialog aufpoppt. Erhöhte Rechte müssen in diesem
            // Fall über den Aufrufer bereitgestellt werden (z. B. geplante Aufgabe mit RunLevel Highest).
            bool isAutostart = IsAutostart(args);

            // Protokollierungsstufe aus den Einstellungen übernehmen: ausführliches
            // Logging (Info) nur bei Bedarf, ansonsten erst ab Warnungen protokollieren.
            Services.Logger.MinimumLevel = Properties.Settings.Default.verboseLogging
                ? Services.LogLevel.Info
                : Services.LogLevel.Warn;

            // Kommandozeilenoptionen auswerten. --help beendet die Anwendung sofort.
            if (HasFlag(args, "help", "?")) {
                ShowUsage();
                return;
            }

            ApplyCommandLineOverrides(args);

            if (!isAutostart && !IsElevated() && IsMemberOfAdministratorsGroup())
            {
                if (TryRestartElevated())
                {
                    return;
                }
            }

            Application.Run(new DeleteForm(isAutostart));
        }

        private static bool IsAutostart(string[] args)
        {
            if (args == null)
            {
                return false;
            }

            foreach (var arg in args)
            {
                if (string.IsNullOrEmpty(arg))
                {
                    continue;
                }

                string normalized = arg.TrimStart('/', '-');
                if (normalized.Equals("autostart", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        // Prüft, ob eines der angegebenen Flags (ohne führende / bzw. -) vorhanden ist.
        private static bool HasFlag(string[] args, params string[] names) {
            if (args == null) {
                return false;
            }

            foreach (var arg in args) {
                if (string.IsNullOrEmpty(arg)) {
                    continue;
                }
                string normalized = arg.TrimStart('/', '-');
                int eq = normalized.IndexOf('=');
                if (eq >= 0) {
                    normalized = normalized.Substring(0, eq);
                }
                foreach (var name in names) {
                    if (normalized.Equals(name, StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }
            }
            return false;
        }

        // Liefert den Wert einer Option der Form --name=wert bzw. /name:wert.
        private static string GetOptionValue(string[] args, string name) {
            if (args == null) {
                return null;
            }

            foreach (var arg in args) {
                if (string.IsNullOrEmpty(arg)) {
                    continue;
                }
                string normalized = arg.TrimStart('/', '-');
                int sep = normalized.IndexOfAny(new[] { '=', ':' });
                if (sep < 0) {
                    continue;
                }
                string key = normalized.Substring(0, sep);
                if (key.Equals(name, StringComparison.OrdinalIgnoreCase)) {
                    return normalized.Substring(sep + 1).Trim('"');
                }
            }
            return null;
        }

        // Wendet Kommandozeilenoptionen auf die (nur zur Laufzeit gültigen) Einstellungen an,
        // ohne sie dauerhaft zu speichern.
        private static void ApplyCommandLineOverrides(string[] args) {
            try {
                var s = Properties.Settings.Default;

                string profileName = GetOptionValue(args, "profile");
                if (!string.IsNullOrWhiteSpace(profileName)) {
                    var profile = Services.CleanupProfileStore.Get(profileName);
                    if (profile != null) {
                        var o = CleanupOptions.FromSettings();
                        profile.ApplyTo(o);
                        ApplyOptionsInMemory(o);
                        Services.Logger.Info($"Profil '{profileName}' per Kommandozeile geladen.", "Program");
                    } else {
                        Services.Logger.Warn($"Profil '{profileName}' nicht gefunden.", "Program");
                    }
                }

                if (HasFlag(args, "dry-run", "dryrun")) {
                    s.dryRun = true;
                }

                if (HasFlag(args, "report")) {
                    s.generateReport = true;
                }

                if (HasFlag(args, "verbose")) {
                    s.verboseLogging = true;
                    Services.Logger.MinimumLevel = Services.LogLevel.Info;
                }
            } catch (Exception ex) {
                Services.Logger.Error(ex, "Program.ApplyCommandLineOverrides");
            }
        }

        // Überträgt Profilwerte in die laufzeitgültigen Einstellungen (ohne Save()).
        private static void ApplyOptionsInMemory(CleanupOptions o) {
            var s = Properties.Settings.Default;
            s.dryRun = o.DryRun;
            s.createRestorePoint = o.CreateRestorePoint;
            s.deleteToRecycleBin = o.DeleteToRecycleBin;
            s.cleanBrowserCaches = o.CleanBrowserCaches;
            s.dismComponentCleanup = o.DismComponentCleanup;
            s.cleanMinidumps = o.CleanMinidumps;
            s.cleanPrefetch = o.CleanPrefetch;
            s.cleanDeliveryOptimization = o.CleanDeliveryOptimization;
            s.cleanOldLogs = o.CleanOldLogs;
            s.oldLogAgeDays = o.OldLogAgeDays;
            s.emptyAllRecycleBins = o.EmptyAllRecycleBins;
            s.confirmBeforeDelete = o.ConfirmBeforeDelete;
            s.generateReport = o.GenerateReport;
            s.reportPath = o.ReportPath ?? string.Empty;
            s.reportFormat = string.IsNullOrWhiteSpace(o.ReportFormat) ? "HTML" : o.ReportFormat;
            s.emailReport = o.EmailReport;
            s.excludePatterns = o.ExcludePatterns ?? string.Empty;
            s.skipIfEnoughFreeSpace = o.SkipIfEnoughFreeSpace;
            s.schedulePendingDeletes = o.SchedulePendingDeletes;
        }

        private static void ShowUsage() {
            string text =
                "DataClean – Kommandozeilenoptionen" + Environment.NewLine + Environment.NewLine +
                "  /autostart            Unbeaufsichtigter Lauf ohne UAC-Abfrage" + Environment.NewLine +
                "  /dry-run              Vorschau-Modus (es wird nichts gelöscht)" + Environment.NewLine +
                "  /profile=<Name>       Benanntes Profil für diesen Lauf anwenden" + Environment.NewLine +
                "  /report               Report nach dem Lauf erzeugen" + Environment.NewLine +
                "  /verbose              Ausführliches Protokoll aktivieren" + Environment.NewLine +
                "  /help, /?             Diese Hilfe anzeigen";

            UI.AppDialog.Show(text, "DataClean – Hilfe", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static bool IsElevated()
        {
            // Prüft, ob das aktuelle Token bereits erhöht ist (läuft als Administrator).
            try
            {
                using (var identity = WindowsIdentity.GetCurrent())
                {
                    var principal = new WindowsPrincipal(identity);
                    return principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool IsMemberOfAdministratorsGroup()
        {
            // Prüft die Gruppenmitgliedschaft unabhängig von der Elevation. Bei einem nicht
            // erhöhten Admin-Token ist die Administratoren-SID weiterhin in den Token-Gruppen
            // enthalten – nur deaktiviert. Bei einem echten Standardbenutzer fehlt sie.
            try
            {
                using (var identity = WindowsIdentity.GetCurrent())
                {
                    var adminSid = new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null);

                    if (identity.Groups != null)
                    {
                        foreach (var group in identity.Groups)
                        {
                            if (group == adminSid)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Im Zweifel keine UAC-Abfrage erzwingen.
            }

            return false;
        }

        private static bool TryRestartElevated()
        {
            try
            {
                var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;

                var startInfo = new System.Diagnostics.ProcessStartInfo(exePath)
                {
                    UseShellExecute = true,
                    Verb = "runas"
                };

                System.Diagnostics.Process.Start(startInfo);
                return true;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // Benutzer hat die UAC-Abfrage abgebrochen – ohne Adminrechte fortfahren.
                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}

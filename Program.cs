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
            // Das Bereinigen von Systemordnern wie SoftwareDistribution erfordert Administratorrechte.
            // Eine UAC-Abfrage wird aber nur ausgelöst, wenn der Benutzer auch Mitglied der
            // Administratorengruppe ist. Ein echter Standardbenutzer (z. B. beim Autostart) wird
            // nicht mit einer zwecklosen Abfrage behelligt und startet direkt ohne erhöhte Rechte.
            //
            // Beim Autostart (Argument /autostart) wird grundsätzlich KEINE UAC-Abfrage angezeigt,
            // damit im Hintergrund kein störender Dialog aufpoppt. Erhöhte Rechte müssen in diesem
            // Fall über den Aufrufer bereitgestellt werden (z. B. geplante Aufgabe mit RunLevel Highest).
            bool isAutostart = IsAutostart(args);

            if (!isAutostart && !IsElevated() && IsMemberOfAdministratorsGroup())
            {
                if (TryRestartElevated())
                {
                    return;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new DeleteForm());
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

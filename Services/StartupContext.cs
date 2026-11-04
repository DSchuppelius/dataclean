using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace DataClean.Services {
    /// <summary>
    /// Ermittelt, wie die Anwendung gestartet wurde. Eine geplante Aufgabe mit interaktivem Token
    /// läuft in der Benutzersitzung und ist daran nicht von einem Start von Hand zu unterscheiden;
    /// sie gilt trotzdem als unbeaufsichtigt, auch wenn sie kein /autostart übergibt.
    /// </summary>
    public static class StartupContext {
        // Die Aufgabenplanung läuft als Dienst in svchost.exe, unter älteren Windows-Versionen als taskeng.exe.
        private static readonly string[] SchedulerProcessNames = { "svchost", "taskeng", "taskhost", "taskhostw" };

        /// <summary>
        /// True, wenn der Elternprozess die Aufgabenplanung ist. Im Zweifel false: Dann fragt die
        /// Anwendung nach, statt ungefragt zu löschen.
        /// </summary>
        public static bool IsStartedByTaskScheduler() {
            try {
                using (Process current = Process.GetCurrentProcess()) {
                    var info = new NativeMethods.PROCESS_BASIC_INFORMATION();
                    int returned;
                    int status = NativeMethods.NtQueryInformationProcess(current.Handle, 0, ref info,
                        Marshal.SizeOf(typeof(NativeMethods.PROCESS_BASIC_INFORMATION)), out returned);
                    if (status != 0) {
                        return false;
                    }

                    using (Process parent = Process.GetProcessById(info.InheritedFromUniqueProcessId.ToInt32())) {
                        return IsSchedulerProcess(parent.ProcessName);
                    }
                }
            } catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException ||
                                         ex is System.ComponentModel.Win32Exception ||
                                         ex is DllNotFoundException || ex is EntryPointNotFoundException) {
                // Elternprozess bereits beendet oder nicht abfragbar.
                return false;
            }
        }

        /// <summary>Prüft, ob ein Prozessname zur Aufgabenplanung gehört.</summary>
        public static bool IsSchedulerProcess(string processName) {
            return SchedulerProcessNames.Contains(processName ?? string.Empty, StringComparer.OrdinalIgnoreCase);
        }
    }
}

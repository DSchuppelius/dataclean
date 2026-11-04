using System;
using System.IO;
using System.Runtime.InteropServices;

namespace DataClean.Services {
    /// <summary>
    /// Markiert Dateien, die aktuell gesperrt sind und daher nicht sofort gelöscht
    /// werden können, für die Löschung beim nächsten Systemneustart
    /// (<c>MoveFileEx</c> mit <c>MOVEFILE_DELAY_UNTIL_REBOOT</c>). Erfordert in der
    /// Regel Administratorrechte.
    /// </summary>
    public static class PendingDeleteService {
        [Flags]
        private enum MoveFileFlags : uint {
            DelayUntilReboot = 0x00000004
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, MoveFileFlags dwFlags);

        /// <summary>
        /// Plant das Löschen der angegebenen Datei beim nächsten Neustart ein.
        /// Gibt true zurück, wenn die Markierung erfolgreich gesetzt wurde.
        /// </summary>
        public static bool ScheduleDeleteOnReboot(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            try {
                if (!File.Exists(path)) {
                    return false;
                }

                // Zielname null bedeutet: Datei beim Neustart löschen.
                bool ok = MoveFileEx(path, null, MoveFileFlags.DelayUntilReboot);
                if (ok) {
                    Logger.Info($"Zum Löschen beim Neustart markiert: {path}", "PendingDeleteService");
                } else {
                    int err = Marshal.GetLastWin32Error();
                    Logger.Warn($"MoveFileEx fehlgeschlagen (Code {err}) für: {path}", "PendingDeleteService");
                }
                return ok;
            } catch (Exception ex) {
                Logger.Error(ex, "PendingDeleteService.ScheduleDeleteOnReboot");
                return false;
            }
        }
    }
}

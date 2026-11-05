using System;
using System.IO;
using System.Security;
using Microsoft.Win32;

namespace DataClean.Services {
    /// <summary>
    /// Entscheidet, ob die Oberfläche dunkel erscheint: "Light" und "Dark" aus den Einstellungen
    /// gelten unverändert, alles andere ("System", leer) folgt dem App-Modus von Windows.
    /// </summary>
    public static class ThemeMode {
        private const string PersonalizeKey = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
        private const string AppsUseLightThemeValue = "AppsUseLightTheme";

        public static bool IsDark(string themeOverride, bool systemIsDark) {
            if (string.Equals(themeOverride, "Light", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }
            if (string.Equals(themeOverride, "Dark", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
            return systemIsDark;
        }

        /// <summary>
        /// True, wenn in Windows der dunkle App-Modus eingestellt ist. Windows-Versionen ohne
        /// Dunkelmodus kennen den Wert nicht und gelten damit als hell.
        /// </summary>
        public static bool SystemUsesDarkMode() {
            try {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(PersonalizeKey)) {
                    return key?.GetValue(AppsUseLightThemeValue) is int value && value == 0;
                }
            } catch (Exception ex) when (ex is SecurityException || ex is UnauthorizedAccessException || ex is IOException) {
                return false;
            }
        }
    }
}

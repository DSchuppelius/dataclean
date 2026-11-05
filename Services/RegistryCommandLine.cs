using System.Text.RegularExpressions;

namespace DataClean.Services {
    /// <summary>
    /// Liest aus Registry-Werten den Pfad heraus, auf den sie zeigen. Lässt sich das nicht sicher
    /// entscheiden, ist das Ergebnis null: Ein Fund muss belegt sein, im Zweifel gibt es keinen.
    /// </summary>
    public static class RegistryCommandLine {
        private const string Extensions = @"exe|com|bat|cmd|dll|ocx|ax|cpl|scr|sys|msi|msc|vbs|vbe|js|wsf|ps1|lnk|tlb|olb";

        // Beginnt mit Laufwerk oder Umgebungsvariable und endet an der ersten bekannten Dateiendung.
        private static readonly Regex UnquotedPath = new Regex(
            @"^(?<path>(?:[A-Za-z]:\\|%[^%\\]+%\\).*?\.(?:" + Extensions + @"))(?=\s|,|$|\\\d+$)(?<rest>.*)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline);

        private static readonly Regex AnotherPath = new Regex(
            @"\.(?:" + Extensions + @")(?:\s|,|$)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // Beschreibt den Pfad einer MUI-Ressource: @C:\x.dll,-101 oder C:\x.exe.FriendlyAppName
        private static readonly Regex MuiName = new Regex(
            @"^@?(?<path>[^,;]+?)(?:\.FriendlyAppName|\.ApplicationCompany|,-?\d+.*|;.*)?$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>Der Pfad, den eine Befehlszeile startet, oder null, wenn er nicht sicher feststeht.</summary>
        public static string ExtractPath(string commandLine) {
            string text = commandLine?.Trim();
            if (string.IsNullOrEmpty(text)) {
                return null;
            }

            if (text[0] == '"') {
                int end = text.IndexOf('"', 1);
                return end > 1 ? text.Substring(1, end - 1) : null;
            }

            Match match = UnquotedPath.Match(text);
            if (!match.Success) {
                return null;
            }

            // Ohne Anführungszeichen ist die Grenze zwischen Pfad und Argumenten nur eine Vermutung.
            // Folgt noch etwas, das wie ein Programm aussieht, könnte auch das gemeint sein.
            return AnotherPath.IsMatch(match.Groups["rest"].Value) ? null : match.Groups["path"].Value;
        }

        /// <summary>Der Dateipfad aus einem Wertnamen des MUI-Caches, oder null bei allem anderen.</summary>
        public static string ExtractMuiPath(string valueName) {
            if (string.IsNullOrEmpty(valueName)) {
                return null;
            }
            Match match = MuiName.Match(valueName);
            return match.Success ? match.Groups["path"].Value : null;
        }

        /// <summary>Entfernt die Ressourcennummer, die Typbibliotheken anhängen können (x.dll\2).</summary>
        public static string StripResourceIndex(string path) {
            return path == null ? null : Regex.Replace(path, @"\\\d+$", string.Empty);
        }
    }
}

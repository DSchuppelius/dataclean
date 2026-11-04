using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DataClean.Services {
    /// <summary>
    /// Prüft Dateipfade gegen eine Liste von Glob-Ausschlussmustern.
    /// Unterstützt die Platzhalter <c>*</c> (beliebig viele Zeichen, keine
    /// Trennzeichen), <c>**</c> (beliebig viele Zeichen inkl. Trennzeichen) und
    /// <c>?</c> (genau ein Zeichen). Der Vergleich erfolgt ohne Groß-/Kleinschreibung
    /// (Windows-Dateisystem) und sowohl gegen den vollständigen Pfad als auch gegen
    /// den reinen Dateinamen.
    /// </summary>
    public class ExcludeRuleMatcher {
        private readonly List<Regex> _patterns = new List<Regex>();

        public ExcludeRuleMatcher(IEnumerable<string> globPatterns) {
            if (globPatterns == null) {
                return;
            }

            foreach (var pattern in globPatterns) {
                if (string.IsNullOrWhiteSpace(pattern)) {
                    continue;
                }

                try {
                    _patterns.Add(new Regex(GlobToRegex(pattern.Trim()),
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
                } catch (Exception ex) {
                    Logger.Warn($"Ungültiges Ausschlussmuster \"{pattern}\": {ex.Message}", "ExcludeRuleMatcher");
                }
            }
        }

        /// <summary>
        /// Gibt true zurück, wenn der Pfad von mindestens einem Muster erfasst wird.
        /// </summary>
        public bool IsExcluded(string path) {
            if (string.IsNullOrEmpty(path) || _patterns.Count == 0) {
                return false;
            }

            string normalized = path.Replace('/', '\\');
            string fileName = normalized.Substring(normalized.LastIndexOf('\\') + 1);

            return _patterns.Any(r => r.IsMatch(normalized) || r.IsMatch(fileName));
        }

        public bool HasRules => _patterns.Count > 0;

        /// <summary>
        /// Übersetzt ein Glob-Muster in einen regulären Ausdruck.
        /// </summary>
        internal static string GlobToRegex(string glob) {
            var sb = new StringBuilder();
            sb.Append('^');

            for (int i = 0; i < glob.Length; i++) {
                char c = glob[i];
                switch (c) {
                    case '*':
                        if (i + 1 < glob.Length && glob[i + 1] == '*') {
                            sb.Append(".*");
                            i++;
                        } else {
                            sb.Append("[^\\\\/]*");
                        }
                        break;
                    case '?':
                        sb.Append("[^\\\\/]");
                        break;
                    case '/':
                    case '\\':
                        sb.Append("[\\\\/]");
                        break;
                    default:
                        sb.Append(Regex.Escape(c.ToString()));
                        break;
                }
            }

            sb.Append('$');
            return sb.ToString();
        }
    }
}

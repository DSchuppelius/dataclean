using System.Globalization;
using System.Text.RegularExpressions;

namespace DataClean.Services {
    /// <summary>Liest den Prozentwert aus dem Fortschrittsbalken von DISM, z. B. <c>[==== 42.5% ====]</c>.</summary>
    public static class DismProgressParser {
        private static readonly Regex Progress = new Regex(@"\[[=\s\-]*(\d{1,3}(?:[.,]\d+)?)\s*%", RegexOptions.Compiled);

        public static bool TryParsePercent(string line, out double percent) {
            percent = 0d;
            if (string.IsNullOrEmpty(line)) {
                return false;
            }

            Match match = Progress.Match(line);
            if (!match.Success) {
                return false;
            }

            string number = match.Groups[1].Value.Replace(',', '.');
            if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out percent) || percent > 100d) {
                percent = 0d;
                return false;
            }
            return true;
        }
    }
}

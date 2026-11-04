using System;
using System.Text;

namespace DataClean.Services {
    /// <summary>
    /// Enthält die für einen Report benötigten Lauf-Kennzahlen.
    /// </summary>
    public class ReportData {
        public DateTime RunStartTime { get; set; }
        public string MachineName { get; set; }
        public bool DryRun { get; set; }
        public int DeletedItemCount { get; set; }
        public long FreedBytes { get; set; }
        public int SkippedItemCount { get; set; }
        public int FailedItemCount { get; set; }
        public TimeSpan Duration { get; set; }
        public string Tasks { get; set; }
    }

    /// <summary>
    /// Erzeugt den Report-Inhalt (HTML oder CSV) aus den Lauf-Kennzahlen.
    /// Reine Zeichenketten-Erzeugung ohne Datei- oder UI-Zugriff.
    /// </summary>
    public static class ReportBuilder {
        public static string BuildFileName(DateTime runStartTime, bool isHtml) {
            string extension = isHtml ? "html" : "csv";
            return $"DataClean-Report_{runStartTime:yyyyMMdd_HHmmss}.{extension}";
        }

        public static string BuildHtml(ReportData data) {
            return "<!DOCTYPE html><html lang=\"de\"><head><meta charset=\"utf-8\"><title>DataClean Report</title>" +
                "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:2rem;}table{border-collapse:collapse;}td,th{border:1px solid #ccc;padding:6px 12px;text-align:left;}th{background:#f0f0f0;}</style>" +
                "</head><body><h1>DataClean Report</h1><table>" +
                $"<tr><th>Zeitpunkt</th><td>{data.RunStartTime:yyyy-MM-dd HH:mm:ss}</td></tr>" +
                $"<tr><th>Computer</th><td>{System.Net.WebUtility.HtmlEncode(data.MachineName)}</td></tr>" +
                $"<tr><th>Modus</th><td>{(data.DryRun ? "Vorschau (Dry-Run)" : "Löschen")}</td></tr>" +
                $"<tr><th>Gelöschte Elemente</th><td>{data.DeletedItemCount}</td></tr>" +
                $"<tr><th>Freigegebener Speicher</th><td>{ByteFormatter.Format(data.FreedBytes)}</td></tr>" +
                $"<tr><th>Übersprungen</th><td>{data.SkippedItemCount}</td></tr>" +
                $"<tr><th>Fehler</th><td>{data.FailedItemCount}</td></tr>" +
                $"<tr><th>Dauer</th><td>{data.Duration:hh\\:mm\\:ss}</td></tr>" +
                $"<tr><th>Zusatzaufgaben</th><td>{System.Net.WebUtility.HtmlEncode(data.Tasks)}</td></tr>" +
                "</table></body></html>";
        }

        public static string BuildCsv(ReportData data) {
            var sb = new StringBuilder();
            sb.AppendLine("Feld;Wert");
            sb.AppendLine($"Zeitpunkt;{data.RunStartTime:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Computer;{data.MachineName}");
            sb.AppendLine($"Modus;{(data.DryRun ? "Vorschau (Dry-Run)" : "Löschen")}");
            sb.AppendLine($"Gelöschte Elemente;{data.DeletedItemCount}");
            sb.AppendLine($"Freigegebener Speicher;{ByteFormatter.Format(data.FreedBytes)}");
            sb.AppendLine($"Übersprungen;{data.SkippedItemCount}");
            sb.AppendLine($"Fehler;{data.FailedItemCount}");
            sb.AppendLine($"Dauer;{data.Duration:hh\\:mm\\:ss}");
            sb.AppendLine($"Zusatzaufgaben;{data.Tasks}");
            return sb.ToString();
        }
    }
}

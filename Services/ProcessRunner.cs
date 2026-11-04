using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace DataClean.Services {
    /// <summary>Ergebnis eines externen Programmlaufs.</summary>
    public sealed class ProcessResult {
        public ProcessResult(int exitCode, string errorOutput, bool timedOut) {
            ExitCode = exitCode;
            ErrorOutput = errorOutput;
            TimedOut = timedOut;
        }

        public int ExitCode { get; }
        public string ErrorOutput { get; }

        /// <summary>Das Programm war nach Ablauf der Wartezeit noch aktiv und läuft weiter.</summary>
        public bool TimedOut { get; }
    }

    /// <summary>
    /// Startet Konsolenprogramme ohne Fenster und liest beide Ausgaben parallel. Bleibt eine davon
    /// ungelesen, blockiert das Programm, sobald der Puffer voll ist, und die Wartezeit greift nie.
    /// </summary>
    public static class ProcessRunner {
        private const int ReaderDrainMs = 2000;

        /// <param name="onOutputLine">Wird für jede Zeile der Standardausgabe aufgerufen (auf einem Hintergrund-Thread).</param>
        /// <param name="timeout">Maximale Wartezeit. Das Programm wird danach nicht beendet, weil
        /// ein Abbruch laufender Systemwartung mehr Schaden anrichten kann als das Warten.</param>
        /// <exception cref="Win32Exception">Das Programm konnte nicht gestartet werden.</exception>
        public static ProcessResult Run(string fileName, string arguments, Action<string> onOutputLine, TimeSpan timeout) {
            Encoding consoleEncoding = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);
            var process = new Process {
                StartInfo = new ProcessStartInfo(fileName, arguments) {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = consoleEncoding,
                    StandardErrorEncoding = consoleEncoding
                }
            };

            process.Start();

            var errors = new StringBuilder();
            Task outputReader = Task.Run(() => ReadSafely(process.StandardOutput, onOutputLine));
            Task errorReader = Task.Run(() => ReadSafely(process.StandardError, line => {
                lock (errors) {
                    errors.AppendLine(line);
                }
            }));

            if (!process.WaitForExit((int)Math.Min(timeout.TotalMilliseconds, int.MaxValue))) {
                // Nicht freigeben: Die Leser halten den Ausgabepuffer leer, solange das Programm läuft.
                return new ProcessResult(-1, ReadText(errors), true);
            }

            Task.WaitAll(new[] { outputReader, errorReader }, ReaderDrainMs);
            int exitCode = process.ExitCode;
            process.Dispose();
            return new ProcessResult(exitCode, ReadText(errors), false);
        }

        /// <summary>
        /// Zerlegt eine Ausgabe in Zeilen und trennt auch an einem einzelnen Wagenrücklauf, weil
        /// Fortschrittsbalken ihre Zeile damit überschreiben. Ein Balken endet mit ] und ohne
        /// Zeilenende; ohne diese Trennung käme jeder Wert erst mit dem nächsten an.
        /// </summary>
        public static void ReadLines(TextReader reader, Action<string> onLine) {
            var line = new StringBuilder();
            int value;
            while ((value = reader.Read()) >= 0) {
                char c = (char)value;
                if (c != '\r' && c != '\n') {
                    line.Append(c);
                }
                if ((c == '\r' || c == '\n' || c == ']') && line.Length > 0) {
                    onLine?.Invoke(line.ToString());
                    line.Clear();
                }
            }
            if (line.Length > 0) {
                onLine?.Invoke(line.ToString());
            }
        }

        private static void ReadSafely(TextReader reader, Action<string> onLine) {
            try {
                ReadLines(reader, onLine);
            } catch (Exception ex) when (ex is IOException || ex is ObjectDisposedException || ex is InvalidOperationException) {
                // Ausgabe geschlossen oder Empfänger nicht mehr vorhanden, z. B. beim Beenden der Anwendung.
            }
        }

        private static string ReadText(StringBuilder errors) {
            lock (errors) {
                return errors.ToString().Trim();
            }
        }
    }
}

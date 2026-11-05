using System;
using System.Collections.Generic;
using System.IO;
using System.Security;

namespace DataClean.Services {
    public enum RegistryPathState {
        Exists,
        Missing,

        /// <summary>Nicht entscheidbar, etwa bei Netzpfaden, relativen Namen oder fehlendem Zugriff.</summary>
        Unknown
    }

    /// <summary>
    /// Entscheidet, ob eine Datei nachweislich fehlt. Das gilt nur auf einem lokalen, bereiten Laufwerk
    /// und nur, wenn Windows das ausdrücklich meldet: Ein Netzlaufwerk kann gerade nicht verbunden sein,
    /// und ein verweigerter Zugriff sagt nichts über die Datei aus.
    /// </summary>
    public class RegistryPathProbe {
        private static readonly string WindowsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        private readonly Dictionary<char, bool> fixedDrives = new Dictionary<char, bool>();

        /// <param name="rawPath">Pfad, auch mit Umgebungsvariablen.</param>
        /// <param name="judgeWindowsFolder">
        /// Ob auch Pfade im Windows-Ordner als fehlend gelten dürfen. Windows führt dort Einträge für
        /// optionale Komponenten, die zur Zeit nicht installiert sind; das sind keine Reste.
        /// </param>
        public RegistryPathState Check(string rawPath, bool judgeWindowsFolder) {
            if (string.IsNullOrWhiteSpace(rawPath)) {
                return RegistryPathState.Unknown;
            }

            string primary = Expand(rawPath);
            if (primary == null || !IsLocalAbsolute(primary)) {
                return RegistryPathState.Unknown;
            }
            if (primary.IndexOf(@"\WindowsApps\", StringComparison.OrdinalIgnoreCase) >= 0) {
                return RegistryPathState.Unknown;
            }
            if (!judgeWindowsFolder && IsInWindowsFolder(primary)) {
                return RegistryPathState.Unknown;
            }

            RegistryPathState result = CheckFile(primary);
            if (result != RegistryPathState.Missing) {
                return result;
            }

            // 32-Bit-Programme sehen andere Ordner als dieser Prozess. Erst wenn die Datei auch dort
            // fehlt, ist sie wirklich weg.
            foreach (string alternative in Alternatives(rawPath, primary)) {
                RegistryPathState state = CheckFile(alternative);
                if (state != RegistryPathState.Missing) {
                    return state;
                }
            }
            return RegistryPathState.Missing;
        }

        /// <summary>Pfad mit aufgelösten Umgebungsvariablen, oder null, wenn eine davon unbekannt ist.</summary>
        public static string Expand(string rawPath) {
            string expanded = Environment.ExpandEnvironmentVariables(rawPath.Trim());
            return expanded.IndexOf('%') >= 0 ? null : expanded;
        }

        private static IEnumerable<string> Alternatives(string rawPath, string primary) {
            var variables = new[] {
                new[] { "%ProgramFiles%", "%ProgramFiles(x86)%" },
                new[] { "%CommonProgramFiles%", "%CommonProgramFiles(x86)%" }
            };
            foreach (string[] pair in variables) {
                if (rawPath.IndexOf(pair[0], StringComparison.OrdinalIgnoreCase) >= 0) {
                    string alternative = Expand(Replace(rawPath, pair[0], pair[1]));
                    if (alternative != null) {
                        yield return alternative;
                    }
                }
            }

            var folders = new[] {
                new[] { @"\System32\", @"\SysWOW64\" },
                new[] { @"\SysWOW64\", @"\System32\" },
                new[] { @"\Program Files\", @"\Program Files (x86)\" },
                new[] { @"\Program Files (x86)\", @"\Program Files\" }
            };
            foreach (string[] pair in folders) {
                if (primary.IndexOf(pair[0], StringComparison.OrdinalIgnoreCase) >= 0) {
                    yield return Replace(primary, pair[0], pair[1]);
                }
            }
        }

        private static string Replace(string text, string oldValue, string newValue) {
            int index = text.IndexOf(oldValue, StringComparison.OrdinalIgnoreCase);
            return index < 0 ? text : text.Substring(0, index) + newValue + text.Substring(index + oldValue.Length);
        }

        private bool IsLocalAbsolute(string path) {
            if (path.Length < 3 || path[1] != ':' || path[2] != '\\' || !char.IsLetter(path[0])) {
                return false;
            }

            char letter = char.ToUpperInvariant(path[0]);
            bool isFixed;
            if (!fixedDrives.TryGetValue(letter, out isFixed)) {
                try {
                    var drive = new DriveInfo(letter.ToString());
                    isFixed = drive.DriveType == DriveType.Fixed && drive.IsReady;
                } catch (Exception ex) when (ex is ArgumentException || ex is IOException || ex is UnauthorizedAccessException) {
                    isFixed = false;
                }
                fixedDrives[letter] = isFixed;
            }
            return isFixed;
        }

        private static bool IsInWindowsFolder(string path) {
            return path.StartsWith(WindowsFolder + @"\", StringComparison.OrdinalIgnoreCase);
        }

        private static RegistryPathState CheckFile(string path) {
            try {
                File.GetAttributes(path);
                return RegistryPathState.Exists;
            } catch (FileNotFoundException) {
                return RegistryPathState.Missing;
            } catch (DirectoryNotFoundException) {
                return RegistryPathState.Missing;
            } catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException || ex is ArgumentException
                                         || ex is NotSupportedException || ex is SecurityException) {
                return RegistryPathState.Unknown;
            }
        }
    }
}

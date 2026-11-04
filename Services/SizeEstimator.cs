using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace DataClean.Services {
    /// <summary>
    /// Schätzt den insgesamt belegten Speicherplatz einer Menge von Verzeichnissen.
    /// Nicht zugängliche Dateien/Ordner werden übersprungen, sodass die Schätzung
    /// auch bei eingeschränkten Rechten robust bleibt.
    /// </summary>
    public static class SizeEstimator {
        /// <summary>
        /// Berechnet die Gesamtgröße (in Bytes) aller Dateien unterhalb der angegebenen Verzeichnisse.
        /// </summary>
        public static long EstimateSize(IEnumerable<DirectoryInfo> directories, CancellationToken cancellationToken = default(CancellationToken)) {
            long total = 0;

            if (directories == null) {
                return 0;
            }

            foreach (var dir in directories) {
                if (dir == null || !dir.Exists) {
                    continue;
                }
                total += EstimateDirectorySize(dir.FullName, cancellationToken);
            }

            return total;
        }

        private static long EstimateDirectorySize(string root, CancellationToken cancellationToken) {
            long total = 0;
            var stack = new Stack<string>();
            stack.Push(root);

            while (stack.Count > 0) {
                cancellationToken.ThrowIfCancellationRequested();
                string current = stack.Pop();

                try {
                    foreach (var file in Directory.EnumerateFiles(current)) {
                        try {
                            total += new FileInfo(file).Length;
                        } catch (IOException) {
                        } catch (System.UnauthorizedAccessException) {
                        }
                    }
                } catch (IOException) {
                } catch (System.UnauthorizedAccessException) {
                }

                try {
                    foreach (var sub in Directory.EnumerateDirectories(current)) {
                        stack.Push(sub);
                    }
                } catch (IOException) {
                } catch (System.UnauthorizedAccessException) {
                }
            }

            return total;
        }
    }
}

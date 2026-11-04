using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DataClean.Services {
    /// <summary>
    /// Durchläuft einen Ordnerbaum einzeln statt mit <see cref="SearchOption.AllDirectories"/>:
    /// Ein nicht lesbarer Unterordner wird übersprungen, statt den gesamten Bereich abzubrechen.
    /// Verknüpfungen (Junctions, symbolische Links) werden nicht verfolgt, weil ihr Ziel außerhalb
    /// des zu bereinigenden Bereichs liegt und dort nichts gelöscht werden darf.
    /// </summary>
    public static class DirectoryWalker {
        /// <summary>
        /// Liefert alle lesbaren Ordner ab <paramref name="root"/> (einschließlich root).
        /// </summary>
        /// <param name="root">Startordner.</param>
        /// <param name="onUnreadable">Wird für jeden nicht lesbaren Ordner aufgerufen.</param>
        public static IEnumerable<DirectoryInfo> Walk(DirectoryInfo root, Action<DirectoryInfo> onUnreadable = null) {
            var pending = new Stack<DirectoryInfo>();
            pending.Push(root);

            while (pending.Count > 0) {
                DirectoryInfo current = pending.Pop();

                List<DirectoryInfo> children;
                try {
                    children = current.EnumerateDirectories().ToList();
                } catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException) {
                    onUnreadable?.Invoke(current);
                    continue;
                }

                yield return current;

                foreach (DirectoryInfo child in children) {
                    if ((child.Attributes & FileAttributes.ReparsePoint) == 0) {
                        pending.Push(child);
                    }
                }
            }
        }

        /// <summary>
        /// Liefert die Dateien eines einzelnen Ordners; bei fehlendem Zugriff eine leere Liste.
        /// Die Dateiinformationen stammen aus der Enumeration und erfordern keinen weiteren Zugriff.
        /// </summary>
        public static IEnumerable<FileInfo> EnumerateFiles(DirectoryInfo directory) {
            try {
                return directory.EnumerateFiles().ToList();
            } catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException) {
                return Enumerable.Empty<FileInfo>();
            }
        }
    }
}

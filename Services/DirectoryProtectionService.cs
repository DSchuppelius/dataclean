using System;
using System.Collections.Generic;
using System.Linq;

namespace DataClean.Services {
    /// <summary>
    /// Entscheidet, ob ein Verzeichnis geschützt ist (nicht gelöscht werden darf).
    /// Erlaubte Unterverzeichnisse heben den Schutz gezielt wieder auf.
    /// </summary>
    public static class DirectoryProtectionService {
        public static bool IsProtected(string directory, IEnumerable<string> protectedPaths, IEnumerable<string> allowedPaths) {
            if (string.IsNullOrEmpty(directory)) {
                return false;
            }

            if (allowedPaths != null && allowedPaths.Any(allowedDir => directory.StartsWith(allowedDir, StringComparison.OrdinalIgnoreCase))) {
                return false; // Verzeichnis ist erlaubt, also nicht geschützt
            }

            return protectedPaths != null && protectedPaths.Any(protectedDir => directory.StartsWith(protectedDir, StringComparison.OrdinalIgnoreCase));
        }
    }
}

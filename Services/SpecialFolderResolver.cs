using System;
using System.IO;
using System.Linq;

namespace DataClean.Services {
    /// <summary>
    /// Löst Platzhalter-Token wie %TEMP% oder %FIREFOX_CACHE% in konkrete Verzeichnispfade auf.
    /// Unbekannte Werte werden unverändert zurückgegeben.
    /// </summary>
    public static class SpecialFolderResolver {
        public static string Resolve(string path) {
            switch (path) {
                case "%TEMP%":
                    return Path.GetTempPath();
                case "%INTERNET_CACHE%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.InternetCache);
                case "%HISTORY%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.History);
                case "%DOWNLOADS%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                case "%DESKTOP%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                case "%DOCUMENTS%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                case "%MUSIC%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
                case "%PICTURES%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                case "%VIDEOS%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
                case "%WINDOWS%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                case "%SYSTEM%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.System);
                case "%PROGRAM_FILES%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                case "%PROGRAM_FILES_X86%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                case "%USERPROFILE%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                case "%WINDOWS_TEMP%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp");
                case "%PREFETCH%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Prefetch");
                case "%SOFTWARE_DISTRIBUTION%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "Download");
                case "%CRASH_DUMPS%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CrashDumps");
                case "%THUMBNAIL_CACHE%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Windows", "Explorer");
                case "%EDGE_CACHE%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Edge", "User Data", "Default", "Cache");
                case "%CHROME_CACHE%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google", "Chrome", "User Data", "Default", "Cache");
                case "%FIREFOX_CACHE%":
                    return ResolveFirefoxCache();
                default:
                    return path;
            }
        }

        public static string ResolveFirefoxCache() {
            string firefoxProfiles = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Mozilla", "Firefox", "Profiles");
            if (!Directory.Exists(firefoxProfiles)) {
                return firefoxProfiles;
            }

            // Standardprofil (*.default-release) bevorzugen, sonst erstes Profil mit Cache verwenden.
            var profile = Directory.EnumerateDirectories(firefoxProfiles)
                .FirstOrDefault(p => p.EndsWith("default-release", StringComparison.OrdinalIgnoreCase))
                ?? Directory.EnumerateDirectories(firefoxProfiles).FirstOrDefault();

            return profile != null ? Path.Combine(profile, "cache2") : firefoxProfiles;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Threading;
using Microsoft.Win32;

namespace DataClean.Services {
    /// <summary>
    /// Sucht verwaiste Registry-Einträge. Liest nur und ändert nichts. Gemeldet wird ein Eintrag nur,
    /// wenn die Datei, auf die er zeigt, nachweislich fehlt.
    /// </summary>
    public sealed class RegistryScanner {
        private readonly RegistryScope scope;
        private readonly RegistryPathProbe probe;

        public RegistryScanner() : this(RegistryScope.System, new RegistryPathProbe()) {
        }

        public RegistryScanner(RegistryScope scope, RegistryPathProbe probe) {
            this.scope = scope;
            this.probe = probe;
        }

        public List<RegistryIssue> Scan(IEnumerable<RegistryCategory> categories, IProgress<RegistryCategory> progress, CancellationToken cancellationToken) {
            var issues = new List<RegistryIssue>();
            foreach (RegistryCategory category in categories) {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(category);
                foreach (KeyValuePair<RegistryHive, string> location in RegistryLocations.For(category)) {
                    try {
                        using (RegistryKey key = scope.Open(location.Key, location.Value)) {
                            if (key != null) {
                                ScanLocation(category, location.Key, location.Value, key, issues, cancellationToken);
                            }
                        }
                    } catch (Exception ex) when (IsAccessProblem(ex)) {
                        Logger.Warn($"Registry-Prüfung übersprungen ({location.Value}): {ex.Message}", "Registry");
                    }
                }
            }
            return issues;
        }

        private void ScanLocation(RegistryCategory category, RegistryHive hive, string path, RegistryKey key, List<RegistryIssue> issues, CancellationToken token) {
            switch (category) {
                case RegistryCategory.StartupEntries:
                    ScanStartup(hive, path, key, issues);
                    break;
                case RegistryCategory.ApplicationPaths:
                    ForEachSubKey(key, token, (name, sub) => {
                        string target = RegistryCommandLine.ExtractPath(ReadString(sub, string.Empty));
                        if (IsMissing(target, false)) {
                            issues.Add(new RegistryIssue(category, hive, path + @"\" + name, null, name, Display(target)));
                        }
                    });
                    break;
                case RegistryCategory.ObsoleteSoftware:
                    ForEachSubKey(key, token, (name, sub) => ScanUninstallEntry(hive, path, name, sub, issues));
                    break;
                case RegistryCategory.Applications:
                    ForEachSubKey(key, token, (name, sub) => ScanApplication(hive, path, name, sub, issues));
                    break;
                case RegistryCategory.SharedDlls:
                    ScanSharedDlls(hive, path, key, issues);
                    break;
                case RegistryCategory.Fonts:
                    ScanFonts(hive, path, key, issues);
                    break;
                case RegistryCategory.HelpFiles:
                    ScanHelpFiles(hive, path, key, issues);
                    break;
                case RegistryCategory.SoundEvents:
                    ScanSounds(hive, path, key, issues, token);
                    break;
                case RegistryCategory.MuiCache:
                    ScanMuiCache(hive, path, key, issues);
                    break;
                case RegistryCategory.ComClasses:
                    ForEachSubKey(key, token, (name, sub) => ScanComClass(hive, path, name, sub, issues));
                    break;
                case RegistryCategory.TypeLibraries:
                    ForEachSubKey(key, token, (name, sub) => ScanTypeLibrary(hive, path, name, sub, issues));
                    break;
            }
        }

        private void ScanStartup(RegistryHive hive, string path, RegistryKey key, List<RegistryIssue> issues) {
            foreach (string name in key.GetValueNames()) {
                string target = RegistryCommandLine.ExtractPath(ReadString(key, name));
                if (IsMissing(target, false)) {
                    issues.Add(new RegistryIssue(RegistryCategory.StartupEntries, hive, path, name, name, Display(target)));
                }
            }
        }

        private void ScanUninstallEntry(RegistryHive hive, string path, string name, RegistryKey entry, List<RegistryIssue> issues) {
            if (entry.GetValue("SystemComponent") is int hidden && hidden == 1) {
                return;
            }
            string target = RegistryCommandLine.ExtractPath(ReadString(entry, "UninstallString"));
            if (IsMissing(target, false)) {
                string title = ReadString(entry, "DisplayName");
                issues.Add(new RegistryIssue(RegistryCategory.ObsoleteSoftware, hive, path + @"\" + name, null, string.IsNullOrEmpty(title) ? name : title, Display(target)));
            }
        }

        // Ein Eintrag für "Öffnen mit": verwaist, wenn alle hinterlegten Befehle ins Leere zeigen.
        private void ScanApplication(RegistryHive hive, string path, string name, RegistryKey application, List<RegistryIssue> issues) {
            using (RegistryKey shell = application.OpenSubKey("shell")) {
                if (shell == null) {
                    return;
                }
                var targets = new List<string>();
                foreach (string verb in shell.GetSubKeyNames()) {
                    using (RegistryKey command = shell.OpenSubKey(verb + @"\command")) {
                        if (command != null) {
                            targets.Add(RegistryCommandLine.ExtractPath(ReadString(command, string.Empty)));
                        }
                    }
                }
                string missing = AllMissing(targets);
                if (missing != null) {
                    issues.Add(new RegistryIssue(RegistryCategory.Applications, hive, path + @"\" + name, null, name, Display(missing)));
                }
            }
        }

        private void ScanSharedDlls(RegistryHive hive, string path, RegistryKey key, List<RegistryIssue> issues) {
            foreach (string name in key.GetValueNames()) {
                if (IsMissing(name, false)) {
                    issues.Add(new RegistryIssue(RegistryCategory.SharedDlls, hive, path, name, FileNameOf(name), Display(name)));
                }
            }
        }

        private void ScanFonts(RegistryHive hive, string path, RegistryKey key, List<RegistryIssue> issues) {
            string fontsFolder = hive == RegistryHive.LocalMachine
                ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts")
                : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Windows\Fonts");
            foreach (string name in key.GetValueNames()) {
                string file = ReadString(key, name);
                if (string.IsNullOrWhiteSpace(file) || file.IndexOfAny(Path.GetInvalidPathChars()) >= 0) {
                    continue;
                }
                string target = Path.IsPathRooted(file) ? file : Path.Combine(fontsFolder, file);
                if (IsMissing(target, true)) {
                    issues.Add(new RegistryIssue(RegistryCategory.Fonts, hive, path, name, name, Display(target)));
                }
            }
        }

        private void ScanHelpFiles(RegistryHive hive, string path, RegistryKey key, List<RegistryIssue> issues) {
            foreach (string name in key.GetValueNames()) {
                string folder = ReadString(key, name);
                if (string.IsNullOrWhiteSpace(folder) || name.Length == 0
                    || folder.IndexOfAny(Path.GetInvalidPathChars()) >= 0 || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
                    continue;
                }
                string target = Path.Combine(folder.Trim(), name);
                if (IsMissing(target, true)) {
                    issues.Add(new RegistryIssue(RegistryCategory.HelpFiles, hive, path, name, name, Display(target)));
                }
            }
        }

        private void ScanSounds(RegistryHive hive, string path, RegistryKey apps, List<RegistryIssue> issues, CancellationToken token) {
            ForEachSubKey(apps, token, (app, appKey) => {
                ForEachSubKey(appKey, token, (eventName, eventKey) => {
                    using (RegistryKey current = eventKey.OpenSubKey(".Current")) {
                        string target = current == null ? null : ReadString(current, string.Empty);
                        if (IsMissing(target, true)) {
                            string keyPath = path + @"\" + app + @"\" + eventName + @"\.Current";
                            issues.Add(new RegistryIssue(RegistryCategory.SoundEvents, hive, keyPath, string.Empty, eventName, Display(target)));
                        }
                    }
                });
            });
        }

        private void ScanMuiCache(RegistryHive hive, string path, RegistryKey key, List<RegistryIssue> issues) {
            foreach (string name in key.GetValueNames()) {
                string target = RegistryCommandLine.ExtractMuiPath(name);
                if (IsMissing(target, true)) {
                    issues.Add(new RegistryIssue(RegistryCategory.MuiCache, hive, path, name, MuiLabel(name, target), Display(target)));
                }
            }
        }

        // Windows merkt sich je Programm Name und Firma: zwei Werte, die sonst gleich aussähen.
        private static string MuiLabel(string valueName, string target) {
            string file = FileNameOf(target);
            if (valueName.EndsWith(".FriendlyAppName", StringComparison.OrdinalIgnoreCase)) {
                return file + " (Name)";
            }
            return valueName.EndsWith(".ApplicationCompany", StringComparison.OrdinalIgnoreCase) ? file + " (Firma)" : file;
        }

        private void ScanComClass(RegistryHive hive, string path, string name, RegistryKey cls, List<RegistryIssue> issues) {
            var targets = new List<string>();
            foreach (string server in new[] { "InprocServer32", "LocalServer32" }) {
                using (RegistryKey key = cls.OpenSubKey(server)) {
                    if (key == null) {
                        continue;
                    }
                    // Ein gleichnamiger Wert kennzeichnet über den Windows Installer angebotene Klassen.
                    // Der Installer stellt deren Datei bei Bedarf selbst wieder her.
                    if (key.GetValue(server) != null) {
                        return;
                    }
                    targets.Add(RegistryCommandLine.ExtractPath(ReadString(key, string.Empty)));
                }
            }
            string missing = AllMissing(targets);
            if (missing != null) {
                string title = ReadString(cls, string.Empty);
                issues.Add(new RegistryIssue(RegistryCategory.ComClasses, hive, path + @"\" + name, null, string.IsNullOrEmpty(title) ? name : title, Display(missing)));
            }
        }

        private void ScanTypeLibrary(RegistryHive hive, string path, string name, RegistryKey library, List<RegistryIssue> issues) {
            foreach (string version in library.GetSubKeyNames()) {
                using (RegistryKey versionKey = library.OpenSubKey(version)) {
                    if (versionKey == null) {
                        continue;
                    }
                    var targets = new List<string>();
                    foreach (string locale in versionKey.GetSubKeyNames()) {
                        foreach (string platform in new[] { "win32", "win64" }) {
                            using (RegistryKey file = versionKey.OpenSubKey(locale + @"\" + platform)) {
                                if (file != null) {
                                    targets.Add(RegistryCommandLine.StripResourceIndex(RegistryCommandLine.ExtractPath(ReadString(file, string.Empty))));
                                }
                            }
                        }
                    }
                    string missing = AllMissing(targets);
                    if (missing != null) {
                        string title = ReadString(versionKey, string.Empty);
                        issues.Add(new RegistryIssue(RegistryCategory.TypeLibraries, hive, path + @"\" + name + @"\" + version, null, string.IsNullOrEmpty(title) ? name : title, Display(missing)));
                    }
                }
            }
        }

        // Liefert den ersten Pfad, wenn alle nachweislich fehlen. Ein einziger gültiger oder
        // unklarer Pfad genügt, damit der Eintrag bleibt.
        private string AllMissing(List<string> targets) {
            if (targets.Count == 0) {
                return null;
            }
            foreach (string target in targets) {
                if (!IsMissing(target, false)) {
                    return null;
                }
            }
            return targets[0];
        }

        private bool IsMissing(string path, bool judgeWindowsFolder) {
            return path != null && probe.Check(path, judgeWindowsFolder) == RegistryPathState.Missing;
        }

        private static string Display(string path) {
            return RegistryPathProbe.Expand(path) ?? path;
        }

        private static string FileNameOf(string path) {
            int index = path.LastIndexOf('\\');
            return index >= 0 ? path.Substring(index + 1) : path;
        }

        private static string ReadString(RegistryKey key, string name) {
            // Unaufgelöst lesen: 32-Bit-Einträge meinen mit %ProgramFiles% einen anderen Ordner.
            return key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
        }

        private static void ForEachSubKey(RegistryKey key, CancellationToken token, Action<string, RegistryKey> action) {
            foreach (string name in key.GetSubKeyNames()) {
                token.ThrowIfCancellationRequested();
                try {
                    using (RegistryKey sub = key.OpenSubKey(name)) {
                        if (sub != null) {
                            action(name, sub);
                        }
                    }
                } catch (Exception ex) when (IsAccessProblem(ex)) {
                    // Ein gesperrter Unterschlüssel beendet die Prüfung nicht; ohne Lesezugriff gibt es keinen Beleg.
                }
            }
        }

        private static bool IsAccessProblem(Exception ex) {
            return ex is SecurityException || ex is UnauthorizedAccessException || ex is IOException;
        }
    }
}

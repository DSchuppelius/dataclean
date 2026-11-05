using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace DataClean.Services {
    /// <summary>
    /// Die Stellen der Registry, die geprüft werden. Außerhalb dieser Stellen wird nie etwas entfernt,
    /// und ein Schlüssel muss darunter liegen: Die Stelle selbst bleibt immer erhalten.
    /// </summary>
    public static class RegistryLocations {
        public const string CurrentVersion = @"Microsoft\Windows\CurrentVersion";

        private static readonly KeyValuePair<RegistryHive, string>[] SoftwareRoots = {
            new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, "SOFTWARE"),
            new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node"),
            new KeyValuePair<RegistryHive, string>(RegistryHive.CurrentUser, "Software")
        };

        private static readonly KeyValuePair<RegistryHive, string>[] MachineRoots = {
            new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, "SOFTWARE"),
            new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node")
        };

        private static readonly KeyValuePair<RegistryHive, string>[] ClassRoots = {
            new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\Classes"),
            new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node\Classes"),
            new KeyValuePair<RegistryHive, string>(RegistryHive.CurrentUser, @"Software\Classes"),
            new KeyValuePair<RegistryHive, string>(RegistryHive.CurrentUser, @"Software\Classes\WOW6432Node")
        };

        public static IEnumerable<KeyValuePair<RegistryHive, string>> For(RegistryCategory category) {
            switch (category) {
                case RegistryCategory.StartupEntries:
                    return Below(SoftwareRoots, CurrentVersion + @"\Run");
                case RegistryCategory.ApplicationPaths:
                    return Below(SoftwareRoots, CurrentVersion + @"\App Paths");
                case RegistryCategory.ObsoleteSoftware:
                    return Below(SoftwareRoots, CurrentVersion + @"\Uninstall");
                case RegistryCategory.Applications:
                    return Below(new[] { ClassRoots[0], ClassRoots[2] }, "Applications");
                case RegistryCategory.SharedDlls:
                    return Below(MachineRoots, CurrentVersion + @"\SharedDLLs");
                case RegistryCategory.Fonts:
                    return new[] {
                        new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
                        new KeyValuePair<RegistryHive, string>(RegistryHive.CurrentUser, @"Software\Microsoft\Windows NT\CurrentVersion\Fonts")
                    };
                case RegistryCategory.HelpFiles:
                    return new[] {
                        new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\Help"),
                        new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\Microsoft\Windows\HTML Help"),
                        new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\Windows\Help"),
                        new KeyValuePair<RegistryHive, string>(RegistryHive.LocalMachine, @"SOFTWARE\WOW6432Node\Microsoft\Windows\HTML Help")
                    };
                case RegistryCategory.SoundEvents:
                    return new[] { new KeyValuePair<RegistryHive, string>(RegistryHive.CurrentUser, @"AppEvents\Schemes\Apps") };
                case RegistryCategory.MuiCache:
                    return new[] { new KeyValuePair<RegistryHive, string>(RegistryHive.CurrentUser, @"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache") };
                case RegistryCategory.ComClasses:
                    return Below(ClassRoots, "CLSID");
                case RegistryCategory.TypeLibraries:
                    return Below(ClassRoots, "TypeLib");
                default:
                    throw new ArgumentOutOfRangeException(nameof(category));
            }
        }

        /// <summary>True, wenn der Fund an einer geprüften Stelle liegt und entfernt werden darf.</summary>
        public static bool IsAllowed(RegistryIssue issue) {
            foreach (KeyValuePair<RegistryHive, string> location in For(issue.Category)) {
                if (location.Key != issue.Hive) {
                    continue;
                }
                bool below = issue.KeyPath.StartsWith(location.Value + @"\", StringComparison.OrdinalIgnoreCase);
                bool same = string.Equals(issue.KeyPath, location.Value, StringComparison.OrdinalIgnoreCase);
                if (below || (same && !issue.IsWholeKey)) {
                    return true;
                }
            }
            return false;
        }

        private static IEnumerable<KeyValuePair<RegistryHive, string>> Below(IEnumerable<KeyValuePair<RegistryHive, string>> roots, string path) {
            foreach (KeyValuePair<RegistryHive, string> root in roots) {
                yield return new KeyValuePair<RegistryHive, string>(root.Key, root.Value + @"\" + path);
            }
        }
    }
}

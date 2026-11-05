using System.Collections.Generic;
using Microsoft.Win32;

namespace DataClean.Services {
    /// <summary>
    /// Bereiche der Registry-Prüfung. Sie entsprechen den Prüfungen bekannter Registry-Cleaner, soweit
    /// sich ein Fund belegen lässt. Bewusst fehlen Windows-Dienste und Windows Installer (ein Fehlgriff
    /// kann das System oder Reparaturen beschädigen) sowie unbenutzte Dateitypen und die
    /// Startmenü-Sortierung (dort lässt sich nicht nachweisen, dass ein Eintrag verwaist ist).
    /// </summary>
    public enum RegistryCategory {
        StartupEntries,
        ApplicationPaths,
        ObsoleteSoftware,
        Applications,
        SharedDlls,
        Fonts,
        HelpFiles,
        SoundEvents,
        MuiCache,
        ComClasses,
        TypeLibraries
    }

    public static class RegistryCategories {
        public static readonly IReadOnlyList<RegistryCategory> All = new[] {
            RegistryCategory.StartupEntries,
            RegistryCategory.ApplicationPaths,
            RegistryCategory.ObsoleteSoftware,
            RegistryCategory.Applications,
            RegistryCategory.SharedDlls,
            RegistryCategory.Fonts,
            RegistryCategory.HelpFiles,
            RegistryCategory.SoundEvents,
            RegistryCategory.MuiCache,
            RegistryCategory.ComClasses,
            RegistryCategory.TypeLibraries
        };

        public static string Title(RegistryCategory category) {
            switch (category) {
                case RegistryCategory.StartupEntries: return "Autostart";
                case RegistryCategory.ApplicationPaths: return "Anwendungspfade";
                case RegistryCategory.ObsoleteSoftware: return "Veraltete Software";
                case RegistryCategory.Applications: return "Anwendungen";
                case RegistryCategory.SharedDlls: return "Gemeinsame DLLs";
                case RegistryCategory.Fonts: return "Schriftarten";
                case RegistryCategory.HelpFiles: return "Hilfedateien";
                case RegistryCategory.SoundEvents: return "Soundereignisse";
                case RegistryCategory.MuiCache: return "MUI-Cache";
                case RegistryCategory.ComClasses: return "ActiveX und COM";
                default: return "Typbibliotheken";
            }
        }

        /// <summary>
        /// Bereiche, in denen ein Eintrag noch von anderen Programmen abhängen kann. Sie werden nur
        /// auf ausdrücklichen Wunsch geprüft, ihre Funde sind nicht vorausgewählt.
        /// </summary>
        public static bool IsCautious(RegistryCategory category) {
            return category == RegistryCategory.ComClasses || category == RegistryCategory.TypeLibraries;
        }
    }

    /// <summary>Ein verwaister Registry-Eintrag: entweder ein ganzer Schlüssel oder ein einzelner Wert.</summary>
    public sealed class RegistryIssue {
        public RegistryIssue(RegistryCategory category, RegistryHive hive, string keyPath, string valueName, string name, string missingPath) {
            Category = category;
            Hive = hive;
            KeyPath = keyPath;
            ValueName = valueName;
            Name = name;
            MissingPath = missingPath;
        }

        public RegistryCategory Category { get; }
        public RegistryHive Hive { get; }
        public string KeyPath { get; }

        /// <summary>Name des Werts ("" = Standardwert); null bedeutet: der ganze Schlüssel ist verwaist.</summary>
        public string ValueName { get; }

        public bool IsWholeKey {
            get { return ValueName == null; }
        }

        /// <summary>Bezeichnung für die Anzeige.</summary>
        public string Name { get; }

        /// <summary>Die Datei, deren Fehlen den Fund begründet.</summary>
        public string MissingPath { get; }
    }
}

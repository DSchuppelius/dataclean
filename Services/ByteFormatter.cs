namespace DataClean.Services {
    /// <summary>
    /// Formatiert Byte-Größen in eine menschenlesbare Darstellung (B, KB, MB, GB, TB).
    /// </summary>
    public static class ByteFormatter {
        private static readonly string[] Units = { "B", "KB", "MB", "GB", "TB" };

        public static string Format(long bytes) {
            double size = bytes;
            int unit = 0;
            while (size >= 1024 && unit < Units.Length - 1) {
                size /= 1024;
                unit++;
            }
            return $"{size:0.##} {Units[unit]}";
        }
    }
}

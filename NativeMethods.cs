using System;
using System.Runtime.InteropServices;

namespace DataClean {
    /// <summary>
    /// Gebündelte native Win32-API-Aufrufe (P/Invoke) für die Anwendung.
    /// Entspricht der CA1060-Empfehlung, P/Invoke-Deklarationen zentral zu halten.
    /// </summary>
    internal static class NativeMethods {
        // Elternprozess ermitteln, um den Start durch die Aufgabenplanung zu erkennen.
        [StructLayout(LayoutKind.Sequential)]
        internal struct PROCESS_BASIC_INFORMATION {
            public IntPtr ExitStatus;
            public IntPtr PebBaseAddress;
            public IntPtr AffinityMask;
            public IntPtr BasePriority;
            public IntPtr UniqueProcessId;
            public IntPtr InheritedFromUniqueProcessId;
        }

        [DllImport("ntdll.dll")]
        internal static extern int NtQueryInformationProcess(IntPtr processHandle, int informationClass,
            ref PROCESS_BASIC_INFORMATION information, int informationLength, out int returnLength);

        // Papierkorb leeren (shell32.dll)
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

        // Weist einem Steuerelement ein Windows-Theme zu (z. B. dunkle Scrollbars/ProgressBar im Dark-Mode).
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        internal static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        // Fensterverschiebung/-handling für randlose Fenster (Custom-Chrome).
        [DllImport("user32.dll")]
        internal static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        internal const int WM_NCLBUTTONDOWN = 0xA1;
        internal const int HT_CAPTION = 0x2;

        // Zeichnen des Fensterrahmens für themenkonforme Eingabefelder.
        [DllImport("user32.dll")]
        internal static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        internal static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        internal static extern int GetWindowLong(IntPtr hWnd, int index);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        internal static extern int SetWindowLong(IntPtr hWnd, int index, int value);

        internal const int GWL_STYLE = -16;
        internal const int WS_HSCROLL = 0x00100000;

        internal const uint RDW_INVALIDATE = 0x0001;
        internal const uint RDW_FRAME = 0x0400;

        // Abgerundete Fensterecken und Rahmenfarbe unter Windows 11 (auf älteren Systemen wirkungslos).
        [DllImport("dwmapi.dll")]
        internal static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        internal const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        internal const int DWMWA_BORDER_COLOR = 34;
        internal const int DWMWCP_ROUND = 2;

        // SHFileOperation zum Verschieben in den Papierkorb
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct SHFILEOPSTRUCT {
            public IntPtr hwnd;
            public uint wFunc;
            public string pFrom;
            public string pTo;
            public ushort fFlags;
            public int fAnyOperationsAborted;
            public IntPtr hNameMappings;
            public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);

        internal const uint FO_DELETE = 0x0003;
        internal const ushort FOF_ALLOWUNDO = 0x0040;
        internal const ushort FOF_NOCONFIRMATION = 0x0010;
        internal const ushort FOF_NOERRORUI = 0x0400;
        internal const ushort FOF_SILENT = 0x0004;

        // Flags für die SHEmptyRecycleBin-Funktion
        internal const uint SHERB_NOCONFIRMATION = 0x00000001;  // Kein Bestätigungsdialog
        internal const uint SHERB_NOPROGRESSUI = 0x00000002;    // Keine Fortschrittsanzeige
        internal const uint SHERB_NOSOUND = 0x00000004;         // Kein Sound nach dem Löschen
    }
}

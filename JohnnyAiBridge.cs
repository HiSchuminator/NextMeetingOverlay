using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace NextMeetingOverlay;

/// <summary>
/// Toggles JohnnyAI recording by posting the same WM_HOTKEY message its global
/// Win+Shift+R hotkey produces (HOTKEY_ID 9001 in JohnnyAI MainWindow.xaml.cs).
/// It is a toggle: JohnnyAI exposes no recording state to other processes.
/// </summary>
public static class JohnnyAiBridge
{
    private const int WM_HOTKEY = 0x0312;
    private const int JohnnyHotkeyId = 9001;

    public static bool IsRunning => FindWpfWindows().Count > 0;

    /// <summary>
    /// Posts to every WPF window of JohnnyAI; only its MainWindow hook handles hotkey id 9001,
    /// so this works regardless of the main window's current title or visibility.
    /// </summary>
    public static bool ToggleRecording()
    {
        bool posted = false;
        foreach (var hwnd in FindWpfWindows())
            posted |= PostMessage(hwnd, WM_HOTKEY, (IntPtr)JohnnyHotkeyId, IntPtr.Zero);
        return posted;
    }

    private static List<IntPtr> FindWpfWindows()
    {
        var result = new List<IntPtr>();
        var pids = new HashSet<uint>();
        foreach (var p in Process.GetProcessesByName("JohnnyAI"))
        {
            pids.Add((uint)p.Id);
            p.Dispose();
        }
        if (pids.Count == 0) return result;

        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var pid);
            if (!pids.Contains(pid)) return true;
            var cls = new StringBuilder(256);
            GetClassName(h, cls, cls.Capacity);
            if (cls.ToString().StartsWith("HwndWrapper", StringComparison.Ordinal))
                result.Add(h);
            return true;
        }, IntPtr.Zero);
        return result;
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder name, int max);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}

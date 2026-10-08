using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Win32;

namespace NextMeetingOverlay;

public enum DockMode { Free, TopLeft, TopCenter, TopRight }

/// <summary>What happens with JohnnyAI when a meeting begins.</summary>
public enum JohnnyStartMode { Off, Ask, Auto }

/// <summary>Persisted in %APPDATA%\NextMeetingOverlay\settings.json.</summary>
public sealed class Settings
{
    public DockMode Dock { get; set; } = DockMode.TopRight;
    public string? ScreenDevice { get; set; }
    public double Left { get; set; } = 100;
    public double Top { get; set; } = 100;
    public double Opacity { get; set; } = 0.95;

    // Filters
    public bool HideFree { get; set; } = true;
    public bool HideKeywordMatches { get; set; } = true;
    public List<string> ExcludeKeywords { get; set; } = new()
    {
        "Brunch", "Mittag", "Lunch", "Pause", "Blocker", "Fokuszeit", "Focus time"
    };
    public bool HideAllDay { get; set; } = true;
    public bool HideTentative { get; set; } = false;
    public bool HideDeclined { get; set; } = true;
    public bool HidePrivate { get; set; } = false;

    // Display
    public bool ShowFollowUp { get; set; } = true;
    public bool Compact { get; set; } = false;
    public bool DetectJoinLink { get; set; } = true;
    public bool ShowJohnnyButton { get; set; } = true;
    public bool ShowDayOverview { get; set; } = true;
    public bool ShowOverlaps { get; set; } = true;

    // Mini mode: small traffic-light pill, details on hover/click
    public bool MiniMode { get; set; } = true;
    public bool AutoExpandAtStart { get; set; } = true;
    public int YellowMinutes { get; set; } = 120;
    public int RedMinutes { get; set; } = 15;
    public int LookaheadDays { get; set; } = 7;
    public int RefreshSeconds { get; set; } = 60;

    // JohnnyAI
    public JohnnyStartMode JohnnyStart { get; set; } = JohnnyStartMode.Ask;
    public bool JohnnyOnlyOnlineMeetings { get; set; } = true;

    // Visibility & behaviour
    public bool HideFromScreenShare { get; set; } = true;
    public bool HideInFullscreen { get; set; } = true;
    public bool ClickThrough { get; set; } = false;
    public bool HotkeysEnabled { get; set; } = true;

    // Calendar source: empty = classic Outlook via COM, otherwise a published ICS link.
    public string? IcsUrlProtected { get; set; }
    public int IcsRefreshMinutes { get; set; } = 5;

    /// <summary>The ICS link is a secret (anyone with it can read the calendar) -> DPAPI, current user.</summary>
    [JsonIgnore]
    public string? IcsUrl
    {
        get
        {
            if (string.IsNullOrEmpty(IcsUrlProtected)) return null;
            try
            {
                var bytes = ProtectedData.Unprotect(Convert.FromBase64String(IcsUrlProtected), null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(bytes);
            }
            catch { return null; }
        }
        set
        {
            IcsUrlProtected = string.IsNullOrWhiteSpace(value)
                ? null
                : Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(value.Trim()), null, DataProtectionScope.CurrentUser));
        }
    }

    /// <summary>"--profile name" runs a separate instance with its own settings (testing).</summary>
    public static string? Profile { get; set; }

    private static string Dir => Path.Combine(BaseDir.Value, Profile == null ? "" : "profiles\\" + Profile);

    /// <summary>
    /// Portable build: settings live in a folder next to the exe so they travel with it
    /// (USB stick, network share). Falls back to %APPDATA% when that folder isn't writable.
    /// </summary>
    private static readonly Lazy<string> BaseDir = new(() =>
    {
        var appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "NextMeetingOverlay");
#if PORTABLE
        var portable = Path.Combine(AppContext.BaseDirectory, "NextMeetingOverlay-Daten");
        try
        {
            Directory.CreateDirectory(portable);
            var probe = Path.Combine(portable, ".write-test");
            File.WriteAllText(probe, "");
            File.Delete(probe);
            return portable;
        }
        catch { return appData; }
#else
        return appData;
#endif
    });
    private static string FilePath => Path.Combine(Dir, "settings.json");

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public static Settings Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath), Options) ?? new Settings();
        }
        catch { /* corrupt file -> defaults */ }
        return new Settings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Dir);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(this, Options));
        }
        catch { /* not critical */ }
    }

    /// <summary>Snapshot handed to the Outlook thread so UI toggles can't race with a running read.</summary>
    public Settings Clone() => JsonSerializer.Deserialize<Settings>(JsonSerializer.Serialize(this, Options), Options)!;
}

public static class Autostart
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "NextMeetingOverlay";

    public static bool IsEnabled
    {
        get
        {
            using var key = Registry.CurrentUser.OpenSubKey(RunKey);
            return key?.GetValue(ValueName) is string;
        }
    }

    public static void Set(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKey);
        if (enabled) key.SetValue(ValueName, $"\"{Environment.ProcessPath}\"");
        else key.DeleteValue(ValueName, false);
    }
}

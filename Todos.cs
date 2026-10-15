using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NextMeetingOverlay;

public enum TodoReminder { None, AfterMeeting, At }

/// <summary>A personal reminder ("was ich erledigen möchte"), optionally due after a meeting or at a time.</summary>
public sealed class TodoItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Text { get; set; } = "";
    public DateTime Created { get; set; } = DateTime.Now;
    public bool Done { get; set; }
    public DateTime? DoneAt { get; set; }
    public TodoReminder Reminder { get; set; }
    /// <summary>Chosen time, or the meeting's end for <see cref="TodoReminder.AfterMeeting"/>.</summary>
    public DateTime? RemindAt { get; set; }
    public string? MeetingSubject { get; set; }
    /// <summary>The reminder balloon was shown already.</summary>
    public bool Reminded { get; set; }

    [JsonIgnore]
    public bool IsDue => !Done && RemindAt is DateTime due && due <= DateTime.Now;

    [JsonIgnore]
    public string ReminderLabel
    {
        get
        {
            if (Reminder == TodoReminder.None || RemindAt is not DateTime at) return "";
            if (IsDue) return "fällig";
            if (Reminder == TodoReminder.AfterMeeting) return $"nach {MeetingSubject}";
            var today = DateTime.Today;
            return at.Date == today ? $"{at:HH:mm}"
                 : at.Date == today.AddDays(1) ? $"morgen {at:HH:mm}"
                 : at.ToString("ddd dd.MM. HH:mm", CultureInfo.CurrentCulture);
        }
    }
}

/// <summary>todos.json next to settings.json (portable build: next to the exe).</summary>
public static class TodoStore
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    private static string FilePath => Path.Combine(Settings.DataDirectory, "todos.json");

    public static List<TodoItem> Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var todos = JsonSerializer.Deserialize<List<TodoItem>>(File.ReadAllText(FilePath), Options) ?? new();
                // Done items are only kept a week (undo / "what did I do"), then forgotten.
                todos.RemoveAll(t => t.Done && (t.DoneAt ?? t.Created) < DateTime.Now.AddDays(-7));
                return todos;
            }
        }
        catch { /* corrupt file -> start empty, file is rewritten on the next change */ }
        return new();
    }

    public static void Save(List<TodoItem> todos)
    {
        try
        {
            Directory.CreateDirectory(Settings.DataDirectory);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(todos, Options));
        }
        catch { /* not critical */ }
    }
}

/// <summary>
/// Builds a calendar file (.ics) for a todo and opens it, so Outlook (classic or new) shows a
/// prefilled appointment - the user picks the calendar and saves. Nothing is written to Outlook
/// directly (the new Outlook has no API for that) and nothing is sent to anyone.
/// </summary>
public static class OutlookReminderFile
{
    public static string Create(TodoItem todo, DateTime start, bool isPrivate, bool showAsFree)
    {
        var end = start.AddMinutes(15);
        string Utc(DateTime d) => d.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);

        var ics = new StringBuilder()
            .Append("BEGIN:VCALENDAR\r\n")
            .Append("VERSION:2.0\r\n")
            .Append("PRODID:-//NextMeetingOverlay//Todo//DE\r\n")
            .Append("METHOD:PUBLISH\r\n")
            .Append("BEGIN:VEVENT\r\n")
            .Append($"UID:todo-{todo.Id}@nextmeetingoverlay\r\n")
            .Append($"DTSTAMP:{Utc(DateTime.Now)}\r\n")
            .Append($"DTSTART:{Utc(start)}\r\n")
            .Append($"DTEND:{Utc(end)}\r\n")
            .Append($"SUMMARY:{Escape("Todo: " + todo.Text)}\r\n")
            .Append($"DESCRIPTION:{Escape("Erinnerung aus „Nächster Termin“" + (todo.MeetingSubject != null ? $" (nach {todo.MeetingSubject})" : ""))}\r\n")
            .Append(isPrivate ? "CLASS:PRIVATE\r\n" : "CLASS:PUBLIC\r\n")
            .Append(showAsFree ? "TRANSP:TRANSPARENT\r\nX-MICROSOFT-CDO-BUSYSTATUS:FREE\r\n" : "TRANSP:OPAQUE\r\nX-MICROSOFT-CDO-BUSYSTATUS:BUSY\r\n")
            .Append("BEGIN:VALARM\r\n")
            .Append("ACTION:DISPLAY\r\n")
            .Append($"DESCRIPTION:{Escape(todo.Text)}\r\n")
            .Append("TRIGGER:-PT0M\r\n")
            .Append("END:VALARM\r\n")
            .Append("END:VEVENT\r\n")
            .Append("END:VCALENDAR\r\n")
            .ToString();

        var dir = Path.Combine(Path.GetTempPath(), "NextMeetingOverlay");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, $"Todo-{todo.Id[..8]}.ics");
        File.WriteAllText(path, ics, new UTF8Encoding(false));
        return path;
    }

    private static string Escape(string text) =>
        text.Replace("\\", "\\\\").Replace(";", "\\;").Replace(",", "\\,").Replace("\r\n", "\\n").Replace("\n", "\\n");
}

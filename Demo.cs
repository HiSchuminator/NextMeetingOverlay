using System;
using System.Collections.Generic;

namespace NextMeetingOverlay;

/// <summary>
/// "--demo-screenshot file.png": runs with made-up meetings (no calendar access, no personal
/// data), renders the mini pill, the expanded card and the context menu into one PNG and exits.
/// Used for the README screenshot.
/// </summary>
internal static class Demo
{
    public static string? ExportPath { get; set; }
    public static bool Active => ExportPath != null;

    /// <summary>JohnnyAI shown as recording, so the pill and card show the recording state.</summary>
    public static JohnnyStatus Johnny => new(true, "04:51", "Projekt-Daily", false);

    public static List<TodoItem> Todos()
    {
        var now = DateTime.Now;
        return new List<TodoItem>
        {
            new() { Text = "Angebot an Kunden schicken", Reminder = TodoReminder.AfterMeeting, MeetingSubject = "Projekt-Daily", RemindAt = now.AddMinutes(10) },
            new() { Text = "Reisekosten einreichen", Kind = TodoKind.Appointment, Reminder = TodoReminder.At, RemindAt = now.Date.AddDays(1).AddHours(9) },
        };
    }

    public static List<Meeting> Meetings()
    {
        var now = DateTime.Now;
        // Round to 5 minutes: a running daily, a workshop starting in parallel, something later.
        var s = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute / 5 * 5, 0).AddMinutes(-5);
        const string teams = "https://teams.microsoft.com/l/meetup-join/demo";
        return new List<Meeting>
        {
            new("Projekt-Daily", s, s.AddMinutes(15), "Microsoft Teams-Besprechung", teams, "", "", MeetingSource.Ics),
            new("Workshop",s.AddMinutes(10), s.AddMinutes(70), "Raum 2.04", teams, "", "", MeetingSource.Ics),
            new("Release-Planung", s.AddMinutes(90), s.AddMinutes(120), "Microsoft Teams-Besprechung", teams, "", "", MeetingSource.Ics),
        };
    }
}

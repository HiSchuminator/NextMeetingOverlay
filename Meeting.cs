using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace NextMeetingOverlay;

public enum MeetingSource { Outlook, Ics }

public sealed record Meeting(
    string Subject, DateTime Start, DateTime End, string Location,
    string? JoinUrl, string EntryId, string StoreId, MeetingSource Source, string CalendarName = "")
{
    /// <summary>Identifies one occurrence (used to ask/auto-start JohnnyAI only once per meeting).</summary>
    public string Key => $"{Start:O}|{Subject}";
}

/// <summary>Filter rules shared by the Outlook (COM) and ICS sources.</summary>
public static class MeetingFilter
{
    // Outlook marks a cancelled occurrence of a series only through its subject.
    private static readonly string[] CancelledPrefixes = { "Abgesagt:", "Canceled:", "Cancelled:" };

    private static readonly Regex JoinLinkRx = new(
        @"https://(?:teams\.microsoft\.com/(?:l/meetup-join|meet)|teams\.live\.com/meet|[\w.-]*zoom\.us/j|meet\.google\.com|[\w.-]*webex\.com/)[^\s""<>]*",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool HasCancelledPrefix(string subject) =>
        CancelledPrefixes.Any(p => subject.TrimStart().StartsWith(p, StringComparison.OrdinalIgnoreCase));

    public static bool MatchesKeyword(Settings s, string subject) =>
        s.HideKeywordMatches && s.ExcludeKeywords.Any(k =>
            !string.IsNullOrWhiteSpace(k) && subject.Contains(k.Trim(), StringComparison.OrdinalIgnoreCase));

    public static string? FindJoinLink(string text)
    {
        var match = JoinLinkRx.Match(text);
        return match.Success ? match.Value.TrimEnd('>', ')', '.', ',') : null;
    }
}

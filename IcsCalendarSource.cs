using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;

namespace NextMeetingOverlay;

/// <summary>
/// Reads a published Outlook/Exchange calendar (ICS link from Outlook on the web →
/// Settings → Calendar → Shared calendars → Publish a calendar). Works with the new
/// Outlook, which has no COM interface. The feed is downloaded every few minutes and
/// expanded (incl. recurring series) locally on each refresh.
/// </summary>
public sealed class IcsCalendarSource
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(30) };

    private Calendar? _calendar;
    private string? _url;
    private DateTime _fetchedAt;

    public void Invalidate() => _calendar = null;

    public async Task<List<Meeting>> ReadAsync(Settings s, string url)
    {
        url = url.Trim();
        if (url.StartsWith("webcal://", StringComparison.OrdinalIgnoreCase))
            url = "https://" + url["webcal://".Length..];

        bool stale = _calendar == null || _url != url ||
                     DateTime.Now - _fetchedAt > TimeSpan.FromMinutes(Math.Max(1, s.IcsRefreshMinutes));
        if (stale)
        {
            try
            {
                var text = await Http.GetStringAsync(url);
                _calendar = await Task.Run(() => Calendar.Load(text));
                _url = url;
                _fetchedAt = DateTime.Now;
            }
            catch when (_calendar != null && _url == url)
            {
                // Network hiccup: keep working with the last downloaded feed.
            }
        }

        var calendar = _calendar!;
        return await Task.Run(() => Expand(calendar, s));
    }

    private static List<Meeting> Expand(Calendar calendar, Settings s)
    {
        var now = DateTime.Now;
        var from = now.Date.AddDays(-1);
        var until = now.Date.AddDays(s.LookaheadDays + 1);

        // Modified occurrences of a series come as separate events with RECURRENCE-ID;
        // their original slot must not be produced again by the series' rule.
        var overridden = new HashSet<(string, DateTime)>();
        foreach (var ev in calendar.Events)
            if (ev.RecurrenceId != null)
                overridden.Add((ev.Uid, ev.RecurrenceId.AsSystemLocal));

        var result = new List<Meeting>();
        foreach (var ev in calendar.Events)
        {
            if (ev.DtStart == null) continue;
            try
            {
                var duration = ev.DtEnd != null ? ev.DtEnd.AsSystemLocal - ev.DtStart.AsSystemLocal : ev.Duration;
                if (ev.RecurrenceRules.Count > 0 && ev.RecurrenceId == null)
                {
                    foreach (var occurrence in ev.GetOccurrences(new CalDateTime(from), new CalDateTime(until)))
                    {
                        var start = occurrence.Period.StartTime.AsSystemLocal;
                        if (overridden.Contains((ev.Uid, start))) continue;
                        Add(result, ev, start, start + duration, s, now, until);
                    }
                }
                else
                {
                    var start = ev.DtStart.AsSystemLocal;
                    Add(result, ev, start, start + duration, s, now, until);
                }
            }
            catch { /* one broken event must not hide the whole calendar */ }
        }

        return result.OrderBy(m => m.Start).Take(OutlookReader.MaxMeetings).ToList();
    }

    private static void Add(List<Meeting> result, CalendarEvent ev, DateTime start, DateTime end,
        Settings s, DateTime now, DateTime until)
    {
        if (end <= now || start >= until) return;

        string subject = (ev.Summary ?? "").Trim();
        string busy = (Prop(ev, "X-MICROSOFT-CDO-BUSYSTATUS") ?? "").ToUpperInvariant();
        bool free = busy == "FREE" || (busy.Length == 0 && string.Equals(Prop(ev, "TRANSP"), "TRANSPARENT", StringComparison.OrdinalIgnoreCase));
        bool allDay = ev.IsAllDay || !ev.DtStart.HasTime ||
                      string.Equals(Prop(ev, "X-MICROSOFT-CDO-ALLDAYEVENT"), "TRUE", StringComparison.OrdinalIgnoreCase);
        bool cancelled = string.Equals(Prop(ev, "STATUS"), "CANCELLED", StringComparison.OrdinalIgnoreCase) ||
                         MeetingFilter.HasCancelledPrefix(subject);
        bool isPrivate = string.Equals(Prop(ev, "CLASS"), "PRIVATE", StringComparison.OrdinalIgnoreCase);

        if (s.HideAllDay && allDay) return;
        if (s.HideFree && free) return;
        if (s.HideTentative && busy == "TENTATIVE") return;
        if (s.HideDeclined && cancelled) return;
        if (s.HidePrivate && isPrivate) return;
        if (MeetingFilter.MatchesKeyword(s, subject)) return;

        string location = (ev.Location ?? "").Trim();
        string? joinUrl = null;
        if (s.DetectJoinLink)
            joinUrl = Prop(ev, "X-MICROSOFT-SKYPETEAMSMEETINGURL")
                      ?? MeetingFilter.FindJoinLink(location)
                      ?? MeetingFilter.FindJoinLink(ev.Description ?? "");

        result.Add(new Meeting(subject, start, end, location, joinUrl, "", "", MeetingSource.Ics));
    }

    private static string? Prop(CalendarEvent ev, string name) =>
        ev.Properties.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase))?.Value?.ToString();
}

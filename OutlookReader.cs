using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace NextMeetingOverlay;

public sealed class OutlookNotRunningException : Exception { }

/// <summary>
/// Reads the default Outlook calendar via late-bound COM on a dedicated STA thread.
/// Only attaches to an already running Outlook (never spawns a new headless "OUTLOOK -Embedding").
/// No reference to Outlook is kept between reads, so closing Outlook is never blocked.
/// </summary>
public sealed class OutlookReader : IDisposable
{
    private const int OlFolderCalendar = 9;
    private const int OlAppointment = 26;
    private const int OlFree = 0, OlTentative = 1;
    private const int OlPrivate = 2;
    private const int OlResponseDeclined = 4;
    private const int OlMeetingCanceled = 5, OlMeetingReceivedAndCanceled = 7;
    /// <summary>Enough for the rest of a busy day (hover overview) plus the next days.</summary>
    public const int MaxMeetings = 30;

    private readonly BlockingCollection<Action> _queue = new();

    public OutlookReader()
    {
        var thread = new Thread(() =>
        {
            foreach (var work in _queue.GetConsumingEnumerable())
                work();
        }) { IsBackground = true, Name = "Outlook COM" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public Task<T> Run<T>(Func<T> func)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add(() =>
        {
            try { tcs.SetResult(func()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        return tcs.Task;
    }

    public List<Meeting> ReadMeetings(Settings s)
    {
        object app = GetRunningOutlook();
        object? ns = null, folder = null, items = null, restricted = null;
        try
        {
            ns = ((dynamic)app).GetNamespace("MAPI");
            folder = ((dynamic)ns).GetDefaultFolder(OlFolderCalendar);
            string storeId = ((dynamic)folder).StoreID;
            items = ((dynamic)folder).Items;

            // Order matters: Sort before IncludeRecurrences, then Restrict.
            ((dynamic)items).Sort("[Start]");
            ((dynamic)items).IncludeRecurrences = true;

            var now = DateTime.Now;
            var until = now.Date.AddDays(s.LookaheadDays + 1);
            string filter =
                $"[Start] < '{until.ToString("g", CultureInfo.CurrentCulture)}' AND " +
                $"[End] > '{now.AddMinutes(-1).ToString("g", CultureInfo.CurrentCulture)}'";
            restricted = ((dynamic)items).Restrict(filter);

            var result = new List<Meeting>();
            object? item = ((dynamic)restricted).GetFirst();
            int guard = 0;
            while (item != null && guard++ < 1000 && result.Count < MaxMeetings)
            {
                try
                {
                    DateTime start = ((dynamic)item).Start;
                    if (start >= until) break;
                    var meeting = ToMeeting(item, s, now, storeId);
                    if (meeting != null) result.Add(meeting);
                }
                catch { /* skip unreadable item */ }
                finally { Release(item); }
                item = ((dynamic)restricted).GetNext();
            }
            return result.OrderBy(m => m.Start).ToList();
        }
        finally
        {
            Release(restricted);
            Release(items);
            Release(folder);
            Release(ns);
            Release(app);
        }
    }

    private static Meeting? ToMeeting(object item, Settings s, DateTime now, string storeId)
    {
        dynamic it = item;
        if ((int)it.Class != OlAppointment) return null;

        DateTime start = it.Start, end = it.End;
        if (end <= now) return null;

        string subject = (string?)it.Subject ?? "";
        bool allDay = it.AllDayEvent;
        int busy = it.BusyStatus;
        int sensitivity = it.Sensitivity;
        int response = it.ResponseStatus;
        int meetingStatus = it.MeetingStatus;

        if (s.HideAllDay && allDay) return null;
        if (s.HideFree && busy == OlFree) return null;
        if (s.HideTentative && busy == OlTentative) return null;
        if (s.HideDeclined && (response == OlResponseDeclined
                               || meetingStatus == OlMeetingCanceled
                               || meetingStatus == OlMeetingReceivedAndCanceled
                               || MeetingFilter.HasCancelledPrefix(subject))) return null;
        if (s.HidePrivate && sensitivity == OlPrivate) return null;
        if (MeetingFilter.MatchesKeyword(s, subject)) return null;

        string location = (string?)it.Location ?? "";
        string? joinUrl = null;
        if (s.DetectJoinLink)
        {
            joinUrl = MeetingFilter.FindJoinLink(location);
            if (joinUrl == null)
            {
                try { joinUrl = MeetingFilter.FindJoinLink((string?)it.Body ?? ""); }
                catch { /* Body blocked by Outlook security -> no link */ }
            }
        }

        string entryId = it.EntryID;
        return new Meeting(subject.Trim(), start, end, location.Trim(), joinUrl, entryId, storeId, MeetingSource.Outlook);
    }

    /// <summary>Opens the meeting (the specific occurrence for recurring series) in Outlook.</summary>
    public bool Open(Meeting m)
    {
        object app = GetRunningOutlook();
        object? ns = null, item = null, pattern = null, occurrence = null;
        try
        {
            ns = ((dynamic)app).GetNamespace("MAPI");
            item = ((dynamic)ns).GetItemFromID(m.EntryId, m.StoreId);
            if ((bool)((dynamic)item).IsRecurring)
            {
                try
                {
                    pattern = ((dynamic)item).GetRecurrencePattern();
                    occurrence = ((dynamic)pattern).GetOccurrence(m.Start);
                    ((dynamic)occurrence).Display();
                    return true;
                }
                catch { /* fall back to the series */ }
            }
            ((dynamic)item).Display();
            return true;
        }
        finally
        {
            Release(occurrence);
            Release(pattern);
            Release(item);
            Release(ns);
            Release(app);
        }
    }

    private static object GetRunningOutlook()
    {
        var processes = Process.GetProcessesByName("OUTLOOK");
        bool anyRunning = processes.Length > 0;
        bool hasWindow = processes.Any(p => p.MainWindowHandle != IntPtr.Zero);
        foreach (var p in processes) p.Dispose();
        if (!anyRunning) throw new OutlookNotRunningException();

        if (CLSIDFromProgID("Outlook.Application", out var clsid) != 0)
            throw new OutlookNotRunningException();

        if (GetActiveObject(ref clsid, IntPtr.Zero, out var active) == 0 && active != null)
            return active;

        // A background-only "-Embedding" instance is often unreachable and CreateInstance would
        // block ~30 s - only fall back when a real Outlook window exists (then it just attaches).
        if (!hasWindow) throw new OutlookNotRunningException();
        var type = Type.GetTypeFromCLSID(clsid) ?? throw new OutlookNotRunningException();
        return Activator.CreateInstance(type) ?? throw new OutlookNotRunningException();
    }

    private static void Release(object? o)
    {
        if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
    }

    public void Dispose() => _queue.CompleteAdding();

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid clsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? obj);
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace NextMeetingOverlay;

/// <summary>Snapshot of today's tracked time. The running timer is added live by <see cref="TotalAt"/>.</summary>
public sealed class TmetricSnapshot
{
    /// <summary>Finished entries today (running one excluded).</summary>
    public TimeSpan Completed { get; init; }
    /// <summary>Start of the running timer (clipped to today), null = nothing running.</summary>
    public DateTime? RunningSince { get; init; }
    public string? RunningLabel { get; init; }
    /// <summary>Short title of the running entry: "#27791 Ticket title", or the note, or the project.</summary>
    public string? RunningTitle { get; init; }
    /// <summary>All entries TMetric returned for the read window (yesterday to tomorrow), for the optional log file.</summary>
    public IReadOnlyList<TmetricEntry> Entries { get; init; } = Array.Empty<TmetricEntry>();

    public TimeSpan TotalAt(DateTime now) => Completed + (RunningSince is DateTime s && now > s ? now - s : TimeSpan.Zero);
}

/// <summary>One time entry as TMetric reports it (times in local time). <see cref="End"/> null = the timer is still running.</summary>
public sealed record TmetricEntry(long Id, DateTime Start, DateTime? End, string? Project, string? TaskNumber, string? Task, string? Note);

/// <summary>
/// Read-only access to the TMetric web API (v3, personal API token as bearer):
/// GET /user -> own id + active account, GET /accounts/{id}/timeentries -> today's entries.
/// An entry without end time is the running timer. Nothing is ever written to TMetric.
/// </summary>
public sealed class TmetricClient
{
    private const string BaseUrl = "https://app.tmetric.com/api/v3";
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(20) };

    private long? _userId;
    private long? _accountId;
    private string? _tokenForIds;

    public string? UserName { get; private set; }

    /// <summary>Forget the cached ids (token changed).</summary>
    public void Reset() { _userId = _accountId = null; _tokenForIds = null; UserName = null; }

    /// <summary>One read request (GET). Returns null for an empty answer.</summary>
    private static async Task<JsonDocument?> SendAsync(string path, string token, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, BaseUrl + path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        using var response = await Http.SendAsync(request, ct);
        if (response.StatusCode is System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden)
            throw new TmetricAuthException();
        var text = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
        {
            // Keep TMetric's own explanation - "400 Bad Request" alone says nothing.
            var detail = text.Length > 300 ? text[..300] : text;
            throw new HttpRequestException($"HTTP {(int)response.StatusCode} bei GET {path.Split('?')[0]}: {detail}");
        }
        return string.IsNullOrWhiteSpace(text) ? null : JsonDocument.Parse(text);
    }

    private static async Task<JsonDocument> GetAsync(string path, string token, CancellationToken ct) =>
        await SendAsync(path, token, ct) ?? JsonDocument.Parse("[]");

    /// <summary>Checks the token and returns the user's name (also caches ids).</summary>
    public async Task<string> ConnectAsync(string token, CancellationToken ct = default)
    {
        if (_tokenForIds == token && _userId != null && _accountId != null) return UserName ?? "";
        using var doc = await GetAsync("/user", token, ct);
        var root = doc.RootElement;
        _userId = root.GetProperty("id").GetInt64();
        _accountId = root.GetProperty("activeAccountId").GetInt64();
        UserName = root.TryGetProperty("name", out var name) ? name.GetString() : null;
        _tokenForIds = token;
        return UserName ?? "";
    }

    public async Task<TmetricSnapshot> ReadTodayAsync(string token, CancellationToken ct = default)
    {
        await ConnectAsync(token, ct);

        var now = DateTime.Now;
        var dayStart = now.Date;
        var dayEnd = dayStart.AddDays(1);
        string D(DateTime d) => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        // One day of slack on both sides - entries are clipped to today below.
        var path = $"/accounts/{_accountId}/timeentries?userId={_userId}&startDate={D(dayStart.AddDays(-1))}&endDate={D(dayEnd)}";

        using var doc = await GetAsync(path, token, ct);
        var completed = TimeSpan.Zero;
        DateTime? runningSince = null;
        string? label = null;
        string? runningTitle = null;
        var entries = new List<TmetricEntry>();

        foreach (var entry in doc.RootElement.EnumerateArray())
        {
            if (!TryTime(entry, "startTime", out var start)) continue;
            bool running = !TryTime(entry, "endTime", out var end);
            if (running) end = now;

            if (ToEntry(entry) is { } logged) entries.Add(logged);

            var from = start < dayStart ? dayStart : start;
            var to = end > dayEnd ? dayEnd : end;
            if (to <= from) continue;

            if (running)
            {
                runningSince = from;
                label = Describe(entry);
                runningTitle = EntryTitle(entry);
            }
            else
            {
                completed += to - from;
            }
        }
        return new TmetricSnapshot { Completed = completed, RunningSince = runningSince, RunningLabel = label, RunningTitle = runningTitle, Entries = entries };
    }

    /// <summary>All entries between two days (inclusive), e.g. a whole month for the log. Read only.</summary>
    public async Task<List<TmetricEntry>> ReadEntriesAsync(string token, DateTime from, DateTime to, CancellationToken ct = default)
    {
        await ConnectAsync(token, ct);
        string D(DateTime d) => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        using var doc = await GetAsync($"/accounts/{_accountId}/timeentries?userId={_userId}&startDate={D(from)}&endDate={D(to)}", token, ct);
        var list = new List<TmetricEntry>();
        if (doc.RootElement.ValueKind != JsonValueKind.Array) return list;
        foreach (var entry in doc.RootElement.EnumerateArray())
            if (ToEntry(entry) is { } e) list.Add(e);
        return list;
    }

    /// <summary>One API entry as <see cref="TmetricEntry"/>; End stays null while the timer runs. Null if the entry is unusable.</summary>
    private static TmetricEntry? ToEntry(JsonElement entry)
    {
        if (!TryTime(entry, "startTime", out var start)) return null;
        if (!entry.TryGetProperty("id", out var idElement) || !idElement.TryGetInt64(out var id)) return null;
        DateTime? end = TryTime(entry, "endTime", out var stop) ? stop : null;
        string? note = entry.TryGetProperty("note", out var n) && n.ValueKind == JsonValueKind.String ? n.GetString() : null;
        return new TmetricEntry(id, start, end, Name(entry, "project"), TaskNumber(entry), Name(entry, "task"), note);
    }

    /// <summary>Number of the ticket in the connected ticket system (e.g. "27791"), if the task is linked to one.</summary>
    private static string? TaskNumber(JsonElement entry) =>
        entry.TryGetProperty("task", out var t) && t.ValueKind == JsonValueKind.Object
        && t.TryGetProperty("externalLink", out var link) && link.ValueKind == JsonValueKind.Object
        && link.TryGetProperty("caption", out var cap) && cap.ValueKind == JsonValueKind.String
            ? cap.GetString() : null;

    private static bool TryTime(JsonElement entry, string name, out DateTime local)
    {
        local = default;
        if (!entry.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String) return false;
        // TMetric returns times in the user's time zone without an offset ("2026-11-13T11:10:00"):
        // read them as local time (a value with Z/offset keeps its own).
        if (!DateTimeOffset.TryParse(value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed)) return false;
        local = parsed.LocalDateTime;
        return true;
    }

    /// <summary>Short title of an entry: "#number task", else the note, else the project name.</summary>
    private static string? EntryTitle(JsonElement entry)
    {
        var task = Name(entry, "task");
        if (!string.IsNullOrWhiteSpace(task))
        {
            string? number = entry.TryGetProperty("task", out var t) && t.TryGetProperty("externalLink", out var link)
                && link.ValueKind == JsonValueKind.Object && link.TryGetProperty("caption", out var cap)
                && cap.ValueKind == JsonValueKind.String ? cap.GetString() : null;
            return string.IsNullOrWhiteSpace(number) ? task : $"#{number} {task}";
        }
        if (entry.TryGetProperty("note", out var n) && n.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(n.GetString()))
            return n.GetString()!.Trim();
        return Name(entry, "project");
    }

    private static string? Describe(JsonElement entry)
    {
        string? project = Name(entry, "project");
        string? task = Name(entry, "task");
        string? note = entry.TryGetProperty("note", out var n) && n.ValueKind == JsonValueKind.String ? n.GetString() : null;
        var parts = new[] { project, string.IsNullOrWhiteSpace(task) ? note : task }
            .Where(p => !string.IsNullOrWhiteSpace(p));
        var text = string.Join(" – ", parts);
        return text.Length == 0 ? null : text;
    }

    private static string? Name(JsonElement entry, string property) =>
        entry.TryGetProperty(property, out var obj) && obj.ValueKind == JsonValueKind.Object
        && obj.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String
            ? name.GetString() : null;
}

public sealed class TmetricAuthException : Exception
{
    public TmetricAuthException() : base("Token ungültig oder gesperrt") { }
}

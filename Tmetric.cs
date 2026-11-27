using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
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
    /// <summary>Identity of the running entry (project|task|note), to hide it from the "recent" shortcuts.</summary>
    public string? RunningKey { get; init; }

    public TimeSpan TotalAt(DateTime now) => Completed + (RunningSince is DateTime s && now > s ? now - s : TimeSpan.Zero);
}

/// <summary>
/// Access to the TMetric web API (v3, personal API token as bearer):
/// GET /user -> own id + active account, GET /accounts/{id}/timeentries -> today's entries,
/// GET .../timeentries/recent -> recent project/ticket/note combinations.
/// An entry without end time is the running timer. TMetric is only written to when the user clicks
/// "Timer starten" or a "Zuletzt" shortcut (stop the running timer, start the new one).
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

    /// <summary>One API call. Returns null for an empty answer (e.g. "no latest entry").</summary>
    private static async Task<JsonDocument?> SendAsync(HttpMethod method, string path, string token, object? body, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(method, BaseUrl + path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (body != null)
            request.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
        using var response = await Http.SendAsync(request, ct);
        if (response.StatusCode is System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden)
            throw new TmetricAuthException();
        var text = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
        {
            // Keep TMetric's own explanation - "400 Bad Request" alone says nothing.
            var detail = text.Length > 300 ? text[..300] : text;
            throw new HttpRequestException($"HTTP {(int)response.StatusCode} bei {method} {path.Split('?')[0]}: {detail}");
        }
        return string.IsNullOrWhiteSpace(text) ? null : JsonDocument.Parse(text);
    }

    private static async Task<JsonDocument> GetAsync(string path, string token, CancellationToken ct) =>
        await SendAsync(HttpMethod.Get, path, token, null, ct) ?? JsonDocument.Parse("[]");

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
        string? runningKey = null;
        string? runningTitle = null;

        foreach (var entry in doc.RootElement.EnumerateArray())
        {
            if (!TryTime(entry, "startTime", out var start)) continue;
            bool running = !TryTime(entry, "endTime", out var end);
            if (running) end = now;

            var from = start < dayStart ? dayStart : start;
            var to = end > dayEnd ? dayEnd : end;
            if (to <= from) continue;

            if (running)
            {
                runningSince = from;
                label = Describe(entry);
                runningKey = EntryKey(entry);
                runningTitle = EntryTitle(entry);
            }
            else
            {
                completed += to - from;
            }
        }
        return new TmetricSnapshot { Completed = completed, RunningSince = runningSince, RunningLabel = label, RunningTitle = runningTitle, RunningKey = runningKey };
    }

    /// <summary>
    /// "Timer starten": stops the running timer (if any) and starts a new one on the given project -
    /// without task and note, so no gap (= no break) appears in TMetric. This is the only write access.
    /// </summary>
    public async Task<TmetricSwitchResult> StartOnProjectAsync(string token, string projectName, CancellationToken ct = default)
    {
        await ConnectAsync(token, ct);
        var wanted = projectName.Trim();

        // 1. project id by name
        var projectId = await FindProjectIdAsync(token, wanted, ct);
        if (projectId == null) throw new TmetricProjectNotFoundException(wanted);

        // 2. + 3. nothing to do if exactly this (project without task/note) already runs, else switch
        var running = await GetRunningAsync(token, ct);
        if (running?.Key == EntryKey(projectId, null, null)) return TmetricSwitchResult.AlreadyRunning;
        return await SwitchAsync(token, running?.Id, new JsonObject { ["project"] = new JsonObject { ["id"] = projectId } }, ct);
    }

    /// <summary>Id of a project available for time tracking, found by name (case, spaces and hyphens ignored), or null.</summary>
    public async Task<long?> FindProjectIdAsync(string token, string projectName, CancellationToken ct = default)
    {
        await ConnectAsync(token, ct);
        var wanted = Normalize(projectName);
        using var projects = await GetAsync($"/accounts/{_accountId}/timeentries/projects?userId={_userId}", token, ct);
        if (projects.RootElement.ValueKind != JsonValueKind.Array) return null;
        foreach (var p in projects.RootElement.EnumerateArray())
        {
            var item = p.TryGetProperty("project", out var inner) && inner.ValueKind == JsonValueKind.Object ? inner : p;
            if (item.TryGetProperty("name", out var n) && n.ValueKind == JsonValueKind.String
                && string.Equals(Normalize(n.GetString()), wanted, StringComparison.Ordinal)
                && item.TryGetProperty("id", out var id) && id.TryGetInt64(out var value))
                return value;
        }
        return null;
    }

    private sealed record RunningEntry(long Id, string Key);

    /// <summary>The latest entry, if it has no end time (= the timer is running).</summary>
    private async Task<RunningEntry?> GetRunningAsync(string token, CancellationToken ct)
    {
        using var latest = await SendAsync(HttpMethod.Get, $"/accounts/{_accountId}/timeentries/latest?userId={_userId}", token, null, ct);
        if (latest == null || latest.RootElement.ValueKind != JsonValueKind.Object) return null;
        var entry = latest.RootElement;
        if (TryTime(entry, "endTime", out _)) return null;
        if (!entry.TryGetProperty("id", out var id) || !id.TryGetInt64(out var idValue)) return null;
        return new RunningEntry(idValue, EntryKey(entry));
    }

    /// <summary>Stops the running timer (if any) and starts the new one at exactly the same moment, so no break appears.</summary>
    private async Task<TmetricSwitchResult> SwitchAsync(string token, long? runningId, JsonObject newEntry, CancellationToken ct)
    {
        // One instant, two spellings: TMetric reads a "Z" time correctly when starting (POST), but when stopping
        // (PUT) it compares the clock value as it is with the entry's local start time ("StartTime cannot be higher
        // than EndTime" if we sent UTC). So: PUT gets local time without offset, POST gets UTC with "Z".
        var now = DateTime.Now;
        now = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second, DateTimeKind.Local);
        var localStamp = now.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
        var stamp = now.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

        if (runningId != null)
        {
            using var _ = await SendAsync(HttpMethod.Put, $"/accounts/{_accountId}/timeentries/{runningId}", token, new { endTime = localStamp }, ct);
        }
        newEntry["startTime"] = stamp;
        try
        {
            using var _ = await SendAsync(HttpMethod.Post, $"/accounts/{_accountId}/timeentries?userId={_userId}", token, newEntry, ct);
        }
        catch (Exception ex) when (runningId != null && ex is not TmetricAuthException)
        {
            throw new TmetricStartAfterStopException(ex.Message);
        }
        return runningId != null ? TmetricSwitchResult.Switched : TmetricSwitchResult.Started;
    }

    /// <summary>Identity of "what is tracked": project | task | note. Equal keys = the same shortcut.</summary>
    internal static string EntryKey(long? projectId, long? taskId, string? note) =>
        $"{projectId}|{taskId}|{note?.Trim()}";

    private static string EntryKey(JsonElement entry)
    {
        long? projectId = entry.TryGetProperty("project", out var p) && p.ValueKind == JsonValueKind.Object
            && p.TryGetProperty("id", out var pid) && pid.TryGetInt64(out var pv) ? pv : null;
        long? taskId = entry.TryGetProperty("task", out var t) && t.ValueKind == JsonValueKind.Object
            && t.TryGetProperty("id", out var tid) && tid.TryGetInt64(out var tv) ? tv : null;
        string? note = entry.TryGetProperty("note", out var n) && n.ValueKind == JsonValueKind.String ? n.GetString() : null;
        return EntryKey(projectId, taskId, note);
    }

    /// <summary>"My Project", "My-Project" and "my project" are the same project for the user.</summary>
    private static string Normalize(string? name) =>
        new string((name ?? "").Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();

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

    /// <summary>Short title like in the shortcut list: "#number task", else the note, else the project name.</summary>
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

public enum TmetricSwitchResult { Started, Switched, AlreadyRunning }

public sealed class TmetricAuthException : Exception
{
    public TmetricAuthException() : base("Token ungültig oder gesperrt") { }
}

public sealed class TmetricProjectNotFoundException : Exception
{
    public TmetricProjectNotFoundException(string name) : base($"Projekt „{name}“ nicht gefunden") { }
}

/// <summary>The old timer was stopped but the new one could not be started - the user has to start one by hand.</summary>
public sealed class TmetricStartAfterStopException : Exception
{
    public TmetricStartAfterStopException(string reason)
        : base("Der laufende Timer wurde beendet, der neue konnte aber nicht gestartet werden (" + reason + ")") { }
}

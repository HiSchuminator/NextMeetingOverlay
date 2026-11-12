using System;
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

    public TimeSpan TotalAt(DateTime now) => Completed + (RunningSince is DateTime s && now > s ? now - s : TimeSpan.Zero);
}

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

    private static async Task<JsonDocument> GetAsync(string path, string token, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, BaseUrl + path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        using var response = await Http.SendAsync(request, ct);
        if (response.StatusCode is System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden)
            throw new TmetricAuthException();
        response.EnsureSuccessStatusCode();
        return await JsonDocument.ParseAsync(await response.Content.ReadAsStreamAsync(ct), cancellationToken: ct);
    }

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
            }
            else
            {
                completed += to - from;
            }
        }
        return new TmetricSnapshot { Completed = completed, RunningSince = runningSince, RunningLabel = label };
    }

    private static bool TryTime(JsonElement entry, string name, out DateTime local)
    {
        local = default;
        if (!entry.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String) return false;
        if (!DateTimeOffset.TryParse(value.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed)) return false;
        local = parsed.LocalDateTime;
        return true;
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

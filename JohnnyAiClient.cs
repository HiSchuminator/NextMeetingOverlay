using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace NextMeetingOverlay;

public sealed record JohnnyStatus(bool Recording, string Time, string Subject, bool Busy);

/// <summary>
/// Talks to JohnnyAI's control pipe (JohnnyAI Services/ControlPipeServer.cs):
/// STATUS / START [event-json] / STOP, one line each way.
/// </summary>
public static class JohnnyAiClient
{
    private const string PipeName = "JohnnyAI.Control.v1";
    private static readonly Encoding Utf8 = new UTF8Encoding(false);
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    public static bool IsRunning
    {
        get
        {
            var processes = Process.GetProcessesByName("JohnnyAI");
            foreach (var p in processes) p.Dispose();
            return processes.Length > 0;
        }
    }

    /// <summary>Null when JohnnyAI is not running or is an older version without the control pipe.</summary>
    public static async Task<JohnnyStatus?> GetStatusAsync()
    {
        var reply = await SendAsync("STATUS");
        if (reply == null || !reply.StartsWith('{')) return null;
        try { return JsonSerializer.Deserialize<JohnnyStatus>(reply, JsonOptions); }
        catch (JsonException) { return null; }
    }

    /// <summary>Starts a recording linked to the meeting (title, time, location) if one is given.</summary>
    public static Task<string?> StartAsync(Meeting? meeting)
    {
        if (meeting == null) return SendAsync("START");
        var ev = JsonSerializer.Serialize(new
        {
            title = meeting.Subject,
            startDate = meeting.Start,
            endDate = meeting.End,
            location = meeting.Location,
            isOnlineMeeting = meeting.JoinUrl != null,
        });
        return SendAsync("START " + ev);
    }

    public static Task<string?> StopAsync() => SendAsync("STOP");

    private static async Task<string?> SendAsync(string request)
    {
        if (!IsRunning) return null;
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(1500));
            await using var client = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut,
                PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
            await client.ConnectAsync(cts.Token);
            await using var writer = new StreamWriter(client, Utf8, 1024, leaveOpen: true) { AutoFlush = true };
            using var reader = new StreamReader(client, Utf8, false, 1024, leaveOpen: true);
            await writer.WriteLineAsync(request.AsMemory(), cts.Token);
            return await reader.ReadLineAsync(cts.Token);
        }
        catch
        {
            return null;
        }
    }
}

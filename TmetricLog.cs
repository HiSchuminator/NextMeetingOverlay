using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace NextMeetingOverlay;

/// <summary>
/// Optional log of the TMetric entries: one CSV file per month (semicolon separated, UTF-8 with BOM, so Excel opens it
/// as a table). Every read of TMetric merges the entries it returned into the file: entries are matched by their
/// TMetric id (changed ones are replaced, the running one grows), entries of the read window that TMetric no longer
/// returns (deleted there) are dropped, older days stay untouched.
/// </summary>
public static class TmetricLog
{
    private static readonly string[] Header =
        { "Datum", "Start", "Ende", "Dauer (min)", "Dauer (h:mm)", "Projekt", "Ticket", "Aufgabe", "Notiz", "ID" };

    private static readonly object Gate = new();

    public static string FolderFor(string dataDirectory) => Path.Combine(dataDirectory, "TMetric-Log");

    public static string FileFor(string folder, DateTime month) => Path.Combine(folder, $"TMetric-{month:yyyy-MM}.csv");

    /// <summary>
    /// Merges <paramref name="entries"/> into the month files. They stand for TMetric's answer to a read of the days
    /// <paramref name="windowFrom"/>..<paramref name="windowTo"/> (default: yesterday to tomorrow, like the regular read).
    /// </summary>
    public static void Write(string folder, IReadOnlyList<TmetricEntry> entries, DateTime now, DateTime? windowFrom = null, DateTime? windowTo = null)
    {
        if (entries.Count == 0) return;
        var from = windowFrom ?? now.Date.AddDays(-1);
        var to = windowTo ?? now.Date.AddDays(1);
        lock (Gate)
        {
            Directory.CreateDirectory(folder);
            foreach (var month in entries.GroupBy(e => new DateTime(e.Start.Year, e.Start.Month, 1)))
                WriteMonth(FileFor(folder, month.Key), month.ToList(), now, from, to);
        }
    }

    private static void WriteMonth(string path, List<TmetricEntry> fresh, DateTime now, DateTime windowFrom, DateTime windowTo)
    {
        var freshIds = fresh.Select(e => e.Id.ToString(CultureInfo.InvariantCulture)).ToHashSet();

        var rows = new List<string>();
        string? before = null;
        if (File.Exists(path))
        {
            before = File.ReadAllText(path, Encoding.UTF8);
            foreach (var line in before.Split('\n').Select(l => l.TrimEnd('\r')).Skip(1).Where(l => l.Length > 12))
            {
                var id = line[(line.LastIndexOf(';') + 1)..];
                if (freshIds.Contains(id)) continue;   // replaced by the current version below
                // Inside the read window TMetric is the truth: a row it does not return any more was deleted there.
                if (DateTime.TryParseExact(line[..10], "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var day)
                    && day >= windowFrom && day <= windowTo) continue;
                rows.Add(line);
            }
        }
        rows.AddRange(fresh.Select(e => Row(e, now)));

        // Sort by date, then start time (both are the first two fields and never contain a semicolon).
        string SortKey(string line)
        {
            var parts = line.Split(';', 3);
            var date = DateTime.TryParseExact(parts[0], "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : DateTime.MinValue;
            return date.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + (parts.Length > 1 ? parts[1] : "");
        }
        var text = string.Join("\r\n", new[] { string.Join(";", Header) }.Concat(rows.OrderBy(SortKey, StringComparer.Ordinal))) + "\r\n";

        if (text == before) return;   // nothing changed (the running entry changes every minute, the rest rarely)
        var tmp = path + ".tmp";
        try
        {
            File.WriteAllText(tmp, text, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            File.Move(tmp, path, overwrite: true);
        }
        catch
        {
            // Typically the file is open in Excel (locked): try again with the next read, nothing is lost.
            try { File.Delete(tmp); } catch { }
            throw;
        }
    }

    private static string Row(TmetricEntry e, DateTime now)
    {
        var end = e.End ?? now;
        int minutes = Math.Max(0, (int)(end - e.Start).TotalMinutes);
        return string.Join(";",
            e.Start.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture),
            e.Start.ToString("HH:mm", CultureInfo.InvariantCulture),
            e.End is DateTime stop ? stop.ToString("HH:mm", CultureInfo.InvariantCulture) : "läuft",
            minutes.ToString(CultureInfo.InvariantCulture),
            $"{minutes / 60}:{minutes % 60:00}",
            Csv(e.Project), Csv(e.TaskNumber), Csv(e.Task), Csv(e.Note),
            e.Id.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>One line per row: line breaks become spaces. Quoted when needed. Text that starts like a formula is defused.</summary>
    private static string Csv(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        text = text.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ').Trim();
        // Ticket titles come from customers: =, +, -, @ at the start would be run as a formula by Excel.
        if (text[0] is '=' or '+' or '-' or '@') text = "'" + text;
        return text.Contains(';') || text.Contains('"') ? "\"" + text.Replace("\"", "\"\"") + "\"" : text;
    }
}

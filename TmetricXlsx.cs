using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace NextMeetingOverlay;

/// <summary>
/// Builds <c>TMetric-Log.xlsx</c> from the monthly CSV files of <see cref="TmetricLog"/>: one sheet per year, in each
/// sheet one block per month (heading, the entries by date, a grey total after every day, a total for the month).
/// An .xlsx is a ZIP of XML files, so it is written with the framework alone - no extra library.
/// The CSV files stay the data source; this workbook is only a view and is rebuilt whenever a CSV is newer.
/// </summary>
public static class TmetricXlsx
{
    private sealed record LogRow(DateTime Date, TimeSpan Start, string End, int Minutes, string Client, string Project, string Ticket, string Task,
                                 string Note, string Tags, string Billable, string Invoiced);

    private static readonly CultureInfo German = new("de-DE");
    private static readonly string[] Titles =
        { "Datum", "Start", "Ende", "Dauer (min)", "Dauer (h:mm)", "Kunde", "Projekt", "Ticket", "Aufgabe", "Notiz", "Tags", "Abrechenbar", "Fakturiert" };
    private static readonly double[] Widths = { 13, 8, 9, 12, 13, 18, 18, 9, 52, 42, 16, 12, 11 };

    // Cell styles (indexes into cellXfs of styles.xml below).
    private const int Plain = 0, Head = 1, DateS = 2, TimeS = 3, Dur = 4, MonthHead = 5, MonthLabel = 6, MonthMin = 7, MonthDur = 8, DayLabel = 9, DayMin = 10, DayDur = 11;

    public static string FileFor(string folder) => Path.Combine(folder, "TMetric-Log.xlsx");

    /// <summary>True when the workbook is missing or older than one of the CSV files.</summary>
    public static bool NeedsBuild(string folder)
    {
        var csvs = Directory.GetFiles(folder, "TMetric-????-??.csv");
        if (csvs.Length == 0) return false;
        var xlsx = FileFor(folder);
        return !File.Exists(xlsx) || csvs.Any(c => File.GetLastWriteTimeUtc(c) > File.GetLastWriteTimeUtc(xlsx));
    }

    /// <summary>Rebuilds the workbook. Throws if it cannot be written (typically: it is open in Excel) - the caller tries again later.</summary>
    public static void Build(string folder)
    {
        var rows = ReadAll(folder);
        if (rows.Count == 0) return;
        var bytes = Create(rows);
        var path = FileFor(folder);
        var tmp = path + ".tmp";
        try
        {
            File.WriteAllBytes(tmp, bytes);
            File.Move(tmp, path, overwrite: true);
        }
        catch
        {
            try { File.Delete(tmp); } catch { }
            throw;
        }
    }

    // ───────────────────────── Reading the CSV files ─────────────────────────

    private static List<LogRow> ReadAll(string folder)
    {
        var rows = new List<LogRow>();
        foreach (var file in Directory.GetFiles(folder, "TMetric-????-??.csv").OrderBy(f => f, StringComparer.Ordinal))
        {
            foreach (var line in File.ReadAllLines(file, Encoding.UTF8).Skip(1))
            {
                var f = SplitCsv(line);
                if (f.Count < 14) continue;   // file in an older layout: skipped until the log rebuilds it
                if (!DateTime.TryParseExact(f[0], "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) continue;
                if (!TimeSpan.TryParseExact(f[1], @"hh\:mm", CultureInfo.InvariantCulture, out var start)) continue;
                if (!int.TryParse(f[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out var minutes)) continue;
                rows.Add(new LogRow(date, start, f[2], minutes, Clean(f[5]), Clean(f[6]), Clean(f[7]), Clean(f[8]), Clean(f[9]), Clean(f[10]), f[11], f[12]));
            }
        }
        return rows.OrderBy(r => r.Date).ThenBy(r => r.Start).ToList();
    }

    /// <summary>Quote-aware split of one CSV line (semicolon separated, quotes doubled inside quoted fields).</summary>
    private static List<string> SplitCsv(string line)
    {
        var fields = new List<string>();
        var sb = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (quoted)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else quoted = false;
                }
                else sb.Append(c);
            }
            else if (c == '"') quoted = true;
            else if (c == ';') { fields.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }
        fields.Add(sb.ToString());
        return fields;
    }

    /// <summary>The CSV defuses text that starts like a formula with a leading apostrophe; in the workbook it is plain text anyway.</summary>
    private static string Clean(string text) =>
        text.Length > 1 && text[0] == '\'' && text[1] is '=' or '+' or '-' or '@' ? text[1..] : text;

    // ───────────────────────── Writing the workbook ─────────────────────────

    private static byte[] Create(List<LogRow> rows)
    {
        var years = rows.GroupBy(r => r.Date.Year).OrderBy(g => g.Key).ToList();
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            Add(zip, "[Content_Types].xml", ContentTypes(years.Count));
            Add(zip, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>");
            Add(zip, "xl/workbook.xml", Workbook(years.Select(g => g.Key).ToList()));
            Add(zip, "xl/_rels/workbook.xml.rels", WorkbookRels(years.Count));
            Add(zip, "xl/styles.xml", Styles);
            for (int i = 0; i < years.Count; i++)
                Add(zip, $"xl/worksheets/sheet{i + 1}.xml", Sheet(years[i].ToList()));
        }
        return ms.ToArray();
    }

    private static void Add(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static string ContentTypes(int sheets)
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        for (int i = 1; i <= sheets; i++)
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        return sb.Append("</Types>").ToString();
    }

    private static string Workbook(List<int> years)
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<bookViews><workbookView activeTab=\"" + (years.Count - 1) + "\"/></bookViews><sheets>");
        for (int i = 0; i < years.Count; i++)
            sb.Append($"<sheet name=\"{years[i]}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
        return sb.Append("</sheets></workbook>").ToString();
    }

    private static string WorkbookRels(int sheets)
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (int i = 1; i <= sheets; i++)
            sb.Append($"<Relationship Id=\"rId{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{sheets + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        return sb.Append("</Relationships>").ToString();
    }

    /// <summary>fonts: 0 normal, 1 bold, 2 grey italic · fills: 0 none, 1 (required), 2 header blue, 3 month green, 4 day grey · formats: 164 date, 165 time, 166 hours:minutes.</summary>
    private const string Styles =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<numFmts count=\"3\"><numFmt numFmtId=\"164\" formatCode=\"dd\\.mm\\.yyyy\"/><numFmt numFmtId=\"165\" formatCode=\"hh:mm\"/><numFmt numFmtId=\"166\" formatCode=\"[h]:mm\"/></numFmts>" +
        "<fonts count=\"3\">" +
        "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
        "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
        "<font><i/><sz val=\"10\"/><color rgb=\"FF595959\"/><name val=\"Calibri\"/></font>" +
        "</fonts>" +
        "<fills count=\"5\">" +
        "<fill><patternFill patternType=\"none\"/></fill>" +
        "<fill><patternFill patternType=\"gray125\"/></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFD9E1F2\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE2EFDA\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF2F2F2\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
        "</fills>" +
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
        "<cellXfs count=\"12\">" +
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +                                                              // 0 plain
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +                              // 1 header
        "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +                                    // 2 date
        "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +                                    // 3 time
        "<xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +                                    // 4 duration
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +                              // 5 month heading
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +                              // 6 month sum label
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +                              // 7 month sum minutes
        "<xf numFmtId=\"166\" fontId=\"1\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\"/>" +    // 8 month sum h:mm
        "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"4\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +                              // 9 day sum label
        "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"4\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +                              // 10 day sum minutes
        "<xf numFmtId=\"166\" fontId=\"2\" fillId=\"4\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\"/>" +    // 11 day sum h:mm
        "</cellXfs>" +
        "<cellStyles count=\"1\"><cellStyle name=\"Standard\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
        "</styleSheet>";

    private static string Sheet(List<LogRow> year)
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>" +
            "<sheetFormatPr defaultRowHeight=\"15\"/><cols>");
        for (int i = 0; i < Widths.Length; i++)
            sb.Append($"<col min=\"{i + 1}\" max=\"{i + 1}\" width=\"{Widths[i].ToString(CultureInfo.InvariantCulture)}\" customWidth=\"1\"/>");
        sb.Append("</cols><sheetData>");

        int r = 1;
        Row(sb, r++, Titles.Select(t => Text(t, Head)).ToArray());

        foreach (var month in year.GroupBy(x => x.Date.Month).OrderBy(g => g.Key))
        {
            var first = month.First().Date;
            var monthName = first.ToString("MMMM yyyy", German);
            Row(sb, r++, Pad(new[] { Text(monthName, MonthHead) }, MonthHead));

            foreach (var day in month.GroupBy(x => x.Date).OrderBy(g => g.Key))
            {
                foreach (var e in day)
                {
                    Row(sb, r++, new[]
                    {
                        Num(e.Date.ToOADate(), DateS),
                        Num(e.Start.TotalDays, TimeS),
                        TimeSpan.TryParseExact(e.End, @"hh\:mm", CultureInfo.InvariantCulture, out var end) ? Num(end.TotalDays, TimeS) : Text(e.End, Plain),
                        Num(e.Minutes, Plain),
                        Num(e.Minutes / 1440.0, Dur),
                        Text(e.Client, Plain), Text(e.Project, Plain), Text(e.Ticket, Plain), Text(e.Task, Plain), Text(e.Note, Plain),
                        Text(e.Tags, Plain), Text(e.Billable, Plain), Text(e.Invoiced, Plain),
                    });
                }
                int dayMinutes = day.Sum(x => x.Minutes);
                Row(sb, r++, Pad(new[] { Text("Tagessumme " + day.Key.ToString("ddd dd.MM.", German), DayLabel), Blank(DayLabel), Blank(DayLabel),
                                         Num(dayMinutes, DayMin), Num(dayMinutes / 1440.0, DayDur) }, DayLabel));
            }

            int monthMinutes = month.Sum(x => x.Minutes);
            Row(sb, r++, Pad(new[] { Text("Summe " + monthName, MonthLabel), Blank(MonthLabel), Blank(MonthLabel),
                                     Num(monthMinutes, MonthMin), Num(monthMinutes / 1440.0, MonthDur) }, MonthLabel));
            r++;   // one empty row between months
        }
        return sb.Append("</sheetData></worksheet>").ToString();
    }

    // A cell is (kind, value, style). Kinds: 't' text, 'n' number, 'b' blank (only styled).
    private readonly record struct Cell(char Kind, string Value, int Style);

    private static Cell Text(string? value, int style) => string.IsNullOrEmpty(value) ? new Cell('b', "", style) : new Cell('t', value, style);
    private static Cell Num(double value, int style) => new('n', value.ToString("R", CultureInfo.InvariantCulture), style);
    private static Cell Blank(int style) => new('b', "", style);

    /// <summary>Fills the row up to all columns with styled blanks, so a coloured row is coloured across the whole table.</summary>
    private static Cell[] Pad(Cell[] cells, int style) =>
        cells.Concat(Enumerable.Range(0, Math.Max(0, Titles.Length - cells.Length)).Select(_ => Blank(style))).ToArray();

    private static void Row(StringBuilder sb, int number, Cell[] cells)
    {
        sb.Append($"<row r=\"{number}\">");
        for (int i = 0; i < cells.Length; i++)
        {
            var c = cells[i];
            string reference = ((char)('A' + i)).ToString() + number;
            switch (c.Kind)
            {
                case 'n': sb.Append($"<c r=\"{reference}\" s=\"{c.Style}\"><v>{c.Value}</v></c>"); break;
                case 't': sb.Append($"<c r=\"{reference}\" s=\"{c.Style}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{Escape(c.Value)}</t></is></c>"); break;
                default: sb.Append($"<c r=\"{reference}\" s=\"{c.Style}\"/>"); break;
            }
        }
        sb.Append("</row>");
    }

    /// <summary>XML text: escapes the markup characters and drops characters XML does not allow.</summary>
    private static string Escape(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (char ch in text)
        {
            if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') continue;
            sb.Append(ch switch { '&' => "&amp;", '<' => "&lt;", '>' => "&gt;", '"' => "&quot;", _ => ch.ToString() });
        }
        return sb.ToString();
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace NextMeetingOverlay;

public partial class MainWindow : Window
{
    private const double DockMargin = 2;     // Card already has a 6 px shadow margin
    private const double SnapDistance = 40;  // drop this close to the top edge -> dock
    private static readonly TimeSpan JohnnyLeadTime = TimeSpan.FromMinutes(10);   // button appears this early
    private static readonly TimeSpan StartWindowBefore = TimeSpan.FromMinutes(1); // ask/auto-start window
    private static readonly TimeSpan StartWindowAfter = TimeSpan.FromMinutes(10);
    private static readonly TimeSpan MinFreeGap = TimeSpan.FromMinutes(10);

    private const int HotkeyToggle = 1, HotkeyJoin = 2, HotkeyJohnny = 3, HotkeyClickThrough = 4, HotkeyTodo = 5;
    private const string HotkeyHint = "Strg+Alt+Umschalt";

    private static readonly Brush Orange = Frozen("#FFB020");
    private static readonly Brush Yellow = Frozen("#FFD23F");
    private static readonly Brush Red = Frozen("#FF5A5A");
    private static readonly Brush RecordRed = Frozen("#D93636");
    private static readonly Brush Green = Frozen("#3DDC84");
    private static readonly Brush Gray = Frozen("#80FFFFFF");
    private static readonly Brush DefaultBorder = Frozen("#33FFFFFF");
    private static readonly Brush JohnnyBorder = Frozen("#66FF5A5A");
    private static readonly Brush SkipBorder = Frozen("#55FFFFFF");

    private readonly Settings _s = Settings.Load();
    private readonly OutlookReader _reader = new();
    private readonly DispatcherTimer _tick = new() { Interval = TimeSpan.FromSeconds(1) };
    private readonly DispatcherTimer _refreshTimer = new();
    private readonly DispatcherTimer _johnnyTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private readonly List<(MenuItem Item, Func<bool> IsChecked)> _checkItems = new();
    private readonly HashSet<string> _johnnyHandled = new();   // meetings already asked about / auto-started

    private List<Meeting> _meetings = new();
    private string? _status = "Lade Termine …";
    private Meeting? _shown;
    private bool _refreshing;
    private bool _blink;
    private ContextMenu _menu = null!;
    private MenuItem _keywordItem = null!;
    private Forms.NotifyIcon? _tray;
    private IntPtr _hwnd;
    private bool _userHidden;
    private bool _autoHidden;
    private string? _toolTipKey;

    // Mini mode: the card is shown while hovered, pinned by a click, or around a meeting's start.
    private readonly DispatcherTimer _expandTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
    private readonly DispatcherTimer _collapseTimer = new() { Interval = TimeSpan.FromMilliseconds(900) };
    private bool _hoverExpanded;
    private bool _pinnedExpanded;
    private bool _autoExpanded;

    private bool _johnnyRunning;
    private bool _johnnyAvailable = (Demo.Active && !JohnnyAiClient.Disabled) || JohnnyAiClient.IsInstalled;   // false -> no JohnnyAI UI at all
    private JohnnyStatus? _johnny;       // null = not running or JohnnyAI without control pipe
    private bool _johnnyPolling;
    private Func<Task>? _balloonAction;  // what a click on the current tray balloon does

    public MainWindow()
    {
        InitializeComponent();
        Opacity = _s.Opacity;
        _menu = BuildMenu();
        ContextMenu = _menu;   // on the window, so it works on the mini pill and the card

        _expandTimer.Tick += (_, _) => { _expandTimer.Stop(); _hoverExpanded = true; ApplyExpansion(); };
        _collapseTimer.Tick += (_, _) =>
        {
            _collapseTimer.Stop();
            if (IsMouseOver || _menu.IsOpen) return;
            _hoverExpanded = false;
            ApplyExpansion();
        };
        MouseEnter += (_, _) => { _collapseTimer.Stop(); if (!_hoverExpanded) _expandTimer.Start(); };
        MouseLeave += (_, _) => { _expandTimer.Stop(); if (_hoverExpanded) _collapseTimer.Start(); };
        _menu.Closed += (_, _) => { if (!IsMouseOver && _hoverExpanded) _collapseTimer.Start(); };

        SourceInitialized += (_, _) =>
        {
            _hwnd = new WindowInteropHelper(this).Handle;
            HwndSource.FromHwnd(_hwnd)?.AddHook(WndProc);
            ApplyWindowFlags();
            RegisterHotkeys();
        };
        Loaded += (_, _) =>
        {
            RestorePosition();
            if (!Demo.Active) BuildTray();
            UpdateDisplay();
            _ = RefreshAsync();
            _ = PollJohnnyAsync();
            if (Demo.Active) Dispatcher.BeginInvoke(async () => await ExportDemoScreenshotAsync(), DispatcherPriority.ApplicationIdle);
        };
        SizeChanged += (_, _) => { if (_s.Dock != DockMode.Free) ApplyDock(); };

        _tick.Tick += (_, _) => { UpdateAutoHide(); UpdateDisplay(); };
        _tick.Start();

        _refreshTimer.Interval = TimeSpan.FromSeconds(Math.Max(15, _s.RefreshSeconds));
        _refreshTimer.Tick += async (_, _) => await RefreshAsync();
        _refreshTimer.Start();

        _johnnyTimer.Tick += async (_, _) => await PollJohnnyAsync();
        _johnnyTimer.Start();

        SystemEvents.DisplaySettingsChanged += OnDisplaySettingsChanged;
        SystemEvents.PowerModeChanged += OnPowerModeChanged;
    }

    // ───────────────────────── Data ─────────────────────────

    private async Task RefreshAsync()
    {
        if (Demo.Active)
        {
            _meetings = Demo.Meetings();
            _status = null;
            UpdateDisplay();
            return;
        }
        if (_refreshing) return;
        _refreshing = true;
        try
        {
            var snapshot = _s.Clone();
            var sources = snapshot.Calendars.Where(c => c.Enabled).ToList();
            if (sources.Count == 0)
            {
                _meetings = new();
                _calendarErrors = new();
                _status = "Kein Kalender aktiv (Rechtsklick → Kalender …)";
                return;
            }

            // All calendars in parallel; one failing calendar must not hide the others.
            var results = await Task.WhenAll(sources.Select(src => ReadCalendarAsync(src, snapshot)));

            bool showNames = sources.Count > 1;
            var merged = results
                .Where(r => r.Meetings != null)
                .SelectMany(r => r.Meetings!.Select(m => m with { CalendarName = showNames ? r.Source.Name : "" }))
                .ToList();
            _calendarErrors = results.Where(r => r.Error != null).Select(r => $"{r.Source.Name}: {r.Error}").ToList();

            if (results.Any(r => r.Meetings != null))
            {
                _meetings = Deduplicate(merged);
                _status = null;
            }
            else if (results.All(r => r.NotRunning))
            {
                _meetings = new();
                _status = IsNewOutlookRunning()
                    ? "Neues Outlook: bitte Kalender-Link hinterlegen (Rechtsklick → Kalender …)"
                    : "Outlook ist nicht geöffnet – bitte Outlook starten";
            }
            else if (_meetings.Count == 0)
            {
                // Keep showing the last known meetings on transient errors (Outlook busy, network).
                _status = results.Length == 1 ? "Kalender-Fehler: " + results[0].Error : "Kein Kalender erreichbar";
            }
        }
        finally
        {
            _refreshing = false;
            _toolTipKey = null;
            UpdateDisplay();
        }
    }

    private List<string> _calendarErrors = new();
    private readonly Dictionary<string, IcsCalendarSource> _icsSources = new();

    private sealed record CalendarResult(CalendarSource Source, List<Meeting>? Meetings, string? Error, bool NotRunning = false);

    private async Task<CalendarResult> ReadCalendarAsync(CalendarSource source, Settings snapshot)
    {
        try
        {
            if (source.Kind == CalendarKind.Outlook)
                return new(source, await _reader.Run(() => _reader.ReadMeetings(snapshot)), null);

            var url = source.Url;
            if (string.IsNullOrWhiteSpace(url))
                return new(source, null, "kein Link hinterlegt (auf einem anderen PC eingetragen?)");
            if (!_icsSources.TryGetValue(url, out var ics))
                _icsSources[url] = ics = new IcsCalendarSource();   // own download cache per link
            return new(source, await ics.ReadAsync(snapshot, url), null);
        }
        catch (OutlookNotRunningException)
        {
            return new(source, null, "Outlook ist nicht geöffnet", NotRunning: true);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            return new(source, null, "Link nicht erreichbar");
        }
        catch (Exception ex)
        {
            return new(source, null, ex.Message);
        }
    }

    /// <summary>
    /// The same meeting often sits in two calendars (e.g. work calendar + shared team calendar):
    /// same start, end and subject = one entry. The copy with a join link wins.
    /// </summary>
    private static List<Meeting> Deduplicate(IEnumerable<Meeting> meetings) =>
        meetings
            .GroupBy(m => (m.Start, m.End, m.Subject.Trim().ToLowerInvariant()))
            .Select(g => g.OrderByDescending(m => m.JoinUrl != null).First())
            .OrderBy(m => m.Start).ThenBy(m => m.End)
            .ToList();

    private void InvalidateCalendarCaches()
    {
        foreach (var ics in _icsSources.Values) ics.Invalidate();
    }

    private static bool IsNewOutlookRunning()
    {
        var processes = Process.GetProcessesByName("olk");
        foreach (var p in processes) p.Dispose();
        return processes.Length > 0;
    }

    // ───────────────────────── Display ─────────────────────────

    private void UpdateDisplay()
    {
        var now = DateTime.Now;
        var upcoming = VisibleMeetings().Where(m => m.End > now).ToList();
        var main = upcoming.FirstOrDefault();
        _shown = main;
        _blink = !_blink;

        UpdateJohnnyButton(main, now);
        UpdateMini(main, now);
        UpdateTodos(now);

        if (main == null)
        {
            UpdateOverlapRows(new List<Meeting>(), now);
            Accent.Background = Gray;
            Card.BorderBrush = DefaultBorder;
            TimeText.Visibility = Visibility.Collapsed;
            SubjectText.Text = _status ?? $"Keine Termine in den nächsten {_s.LookaheadDays} Tagen";
            SubjectText.Foreground = Gray;
            LocationText.Visibility = Visibility.Collapsed;
            FollowUpText.Visibility = Visibility.Collapsed;
            JoinButton.Visibility = Visibility.Collapsed;
            SkipButton.Visibility = Visibility.Collapsed;
            Card.ToolTip = "Rechtsklick: Einstellungen";
            SetTrayText(SubjectText.Text);
            return;
        }

        bool running = main.Start <= now;
        var untilStart = main.Start - now;
        bool imminent = !running && untilStart <= TimeSpan.FromMinutes(1);

        Accent.Background = TrafficLight(main, now);
        Card.BorderBrush = imminent && _blink ? Red : DefaultBorder;
        SubjectText.Foreground = Brushes.White;

        string relative = running
            ? $"läuft · noch {FormatSpan(main.End - now)}"
            : $"in {FormatSpan(untilStart)}";
        string range = $"{DayPrefix(main.Start, now)}{main.Start:HH:mm}–{main.End:HH:mm}";
        string subject = main.Subject.Length > 0 ? main.Subject : "(ohne Betreff)";

        // Meetings that start before the main one ends run (partly) in parallel.
        var overlaps = _s.ShowOverlaps
            ? upcoming.Skip(1).Where(m => m.Start < main.End).ToList()
            : new List<Meeting>();
        var groupEnd = overlaps.Select(m => m.End).Append(main.End).Max();

        if (_s.Compact)
        {
            UpdateOverlapRows(new List<Meeting>(), now);
            TimeText.Visibility = Visibility.Collapsed;
            SubjectText.Text = $"{main.Start:HH:mm}  {subject}  · {relative}"
                               + (overlaps.Count > 0 ? $"  (+{overlaps.Count} parallel)" : "");
            LocationText.Visibility = Visibility.Collapsed;
            FollowUpText.Visibility = Visibility.Collapsed;
        }
        else
        {
            TimeText.Visibility = Visibility.Visible;
            TimeText.Text = $"{range} · {relative}";
            SubjectText.Text = subject;
            LocationText.Text = WithCalendar(main.Location, main);
            LocationText.Visibility = LocationText.Text.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
            UpdateOverlapRows(overlaps, now);

            var followUp = upcoming.Skip(1).FirstOrDefault(m => m.Start >= main.Start && !overlaps.Contains(m));
            if (_s.ShowFollowUp && followUp != null && followUp.Start.Date == now.Date)
            {
                var gap = followUp.Start - groupEnd;
                string gapText = gap >= MinFreeGap ? $" · {FormatSpan(gap)} frei davor"
                               : gap <= TimeSpan.Zero ? " · direkt im Anschluss"
                               : "";
                FollowUpText.Text = $"Danach: {followUp.Start:HH:mm} {followUp.Subject}{gapText}";
                FollowUpText.Visibility = Visibility.Visible;
            }
            else FollowUpText.Visibility = Visibility.Collapsed;
        }

        JoinButton.Visibility = main.JoinUrl != null && !_joined.Contains(main.Key) ? Visibility.Visible : Visibility.Collapsed;
        JoinButton.ToolTip = $"Online-Meeting beitreten ({HotkeyHint}+J)";
        SkipButton.Visibility = Visibility.Visible;
        SkipButton.ToolTip = $"Nicht teilnehmen – „{main.Subject}“ im Overlay ausblenden\n(in Outlook bleibt der Termin unverändert, es geht keine Absage raus)";

        UpdateToolTip(main, upcoming, range, now);
        SetTrayText($"{main.Start:HH:mm} {subject} ({relative})");
    }

    /// <summary>
    /// Traffic light shared by the mini pill and the card accent:
    /// green = plenty of time, yellow = starts soon, red = starts in a few minutes or is running.
    /// </summary>
    private Brush TrafficLight(Meeting? m, DateTime now)
    {
        if (m == null) return Gray;
        if (m.Start <= now) return Red;
        var untilStart = m.Start - now;
        if (untilStart <= TimeSpan.FromMinutes(1)) return _blink ? Red : Gray;
        if (untilStart <= TimeSpan.FromMinutes(_s.RedMinutes)) return Red;
        if (untilStart <= TimeSpan.FromMinutes(_s.YellowMinutes)) return Yellow;
        return Green;
    }

    private void UpdateMini(Meeting? main, DateTime now)
    {
        MiniDot.Fill = TrafficLight(main, now);
        MiniRecording.Visibility = _johnny?.Recording == true ? Visibility.Visible : Visibility.Collapsed;

        if (main == null)
        {
            MiniText.Text = _status != null ? "!" : "frei";
        }
        else if (main.Start <= now)
        {
            MiniText.Text = $"noch {FormatSpan(main.End - now)}";
        }
        else
        {
            var untilStart = main.Start - now;
            MiniText.Text = untilStart <= TimeSpan.FromMinutes(60) ? $"in {FormatSpan(untilStart)}"
                          : main.Start.Date == now.Date ? $"{main.Start:HH:mm}"
                          : $"{DayPrefix(main.Start, now).Trim()} {main.Start:HH:mm}";
        }
        MiniPill.ToolTip = main == null ? _status ?? "Keine anstehenden Termine" : null;

        // "+1": meetings overlapping the shown one (same rule as the parallel rows in the card).
        int overlapCount = main == null || !_s.ShowOverlaps ? 0
            : VisibleMeetings().Count(m => m != main && m.End > now && m.Start >= main.Start && m.Start < main.End);
        MiniOverlapText.Text = $"+{overlapCount}";
        MiniOverlap.Visibility = overlapCount > 0 ? Visibility.Visible : Visibility.Collapsed;

        // Briefly show the details around a meeting's start (join button, JohnnyAI prompt).
        _autoExpanded = _s.AutoExpandAtStart && main != null &&
                        now >= main.Start - TimeSpan.FromMinutes(1) && now <= main.Start + TimeSpan.FromMinutes(2);
        ApplyExpansion();
    }

    private void ApplyExpansion()
    {
        bool expanded = !_s.MiniMode || _hoverExpanded || _pinnedExpanded || _autoExpanded;
        var cardVisibility = expanded ? Visibility.Visible : Visibility.Collapsed;
        if (Card.Visibility == cardVisibility) return;
        Card.Visibility = cardVisibility;
        MiniPill.Visibility = expanded ? Visibility.Collapsed : Visibility.Visible;
    }

    private readonly List<(Meeting Meeting, TextBlock Text, string RowKey)> _overlapRows = new();

    /// <summary>
    /// One row per meeting running in parallel to the main one, with its own join button.
    /// Rows are only rebuilt when the set of meetings changes, so a click on a button is never
    /// lost to the per-second refresh; the time text is updated in place.
    /// </summary>
    private void UpdateOverlapRows(List<Meeting> overlaps, DateTime now)
    {
        // Rebuild when the meetings or their joined state change (join button disappears).
        string RowKey(Meeting m) => m.Key + (_joined.Contains(m.Key) ? "|joined" : "");
        bool sameSet = overlaps.Count == _overlapRows.Count &&
                       overlaps.Select(RowKey).SequenceEqual(_overlapRows.Select(r => r.RowKey));
        if (!sameSet)
        {
            OverlapPanel.Children.Clear();
            _overlapRows.Clear();
            foreach (var m in overlaps)
            {
                var grid = new Grid { Margin = new Thickness(0, 2, 0, 0) };
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var text = new TextBlock
                {
                    FontSize = 12,
                    Foreground = Frozen("#D8FFFFFF"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                };
                grid.Children.Add(text);

                if (m.JoinUrl != null && !_joined.Contains(m.Key))
                {
                    var join = new Button
                    {
                        Template = (ControlTemplate)FindResource("PillButton"),
                        Content = "Beitreten",
                        Background = JoinButton.Background,
                        BorderBrush = JoinButton.BorderBrush,
                        Margin = new Thickness(8, 0, 0, 0),
                        Cursor = Cursors.Hand,
                        Focusable = false,
                        ToolTip = $"„{m.Subject}“ beitreten",
                    };
                    var meeting = m;
                    join.Click += (_, _) => JoinMeeting(meeting);
                    Grid.SetColumn(join, 1);
                    grid.Children.Add(join);
                }

                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                var skip = CreateSkipButton(m);
                Grid.SetColumn(skip, 2);
                grid.Children.Add(skip);

                OverlapPanel.Children.Add(grid);
                _overlapRows.Add((m, text, RowKey(m)));
            }
        }

        OverlapPanel.Visibility = _overlapRows.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        foreach (var (m, text, _) in _overlapRows)
        {
            string state = m.Start <= now ? "läuft" : $"in {FormatSpan(m.Start - now)}";
            text.Text = $"∥ {m.Start:HH:mm}–{m.End:HH:mm} {m.Subject} · {state}";
            text.ToolTip = $"Parallel: {m.Subject} ({m.Start:HH:mm}–{m.End:HH:mm})";
        }
    }

    // ───────────────────────── Todos ─────────────────────────

    private readonly List<TodoItem> _todos = Demo.Active ? Demo.Todos() : TodoStore.Load();
    private string? _todoSignature;
    private static readonly Brush TodoDueBrush = Frozen("#CCE08A1E");
    private static readonly Brush TodoNormalBrush = Frozen("#33FFFFFF");

    private void SaveTodos()
    {
        if (!Demo.Active) TodoStore.Save(_todos);
        _todoSignature = null;
        UpdateDisplay();
    }

    private List<TodoItem> OpenTodos() =>
        _todos.Where(t => !t.Done)
              .OrderByDescending(t => t.IsDue)
              .ThenBy(t => t.RemindAt ?? DateTime.MaxValue)
              .ThenBy(t => t.Created)
              .ToList();

    /// <summary>Counter in the pill + todo rows in the card; rows only rebuilt when something changed.</summary>
    private void UpdateTodos(DateTime now)
    {
        CheckTodoReminders(now);
        var open = OpenTodos();
        bool anyDue = open.Any(t => t.IsDue);

        MiniTodo.Visibility = _s.ShowTodos && open.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        MiniTodoText.Text = $"☐ {open.Count}";
        MiniTodo.Background = anyDue ? TodoDueBrush : TodoNormalBrush;
        MiniTodo.ToolTip = anyDue ? "Ein Todo ist fällig" : $"{open.Count} offene Todos";

        TodoPanel.Visibility = _s.ShowTodos ? Visibility.Visible : Visibility.Collapsed;
        if (!_s.ShowTodos) return;

        string signature = string.Join(";", open.Select(t => $"{t.Id}|{t.Text}|{t.ReminderLabel}"));
        if (signature == _todoSignature) return;
        _todoSignature = signature;

        TodoPanel.Children.Clear();
        var header = new DockPanel { Margin = new Thickness(0, 0, 0, 2) };
        var add = new Button
        {
            Template = (ControlTemplate)FindResource("PillButton"),
            Content = "＋ Todo",
            Background = Brushes.Transparent,
            BorderBrush = SkipBorder,
            Cursor = Cursors.Hand,
            Focusable = false,
            ToolTip = $"Todo hinzufügen ({HotkeyHint}+N)",
        };
        add.Click += (_, _) => AddTodo();
        DockPanel.SetDock(add, Dock.Right);
        header.Children.Add(add);
        header.Children.Add(new TextBlock
        {
            Text = open.Count == 0 ? "Todos – nichts offen" : "Todos",
            FontSize = 11,
            Foreground = Gray,
            VerticalAlignment = VerticalAlignment.Center,
        });
        TodoPanel.Children.Add(header);

        const int maxRows = 8;
        foreach (var todo in open.Take(maxRows))
            TodoPanel.Children.Add(CreateTodoRow(todo));
        if (open.Count > maxRows)
            TodoPanel.Children.Add(new TextBlock { Text = $"… und {open.Count - maxRows} weitere", FontSize = 11, Foreground = Gray });
    }

    private UIElement CreateTodoRow(TodoItem todo)
    {
        var grid = new Grid { Margin = new Thickness(0, 2, 0, 0) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var check = new CheckBox { VerticalAlignment = VerticalAlignment.Center, Focusable = false, ToolTip = "Erledigt" };
        check.Checked += (_, _) => CompleteTodo(todo);
        grid.Children.Add(check);

        var text = new TextBlock
        {
            Text = todo.Text,
            FontSize = 12,
            Foreground = Brushes.White,
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 6, 0),
            ToolTip = todo.Text,
        };
        Grid.SetColumn(text, 1);
        grid.Children.Add(text);

        var label = new TextBlock
        {
            Text = todo.ReminderLabel,
            FontSize = 11,
            Foreground = todo.IsDue ? Orange : Gray,
            FontWeight = todo.IsDue ? FontWeights.SemiBold : FontWeights.Normal,
            VerticalAlignment = VerticalAlignment.Center,
            MaxWidth = 170,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };
        Grid.SetColumn(label, 2);
        grid.Children.Add(label);

        var outlook = SmallButton("📅", "Als Erinnerungstermin in Outlook eintragen (öffnet den Termin in Outlook, Kalender wählen und speichern)");
        outlook.Click += (_, _) => OpenInOutlookAsAppointment(todo, todo.RemindAt ?? DateTime.Now.AddHours(1));
        Grid.SetColumn(outlook, 3);
        grid.Children.Add(outlook);

        var delete = SmallButton("✕", "Todo löschen");
        delete.Click += (_, _) => DeleteTodo(todo);
        Grid.SetColumn(delete, 4);
        grid.Children.Add(delete);
        return grid;
    }

    private Button SmallButton(string content, string tooltip) => new()
    {
        Template = (ControlTemplate)FindResource("PillButton"),
        Content = content,
        Background = Brushes.Transparent,
        BorderBrush = Brushes.Transparent,
        Margin = new Thickness(2, 0, 0, 0),
        Cursor = Cursors.Hand,
        Focusable = false,
        ToolTip = tooltip,
    };

    private void CompleteTodo(TodoItem todo)
    {
        todo.Done = true;
        todo.DoneAt = DateTime.Now;
        SaveTodos();
        ShowActionBalloon("Todo erledigt", $"„{todo.Text}“\nHier klicken zum Rückgängigmachen.", () =>
        {
            todo.Done = false;
            todo.DoneAt = null;
            SaveTodos();
            return Task.CompletedTask;
        }, timeoutMs: 6000);
    }

    private void DeleteTodo(TodoItem todo)
    {
        int index = _todos.IndexOf(todo);
        _todos.Remove(todo);
        SaveTodos();
        ShowActionBalloon("Todo gelöscht", $"„{todo.Text}“\nHier klicken zum Rückgängigmachen.", () =>
        {
            _todos.Insert(Math.Min(Math.Max(index, 0), _todos.Count), todo);
            SaveTodos();
            return Task.CompletedTask;
        }, timeoutMs: 6000);
    }

    /// <summary>One balloon for all todos that became due (several at once are bundled).</summary>
    private void CheckTodoReminders(DateTime now)
    {
        if (Demo.Active) return;
        var due = _todos.Where(t => !t.Done && !t.Reminded && t.RemindAt is DateTime at && at <= now).ToList();
        if (due.Count == 0) return;
        foreach (var t in due) t.Reminded = true;
        TodoStore.Save(_todos);
        _todoSignature = null;

        string text = due.Count == 1
            ? $"{due[0].Text}" + (due[0].Reminder == TodoReminder.AfterMeeting ? $"\n(nach {due[0].MeetingSubject})" : "")
            : string.Join("\n", due.Take(4).Select(t => "• " + t.Text)) + (due.Count > 4 ? $"\n… und {due.Count - 4} weitere" : "");
        ShowActionBalloon(due.Count == 1 ? "Erinnerung" : $"{due.Count} Todos fällig", text + "\nKlicken: Todo-Liste öffnen", () =>
        {
            SetUserHidden(false);
            _pinnedExpanded = true;
            ApplyExpansion();
            return Task.CompletedTask;
        }, timeoutMs: 20000);
    }

    private void OpenInOutlookAsAppointment(TodoItem todo, DateTime start)
    {
        try
        {
            var path = OutlookReminderFile.Create(todo, start, _s.TodoOutlookPrivate, _s.TodoOutlookFree);
            if (!TryStart(path))
                ShowInfo("Kein Programm für Kalenderdateien (.ics) gefunden – ist Outlook installiert?");
        }
        catch (Exception ex)
        {
            ShowInfo("Termin konnte nicht erstellt werden: " + ex.Message);
        }
    }

    /// <summary>Quick capture: text, optional reminder (after a meeting / at a time), optional Outlook appointment.</summary>
    private void AddTodo()
    {
        var now = DateTime.Now;
        var choices = new List<(string Label, TodoReminder Kind, DateTime? At, string? Meeting)>
        {
            ("Keine Erinnerung", TodoReminder.None, null, null),
        };
        foreach (var m in VisibleMeetings().Where(m => m.End > now && m.Start.Date <= now.Date).Take(6))
            choices.Add(($"Nach „{m.Subject}“ (endet {m.End:HH:mm})", TodoReminder.AfterMeeting, m.End, m.Subject));
        choices.Add(("In 30 Minuten", TodoReminder.At, now.AddMinutes(30), null));
        choices.Add(("In 1 Stunde", TodoReminder.At, now.AddHours(1), null));
        if (now.Hour < 16) choices.Add(("Heute 16:00", TodoReminder.At, now.Date.AddHours(16), null));
        choices.Add(("Morgen 09:00", TodoReminder.At, now.Date.AddDays(1).AddHours(9), null));
        int customIndex = choices.Count;
        choices.Add(("Eigene Zeit …", TodoReminder.At, null, null));

        var text = new TextBox { Margin = new Thickness(0, 2, 0, 10), Padding = new Thickness(3) };
        var reminder = new ComboBox { Margin = new Thickness(0, 2, 0, 6) };
        foreach (var c in choices) reminder.Items.Add(c.Label);
        // Default: after the meeting that is running or next - the typical "nach dem Call erledigen".
        reminder.SelectedIndex = choices.FindIndex(c => c.Kind == TodoReminder.AfterMeeting) is int i && i > 0 ? i : 0;

        var date = new DatePicker { SelectedDate = now.Date, Width = 130, Margin = new Thickness(0, 0, 8, 0) };
        var time = new TextBox { Text = now.AddHours(1).ToString("HH:00"), Width = 60, Padding = new Thickness(3) };
        var custom = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10), IsEnabled = false };
        custom.Children.Add(date);
        custom.Children.Add(time);
        custom.Children.Add(new TextBlock { Text = "Uhr", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0, 0, 0) });

        var outlook = new CheckBox { Content = "Zusätzlich als Erinnerungstermin in Outlook eintragen", Margin = new Thickness(0, 0, 0, 4) };
        var isPrivate = new CheckBox { Content = "als privat markieren", IsChecked = _s.TodoOutlookPrivate, Margin = new Thickness(20, 0, 0, 2) };
        var asFree = new CheckBox { Content = "als „frei“ eintragen (blockiert den Kalender nicht)", IsChecked = _s.TodoOutlookFree, Margin = new Thickness(20, 0, 0, 10) };
        var outlookHint = new TextBlock
        {
            Text = "Outlook öffnet den Termin – dort Kalender wählen und speichern.",
            Foreground = Brushes.DimGray,
            FontSize = 11,
            Margin = new Thickness(20, 0, 0, 10),
        };

        void SyncEnabled()
        {
            bool hasTime = reminder.SelectedIndex > 0;
            custom.IsEnabled = reminder.SelectedIndex == customIndex;
            outlook.IsEnabled = hasTime;
            if (!hasTime) outlook.IsChecked = false;
            isPrivate.IsEnabled = asFree.IsEnabled = outlook.IsChecked == true;
        }
        reminder.SelectionChanged += (_, _) => SyncEnabled();
        outlook.Checked += (_, _) => SyncEnabled();
        outlook.Unchecked += (_, _) => SyncEnabled();

        var save = new Button { Content = "Hinzufügen", Width = 100, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new Button { Content = "Abbrechen", Width = 90, IsCancel = true };
        var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        footer.Children.Add(save);
        footer.Children.Add(cancel);

        var panel = new StackPanel { Margin = new Thickness(14) };
        panel.Children.Add(new TextBlock { Text = "Was möchtest du erledigen?" });
        panel.Children.Add(text);
        panel.Children.Add(new TextBlock { Text = "Erinnern:" });
        panel.Children.Add(reminder);
        panel.Children.Add(custom);
        panel.Children.Add(outlook);
        panel.Children.Add(isPrivate);
        panel.Children.Add(asFree);
        panel.Children.Add(outlookHint);
        panel.Children.Add(footer);

        var dialog = new Window
        {
            Title = "Todo hinzufügen",
            Content = panel,
            Width = 440,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize,
            Topmost = true,
        };
        save.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(text.Text)) { text.Focus(); return; }
            if (reminder.SelectedIndex == customIndex && CustomTime() == null)
            {
                MessageBox.Show(dialog, "Bitte eine Uhrzeit wie 14:30 eingeben.", "Todo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            dialog.DialogResult = true;
        };
        dialog.Loaded += (_, _) => { dialog.Activate(); text.Focus(); };
        SyncEnabled();

        DateTime? CustomTime() =>
            date.SelectedDate is DateTime d && TimeSpan.TryParse(time.Text.Trim().Replace('.', ':'), out var t) && t < TimeSpan.FromDays(1)
                ? d.Date + t : null;

        if (dialog.ShowDialog() != true) return;

        var choice = choices[reminder.SelectedIndex];
        var todo = new TodoItem
        {
            Text = text.Text.Trim(),
            Reminder = choice.Kind,
            RemindAt = reminder.SelectedIndex == customIndex ? CustomTime() : choice.At,
            MeetingSubject = choice.Meeting,
        };
        _todos.Add(todo);
        SaveTodos();

        if (outlook.IsChecked == true && todo.RemindAt is DateTime start)
        {
            _s.TodoOutlookPrivate = isPrivate.IsChecked == true;
            _s.TodoOutlookFree = asFree.IsChecked == true;
            _s.Save();
            OpenInOutlookAsAppointment(todo, start);
        }
    }

    // ───────────────────────── Not attending (skip) ─────────────────────────

    private IEnumerable<Meeting> VisibleMeetings() =>
        _s.SkippedMeetings.Count == 0 ? _meetings : _meetings.Where(m => !_s.SkippedMeetings.Contains(m.Key));

    /// <summary>
    /// "Nicht teilnehmen": hides this one occurrence in the overlay (no ampel, no JohnnyAI prompt).
    /// Outlook stays untouched - nothing is declined or sent to the organizer.
    /// </summary>
    private void SkipMeeting(Meeting meeting)
    {
        if (!_s.SkippedMeetings.Contains(meeting.Key)) _s.SkippedMeetings.Add(meeting.Key);
        _johnnyHandled.Add(meeting.Key);
        _s.Save();
        _toolTipKey = null;
        UpdateDisplay();
        ShowActionBalloon("Termin ausgeblendet",
            $"„{meeting.Subject}“ ({meeting.Start:HH:mm}) wird im Overlay nicht mehr angezeigt – in Outlook bleibt er unverändert.\nHier klicken zum Rückgängigmachen.",
            () =>
            {
                _s.SkippedMeetings.Remove(meeting.Key);
                _s.Save();
                _toolTipKey = null;
                UpdateDisplay();
                return Task.CompletedTask;
            }, timeoutMs: 8000);
    }

    private void ShowSkippedMeetingsAgain()
    {
        _s.SkippedMeetings.Clear();
        _s.Save();
        _toolTipKey = null;
        UpdateDisplay();
    }

    private Button CreateSkipButton(Meeting meeting)
    {
        var button = new Button
        {
            Template = (ControlTemplate)FindResource("PillButton"),
            Content = "✕",
            Background = Brushes.Transparent,
            BorderBrush = SkipBorder,
            Margin = new Thickness(6, 0, 0, 0),
            Cursor = Cursors.Hand,
            Focusable = false,
            ToolTip = $"Nicht teilnehmen – „{meeting.Subject}“ im Overlay ausblenden\n(in Outlook bleibt der Termin unverändert, es geht keine Absage raus)",
        };
        button.Click += (_, _) => SkipMeeting(meeting);
        return button;
    }

    private void OnSkipClick(object sender, RoutedEventArgs e)
    {
        if (_shown != null) SkipMeeting(_shown);
    }

    /// <summary>Appends the calendar name ("Raum 2.04 · Privat") when several calendars are active.</summary>
    private static string WithCalendar(string text, Meeting m) =>
        m.CalendarName.Length == 0 ? text
        : text.Length == 0 ? m.CalendarName
        : $"{text} · {m.CalendarName}";

    /// <summary>Hover card: the shown meeting plus the rest of today including free gaps.</summary>
    private void UpdateToolTip(Meeting main, List<Meeting> upcoming, string range, DateTime now)
    {
        string key = $"{now:yyyyMMddHHmm}|{main.Key}|{upcoming.Count}|{_s.ShowDayOverview}|{string.Join(";", _calendarErrors)}";
        if (key == _toolTipKey) return;
        _toolTipKey = key;

        var panel = new StackPanel { MaxWidth = 440 };
        panel.Children.Add(new TextBlock { Text = main.Subject, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(new TextBlock { Text = WithCalendar(range + (main.Location.Length > 0 ? " · " + main.Location : ""), main), TextWrapping = TextWrapping.Wrap });

        if (_s.ShowDayOverview)
        {
            var today = upcoming.Where(m => m.Start.Date <= now.Date).ToList();
            if (today.Count > 0)
            {
                panel.Children.Add(new TextBlock { Text = "Heute noch", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 8, 0, 2) });
                var cursor = now;
                foreach (var m in today)
                {
                    var gap = m.Start - cursor;
                    if (gap >= MinFreeGap)
                        panel.Children.Add(new TextBlock { Text = $"        {FormatSpan(gap)} frei", Foreground = Brushes.SeaGreen, FontStyle = FontStyles.Italic });
                    panel.Children.Add(new TextBlock
                    {
                        Text = $"{m.Start:HH:mm}–{m.End:HH:mm}  {m.Subject}" + (m.JoinUrl != null ? "  (online)" : "")
                               + (m.CalendarName.Length > 0 ? $"  · {m.CalendarName}" : ""),
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        FontWeight = m == main ? FontWeights.SemiBold : FontWeights.Normal,
                    });
                    if (m.End > cursor) cursor = m.End;
                }
                if (cursor.Date == now.Date && cursor.Hour < 23)
                    panel.Children.Add(new TextBlock { Text = $"Ab {cursor:HH:mm} frei", Foreground = Brushes.SeaGreen, FontStyle = FontStyles.Italic });
            }
        }

        if (_calendarErrors.Count > 0)
            panel.Children.Add(new TextBlock
            {
                Text = "Nicht erreichbar – " + string.Join(" · ", _calendarErrors),
                Foreground = Brushes.DarkOrange,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 8, 0, 0),
            });

        panel.Children.Add(new TextBlock
        {
            Text = $"Doppelklick: Outlook öffnen · Rechtsklick: Einstellungen\n{HotkeyHint}+O ein/aus · +J beitreten"
                   + (_johnnyAvailable ? " · +A Aufnahme" : "") + " · +N Todo · +K Klick-durch",
            Foreground = Brushes.Gray,
            FontSize = 11,
            Margin = new Thickness(0, 8, 0, 0),
        });
        Card.ToolTip = new ToolTip { Content = panel };
    }

    private static string FormatSpan(TimeSpan span)
    {
        int minutes = (int)Math.Ceiling(span.TotalMinutes);
        if (minutes < 1) return "< 1 Min";
        if (minutes < 60) return $"{minutes} Min";
        if (minutes < 24 * 60)
        {
            int h = minutes / 60, m = minutes % 60;
            return m == 0 ? $"{h} Std" : $"{h} Std {m} Min";
        }
        int days = (int)Math.Round(span.TotalDays);
        return days == 1 ? "1 Tag" : $"{days} Tagen";
    }

    private static string DayPrefix(DateTime start, DateTime now)
    {
        if (start.Date == now.Date) return "";
        if (start.Date == now.Date.AddDays(1)) return "Morgen ";
        return start.ToString("ddd dd.MM. ");
    }

    // ───────────────────────── JohnnyAI ─────────────────────────

    private async Task PollJohnnyAsync()
    {
        if (Demo.Active)
        {
            _johnnyRunning = !JohnnyAiClient.Disabled;
            _johnny = _johnnyRunning ? Demo.Johnny : null;
            UpdateJohnnyButton(_shown, DateTime.Now);
            return;
        }
        if (_johnnyPolling) return;
        _johnnyPolling = true;
        try
        {
            _johnnyRunning = JohnnyAiClient.IsRunning;
            _johnny = _johnnyRunning ? await JohnnyAiClient.GetStatusAsync() : null;
            HandleMeetingStart();

            // JohnnyAI installed/removed meanwhile: show or hide its menu, hint and hotkey.
            bool available = _johnnyRunning || JohnnyAiClient.IsInstalled;
            if (available != _johnnyAvailable)
            {
                _johnnyAvailable = available;
                _toolTipKey = null;
                RegisterHotkeys();
            }
        }
        finally
        {
            _johnnyPolling = false;
        }
        UpdateJohnnyButton(_shown, DateTime.Now);
    }

    private bool IsInStartWindow(Meeting m, DateTime now) =>
        now >= m.Start - StartWindowBefore && now <= m.Start + StartWindowAfter && now < m.End;

    private bool IsEligibleForJohnny(Meeting m) => !_s.JohnnyOnlyOnlineMeetings || m.JoinUrl != null;

    /// <summary>At meeting start: offer (balloon) or automatically start a JohnnyAI recording once per meeting.</summary>
    private void HandleMeetingStart()
    {
        var m = _shown;
        var now = DateTime.Now;
        if (m == null || _johnny == null || _johnny.Recording || _s.JohnnyStart == JohnnyStartMode.Off) return;
        if (!IsInStartWindow(m, now) || !IsEligibleForJohnny(m) || !_johnnyHandled.Add(m.Key)) return;

        if (_s.JohnnyStart == JohnnyStartMode.Auto)
        {
            _ = StartJohnnyAsync(m);
            ShowInfo($"JohnnyAI-Aufnahme gestartet: {m.Subject}");
        }
        else
        {
            ShowActionBalloon("Termin beginnt – Aufnahme starten?",
                $"{m.Subject}\nHier klicken oder ⏺ im Overlay, um JohnnyAI aufnehmen zu lassen.",
                async () =>
                {
                    if (_johnny == null || _johnny.Recording) return;
                    await StartJohnnyAsync(m);
                    await PollJohnnyAsync();
                });
        }
    }

    private void UpdateJohnnyButton(Meeting? main, DateTime now)
    {
        if (!_s.ShowJohnnyButton || !_johnnyRunning)
        {
            JohnnyButton.Visibility = Visibility.Collapsed;
            return;
        }

        if (_johnny?.Recording == true)
        {
            JohnnyButton.Visibility = Visibility.Visible;
            JohnnyButton.Content = $"⏹ {_johnny.Time}";
            JohnnyButton.Background = RecordRed;
            JohnnyButton.BorderBrush = RecordRed;
            JohnnyButton.ToolTip = "JohnnyAI nimmt auf" + (_johnny.Subject.Length > 0 ? $": {_johnny.Subject}" : "")
                                   + $" – klicken zum Beenden ({HotkeyHint}+A)";
            return;
        }

        var candidate = JohnnyCandidate(main, now);
        if (candidate == null)
        {
            JohnnyButton.Visibility = Visibility.Collapsed;
            return;
        }

        JohnnyButton.Visibility = Visibility.Visible;
        JohnnyButton.BorderBrush = JohnnyBorder;
        if (_johnny == null)
        {
            // JohnnyAI version without control pipe: hotkey toggle only, state unknown.
            JohnnyButton.Content = "⏺";
            JohnnyButton.Background = Brushes.Transparent;
            JohnnyButton.ToolTip = "JohnnyAI-Aufnahme umschalten (wie Win+Shift+R) – Aufnahmestatus erst ab der neuen JohnnyAI-Version sichtbar";
            return;
        }

        bool prompt = _s.JohnnyStart == JohnnyStartMode.Ask && IsInStartWindow(candidate, now) && IsEligibleForJohnny(candidate);
        JohnnyButton.Content = prompt ? "⏺ Aufnahme starten?" : "⏺ Aufnahme";
        JohnnyButton.Background = prompt && _blink ? Orange : Brushes.Transparent;
        JohnnyButton.ToolTip = $"JohnnyAI-Aufnahme starten – verknüpft mit „{candidate.Subject}“ ({HotkeyHint}+A)";
    }

    /// <summary>The meeting a recording would belong to: running or starting within the lead time.</summary>
    private static Meeting? JohnnyCandidate(Meeting? main, DateTime now) =>
        main != null && main.Start - now <= JohnnyLeadTime ? main : null;

    private async void OnJohnnyClick(object sender, RoutedEventArgs e) => await ToggleJohnnyAsync();

    private async Task ToggleJohnnyAsync()
    {
        if (!_johnnyRunning)
        {
            ShowInfo("JohnnyAI läuft nicht.");
            return;
        }
        if (_johnny == null)
        {
            if (!JohnnyAiBridge.ToggleRecording()) ShowInfo("JohnnyAI wurde nicht gefunden.");
            return;
        }

        if (_johnny.Recording) await JohnnyAiClient.StopAsync();
        else await StartJohnnyAsync(JohnnyCandidate(_shown, DateTime.Now));
        await PollJohnnyAsync();
    }

    private async Task StartJohnnyAsync(Meeting? meeting)
    {
        if (meeting != null) _johnnyHandled.Add(meeting.Key);
        if (await JohnnyAiClient.StartAsync(meeting) == null)
            ShowInfo("JohnnyAI hat nicht geantwortet.");
    }

    // ───────────────────────── Actions ─────────────────────────

    private void OnJoinClick(object sender, RoutedEventArgs e) => JoinMeeting(_shown);

    // Meetings joined via the overlay this session: their "Beitreten" button is hidden.
    // Re-joining stays possible via hotkey or the context menu ("Erneut beitreten").
    private readonly HashSet<string> _joined = new();

    private void JoinMeeting(Meeting? meeting)
    {
        if (meeting?.JoinUrl == null) return;
        Join(meeting.JoinUrl);
        if (_joined.Add(meeting.Key)) UpdateDisplay();
    }

    private static void Join(string? url)
    {
        if (url == null) return;

        // Prefer the Teams app directly over the browser hop.
        const string teamsPrefix = "https://teams.microsoft.com";
        if (url.StartsWith(teamsPrefix, StringComparison.OrdinalIgnoreCase) && TryStart("msteams:" + url[teamsPrefix.Length..]))
            return;
        TryStart(url);
    }

    private async void OpenInOutlook()
    {
        var meeting = _shown;
        if (meeting?.Source == MeetingSource.Outlook)
        {
            try
            {
                await _reader.Run(() => _reader.Open(meeting));
                return;
            }
            catch (Exception ex) when (ex is not OutlookNotRunningException)
            {
                ShowInfo("Termin konnte nicht geöffnet werden: " + ex.Message);
                return;
            }
            catch (OutlookNotRunningException) { /* fall through: just open Outlook */ }
        }

        // ICS source (or COM unavailable): open the Outlook calendar itself.
        if (IsNewOutlookRunning())
            TryStart("explorer.exe", @"shell:AppsFolder\Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows");
        else if (!TryStart("outlook.exe", "/select outlook:calendar"))
            ShowInfo("Outlook konnte nicht geöffnet werden.");
    }

    private static bool TryStart(string target, string? arguments = null)
    {
        try
        {
            var psi = new ProcessStartInfo(target) { UseShellExecute = true };
            if (arguments != null) psi.Arguments = arguments;
            Process.Start(psi);
            return true;
        }
        catch { return false; }
    }

    private void ShowInfo(string text)
    {
        _balloonAction = null;   // a plain info balloon must not trigger an older balloon's action
        ShowInfoBalloon(text);
    }

    private void ShowInfoBalloon(string text) =>
        _tray?.ShowBalloonTip(4000, "Nächster Termin", text, Forms.ToolTipIcon.Info);

    // ───────────────────────── Window flags, visibility, hotkeys ─────────────────────────

    private void ApplyWindowFlags()
    {
        if (_hwnd == IntPtr.Zero) return;
        NativeMethods.SetExcludeFromCapture(_hwnd, _s.HideFromScreenShare && !Demo.Active);
        NativeMethods.SetClickThrough(_hwnd, _s.ClickThrough);
    }

    private void UpdateAutoHide()
    {
        bool hide = _s.HideInFullscreen && NativeMethods.IsFullscreenOrPresenting();
        if (hide == _autoHidden) return;
        _autoHidden = hide;
        ApplyVisibility();
    }

    private void ApplyVisibility()
    {
        bool show = !_userHidden && !_autoHidden;
        if (show && !IsVisible) Show();
        else if (!show && IsVisible) Hide();
    }

    private void SetUserHidden(bool hidden)
    {
        _userHidden = hidden;
        ApplyVisibility();
    }

    private void RegisterHotkeys()
    {
        if (_hwnd == IntPtr.Zero) return;
        UnregisterHotkeys();
        if (!_s.HotkeysEnabled) return;

        const uint mods = NativeMethods.MOD_CONTROL | NativeMethods.MOD_ALT | NativeMethods.MOD_SHIFT | NativeMethods.MOD_NOREPEAT;
        var failed = new List<string>();
        foreach (var (id, key) in new[] { (HotkeyToggle, 'O'), (HotkeyJoin, 'J'), (HotkeyJohnny, 'A'), (HotkeyClickThrough, 'K'), (HotkeyTodo, 'N') })
            if ((id != HotkeyJohnny || _johnnyAvailable) && !NativeMethods.RegisterHotKey(_hwnd, id, mods, key))
                failed.Add($"{HotkeyHint}+{key}");
        if (failed.Count > 0)
            Dispatcher.BeginInvoke(() => ShowInfo("Bereits von einem anderen Programm belegt: " + string.Join(", ", failed)),
                DispatcherPriority.ApplicationIdle);
    }

    private void UnregisterHotkeys()
    {
        foreach (var id in new[] { HotkeyToggle, HotkeyJoin, HotkeyJohnny, HotkeyClickThrough, HotkeyTodo })
            NativeMethods.UnregisterHotKey(_hwnd, id);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg != NativeMethods.WM_HOTKEY) return IntPtr.Zero;
        switch (wParam.ToInt32())
        {
            case HotkeyToggle:
                SetUserHidden(!_userHidden);
                break;
            case HotkeyJoin:
                JoinMeeting(_shown);   // also re-joins when the button is already hidden
                break;
            case HotkeyJohnny:
                _ = ToggleJohnnyAsync();
                break;
            case HotkeyTodo:
                AddTodo();
                break;
            case HotkeyClickThrough:
                _s.ClickThrough = !_s.ClickThrough;
                _s.Save();
                ApplyWindowFlags();
                ShowInfo(_s.ClickThrough ? "Klick-durch-Modus an – Einstellungen über das Tray-Symbol." : "Klick-durch-Modus aus.");
                break;
            default:
                return IntPtr.Zero;
        }
        handled = true;
        return IntPtr.Zero;
    }

    // ───────────────────────── Position & docking ─────────────────────────

    private void OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            OpenInOutlook();
            return;
        }
        double left = Left, top = Top;
        try { DragMove(); }
        catch (InvalidOperationException) { return; }

        if (Math.Abs(Left - left) < 3 && Math.Abs(Top - top) < 3)
        {
            // A click, not a drag: pin / unpin the detail view in mini mode.
            if (_s.MiniMode)
            {
                _pinnedExpanded = !_pinnedExpanded;
                if (!_pinnedExpanded) _hoverExpanded = false;
                ApplyExpansion();
            }
            if (_s.Dock != DockMode.Free) ApplyDock();  // undo the few pixels a click may have moved
            return;
        }
        SnapAfterDrag();
    }

    private void SnapAfterDrag()
    {
        var screen = ScreenOfWindow();
        _s.ScreenDevice = screen.DeviceName;
        var area = WorkAreaInDips(screen);

        if (Top - area.Top < SnapDistance)
        {
            double relative = (Left + ActualWidth / 2 - area.Left) / area.Width;
            _s.Dock = relative < 1.0 / 3 ? DockMode.TopLeft
                    : relative < 2.0 / 3 ? DockMode.TopCenter
                    : DockMode.TopRight;
            ApplyDock();
        }
        else
        {
            _s.Dock = DockMode.Free;
            _s.Left = Left;
            _s.Top = Top;
        }
        _s.Save();
    }

    private void SetDock(DockMode dock)
    {
        if (dock == DockMode.Free)
        {
            _s.Dock = DockMode.Free;
            _s.Left = Left;
            _s.Top = Top;
            return;
        }
        _s.ScreenDevice ??= ScreenOfWindow().DeviceName;
        _s.Dock = dock;
        ApplyDock();
    }

    private void ApplyDock()
    {
        var area = WorkAreaInDips(TargetScreen());
        double width = ActualWidth;
        Left = _s.Dock switch
        {
            DockMode.TopLeft => area.Left + DockMargin,
            DockMode.TopCenter => area.Left + (area.Width - width) / 2,
            _ => area.Right - width - DockMargin,
        };
        Top = area.Top + DockMargin;
    }

    private void RestorePosition()
    {
        if (_s.Dock != DockMode.Free)
        {
            ApplyDock();
            return;
        }
        Left = _s.Left;
        Top = _s.Top;

        // Saved position on a monitor that is no longer attached -> fall back to top right.
        var probe = ToDevice(new Point(Left + 20, Top + 20));
        bool visible = Forms.Screen.AllScreens.Any(sc => sc.WorkingArea.Contains((int)probe.X, (int)probe.Y));
        if (!visible)
        {
            _s.Dock = DockMode.TopRight;
            _s.ScreenDevice = null;
            ApplyDock();
        }
    }

    private Forms.Screen TargetScreen() =>
        Forms.Screen.AllScreens.FirstOrDefault(sc => sc.DeviceName == _s.ScreenDevice)
        ?? Forms.Screen.PrimaryScreen!;

    private Forms.Screen ScreenOfWindow()
    {
        var center = ToDevice(new Point(Left + ActualWidth / 2, Top + ActualHeight / 2));
        return Forms.Screen.FromPoint(new Drawing.Point((int)center.X, (int)center.Y));
    }

    private Rect WorkAreaInDips(Forms.Screen screen)
    {
        var wa = screen.WorkingArea;
        var m = PresentationSource.FromVisual(this)?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
        var topLeft = m.Transform(new Point(wa.Left, wa.Top));
        var bottomRight = m.Transform(new Point(wa.Right, wa.Bottom));
        return new Rect(topLeft, bottomRight);
    }

    private Point ToDevice(Point dip)
    {
        var m = PresentationSource.FromVisual(this)?.CompositionTarget?.TransformToDevice ?? Matrix.Identity;
        return m.Transform(dip);
    }

    private void OnDisplaySettingsChanged(object? sender, EventArgs e) =>
        Dispatcher.BeginInvoke(() => { if (_s.Dock != DockMode.Free) ApplyDock(); });

    private void OnPowerModeChanged(object? sender, PowerModeChangedEventArgs e)
    {
        if (e.Mode == PowerModes.Resume)
            Dispatcher.BeginInvoke(async () => await RefreshAsync());
    }

    // ───────────────────────── Menu ─────────────────────────

    private ContextMenu BuildMenu()
    {
        var menu = new ContextMenu();
        menu.Items.Add(Item("In Outlook öffnen", OpenInOutlook));
        menu.Items.Add(Item("Jetzt aktualisieren", () => { InvalidateCalendarCaches(); _ = RefreshAsync(); }));
        menu.Items.Add(Item("Kalender …", EditCalendars));
        menu.Items.Add(Item($"Todo hinzufügen … ({HotkeyHint}+N)", AddTodo));
        _rejoinItem = Item("Erneut beitreten", () => JoinMeeting(_shown));
        menu.Items.Add(_rejoinItem);
        _skipAllItem = Item("Ausgeblendete Termine wieder anzeigen", ShowSkippedMeetingsAgain);
        menu.Items.Add(_skipAllItem);
        menu.Items.Add(new Separator());

        var position = new MenuItem { Header = "Position" };
        position.Items.Add(Check("Oben links", () => _s.Dock == DockMode.TopLeft, () => SetDock(DockMode.TopLeft)));
        position.Items.Add(Check("Oben Mitte", () => _s.Dock == DockMode.TopCenter, () => SetDock(DockMode.TopCenter)));
        position.Items.Add(Check("Oben rechts", () => _s.Dock == DockMode.TopRight, () => SetDock(DockMode.TopRight)));
        position.Items.Add(Check("Frei (dort, wo verschoben)", () => _s.Dock == DockMode.Free, () => SetDock(DockMode.Free)));
        menu.Items.Add(position);

        var hide = new MenuItem { Header = "Termine ausblenden" };
        hide.Items.Add(Check("Als „frei“ markierte Termine (Blocker)", () => _s.HideFree, () => _s.HideFree = !_s.HideFree));
        _keywordItem = Check(KeywordHeader(), () => _s.HideKeywordMatches, () => _s.HideKeywordMatches = !_s.HideKeywordMatches);
        hide.Items.Add(_keywordItem);
        hide.Items.Add(Check("Ganztägige Termine", () => _s.HideAllDay, () => _s.HideAllDay = !_s.HideAllDay));
        hide.Items.Add(Check("Abgelehnte und abgesagte Termine", () => _s.HideDeclined, () => _s.HideDeclined = !_s.HideDeclined));
        hide.Items.Add(Check("Mit Vorbehalt markierte Termine", () => _s.HideTentative, () => _s.HideTentative = !_s.HideTentative));
        hide.Items.Add(Check("Private Termine", () => _s.HidePrivate, () => _s.HidePrivate = !_s.HidePrivate));
        hide.Items.Add(new Separator());
        hide.Items.Add(Item("Stichwörter bearbeiten …", EditKeywords));
        menu.Items.Add(hide);

        var johnny = new MenuItem { Header = "JohnnyAI" };
        _johnnyMenu = johnny;
        johnny.Items.Add(Check("Aufnahmeknopf anzeigen", () => _s.ShowJohnnyButton, () => _s.ShowJohnnyButton = !_s.ShowJohnnyButton));
        johnny.Items.Add(new Separator());
        johnny.Items.Add(Check("Bei Terminbeginn: nichts tun", () => _s.JohnnyStart == JohnnyStartMode.Off, () => _s.JohnnyStart = JohnnyStartMode.Off));
        johnny.Items.Add(Check("Bei Terminbeginn: Aufnahme anbieten", () => _s.JohnnyStart == JohnnyStartMode.Ask, () => _s.JohnnyStart = JohnnyStartMode.Ask));
        johnny.Items.Add(Check("Bei Terminbeginn: automatisch aufnehmen", () => _s.JohnnyStart == JohnnyStartMode.Auto, () => _s.JohnnyStart = JohnnyStartMode.Auto));
        johnny.Items.Add(Check("Nur bei Online-Meetings (Teams-/Zoom-Link)", () => _s.JohnnyOnlyOnlineMeetings, () => _s.JohnnyOnlyOnlineMeetings = !_s.JohnnyOnlyOnlineMeetings));
        menu.Items.Add(johnny);

        var view = new MenuItem { Header = "Anzeige" };
        view.Items.Add(Check("Mini-Modus (Ampel, Details bei Mausberührung oder Klick)", () => _s.MiniMode,
            () => { _s.MiniMode = !_s.MiniMode; _pinnedExpanded = _hoverExpanded = false; ApplyExpansion(); }));
        view.Items.Add(Check("Bei Terminbeginn kurz aufklappen", () => _s.AutoExpandAtStart, () => _s.AutoExpandAtStart = !_s.AutoExpandAtStart));
        view.Items.Add(new Separator());
        view.Items.Add(Check("Folgetermin und freie Zeit anzeigen", () => _s.ShowFollowUp, () => _s.ShowFollowUp = !_s.ShowFollowUp));
        view.Items.Add(Check("Überschneidende Termine anzeigen", () => _s.ShowOverlaps, () => _s.ShowOverlaps = !_s.ShowOverlaps));
        view.Items.Add(Check("Todos anzeigen", () => _s.ShowTodos, () => { _s.ShowTodos = !_s.ShowTodos; _todoSignature = null; }));
        view.Items.Add(Check("Tagesübersicht beim Darüberfahren", () => _s.ShowDayOverview, () => _s.ShowDayOverview = !_s.ShowDayOverview));
        view.Items.Add(Check("Kompakt (eine Zeile)", () => _s.Compact, () => _s.Compact = !_s.Compact));
        view.Items.Add(Check("Teams-/Zoom-Link erkennen („Beitreten“)", () => _s.DetectJoinLink, () => _s.DetectJoinLink = !_s.DetectJoinLink));
        var opacity = new MenuItem { Header = "Deckkraft" };
        foreach (var value in new[] { 1.0, 0.95, 0.85, 0.75, 0.6 })
            opacity.Items.Add(Check($"{value * 100:0} %", () => Math.Abs(_s.Opacity - value) < 0.001,
                () => { _s.Opacity = value; Opacity = value; }));
        view.Items.Add(opacity);
        menu.Items.Add(view);

        var visibility = new MenuItem { Header = "Sichtbarkeit" };
        visibility.Items.Add(Check("Bei Bildschirmfreigabe unsichtbar (Teams, Zoom, Screenshots)", () => _s.HideFromScreenShare, () => _s.HideFromScreenShare = !_s.HideFromScreenShare));
        visibility.Items.Add(Check("Bei Vollbild und Präsentationen ausblenden", () => _s.HideInFullscreen, () => _s.HideInFullscreen = !_s.HideInFullscreen));
        visibility.Items.Add(Check($"Klick-durch-Modus ({HotkeyHint}+K)", () => _s.ClickThrough, () => _s.ClickThrough = !_s.ClickThrough));
        _hotkeyItem = Check(HotkeyMenuHeader(), () => _s.HotkeysEnabled, () => _s.HotkeysEnabled = !_s.HotkeysEnabled);
        visibility.Items.Add(_hotkeyItem);
        menu.Items.Add(visibility);

        menu.Items.Add(Check("Mit Windows starten", () => Autostart.IsEnabled, () => Autostart.Set(!Autostart.IsEnabled)));
        menu.Items.Add(Item($"Overlay ausblenden ({HotkeyHint}+O)", () => SetUserHidden(true)));
        menu.Items.Add(new Separator());
        menu.Items.Add(Item("Beenden", ExitApp));

        menu.Opened += (_, _) =>
        {
            foreach (var (item, isChecked) in _checkItems) item.IsChecked = isChecked();
            _keywordItem.Header = KeywordHeader();
            _johnnyMenu.Visibility = _johnnyAvailable ? Visibility.Visible : Visibility.Collapsed;
            _rejoinItem.Visibility = _shown != null && _shown.JoinUrl != null && _joined.Contains(_shown.Key)
                ? Visibility.Visible : Visibility.Collapsed;
            if (_shown != null) _rejoinItem.Header = $"Erneut beitreten: {_shown.Subject} ({HotkeyHint}+J)";
            int skipped = _s.SkippedMeetings.Count;
            _skipAllItem.Header = skipped > 0 ? $"Ausgeblendete Termine wieder anzeigen ({skipped})" : "Ausgeblendete Termine wieder anzeigen";
            _skipAllItem.IsEnabled = skipped > 0;
            _hotkeyItem.Header = HotkeyMenuHeader();
        };
        return menu;
    }

    private MenuItem _johnnyMenu = null!;
    private MenuItem _skipAllItem = null!;
    private MenuItem _rejoinItem = null!;
    private MenuItem _hotkeyItem = null!;

    private string HotkeyMenuHeader() => $"Tastenkürzel ({HotkeyHint}+O/J{(_johnnyAvailable ? "/A" : "")}/N/K)";

    private string KeywordHeader()
    {
        var preview = string.Join(", ", _s.ExcludeKeywords.Take(3));
        return _s.ExcludeKeywords.Count > 3
            ? $"Termine mit Stichwort ({preview} …)"
            : $"Termine mit Stichwort ({preview})";
    }

    private static MenuItem Item(string header, Action action)
    {
        var item = new MenuItem { Header = header };
        item.Click += (_, _) => action();
        return item;
    }

    private MenuItem Check(string header, Func<bool> isChecked, Action toggle)
    {
        var item = new MenuItem { Header = header };
        item.Click += (_, _) =>
        {
            toggle();
            _s.Save();
            ApplyWindowFlags();
            RegisterHotkeys();
            UpdateAutoHide();
            _toolTipKey = null;
            UpdateDisplay();
            _ = RefreshAsync();
        };
        _checkItems.Add((item, isChecked));
        return item;
    }

    private void EditKeywords()
    {
        var text = ShowTextDialog("Stichwörter ausblenden",
            "Termine, deren Betreff eines dieser Stichwörter enthält, werden ausgeblendet.\nEin Stichwort pro Zeile, Groß-/Kleinschreibung egal.",
            string.Join(Environment.NewLine, _s.ExcludeKeywords), multiline: true);
        if (text == null) return;
        _s.ExcludeKeywords = text
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        _s.Save();
        _ = RefreshAsync();
    }

    /// <summary>
    /// Manage the calendar list: add/edit/remove ICS links, add the classic Outlook calendar,
    /// switch single calendars on/off. Works on a copy; applied on "Speichern".
    /// </summary>
    private void EditCalendars()
    {
        var working = _s.Calendars.Select(c => c.Copy()).ToList();

        var list = new ListBox { Height = 190, Margin = new Thickness(0, 0, 0, 10) };
        var addLink = new Button { Content = "Kalender-Link hinzufügen …", Margin = new Thickness(0, 0, 0, 6), Padding = new Thickness(8, 3, 8, 3) };
        var addOutlook = new Button { Content = "Klassisches Outlook hinzufügen", Margin = new Thickness(0, 0, 0, 6), Padding = new Thickness(8, 3, 8, 3) };
        var edit = new Button { Content = "Bearbeiten …", Margin = new Thickness(0, 0, 0, 6), Padding = new Thickness(8, 3, 8, 3) };
        var remove = new Button { Content = "Entfernen", Padding = new Thickness(8, 3, 8, 3) };
        var save = new Button { Content = "Speichern", Width = 90, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new Button { Content = "Abbrechen", Width = 90, IsCancel = true };

        void Fill()
        {
            int selected = list.SelectedIndex;
            list.Items.Clear();
            foreach (var c in working)
            {
                var check = new CheckBox { IsChecked = c.Enabled, VerticalAlignment = VerticalAlignment.Center, ToolTip = "Kalender ein-/ausschalten" };
                var source = c;
                check.Checked += (_, _) => source.Enabled = true;
                check.Unchecked += (_, _) => source.Enabled = false;
                var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(2) };
                row.Children.Add(check);
                row.Children.Add(new TextBlock { Text = c.Name, FontWeight = FontWeights.SemiBold, Margin = new Thickness(6, 0, 0, 0) });
                string state = c.Kind == CalendarKind.Ics && string.IsNullOrEmpty(c.Url) ? " – Link fehlt" : "";
                row.Children.Add(new TextBlock { Text = $"   {c.KindLabel}{state}", Foreground = Brushes.Gray });
                list.Items.Add(row);
            }
            list.SelectedIndex = Math.Min(Math.Max(selected, working.Count > 0 ? 0 : -1), working.Count - 1);
            addOutlook.IsEnabled = working.All(c => c.Kind != CalendarKind.Outlook);
            edit.IsEnabled = remove.IsEnabled = working.Count > 0;
        }

        Window? dialog = null;
        addLink.Click += (_, _) =>
        {
            var added = EditIcsCalendar(dialog!, new CalendarSource { Name = $"Kalender {working.Count + 1}" });
            if (added == null) return;
            working.Add(added);
            Fill();
            list.SelectedIndex = working.Count - 1;
        };
        addOutlook.Click += (_, _) =>
        {
            working.Add(new CalendarSource { Name = "Outlook", Kind = CalendarKind.Outlook });
            Fill();
            list.SelectedIndex = working.Count - 1;
        };
        edit.Click += (_, _) =>
        {
            int i = list.SelectedIndex;
            if (i < 0) return;
            if (working[i].Kind == CalendarKind.Ics)
            {
                var changed = EditIcsCalendar(dialog!, working[i].Copy());
                if (changed != null) working[i] = changed;
            }
            else
            {
                var name = ShowTextDialog("Kalender umbenennen", "Name für das klassische Outlook:", working[i].Name, multiline: false);
                if (!string.IsNullOrWhiteSpace(name)) working[i].Name = name.Trim();
            }
            Fill();
        };
        list.MouseDoubleClick += (_, _) => edit.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        remove.Click += (_, _) =>
        {
            int i = list.SelectedIndex;
            if (i < 0) return;
            working.RemoveAt(i);
            Fill();
        };

        var buttons = new StackPanel { Margin = new Thickness(10, 0, 0, 0), Width = 210 };
        foreach (var b in new[] { addLink, addOutlook, edit, remove }) buttons.Children.Add(b);
        var body = new DockPanel();
        DockPanel.SetDock(buttons, Dock.Right);
        body.Children.Add(buttons);
        body.Children.Add(list);

        var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        footer.Children.Add(save);
        footer.Children.Add(cancel);

        var panel = new StackPanel { Margin = new Thickness(14) };
        panel.Children.Add(new TextBlock
        {
            Text = "Die Termine aller eingeschalteten Kalender werden zusammen angezeigt. " +
                   "Steht ein Termin in mehreren Kalendern, erscheint er nur einmal.",
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 10),
        });
        panel.Children.Add(body);
        panel.Children.Add(footer);

        dialog = new Window
        {
            Title = "Kalender",
            Content = panel,
            Width = 620,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize,
            Topmost = true,
        };
        save.Click += (_, _) => dialog.DialogResult = true;
        Fill();
        if (dialog.ShowDialog() != true) return;

        _s.Calendars = working;
        _s.Save();
        _icsSources.Clear();
        _meetings = new();
        _status = "Lade Termine …";
        UpdateDisplay();
        _ = RefreshAsync();
    }

    /// <summary>Name + ICS link of one calendar. Returns null on cancel or when no link was entered.</summary>
    private static CalendarSource? EditIcsCalendar(Window owner, CalendarSource source)
    {
        var name = new TextBox { Text = source.Name, Margin = new Thickness(0, 2, 0, 10) };
        var url = new TextBox { Text = source.Url ?? "", Height = 60, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 2, 0, 10) };
        var save = new Button { Content = "Übernehmen", Width = 100, IsDefault = true, Margin = new Thickness(0, 0, 8, 0) };
        var cancel = new Button { Content = "Abbrechen", Width = 90, IsCancel = true };

        var panel = new StackPanel { Margin = new Thickness(14) };
        panel.Children.Add(new TextBlock { Text = "Name (wird beim Termin angezeigt, z. B. „Arbeit“ oder „Privat“):" });
        panel.Children.Add(name);
        panel.Children.Add(new TextBlock { Text = "ICS-Link des Kalenders:" });
        panel.Children.Add(url);
        panel.Children.Add(new TextBlock
        {
            Text = "Outlook: Einstellungen → Kalender → Freigegebene Kalender → „Kalender veröffentlichen“ → Kalender wählen, " +
                   "„Kann alle Details anzeigen“ → Veröffentlichen → ICS-Link kopieren.\n" +
                   "Google Kalender: Einstellungen → Kalender → „Geheime Adresse im iCal-Format“.\n" +
                   "iCloud: Kalender teilen → „Öffentlicher Kalender“ → Link kopieren.\n\n" +
                   "Wer einen solchen Link kennt, kann den Kalender lesen. Er wird hier nur verschlüsselt für deinen Windows-Benutzer gespeichert.",
            TextWrapping = TextWrapping.Wrap,
            Foreground = Brushes.DimGray,
            Margin = new Thickness(0, 0, 0, 10),
        });
        var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        footer.Children.Add(save);
        footer.Children.Add(cancel);
        panel.Children.Add(footer);

        var dialog = new Window
        {
            Title = "Kalender-Link",
            Content = panel,
            Width = 520,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = owner,
            ResizeMode = ResizeMode.NoResize,
        };
        save.Click += (_, _) => dialog.DialogResult = true;
        dialog.Loaded += (_, _) => { url.Focus(); url.SelectAll(); };
        if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(url.Text)) return null;

        source.Name = string.IsNullOrWhiteSpace(name.Text) ? "Kalender" : name.Text.Trim();
        source.Kind = CalendarKind.Ics;
        source.Url = url.Text;
        return source;
    }

    private static string? ShowTextDialog(string title, string intro, string initial, bool multiline)
    {
        var box = new TextBox
        {
            Text = initial,
            AcceptsReturn = multiline,
            TextWrapping = multiline ? TextWrapping.NoWrap : TextWrapping.Wrap,
            Height = multiline ? 180 : 60,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
        };
        var save = new Button { Content = "Speichern", Width = 90, Margin = new Thickness(0, 10, 0, 0), HorizontalAlignment = HorizontalAlignment.Right, IsDefault = !multiline };
        var panel = new StackPanel { Margin = new Thickness(14) };
        panel.Children.Add(new TextBlock { Text = intro, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) });
        panel.Children.Add(box);
        panel.Children.Add(save);

        var dialog = new Window
        {
            Title = title,
            Content = panel,
            Width = multiline ? 380 : 520,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize,
            Topmost = true,
        };
        save.Click += (_, _) => dialog.DialogResult = true;
        dialog.Loaded += (_, _) => { box.Focus(); box.SelectAll(); };
        return dialog.ShowDialog() == true ? box.Text : null;
    }

    // ───────────────────────── Tray ─────────────────────────

    private void BuildTray()
    {
        _tray = new Forms.NotifyIcon { Icon = CreateTrayIcon(), Visible = true, Text = "Nächster Termin" };
        _tray.MouseUp += (_, e) =>
        {
            if (e.Button == Forms.MouseButtons.Left)
            {
                SetUserHidden(IsVisible);
            }
            else if (e.Button == Forms.MouseButtons.Right)
            {
                SetUserHidden(false);
                if (!IsVisible) Show();   // even while auto-hidden, the menu needs a visible owner
                Activate();
                _menu.PlacementTarget = this;
                _menu.Placement = PlacementMode.MousePoint;
                _menu.IsOpen = true;
            }
        };
        _tray.BalloonTipClicked += async (_, _) =>
        {
            var action = _balloonAction;
            _balloonAction = null;
            if (action != null) await action();
        };
        _tray.BalloonTipClosed += (_, _) => _balloonAction = null;
    }

    /// <summary>Balloon whose click runs an action (start recording, undo hiding a meeting).</summary>
    private void ShowActionBalloon(string title, string text, Func<Task> onClick, int timeoutMs = 15000)
    {
        _balloonAction = onClick;
        _tray?.ShowBalloonTip(timeoutMs, title, text, Forms.ToolTipIcon.Info);
    }

    private void SetTrayText(string text)
    {
        if (_tray == null) return;
        _tray.Text = text.Length > 63 ? text[..62] + "…" : text;
    }

    private static Drawing.Icon CreateTrayIcon()
    {
        // Same icon as the exe (app.ico); the drawn fallback only kicks in if extraction fails.
        try
        {
            var exeIcon = Environment.ProcessPath != null ? Drawing.Icon.ExtractAssociatedIcon(Environment.ProcessPath) : null;
            if (exeIcon != null) return new Drawing.Icon(exeIcon, Forms.SystemInformation.SmallIconSize);
        }
        catch { /* fall back to drawing */ }

        using var bmp = new Drawing.Bitmap(32, 32);
        using (var g = Drawing.Graphics.FromImage(bmp))
        {
            g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Drawing.Color.Transparent);
            using var fill = new Drawing.SolidBrush(Drawing.Color.FromArgb(76, 194, 255));
            g.FillEllipse(fill, 1, 1, 30, 30);
            using var pen = new Drawing.Pen(Drawing.Color.White, 3) { StartCap = Drawing.Drawing2D.LineCap.Round, EndCap = Drawing.Drawing2D.LineCap.Round };
            g.DrawLine(pen, 16, 16, 16, 7);
            g.DrawLine(pen, 16, 16, 23, 16);
        }
        return Drawing.Icon.FromHandle(bmp.GetHicon());
    }

    // ───────────────────────── Demo screenshot ─────────────────────────

    /// <summary>Renders mini pill, expanded card and context menu side by side into one PNG, then exits.</summary>
    private async Task ExportDemoScreenshotAsync()
    {
        const double dpi = 192;   // 2x for crisp images on GitHub
        var root = (FrameworkElement)Content;
        _s.AutoExpandAtStart = false;

        async Task Settle()
        {
            UpdateDisplay();
            await Task.Delay(300);
            root.UpdateLayout();
        }

        _pinnedExpanded = false;
        ApplyExpansion();
        await Settle();
        var mini = Snapshot(root, dpi);

        _pinnedExpanded = true;
        ApplyExpansion();
        await Settle();
        var card = Snapshot(root, dpi);

        _menu.PlacementTarget = this;
        _menu.Placement = PlacementMode.Bottom;
        _menu.IsOpen = true;
        await Task.Delay(500);
        _menu.UpdateLayout();
        var menu = Snapshot(_menu, dpi);
        _menu.IsOpen = false;

        SaveComposite(Demo.ExportPath!, dpi,
            (mini, "Klein: Ampel-Pille"),
            (card, "Aufgeklappt: Details"),
            (menu, "Rechtsklick: Einstellungen"));
        ExitApp();
    }

    private static System.Windows.Media.Imaging.BitmapSource Snapshot(FrameworkElement element, double dpi)
    {
        double w = element.ActualWidth, h = element.ActualHeight;
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
            // 1:1 crop of the element's own area - without an absolute viewbox the brush would
            // stretch the effect bounds (drop shadow) into the rectangle and squash the content.
            dc.DrawRectangle(new VisualBrush(element)
            {
                Stretch = Stretch.None,
                ViewboxUnits = BrushMappingMode.Absolute,
                Viewbox = new Rect(0, 0, w, h),
                AlignmentX = AlignmentX.Left,
                AlignmentY = AlignmentY.Top,
            }, null, new Rect(0, 0, w, h));
        var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(
            (int)Math.Ceiling(w * dpi / 96), (int)Math.Ceiling(h * dpi / 96), dpi, dpi, PixelFormats.Pbgra32);
        bitmap.Render(visual);
        bitmap.Freeze();
        return bitmap;
    }

    private static void SaveComposite(string path, double dpi,
        params (System.Windows.Media.Imaging.BitmapSource Image, string Caption)[] parts)
    {
        const double pad = 36, gap = 48, captionHeight = 30;
        var sizes = parts.Select(p => new Size(p.Image.PixelWidth * 96 / dpi, p.Image.PixelHeight * 96 / dpi)).ToList();
        double width = pad * 2 + sizes.Sum(s => s.Width) + gap * (parts.Length - 1);
        double height = pad * 2 + captionHeight + sizes.Max(s => s.Height);

        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
        {
            // Calm desktop-like backdrop so the dark overlay stands out.
            var background = new LinearGradientBrush(
                (Color)ColorConverter.ConvertFromString("#5E7194"),
                (Color)ColorConverter.ConvertFromString("#2D3A56"), 45);
            dc.DrawRectangle(background, null, new Rect(0, 0, width, height));

            var typeface = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);
            double x = pad;
            for (int i = 0; i < parts.Length; i++)
            {
                var caption = new FormattedText(parts[i].Caption, System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight, typeface, 14, Brushes.White, dpi / 96);
                dc.DrawText(caption, new Point(x + 6, pad));
                dc.DrawImage(parts[i].Image, new Rect(new Point(x, pad + captionHeight), sizes[i]));
                x += sizes[i].Width + gap;
            }
        }

        var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(
            (int)Math.Ceiling(width * dpi / 96), (int)Math.Ceiling(height * dpi / 96), dpi, dpi, PixelFormats.Pbgra32);
        bitmap.Render(visual);
        var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
        encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmap));
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
        using var file = System.IO.File.Create(path);
        encoder.Save(file);
    }

    // ───────────────────────── Shutdown ─────────────────────────

    private bool _exiting;

    private void ExitApp()
    {
        if (_exiting) return;
        _exiting = true;
        SystemEvents.DisplaySettingsChanged -= OnDisplaySettingsChanged;
        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        if (_hwnd != IntPtr.Zero) UnregisterHotkeys();
        _s.Save();
        if (_tray != null)
        {
            _tray.Visible = false;
            _tray.Dispose();
        }
        _reader.Dispose();
        Application.Current.Shutdown();
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        ExitApp();
    }

    private static Brush Frozen(string hex)
    {
        var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(hex)!;
        brush.Freeze();
        return brush;
    }
}

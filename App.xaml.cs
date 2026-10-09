using System;
using System.Linq;
using System.Threading;
using System.Windows;

namespace NextMeetingOverlay;

public partial class App : Application
{
    private static Mutex? _singleInstance;

    protected override void OnStartup(StartupEventArgs e)
    {
        int profileArg = Array.IndexOf(e.Args, "--profile");
        if (profileArg >= 0 && profileArg + 1 < e.Args.Length)
            Settings.Profile = string.Concat(e.Args[profileArg + 1].Where(char.IsLetterOrDigit));

        _singleInstance = new Mutex(true, @"Local\NextMeetingOverlay_SingleInstance" + Settings.Profile, out var isFirst);
        if (!isFirst)
        {
            Shutdown();
            return;
        }

        if (Settings.Profile == null) Autostart.Repair();

        base.OnStartup(e);
        new MainWindow().Show();
    }
}

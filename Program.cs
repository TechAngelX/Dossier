// Program.cs

using Avalonia;
using Avalonia.ReactiveUI;
using System;
using System.Diagnostics;
using System.IO;

namespace Dossier;

class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        RelaunchWithoutHomebrewDyldIfNeeded(args);
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    // Some macOS dev machines export DYLD_LIBRARY_PATH=/usr/local/lib in ~/.zshrc,
    // which forces this process to load Homebrew's libpng/libjpeg/libtiff/etc. Apple's
    // ImageIO then binds the wrong libpng and SIGSEGVs (0x…bad4007 in PNGReadPlugin)
    // while decoding ANY png — emoji glyphs, hover cursors — crashing the app at random.
    // A dylib can't be unloaded once mapped, so if that variable is present we relaunch
    // ourselves once with a clean environment, before Avalonia renders anything.
    // No-op on Windows/Linux, when the variable is absent, or after we've already relaunched.
    private static void RelaunchWithoutHomebrewDyldIfNeeded(string[] args)
    {
        try
        {
            if (!OperatingSystem.IsMacOS()) return;
            if (Environment.GetEnvironmentVariable("DOSSIER_DYLD_SANITISED") == "1") return;

            var dyld = Environment.GetEnvironmentVariable("DYLD_LIBRARY_PATH");
            if (string.IsNullOrEmpty(dyld) || !dyld.Contains("/usr/local")) return;

            var exe = Environment.ProcessPath;
            // Only relaunch the real app host, not a "dotnet Dossier.dll" style launch.
            if (string.IsNullOrEmpty(exe) ||
                !Path.GetFileNameWithoutExtension(exe).Equals("Dossier", StringComparison.OrdinalIgnoreCase))
                return;

            var psi = new ProcessStartInfo { FileName = exe, UseShellExecute = false };
            foreach (var a in args) psi.ArgumentList.Add(a);
            psi.Environment.Remove("DYLD_LIBRARY_PATH");
            psi.Environment["DOSSIER_DYLD_SANITISED"] = "1";

            var proc = Process.Start(psi);
            if (proc == null) return;      // relaunch failed — fall through and run in-process
            proc.WaitForExit();
            Environment.Exit(proc.ExitCode);
        }
        catch
        {
            // If anything goes wrong, don't block startup — just run normally.
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace()
            .UseReactiveUI();
}

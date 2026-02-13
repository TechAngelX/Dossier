// Services/SleepInhibitor.cs

using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Dossier.Services;

/// <summary>
/// Prevents the OS from sleeping while batch processing is running.
/// macOS: spawns caffeinate. Windows: SetThreadExecutionState. Linux: systemd-inhibit.
/// </summary>
public sealed class SleepInhibitor : IDisposable
{
    private Process? _caffeinateProcess;
    private bool _disposed;

    // Windows P/Invoke
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SetThreadExecutionState(uint esFlags);

    private const uint ES_CONTINUOUS = 0x80000000;
    private const uint ES_SYSTEM_REQUIRED = 0x00000001;
    private const uint ES_DISPLAY_REQUIRED = 0x00000002;

    public void Activate()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            ActivateMac();
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            ActivateWindows();
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            ActivateLinux();
        }
    }

    public void Release()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            ReleaseMac();
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            ReleaseWindows();
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            ReleaseMac(); // same pattern — kill the child process
        }
    }

    private void ActivateMac()
    {
        try
        {
            // caffeinate -di: prevent idle sleep (-i) and display sleep (-d)
            _caffeinateProcess = Process.Start(new ProcessStartInfo
            {
                FileName = "caffeinate",
                Arguments = "-di",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
            Console.WriteLine($"[SleepInhibitor] caffeinate started (PID {_caffeinateProcess?.Id})");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[SleepInhibitor] Failed to start caffeinate: {ex.Message}");
        }
    }

    private void ReleaseMac()
    {
        if (_caffeinateProcess != null && !_caffeinateProcess.HasExited)
        {
            _caffeinateProcess.Kill();
            _caffeinateProcess.Dispose();
            _caffeinateProcess = null;
            Console.WriteLine("[SleepInhibitor] caffeinate stopped");
        }
    }

    private void ActivateWindows()
    {
        SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED);
        Console.WriteLine("[SleepInhibitor] Windows sleep prevention active");
    }

    private void ReleaseWindows()
    {
        SetThreadExecutionState(ES_CONTINUOUS);
        Console.WriteLine("[SleepInhibitor] Windows sleep prevention released");
    }

    private void ActivateLinux()
    {
        try
        {
            // systemd-inhibit keeps the system awake while the child process runs
            // We spawn a long-sleep process under it as a keep-alive
            _caffeinateProcess = Process.Start(new ProcessStartInfo
            {
                FileName = "systemd-inhibit",
                Arguments = "--what=idle --why=\"Dossier batch processing\" sleep infinity",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
            Console.WriteLine($"[SleepInhibitor] systemd-inhibit started (PID {_caffeinateProcess?.Id})");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[SleepInhibitor] Failed to start systemd-inhibit: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            Release();
            _disposed = true;
        }
    }
}

// Views/IsoView.axaml.cs

using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Threading;
using Dossier.Models;
using Dossier.Services;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Dossier.Views;

public partial class IsoView : UserControl
{
    private bool _isRunning;

    private TextBox _studentNumberBox = null!;
    private CheckBox _headlessCheckbox = null!;
    private Button _downloadButton = null!;
    private Button _resetButton = null!;
    private Button _clearLogButton = null!;
    private TextBox _statusLog = null!;
    private TextBlock _statusLabel = null!;
    private ProgressBar _busyBar = null!;

    public IsoView()
    {
        InitializeComponent();

        _downloadButton.Click += DownloadButton_Click;
        _resetButton.Click += ResetButton_Click;
        _clearLogButton.Click += (_, _) => _statusLog.Text = string.Empty;
        _studentNumberBox.TextChanged += (_, _) => UpdateReadyState();

        UpdateReadyState();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);

        _studentNumberBox = this.FindControl<TextBox>("StudentNumberBox")!;
        _headlessCheckbox = this.FindControl<CheckBox>("HeadlessCheckbox")!;
        _downloadButton   = this.FindControl<Button>("DownloadButton")!;
        _resetButton      = this.FindControl<Button>("ResetButton")!;
        _clearLogButton   = this.FindControl<Button>("ClearLogButton")!;
        _statusLog        = this.FindControl<TextBox>("StatusLog")!;
        _statusLabel      = this.FindControl<TextBlock>("StatusLabel")!;
        _busyBar          = this.FindControl<ProgressBar>("BusyBar")!;
    }

    private string StudentNumber => (_studentNumberBox.Text ?? string.Empty).Trim();

    private void UpdateReadyState()
    {
        if (_isRunning) return;
        bool ready = StudentNumber.Length > 0;
        _downloadButton.IsEnabled = ready;
        _statusLabel.Text = ready ? "Ready" : "Enter a student number to begin";
    }

    private async void DownloadButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_isRunning) return;

        var studentNo = StudentNumber;
        if (string.IsNullOrEmpty(studentNo))
        {
            await ShowMessageBoxAsync("Missing input", "Please enter a student number first.");
            return;
        }

        _isRunning = true;
        _downloadButton.IsEnabled = false;
        _busyBar.IsIndeterminate = true;
        _statusLabel.Text = "Working...";

        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        PorticoAutomationService? service = null;
        EventHandler<string>? statusHandler = null;
        string? savedPath = null;

        try
        {
            var config = new AppConfig { HeadlessMode = _headlessCheckbox.IsChecked == true };

            service = PorticoAutomationService.Shared;
            statusHandler = (_, msg) => LogStatus(msg);
            service.StatusUpdated += statusHandler;

            LogStatus("Launching browser...");
            await service.InitialiseAsync(config);

            LogStatus("Checking Portico session...");
            await service.LoginAsync();

            savedPath = await service.DownloadIndividualStudentOverviewCsvAsync(studentNo, desktop);
        }
        catch (Exception ex)
        {
            LogStatus($"ERROR: {ex.Message}");
        }
        finally
        {
            if (service != null && statusHandler != null)
            {
                service.StatusUpdated -= statusHandler;
            }

            _isRunning = false;
            _busyBar.IsIndeterminate = false;
            UpdateReadyState();
        }

        if (savedPath != null)
        {
            _statusLabel.Text = "Done";
            OpenContainingFolder(savedPath);
            await ShowMessageBoxAsync("Download complete",
                $"Saved to Desktop:\n\n{Path.GetFileName(savedPath)}");
        }
        else
        {
            _statusLabel.Text = "Failed — see log";
        }
    }

    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_isRunning) return;
        _studentNumberBox.Text = string.Empty;
        _statusLog.Text = string.Empty;
        UpdateReadyState();
    }

    private void LogStatus(string message) =>
        Dispatcher.UIThread.Post(() => _statusLog.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n");

    private static void OpenContainingFolder(string filePath)
    {
        try
        {
            var folder = Path.GetDirectoryName(filePath);
            if (string.IsNullOrEmpty(folder)) return;

            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                Process.Start(new ProcessStartInfo { FileName = "open", Arguments = $"\"{folder}\"", UseShellExecute = false });
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = $"\"{folder}\"", UseShellExecute = true });
        }
        catch { /* non-fatal */ }
    }

    private async Task ShowMessageBoxAsync(string title, string message)
    {
        var parentWindow = TopLevel.GetTopLevel(this) as Window;

        var okButton = new Button
        {
            Content = "OK",
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
            Background = SolidColorBrush.Parse("#2563EB"),
            Foreground = Brushes.White,
            Padding = new Avalonia.Thickness(20, 10)
        };

        var win = new Window
        {
            Title = title,
            Width = 460,
            Height = 240,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SystemDecorations = SystemDecorations.BorderOnly,
            ExtendClientAreaToDecorationsHint = true,
            Content = new Border
            {
                BorderBrush = Brushes.Gray,
                BorderThickness = new Avalonia.Thickness(1),
                Padding = new Avalonia.Thickness(30),
                Child = new StackPanel
                {
                    Spacing = 20,
                    Children =
                    {
                        new TextBlock { Text = title, FontWeight = FontWeight.Bold, FontSize = 18 },
                        new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap, FontSize = 13 },
                        okButton
                    }
                }
            }
        };

        okButton.Click += (_, _) => win.Close();

        if (parentWindow != null)
            await win.ShowDialog(parentWindow);
        else
            win.Show();
    }
}

// Views/AdMergerView.axaml.cs

using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Dossier.Configuration;
using Dossier.Models;
using Dossier.Services;
using Dossier.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Dossier.Views;

public class ProcessingItem : INotifyPropertyChanged
{
    private string _status = "";
    private string _ukGrade = "";
    private string _studentNo = "";
    private string _name = "";

    public string StudentNo { get => _studentNo; set { _studentNo = value; OnPropertyChanged(nameof(StudentNo)); } }
    public string Name { get => _name; set { _name = value; OnPropertyChanged(nameof(Name)); } }
    public string ReceivedDate { get; set; } = "";
    public string Status { get => _status; set { _status = value; OnPropertyChanged(nameof(Status)); OnPropertyChanged(nameof(StatusColor)); OnPropertyChanged(nameof(StatusForeColor)); } }
    public string UkGrade { get => _ukGrade; set { _ukGrade = value; OnPropertyChanged(nameof(UkGrade)); OnPropertyChanged(nameof(GradeColor)); OnPropertyChanged(nameof(GradeForeColor)); } }

    public IBrush StatusColor
    {
        get
        {
            if (Status.Contains("✓")) return SolidColorBrush.Parse("#DCFCE7");
            if (Status.Contains("⚠️")) return SolidColorBrush.Parse("#FEF3C7");
            if (Status.Contains("⚙")) return SolidColorBrush.Parse("#DBEAFE");
            return SolidColorBrush.Parse("#F1F5F9");
        }
    }
    public IBrush StatusForeColor
    {
        get
        {
            if (Status.Contains("✓")) return SolidColorBrush.Parse("#166534");
            if (Status.Contains("⚠️")) return SolidColorBrush.Parse("#92400E");
            if (Status.Contains("⚙")) return SolidColorBrush.Parse("#1E40AF");
            return SolidColorBrush.Parse("#475569");
        }
    }
    public IBrush GradeColor
    {
        get
        {
            if (string.IsNullOrEmpty(UkGrade) || UkGrade == "??") return Brushes.Transparent;
            if (UkGrade == "1.0") return SolidColorBrush.Parse("#F0FDF4");
            if (UkGrade == "2.1") return SolidColorBrush.Parse("#EFF6FF");
            if (UkGrade == "2.2") return SolidColorBrush.Parse("#FFFBEB");
            return SolidColorBrush.Parse("#FEF2F2");
        }
    }
    public IBrush GradeForeColor
    {
        get
        {
            if (string.IsNullOrEmpty(UkGrade) || UkGrade == "??") return SolidColorBrush.Parse("#94A3B8");
            if (UkGrade == "1.0") return SolidColorBrush.Parse("#15803D");
            if (UkGrade == "2.1") return SolidColorBrush.Parse("#1D4ED8");
            if (UkGrade == "2.2") return SolidColorBrush.Parse("#B45309");
            return SolidColorBrush.Parse("#B91C1C");
        }
    }
    public FontWeight GradeWeight => (string.IsNullOrEmpty(UkGrade) || UkGrade == "??") ? FontWeight.Normal : FontWeight.Bold;

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public partial class AdMergerView : UserControl
{
    private readonly ICsvService _csvService;
    private readonly IEquivalencyService _equivalencyService;
    private readonly IInstitutionMatchingService _matchingService;
    private readonly IRankingService _rankingService;
    private readonly IGradeClassificationService _gradeService;

    private string _inTrayFilePath = string.Empty;
    private string _resolvedProgrammeName = string.Empty;
    private Dictionary<string, string> _shortToFullProg = new(StringComparer.OrdinalIgnoreCase);
    private readonly string _outputFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isProcessing = false;
    private bool _rankingsLoaded = false;
    public ObservableCollection<ProcessingItem> ProcessingItems { get; } = new ObservableCollection<ProcessingItem>();

    // Control references
    private TextBlock _versionLabel = null!;
    private Border _inTrayCard = null!;
    private Border _programmeCard = null!;
    private Button _browseInTrayButton = null!;
    private TextBlock _inTrayFileLabel = null!;
    private TextBox _programmeCodeBox = null!;
    private TextBlock _programmeLookupLabel = null!;
    private CheckBox _includeGenderCheckbox = null!;
    private CheckBox _includeNationalityCheckbox = null!;
    private CheckBox _includeDOBCheckbox = null!;
    private CheckBox _includeEmailCheckbox = null!;
    private CheckBox _includePaidCheckbox = null!;
    private ListBox _processingList = null!;
    private TextBlock _statusLabel = null!;
    private TextBlock _footerStatus = null!;
    private ProgressBar _mainProgressBar = null!;
    private Button _processButton = null!;
    private Button _resetButton = null!;
    private Button _exitButton = null!;
    private TextBox _statusLog = null!;
    private Button _clearLogButton = null!;

    [DllImport("winmm.dll")]
    private static extern long mciSendString(string strCommand, StringBuilder? strReturn, int iReturnLength, IntPtr hwndCallback);

    public AdMergerView()
    {
        InitializeComponent();

        _csvService = new CsvService();
        _equivalencyService = new EquivalencyService();
        _matchingService = new InstitutionMatchingService();
        _rankingService = new RankingService(_matchingService);
        _gradeService = new GradeClassificationService(_equivalencyService);

        _processingList.ItemsSource = ProcessingItems;
        SetVersion();
        LoadProgCodes();
        InitializeDataAsync();

        _browseInTrayButton.Click += BrowseInTrayButton_Click;
        _programmeCodeBox.TextChanged += ProgrammeCodeBox_TextChanged;
        _processButton.Click += ProcessButton_Click;
        _clearLogButton.Click += ClearLogButton_Click;
        _resetButton.Click += ResetButton_Click;
        _exitButton.Click += ExitButton_Click;

        _inTrayCard.AddHandler(DragDrop.DropEvent, OnInTrayDrop);
        _inTrayCard.AddHandler(DragDrop.DragOverEvent, OnDragOver);
        _inTrayCard.AddHandler(DragDrop.DragEnterEvent, (_, _) => SetCardDragState(_inTrayCard, true));
        _inTrayCard.AddHandler(DragDrop.DragLeaveEvent, (_, _) => SetCardDragState(_inTrayCard, false));
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);

        _versionLabel          = this.FindControl<TextBlock>("VersionLabel")!;
        _inTrayCard            = this.FindControl<Border>("InTrayCard")!;
        _programmeCard         = this.FindControl<Border>("ProgrammeCard")!;
        _browseInTrayButton    = this.FindControl<Button>("BrowseInTrayButton")!;
        _inTrayFileLabel       = this.FindControl<TextBlock>("InTrayFileLabel")!;
        _programmeCodeBox      = this.FindControl<TextBox>("ProgrammeCodeBox")!;
        _programmeLookupLabel  = this.FindControl<TextBlock>("ProgrammeLookupLabel")!;
        _includeGenderCheckbox = this.FindControl<CheckBox>("IncludeGenderCheckbox")!;
        _includeNationalityCheckbox = this.FindControl<CheckBox>("IncludeNationalityCheckbox")!;
        _includeDOBCheckbox    = this.FindControl<CheckBox>("IncludeDOBCheckbox")!;
        _includeEmailCheckbox  = this.FindControl<CheckBox>("IncludeEmailCheckbox")!;
        _includePaidCheckbox   = this.FindControl<CheckBox>("IncludePaidCheckbox")!;
        _processingList        = this.FindControl<ListBox>("ProcessingList")!;
        _statusLabel           = this.FindControl<TextBlock>("StatusLabel")!;
        _footerStatus          = this.FindControl<TextBlock>("FooterStatus")!;
        _mainProgressBar       = this.FindControl<ProgressBar>("MainProgressBar")!;
        _processButton         = this.FindControl<Button>("ProcessButton")!;
        _resetButton           = this.FindControl<Button>("ResetButton")!;
        _exitButton            = this.FindControl<Button>("ExitButton")!;
        _statusLog             = this.FindControl<TextBox>("StatusLog")!;
        _clearLogButton        = this.FindControl<Button>("ClearLogButton")!;
    }

    private OutputSettings GetOutputSettings() => new OutputSettings
    {
        IncludeGender      = _includeGenderCheckbox.IsChecked == true,
        IncludeNationality = _includeNationalityCheckbox.IsChecked == true,
        IncludeDateOfBirth = _includeDOBCheckbox.IsChecked == true,
        IncludeEmail       = _includeEmailCheckbox.IsChecked == true,
        IncludePaid        = _includePaidCheckbox.IsChecked == true
    };

    private void LoadProgCodes()
    {
        _shortToFullProg = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Look for progcodes.csv: first in output dir, then relative to project root (dev)
        var searchPaths = new[]
        {
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "progcodes.csv"),
            Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "data", "progcodes.csv")),
        };

        foreach (var path in searchPaths)
        {
            if (!File.Exists(path)) continue;

            foreach (var line in File.ReadAllLines(path).Skip(1))
            {
                var parts = SplitCsvRow(line);
                if (parts.Length >= 4 && !string.IsNullOrWhiteSpace(parts[2]) && !string.IsNullOrWhiteSpace(parts[3]))
                    _shortToFullProg[parts[2].Trim()] = parts[3].Trim();
            }
            return;
        }

        // Fallback: invert ProgrammeMapping
        foreach (var kvp in ProgrammeMapping.Mappings)
            _shortToFullProg[kvp.Value] = kvp.Key;
    }

    private static string[] SplitCsvRow(string line)
    {
        var fields = new List<string>();
        var current = new StringBuilder();
        bool inQuotes = false;
        foreach (char c in line)
        {
            if (c == '"') { inQuotes = !inQuotes; }
            else if (c == ',' && !inQuotes) { fields.Add(current.ToString()); current.Clear(); }
            else { current.Append(c); }
        }
        fields.Add(current.ToString());
        return fields.ToArray();
    }

    private async void InitializeDataAsync()
    {
        _statusLabel.Text = "Loading baked data...";
        try
        {
            _equivalencyService.LoadEquivalencies();
            await _rankingService.LoadRankingsAsync();
            _rankingsLoaded = true;
            Dispatcher.UIThread.Post(() => {
                _statusLabel.Text = "Ready";
                CheckReadyToProcess();
            });
        }
        catch (Exception ex) { LogStatus($"Load error: {ex.Message}"); }
    }

    private void CheckReadyToProcess()
    {
        bool ready = !string.IsNullOrEmpty(_inTrayFilePath) && !string.IsNullOrEmpty(_resolvedProgrammeName);
        _processButton.IsEnabled = ready && _rankingsLoaded;
        _statusLabel.Text = !_rankingsLoaded ? "Waiting for rankings..." : ready ? "Ready to start" : "Waiting for files...";
    }

    private void ProgrammeCodeBox_TextChanged(object? sender, TextChangedEventArgs e)
    {
        var input = _programmeCodeBox.Text?.Trim() ?? "";

        if (string.IsNullOrEmpty(input))
        {
            _resolvedProgrammeName = string.Empty;
            _programmeLookupLabel.Text = "Enter a programme code above";
            _programmeLookupLabel.Foreground = SolidColorBrush.Parse("#94A3B8");
            CheckReadyToProcess();
            return;
        }

        if (_shortToFullProg.TryGetValue(input, out var fullName))
        {
            _resolvedProgrammeName = fullName;
            _programmeLookupLabel.Text = $"✓ {fullName}";
            _programmeLookupLabel.Foreground = SolidColorBrush.Parse("#15803D");
        }
        else
        {
            _resolvedProgrammeName = string.Empty;
            _programmeLookupLabel.Text = $"Unknown code \"{input}\"";
            _programmeLookupLabel.Foreground = SolidColorBrush.Parse("#B91C1C");
        }

        CheckReadyToProcess();
    }

    private async Task ShowMessageBoxAsync(string title, string message)
    {
        var parentWindow = TopLevel.GetTopLevel(this) as Window;

        var okButton = new Button {
            Content = "OK",
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
            Background = SolidColorBrush.Parse("#2563EB"),
            Foreground = Brushes.White,
            Padding = new Thickness(20, 10)
        };

        var win = new Window {
            Title = title, Width = 500, Height = 350,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SystemDecorations = SystemDecorations.BorderOnly,
            ExtendClientAreaToDecorationsHint = true,
            Content = new Border {
                BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1), Padding = new Thickness(30),
                Child = new StackPanel {
                    Spacing = 20,
                    Children = {
                        new TextBlock { Text = title, FontWeight = FontWeight.Bold, FontSize = 20 },
                        new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap, FontSize = 14 },
                        okButton
                    }
                }
            }
        };

        okButton.Click += (s, e) => {
            win.Close();
        };

        if (parentWindow != null)
            await win.ShowDialog(parentWindow);
        else
            win.Show();
    }

    private async void ProcessButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_isProcessing)
        {
            _cancellationTokenSource?.Cancel();
            return;
        }

        _isProcessing = true;
        _cancellationTokenSource = new CancellationTokenSource();
        var token = _cancellationTokenSource.Token;
        _processButton.Content = "🛑 STOP";
        _processButton.Background = SolidColorBrush.Parse("#DC2626");

        var outputSettings = GetOutputSettings();
        var programmeName = _resolvedProgrammeName;
        var inTrayPath = _inTrayFilePath;
        string? downloadedReportPath = null;
        PorticoAutomationService? porticoService = null;

        try
        {
            // Phase 1 — launch Edge, login, download the department report
            LogStatus($"Launching browser to fetch report for: {programmeName}");
            _statusLabel.Text = "Fetching report from Portico...";

            var config = new AppConfig();
            porticoService = new PorticoAutomationService();
            porticoService.StatusUpdated += (_, msg) => LogStatus(msg);

            await porticoService.InitialiseAsync(config);

            if (token.IsCancellationRequested) return;

            await porticoService.LoginAsync();

            if (token.IsCancellationRequested) return;

            var tempDir = Path.Combine(Path.GetTempPath(), "DossierADMerger");
            downloadedReportPath = await porticoService.DownloadDepartmentReportAsync(programmeName, tempDir);

            await porticoService.CloseAsync();
            porticoService = null;

            if (token.IsCancellationRequested) return;

            LogStatus("Report downloaded. Running data merge...");
            _statusLabel.Text = "Merging data...";

            // Phase 2 — run the merge on a background thread
            await Task.Run(async () =>
            {
                var inTrayRecords = _csvService.LoadInTrayRecords(inTrayPath);
                var appRecords    = _csvService.LoadApplicationRecords(downloadedReportPath);

                await Dispatcher.UIThread.InvokeAsync(() => {
                    ProcessingItems.Clear();
                    foreach (var r in inTrayRecords)
                        ProcessingItems.Add(new ProcessingItem { StudentNo = r.StudentNo, Name = r.Name, Status = "⏳ Pending", ReceivedDate = r.ReceivedOn ?? "" });
                    _footerStatus.Text = $"0/{inTrayRecords.Count} Processed";
                });

                var outputRecords = new List<OutputRecord>();
                int current = 0;

                foreach (var inTray in inTrayRecords)
                {
                    if (token.IsCancellationRequested) return;
                    current++;

                    var app = appRecords.FirstOrDefault(a => a.ApplicantID?.Trim() == inTray.StudentNo?.Trim());
                    string classification = "??";

                    if (app != null)
                    {
                        classification = _gradeService.DetermineUKClassification(
                            app.OverallGradeGPA ?? "", app.EquivalencyNote ?? "",
                            app.CountryOfStudy ?? "", app.QualificationName ?? "");

                        outputRecords.Add(new OutputRecord {
                            ReceivedDate    = DateFormatter.FormatDate(inTray.ReceivedOn ?? ""),
                            DueDate         = DateFormatter.CalculateDueDate(inTray.ReceivedOn ?? ""),
                            StudentNo       = inTray.StudentNo,
                            Programme       = ProgrammeMapping.GetCode(app.Programme ?? ""),
                            Forename        = app.Forename,
                            Surname         = app.Surname,
                            FeeStatus       = app.FeeStatus,
                            QualificationName = app.QualificationName,
                            DegreeSubject   = app.DegreeSubject,
                            InstitutionName = app.InstitutionName,
                            THERanking      = _rankingService.GetRanking(app.InstitutionName ?? ""),
                            CountryOfStudy  = app.CountryOfStudy,
                            EquivalencyNote = app.EquivalencyNote,
                            OverallGradeGPA = _gradeService.GetPreferredGradeDisplay(app.OverallGradeGPA ?? "", app.EquivalencyNote ?? ""),
                            DegreeStatus    = app.GradeAchievedPending,
                            UKGrade         = classification,
                            Gender          = app.Gender,
                            Nationality     = app.Nationality,
                            DateOfBirth     = DateFormatter.FormatDate(app.DateOfBirth ?? ""),
                            Email           = app.Email,
                            Paid            = app.Paid
                        });
                    }

                    await Dispatcher.UIThread.InvokeAsync(() => {
                        var item = ProcessingItems.FirstOrDefault(p => p.StudentNo == inTray.StudentNo);
                        if (item != null) { item.Status = app != null ? "✓ Done" : "⚠️ Missing"; item.UkGrade = classification; }
                        _mainProgressBar.Value = (double)current / inTrayRecords.Count * 100;
                        _footerStatus.Text = $"{current}/{inTrayRecords.Count} Processed";
                    });
                }

                _csvService.GenerateOutputFiles(outputRecords, _outputFolderPath, outputSettings);

                if (!token.IsCancellationRequested)
                {
                    await Dispatcher.UIThread.InvokeAsync(async () =>
                    {
                        PlayConfirmationSound();

                        // Phase 3 — delete downloaded report and in-tray file
                        LogStatus("Cleaning up source files...");
                        TryDeleteFile(downloadedReportPath);
                        TryDeleteFile(inTrayPath);
                        _inTrayFilePath = string.Empty;
                        _inTrayFileLabel.Text = "No file selected";
                        LogStatus("Cleanup done.");

                        int countFirst = outputRecords.Count(r => r.UKGrade == "1.0");
                        int countUpper = outputRecords.Count(r => r.UKGrade == "2.1");
                        int countLower = outputRecords.Count(r => r.UKGrade == "2.2");
                        int countThird = outputRecords.Count(r => r.UKGrade == "3.0");
                        int countOther = outputRecords.Count - (countFirst + countUpper + countLower + countThird);

                        string summaryList = $"{countFirst}\t(First Class)\n{countUpper}\t(Upper Second)\n{countLower}\t(Lower Second)";
                        if (countThird > 0) summaryList += $"\n{countThird}\t(Third Class)";
                        if (countOther > 0) summaryList += $"\n{countOther}\t(Other / Ungraded)";

                        await ShowMessageBoxAsync("Success",
                            $"Processing complete!\n\n{summaryList}\n\nExcel file(s) saved at:\n{_outputFolderPath}");
                    });
                }
            }, token);
        }
        catch (Exception ex) { LogStatus($"Error: {ex.Message}"); }
        finally
        {
            if (porticoService != null)
            {
                try { await porticoService.CloseAsync(); } catch { }
            }
            _isProcessing = false;
            _processButton.Content = "▶  Process Records";
            _processButton.Background = SolidColorBrush.Parse("#10B981");
            CheckReadyToProcess();
        }
    }

    private static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrEmpty(path)) return;
        try { File.Delete(path); } catch { }
    }

    private void SetVersion()
    {
        var v = Assembly.GetExecutingAssembly().GetName().Version;
        if (v != null)
            _versionLabel.Text = $"v{v.Major}.{v.Minor}.{v.Build}.{v.Revision:D4}";
    }

    private async void BrowseInTrayButton_Click(object? sender, RoutedEventArgs e)
    {
        var files = await TopLevel.GetTopLevel(this)!.StorageProvider.OpenFilePickerAsync(
            new FilePickerOpenOptions { Title = "Select InTray File", AllowMultiple = false,
                FileTypeFilter = new[] { new FilePickerFileType("Spreadsheet") { Patterns = new[] { "*.csv", "*.xlsx" } } } });
        if (files.Any()) {
            _inTrayFilePath = files[0].Path.LocalPath;
            _inTrayFileLabel.Text = Path.GetFileName(_inTrayFilePath);
            CheckReadyToProcess();
        }
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        e.DragEffects = e.Data.Contains(DataFormats.Files) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void OnInTrayDrop(object? sender, DragEventArgs e)
    {
        SetCardDragState(_inTrayCard, false);
        var file = e.Data.GetFiles()?.FirstOrDefault();
        if (file == null) return;
        _inTrayFilePath = file.Path.LocalPath;
        _inTrayFileLabel.Text = Path.GetFileName(_inTrayFilePath);
        CheckReadyToProcess();
    }

    private void SetCardDragState(Border card, bool active)
    {
        card.Background      = active ? SolidColorBrush.Parse("#EFF6FF") : SolidColorBrush.Parse("#CCFFFFFF");
        card.BorderBrush     = active ? SolidColorBrush.Parse("#2563EB") : SolidColorBrush.Parse("#80FFFFFF");
        card.BorderThickness = new Thickness(active ? 2 : 1);
    }

    private void PlayConfirmationSound()
    {
        try
        {
            string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "audio", "confirmed.mp3");
            if (!File.Exists(soundPath))
            {
                string devPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "audio", "confirmed.mp3");
                if (File.Exists(devPath)) soundPath = devPath;
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                if (File.Exists(soundPath))
                    Process.Start(new ProcessStartInfo { FileName = "afplay", Arguments = $"\"{soundPath}\"", CreateNoWindow = true, UseShellExecute = false });
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                if (File.Exists(soundPath))
                {
                    string fullPath = Path.GetFullPath(soundPath);
                    mciSendString($"open \"{fullPath}\" type mpegvideo alias confirm", null, 0, IntPtr.Zero);
                    mciSendString("play confirm from 0", null, 0, IntPtr.Zero);
                }
            }
        }
        catch (Exception ex) { Debug.WriteLine($"Audio error: {ex.Message}"); }
    }

    private void LogStatus(string message) =>
        Dispatcher.UIThread.Post(() => _statusLog.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n");

    private void ClearLogButton_Click(object? sender, RoutedEventArgs e) => _statusLog.Text = string.Empty;
    private void ExitButton_Click(object? sender, RoutedEventArgs e) => Environment.Exit(0);

    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        _inTrayFilePath = string.Empty;
        _resolvedProgrammeName = string.Empty;
        _inTrayFileLabel.Text  = "No file selected";
        _programmeCodeBox.Text = string.Empty;
        _programmeLookupLabel.Text = "Enter a programme code above";
        _programmeLookupLabel.Foreground = SolidColorBrush.Parse("#94A3B8");
        ProcessingItems.Clear();
        _mainProgressBar.Value = 0;
        _footerStatus.Text = "Ready";
        _statusLog.Text = string.Empty;
        CheckReadyToProcess();
        LogStatus("Application reset.");
    }
}

// Views/PatView.axaml.cs
//
// PAT — Personal Tutor Allocations tab.
// Loads a spreadsheet (Student ID + a "PAT (PIVOT)" tutor-code column), then for each row drives
// Portico's Personal Tutor Allocations page: types the student number, types the tutor code, and
// clicks "Apply New Criteria". It shares the single browser through PorticoAutomationService.Shared,
// as every tab must, and reuses the same ProcessingWindow progress UI as the other tabs.

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
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Dossier.Views;

public partial class PatView : UserControl
{
    private readonly IExcelService _excelService = new ExcelService();
    private readonly IPorticoAutomationService _automationService = PorticoAutomationService.Shared;
    private readonly AppConfig _config = new AppConfig();

    // PAT reads its data from the "Dashboard-Students" worksheet ONLY — never any other tab.
    private const string PatSheetName = "Dashboard-Students";

    private string _currentFilePath = string.Empty;
    private ObservableCollection<StudentRecord> _students = new();
    private CancellationTokenSource? _cancellationTokenSource;

    private Border _dropZone = null!;
    private TextBlock _dropZoneText = null!;
    private Button _browseButton = null!;
    private Border _sheetSelectionPanel = null!;
    private ComboBox _sheetComboBox = null!;
    private Button _loadSheetButton = null!;
    private Border _studentListPanel = null!;
    private DataGrid _studentGrid = null!;
    private TextBlock _studentCountText = null!;
    private Border _actionPanel = null!;
    private CheckBox _debugModeCheckBox = null!;
    private Button _startButton = null!;
    private Button _stopButton = null!;
    private TextBox _statusLog = null!;
    private Button _clearLogButton = null!;
    private Button _resetButton = null!;
    private ScrollViewer _rootScroller = null!;

    public PatView()
    {
        InitializeComponent();
        SetupEventHandlers();
        SetupDragDrop();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);

        _dropZone = this.FindControl<Border>("DropZone")!;
        _dropZoneText = this.FindControl<TextBlock>("DropZoneText")!;
        _browseButton = this.FindControl<Button>("BrowseButton")!;
        _sheetSelectionPanel = this.FindControl<Border>("SheetSelectionPanel")!;
        _sheetComboBox = this.FindControl<ComboBox>("SheetComboBox")!;
        _loadSheetButton = this.FindControl<Button>("LoadSheetButton")!;
        _studentListPanel = this.FindControl<Border>("StudentListPanel")!;
        _studentGrid = this.FindControl<DataGrid>("StudentGrid")!;
        _studentCountText = this.FindControl<TextBlock>("StudentCountText")!;
        _actionPanel = this.FindControl<Border>("ActionPanel")!;
        _debugModeCheckBox = this.FindControl<CheckBox>("DebugModeCheckBox")!;
        _startButton = this.FindControl<Button>("StartButton")!;
        _stopButton = this.FindControl<Button>("StopButton")!;
        _statusLog = this.FindControl<TextBox>("StatusLog")!;
        _clearLogButton = this.FindControl<Button>("ClearLogButton")!;
        _resetButton = this.FindControl<Button>("ResetButton")!;
        _rootScroller = this.FindControl<ScrollViewer>("RootScroller")!;
    }

    private Window? OwnerWindow => TopLevel.GetTopLevel(this) as Window;

    private void UpdateFooterStatus(string status)
    {
        var footer = OwnerWindow?.FindControl<TextBlock>("FooterStatus");
        if (footer != null)
            footer.Text = $"[PAT] {status}";
    }

    private void SetupEventHandlers()
    {
        _browseButton.Click += BrowseButton_Click;
        _loadSheetButton.Click += LoadSheetButton_Click;
        _startButton.Click += StartButton_Click;
        _stopButton.Click += StopButton_Click;
        _clearLogButton.Click += ClearLogButton_Click;
        _resetButton.Click += ResetButton_Click;
    }

    private void SetupDragDrop()
    {
        _dropZone.AddHandler(DragDrop.DropEvent, OnDrop);
        _dropZone.AddHandler(DragDrop.DragOverEvent, OnDragOver);
        _dropZone.AddHandler(DragDrop.DragEnterEvent, OnDragEnter);
        _dropZone.AddHandler(DragDrop.DragLeaveEvent, OnDragLeave);
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        e.DragEffects = e.Data.Contains(DataFormats.Files)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
    }

    private void OnDragEnter(object? sender, DragEventArgs e)
    {
        _dropZone.BorderBrush = new SolidColorBrush(Color.Parse("#4F46E5"));
        _dropZone.Background = new SolidColorBrush(Color.Parse("#EEF2FF"));
    }

    private void OnDragLeave(object? sender, DragEventArgs e)
    {
        _dropZone.BorderBrush = new SolidColorBrush(Color.Parse("#E2E8F0"));
        _dropZone.Background = new SolidColorBrush(Color.Parse("#FFFFFF"));
    }

    private async void OnDrop(object? sender, DragEventArgs e)
    {
        OnDragLeave(sender, e);

        var files = e.Data.GetFiles();
        var file = files?.FirstOrDefault();
        if (file == null) return;

        var path = file.Path.LocalPath;
        if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
        {
            await LoadExcelFileAsync(path);
        }
        else if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
        {
            LoadCsvFile(path);
        }
        else
        {
            LogStatus("Please drop an Excel (.xlsx/.xls) or CSV (.csv) file");
        }
    }

    private async void BrowseButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select Excel or CSV File",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Spreadsheet Files")
                {
                    Patterns = new[] { "*.xlsx", "*.xls", "*.csv" }
                }
            }
        });

        if (files.Count > 0)
        {
            var path = files[0].Path.LocalPath;
            if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                LoadCsvFile(path);
            else
                await LoadExcelFileAsync(path);
        }
    }

    private async Task LoadExcelFileAsync(string filePath)
    {
        try
        {
            _currentFilePath = filePath;
            _dropZoneText.Text = $"Loaded: {Path.GetFileName(filePath)}";
            LogStatus($"Loaded file: {filePath}");

            // PAT is hard-pinned to the "Dashboard-Students" tab — no sheet picker.
            _sheetSelectionPanel.IsVisible = false;

            var sheets = _excelService.GetSheetNames(filePath);
            var matchedSheet = sheets.FirstOrDefault(
                s => s.Trim().Equals(PatSheetName, StringComparison.OrdinalIgnoreCase));

            if (matchedSheet == null)
            {
                LogStatus($"Error: this spreadsheet has no '{PatSheetName}' tab. "
                    + $"PAT reads from the '{PatSheetName}' tab only. Found tabs: {string.Join(", ", sheets)}");
                UpdateFooterStatus($"No '{PatSheetName}' tab found in {Path.GetFileName(filePath)}");
                return;
            }

            PopulateStudents(_excelService.LoadStudentsFromFile(filePath, matchedSheet));
            UpdateFooterStatus($"File loaded: {Path.GetFileName(filePath)} — read '{matchedSheet}' tab");
            _rootScroller.Offset = new Avalonia.Vector(0, 0);
        }
        catch (Exception ex)
        {
            LogStatus($"Error loading file: {ex.Message}");
        }
    }

    private void LoadCsvFile(string filePath)
    {
        try
        {
            _currentFilePath = filePath;
            _dropZoneText.Text = $"Loaded: {Path.GetFileName(filePath)}";
            LogStatus($"Loaded CSV file: {filePath}");

            PopulateStudents(_excelService.LoadStudentsFromCsv(filePath));
            _rootScroller.Offset = new Avalonia.Vector(0, 0);
        }
        catch (Exception ex)
        {
            LogStatus($"Error loading CSV: {ex.Message}");
        }
    }

    private void LoadSheetButton_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var selectedSheet = _sheetComboBox.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedSheet))
            {
                LogStatus("Please select a worksheet.");
                return;
            }

            PopulateStudents(_excelService.LoadStudentsFromFile(_currentFilePath, selectedSheet));
            _rootScroller.Offset = new Avalonia.Vector(0, 0);
        }
        catch (Exception ex)
        {
            LogStatus($"Error loading students: {ex.Message}");
        }
    }

    // A record is allocated only if it has a tutor code AND "PAT Required" is not "N".
    private static bool IsNotRequired(StudentRecord s) =>
        s.PatRequired.Trim().Equals("N", StringComparison.OrdinalIgnoreCase);

    private static bool ShouldProcess(StudentRecord s) =>
        !string.IsNullOrWhiteSpace(s.PersonalTutor) && !IsNotRequired(s);

    // Derives the Portico programme/route code entered in the Programme filter box, from the Route
    // column (falling back to Programme). Returns null if the value can't be recognised.
    private static string? ResolveProgrammeCode(StudentRecord s)
    {
        // Prefer the Portico code read straight from the spreadsheet's Programme-code column.
        var direct = s.ProgrammeCode?.Trim() ?? "";
        if (direct.Length > 0)
        {
            var token = direct.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? direct;
            if (token.StartsWith("T", StringComparison.OrdinalIgnoreCase))
                return token.ToUpperInvariant();
        }

        // Otherwise map the programme/route NAME (or ADMerger short code) to a Portico code via
        // the shared ProgrammeMapping, so every tab (ADMerger, PAT, Portico, …) stays consistent.
        var raw = (!string.IsNullOrWhiteSpace(s.Route) ? s.Route : s.Programme)?.Trim() ?? "";
        if (raw.Length == 0) return null;
        if (raw.StartsWith("tms", StringComparison.OrdinalIgnoreCase)) return raw.ToUpperInvariant(); // already a code
        return ProgrammeMapping.GetPorticoCode(raw);
    }

    private void PopulateStudents(List<StudentRecord> students)
    {
        _students = new ObservableCollection<StudentRecord>(students);
        _studentGrid.ItemsSource = _students;

        var queue = _students.Count(ShouldProcess);
        var notRequired = _students.Count(IsNotRequired);
        var missing = _students.Count(s => !IsNotRequired(s) && string.IsNullOrWhiteSpace(s.PersonalTutor));

        _dropZone.IsVisible = false;
        _sheetSelectionPanel.IsVisible = false;
        _studentListPanel.IsVisible = true;
        _actionPanel.IsVisible = true;

        _studentCountText.Text =
            $"Loaded: {_students.Count} | To allocate: {queue} | PAT Required = N (skipped): {notRequired} | Missing code: {missing}";

        if (queue == 0)
            LogStatus("WARNING: Nothing to allocate. Check for a 'PAT (PIVOT)' tutor-code column and 'PAT Required' values.");
        else
            LogStatus($"Loaded {_students.Count} students — {queue} to allocate, {notRequired} skipped (PAT Required = N), {missing} missing a tutor code.");

        UpdateFooterStatus($"Ready to allocate {queue} students");
    }

    private async void StartButton_Click(object? sender, RoutedEventArgs e)
    {
        var queue = _students.Where(ShouldProcess).ToList();
        if (queue.Count == 0)
        {
            LogStatus("No students to allocate (need a tutor code and 'PAT Required' ≠ N).");
            return;
        }

        var debugMode = _debugModeCheckBox.IsChecked ?? false;
        _automationService.DebugMode = debugMode;

        _cancellationTokenSource = new CancellationTokenSource();
        _startButton.IsEnabled = false;
        _stopButton.IsEnabled = true;
        _browseButton.IsEnabled = false;
        _loadSheetButton.IsEnabled = false;

        using var sleepInhibitor = new SleepInhibitor();
        sleepInhibitor.Activate();

        // Show the tutor code on each row of the progress window (via the Decision display slot).
        var displayRecords = queue
            .Select(s => new StudentRecord
            {
                StudentNo = s.StudentNo,
                Forename = s.Forename,
                Surname = s.Surname,
                Decision = $"→ {s.PersonalTutor}"
            })
            .ToList();

        var processingWindow = new Views.ProcessingWindow();
        processingWindow.Initialize(displayRecords);
        processingWindow.CancelRequested += (s, e) => _cancellationTokenSource?.Cancel();

        EventHandler<string> statusHandler = (s, msg) => processingWindow.LogMessage(msg);
        EventHandler<StudentRecord> studentHandler = (s, student) =>
            processingWindow.UpdateStudentStatus(student.StudentNo, student.Status, student.ErrorMessage);

        _automationService.StatusUpdated += statusHandler;
        _automationService.StudentProcessed += studentHandler;

        processingWindow.Show();

        try
        {
            processingWindow.LogMessage("PAT — Personal Tutor Allocations.");
            processingWindow.LogMessage("Initialising browser automation...");
            await _automationService.InitialiseAsync(_config);

            processingWindow.LogMessage("Attempting login to Portico...");
            var loginSuccess = await _automationService.LoginAsync();
            if (!loginSuccess)
            {
                processingWindow.LogMessage("Login failed or timed out.");
                processingWindow.UpdateFooterStatus("Login failed");
                return;
            }

            processingWindow.LogMessage("Opening Personal Tutor Allocations...");
            await _automationService.NavigateToPersonalTutorAllocationsAsync();

            // One-off per spreadsheet: set the Department.
            processingWindow.LogMessage("Setting Department = Computer Science (once)...");
            await _automationService.ApplyPatDepartmentAsync("Computer Science");

            // One-off per spreadsheet: set the Programme code, then press "Apply New Criteria" — ONCE.
            // The spreadsheet's Programme column supplies the Portico code directly (e.g. TMSCOMSDDI19).
            var distinctCodes = queue
                .Select(ResolveProgrammeCode)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c!)
                .Distinct()
                .ToList();

            if (distinctCodes.Count == 0)
            {
                foreach (var s in queue)
                {
                    s.Status = ProcessingStatus.Failed;
                    s.ErrorMessage = "No Portico programme code — add a Programme code column (e.g. TMSCOMSDDI19) or a recognised Route.";
                    processingWindow.UpdateStudentStatus(s.StudentNo, s.Status, s.ErrorMessage);
                }
                RefreshStudentGrid();
                processingWindow.LogMessage("No programme code could be resolved — nothing processed.");
            }
            else
            {
                if (distinctCodes.Count > 1)
                    processingWindow.LogMessage(
                        $"WARNING: {distinctCodes.Count} different programme codes found ({string.Join(", ", distinctCodes)}). "
                        + "PAT sets the Programme once per spreadsheet — using the first; students on the others may not appear.");

                var programmeCode = distinctCodes[0];
                processingWindow.LogMessage($"Setting Programme = {programmeCode} and pressing 'Apply New Criteria' (once)...");
                await _automationService.ApplyPatProgrammeAsync(programmeCode);

                foreach (var student in queue)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        processingWindow.LogMessage("Processing cancelled by user.");
                        processingWindow.UpdateFooterStatus("Cancelled");
                        break;
                    }

                    await _automationService.ProcessStudentPatAsync(student);
                    RefreshStudentGrid();

                    if (debugMode)
                    {
                        processingWindow.LogMessage("DEBUG MODE: stopped after the first student — browser paused for inspection.");
                        processingWindow.UpdateFooterStatus("Debug mode complete - browser paused");
                        processingWindow.ProcessingComplete();
                        return;
                    }
                }
            }

            processingWindow.LogMessage("Processing complete.");
            processingWindow.ProcessingComplete();
        }
        catch (Exception ex)
        {
            processingWindow.LogMessage($"Error: {ex.Message}");
            processingWindow.UpdateFooterStatus("Error occurred");
        }
        finally
        {
            _automationService.StatusUpdated -= statusHandler;
            _automationService.StudentProcessed -= studentHandler;

            _startButton.IsEnabled = true;
            _browseButton.IsEnabled = true;
            _loadSheetButton.IsEnabled = true;
            _stopButton.IsEnabled = false;

            processingWindow.LogMessage("Browser left open — use it freely; the next run will take it over automatically.");

            var successCount = queue.Count(s => s.Status == ProcessingStatus.Success);
            var failedCount = queue.Count(s => s.Status == ProcessingStatus.Failed);
            var skippedCount = queue.Count(s => s.Status == ProcessingStatus.Skipped);
            UpdateFooterStatus($"Complete: {successCount} allocated, {skippedCount} already allocated (skipped), {failedCount} failed");

            try
            {
                var csvPath = ExportResultsCsv(queue);
                if (csvPath != null)
                {
                    processingWindow.LogMessage($"Results CSV written to: {csvPath}");
                    LogStatus($"Results CSV written to: {csvPath}");
                }
            }
            catch (Exception ex)
            {
                processingWindow.LogMessage($"Could not write results CSV: {ex.Message}");
                LogStatus($"Could not write results CSV: {ex.Message}");
            }
        }
    }

    // Writes a results CSV to the Desktop: one row per attempted student, six columns.
    // Completed = Y when the allocation succeeded, N otherwise. Returns the file path (or null).
    private static string? ExportResultsCsv(List<StudentRecord> queue)
    {
        if (queue.Count == 0) return null;

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("Student Number,Programme,Surname,PAT Allocated (Staff Name),PAT Number,Completed");

        foreach (var s in queue)
        {
            var programme = !string.IsNullOrWhiteSpace(s.Programme) ? s.Programme : s.Route;
            // Skipped = already allocated in Portico, so still "done" (Y).
            var completed = (s.Status == ProcessingStatus.Success || s.Status == ProcessingStatus.Skipped) ? "Y" : "N";
            sb.AppendLine(string.Join(",",
                CsvField(s.StudentNo),
                CsvField(programme),
                CsvField(s.Surname),
                CsvField(s.PersonalTutorName),
                CsvField(s.PersonalTutor),
                CsvField(completed)));
        }

        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var path = Path.Combine(desktop, $"PAT Allocations {DateTime.Now:yyyy-MM-dd HHmm}.csv");
        File.WriteAllText(path, sb.ToString());
        return path;
    }

    private static string CsvField(string? value)
    {
        var v = value ?? "";
        if (v.Contains('"') || v.Contains(',') || v.Contains('\n') || v.Contains('\r'))
            return "\"" + v.Replace("\"", "\"\"") + "\"";
        return v;
    }

    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        _students.Clear();
        _currentFilePath = string.Empty;

        _dropZoneText.Text = "Drag & Drop Excel or CSV File";
        _dropZone.IsVisible = true;
        _sheetSelectionPanel.IsVisible = false;
        _studentListPanel.IsVisible = false;
        _actionPanel.IsVisible = false;
        _studentGrid.ItemsSource = null;
        _sheetComboBox.ItemsSource = null;
        _statusLog.Text = string.Empty;
        _debugModeCheckBox.IsChecked = false;

        UpdateFooterStatus("Waiting for file...");
        LogStatus("PAT tab reset - ready to load new file");
    }

    private void StopButton_Click(object? sender, RoutedEventArgs e)
    {
        _cancellationTokenSource?.Cancel();
        LogStatus("Stopping after the current student — browser stays open.");
        _stopButton.IsEnabled = false;
    }

    private void ClearLogButton_Click(object? sender, RoutedEventArgs e) => _statusLog.Text = string.Empty;

    private void LogStatus(string message)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            _statusLog.Text += $"[{timestamp}] {message}\n";
            _statusLog.CaretIndex = _statusLog.Text?.Length ?? 0;
        });
    }

    private void RefreshStudentGrid()
    {
        _studentGrid.ItemsSource = null;
        _studentGrid.ItemsSource = _students;
    }
}

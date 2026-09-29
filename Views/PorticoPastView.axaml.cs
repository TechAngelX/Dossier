// Views/PorticoPastView.axaml.cs
//
// A self-contained copy of the main "Portico" tab that targets PAST applicants.
// Every student search auto-selects the 2026/27 academic year in Portico's "Year"
// dropdown (via IPorticoAutomationService.TargetAcademicYear) instead of the default
// "Current Applications". This view is deliberately independent of MainWindow's Portico
// tab so changes here never affect the working current-cohort workflow. It shares the
// single browser through PorticoAutomationService.Shared, as every tab must.

using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
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

public partial class PorticoPastView : UserControl
{
    // The past academic year every search in this tab targets.
    private const string PastYear = "2026/27";

    private readonly IExcelService _excelService = new ExcelService();
    private readonly IPorticoAutomationService _automationService = PorticoAutomationService.Shared;
    private readonly IPdfRenameService _pdfRenameService = new PdfRenameService();
    private readonly AppConfig _config = new AppConfig();

    private string _currentFilePath = string.Empty;
    private List<StudentRecord> _allStudents = new();
    private ObservableCollection<StudentRecord> _students = new();
    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isProcessing;

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
    private RadioButton _processAcceptsCheckBox = null!;
    private RadioButton _processRejectsCheckBox = null!;
    private RadioButton _mergeOverviewCheckBox = null!;
    private CheckBox _debugModeCheckBox = null!;
    private Button _startButton = null!;
    private Button _stopButton = null!;
    private TextBox _statusLog = null!;
    private Button _clearLogButton = null!;
    private Button _resetButton = null!;
    private Button _pdfToolsButton = null!;
    private ScrollViewer _rootScroller = null!;

    public PorticoPastView()
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
        _processAcceptsCheckBox = this.FindControl<RadioButton>("ProcessAcceptsCheckBox")!;
        _processRejectsCheckBox = this.FindControl<RadioButton>("ProcessRejectsCheckBox")!;
        _mergeOverviewCheckBox = this.FindControl<RadioButton>("MergeOverviewCheckBox")!;
        _debugModeCheckBox = this.FindControl<CheckBox>("DebugModeCheckBox")!;
        _startButton = this.FindControl<Button>("StartButton")!;
        _stopButton = this.FindControl<Button>("StopButton")!;
        _statusLog = this.FindControl<TextBox>("StatusLog")!;
        _clearLogButton = this.FindControl<Button>("ClearLogButton")!;
        _resetButton = this.FindControl<Button>("ResetButton")!;
        _pdfToolsButton = this.FindControl<Button>("PdfToolsButton")!;
        _rootScroller = this.FindControl<ScrollViewer>("RootScroller")!;
    }

    // The status footer lives on MainWindow (shared across tabs). Update it so loading a file
    // here shows "Ready to process ..." just like the main Portico tab, instead of leaving the
    // stale "Waiting for file..." that makes it look as though the file wasn't accepted.
    private void UpdateFooterStatus(string status)
    {
        var footer = OwnerWindow?.FindControl<TextBlock>("FooterStatus");
        if (footer != null)
            footer.Text = $"[Past 2026/27] {status}";
    }

    private void SetupEventHandlers()
    {
        _browseButton.Click += BrowseButton_Click;
        _loadSheetButton.Click += LoadSheetButton_Click;
        _startButton.Click += StartButton_Click;
        _stopButton.Click += StopButton_Click;
        _clearLogButton.Click += ClearLogButton_Click;
        _processAcceptsCheckBox.IsCheckedChanged += FilterCheckBox_Changed;
        _processRejectsCheckBox.IsCheckedChanged += FilterCheckBox_Changed;
        _mergeOverviewCheckBox.IsCheckedChanged += FilterCheckBox_Changed;
        _resetButton.Click += ResetButton_Click;
        _pdfToolsButton.Click += PdfToolsButton_Click;
    }

    private Window? OwnerWindow => TopLevel.GetTopLevel(this) as Window;

    private void FilterCheckBox_Changed(object? sender, RoutedEventArgs e) => FilterStudentList();

    private void FilterStudentList()
    {
        if (_allStudents == null || _allStudents.Count == 0)
        {
            LogStatus("No students loaded to filter.");
            return;
        }

        var processAccepts = _processAcceptsCheckBox.IsChecked ?? false;
        var processRejects = _processRejectsCheckBox.IsChecked ?? false;
        var mergeOverview = _mergeOverviewCheckBox.IsChecked ?? false;

        List<StudentRecord> filtered;

        if (mergeOverview)
        {
            filtered = _allStudents.ToList();
        }
        else
        {
            filtered = _allStudents.Where(s =>
            {
                var isAccept = s.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase);
                var isReject = s.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase);

                if (isAccept && processAccepts) return true;
                if (isReject && processRejects) return true;
                return false;
            }).ToList();
        }

        _students = new ObservableCollection<StudentRecord>(filtered);
        _studentGrid.ItemsSource = null;
        _studentGrid.ItemsSource = _students;

        var acceptCount = filtered.Count(s => s.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase));
        var rejectCount = filtered.Count(s => s.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase));

        if (mergeOverview)
        {
            _studentCountText.Text = $"Showing: {filtered.Count} | Merge Overview mode";
            UpdateFooterStatus($"Ready to merge overview for {filtered.Count} students");
            LogStatus($"Merge Overview mode: {filtered.Count} students loaded");
        }
        else
        {
            _studentCountText.Text = $"Showing: {filtered.Count} | Accepts: {acceptCount} | Rejects: {rejectCount}";
            UpdateFooterStatus($"Ready to process {filtered.Count} students");
            LogStatus($"Filtered: {filtered.Count} students to process (Accepts: {acceptCount}, Rejects: {rejectCount})");
        }
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
        _dropZone.BorderBrush = new SolidColorBrush(Color.Parse("#007BFF"));
        _dropZone.Background = new SolidColorBrush(Color.Parse("#E7F3FF"));
    }

    private void OnDragLeave(object? sender, DragEventArgs e)
    {
        _dropZone.BorderBrush = new SolidColorBrush(Color.Parse("#DEE2E6"));
        _dropZone.Background = new SolidColorBrush(Color.Parse("#F8F9FA"));
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

            var sheets = _excelService.GetSheetNames(filePath);
            _sheetComboBox.ItemsSource = sheets;

            var deptInTray = sheets.FirstOrDefault(s =>
                s.Contains("Dept", StringComparison.OrdinalIgnoreCase) &&
                s.Contains("tray", StringComparison.OrdinalIgnoreCase));

            if (deptInTray != null)
                _sheetComboBox.SelectedItem = deptInTray;
            else if (sheets.Count > 0)
                _sheetComboBox.SelectedIndex = 0;

            _sheetSelectionPanel.IsVisible = true;
            UpdateFooterStatus($"File loaded: {Path.GetFileName(filePath)} — choose a worksheet");
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

            _allStudents = _excelService.LoadStudentsFromCsv(filePath);

            var totalAccepts = _allStudents.Count(s => s.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase));
            var totalRejects = _allStudents.Count(s => s.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase));

            LogStatus($"Loaded {_allStudents.Count} students from CSV (Accepts: {totalAccepts}, Rejects: {totalRejects})");

            _dropZone.IsVisible = false;
            _sheetSelectionPanel.IsVisible = false;
            _studentListPanel.IsVisible = true;
            _actionPanel.IsVisible = true;

            FilterStudentList();
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

            _allStudents = _excelService.LoadStudentsFromFile(_currentFilePath, selectedSheet);

            var totalAccepts = _allStudents.Count(s => s.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase));
            var totalRejects = _allStudents.Count(s => s.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase));

            LogStatus($"Loaded {_allStudents.Count} students from '{selectedSheet}' (Accepts: {totalAccepts}, Rejects: {totalRejects})");

            _dropZone.IsVisible = false;
            _sheetSelectionPanel.IsVisible = false;
            _studentListPanel.IsVisible = true;
            _actionPanel.IsVisible = true;

            FilterStudentList();
            _rootScroller.Offset = new Avalonia.Vector(0, 0);
        }
        catch (Exception ex)
        {
            LogStatus($"Error loading students: {ex.Message}");
        }
    }

    private async void StartButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_students.Count == 0)
        {
            LogStatus("No students loaded to process.");
            return;
        }

        var debugMode = _debugModeCheckBox.IsChecked ?? false;
        _automationService.DebugMode = debugMode;

        var mergeOverviewSelected = _mergeOverviewCheckBox.IsChecked ?? false;
        string? batchDownloadPath = null;

        if (mergeOverviewSelected)
        {
            var (_, batchFolderName) = await ShowBatchNumberDialog();
            if (batchFolderName == null) return; // user cancelled

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            batchDownloadPath = Path.Combine(desktop, GetBatchFolderName(), batchFolderName);
            Directory.CreateDirectory(batchDownloadPath);
            LogStatus($"Batch folder: {batchDownloadPath}");
        }

        _isProcessing = true;
        _cancellationTokenSource = new CancellationTokenSource();

        _startButton.IsEnabled = false;
        _stopButton.IsEnabled = true;
        _browseButton.IsEnabled = false;
        _loadSheetButton.IsEnabled = false;

        // Target the past academic year for every search in this run.
        _automationService.TargetAcademicYear = PastYear;

        using var sleepInhibitor = new SleepInhibitor();
        sleepInhibitor.Activate();

        var processingWindow = new Views.ProcessingWindow();
        processingWindow.Initialize(_students.ToList());
        processingWindow.CancelRequested += (s, e) => _cancellationTokenSource?.Cancel();

        EventHandler<string> statusHandler = (s, msg) => processingWindow.LogMessage(msg);
        EventHandler<StudentRecord> studentHandler = (s, student) =>
            processingWindow.UpdateStudentStatus(student.StudentNo, student.Status, student.ErrorMessage);

        _automationService.StatusUpdated += statusHandler;
        _automationService.StudentProcessed += studentHandler;

        processingWindow.Show();

        try
        {
            processingWindow.LogMessage($"Past Applicants mode — targeting {PastYear} academic year.");
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

            await _automationService.NavigateToUclSelectAsync();

            var processAccepts = _processAcceptsCheckBox.IsChecked ?? true;
            var processRejects = _processRejectsCheckBox.IsChecked ?? false;
            var mergeOverview = _mergeOverviewCheckBox.IsChecked ?? false;

            if (mergeOverview)
            {
                var sorted = _students.OrderBy(s => long.TryParse(s.StudentNo, out var n) ? n : long.MaxValue)
                                      .ThenBy(s => s.StudentNo, StringComparer.Ordinal)
                                      .ToList();
                _students = new ObservableCollection<StudentRecord>(sorted);
                _studentGrid.ItemsSource = _students;
                processingWindow.ReorderStudents(sorted);

                var downloadPath = batchDownloadPath!;
                processingWindow.LogMessage($"Download path: {downloadPath}");
                processingWindow.LogMessage("Students sorted by student number (lowest → highest)");

                if (debugMode)
                {
                    processingWindow.LogMessage("DEBUG MODE: Processing only FIRST student for merge overview");
                    var firstStudent = _students.First();

                    var debugExistingFiles = Directory.GetFiles(downloadPath, "*.pdf", SearchOption.TopDirectoryOnly)
                        .Concat(Directory.GetFiles(downloadPath, "*.PDF", SearchOption.TopDirectoryOnly))
                        .Select(f => Path.GetFileName(f))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    var debugAlreadyExists = debugExistingFiles.Any(f =>
                        f.Contains(firstStudent.StudentNo, StringComparison.OrdinalIgnoreCase));

                    if (debugAlreadyExists)
                    {
                        firstStudent.Status = ProcessingStatus.Skipped;
                        processingWindow.LogMessage($"SKIPPED {firstStudent.StudentNo}: PDF already exists");
                        processingWindow.UpdateStudentStatus(firstStudent.StudentNo, ProcessingStatus.Skipped);
                    }
                    else
                    {
                        await _automationService.ProcessStudentMergeOverviewAsync(firstStudent, downloadPath);
                        if (firstStudent.Status == ProcessingStatus.Success)
                            AutoRenameStudentPdf(firstStudent, downloadPath, processingWindow);
                    }

                    processingWindow.LogMessage("DEBUG MODE COMPLETE: Browser paused for inspection.");
                    processingWindow.UpdateFooterStatus("Debug mode complete - browser paused");
                    processingWindow.ProcessingComplete();
                    RefreshStudentGrid();
                    return;
                }

                var existingFiles = Directory.GetFiles(downloadPath, "*.pdf", SearchOption.TopDirectoryOnly)
                    .Concat(Directory.GetFiles(downloadPath, "*.PDF", SearchOption.TopDirectoryOnly))
                    .Select(f => Path.GetFileName(f))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var student in _students)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        processingWindow.LogMessage("Processing cancelled by user.");
                        processingWindow.UpdateFooterStatus("Cancelled");
                        break;
                    }

                    var alreadyExists = existingFiles.Any(f =>
                        f.Contains(student.StudentNo, StringComparison.OrdinalIgnoreCase));

                    if (alreadyExists)
                    {
                        student.Status = ProcessingStatus.Skipped;
                        processingWindow.LogMessage($"SKIPPED {student.StudentNo}: PDF already exists");
                        processingWindow.UpdateStudentStatus(student.StudentNo, ProcessingStatus.Skipped);
                        RefreshStudentGrid();
                        continue;
                    }

                    await _automationService.ProcessStudentMergeOverviewAsync(student, downloadPath);
                    if (student.Status == ProcessingStatus.Success)
                        AutoRenameStudentPdf(student, downloadPath, processingWindow);
                    RefreshStudentGrid();
                }
            }
            else
            {
                if (debugMode)
                {
                    processingWindow.LogMessage("DEBUG MODE: Processing only FIRST student");
                    var firstStudent = _students.First();

                    var isAccept = firstStudent.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase);
                    var isReject = firstStudent.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase);

                    if (isAccept && processAccepts)
                        await _automationService.ProcessStudentAcceptAsync(firstStudent);
                    else if (isReject && processRejects)
                        await _automationService.ProcessStudentRejectAsync(firstStudent);

                    processingWindow.LogMessage("DEBUG MODE COMPLETE: Browser paused for inspection.");
                    processingWindow.LogMessage("Verify that 'Reject' is selected and 'Reason 1' shows option 8");
                    processingWindow.UpdateFooterStatus("Debug mode complete - browser paused");
                    processingWindow.ProcessingComplete();

                    RefreshStudentGrid();
                    return;
                }

                foreach (var student in _students)
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        processingWindow.LogMessage("Processing cancelled by user.");
                        processingWindow.UpdateFooterStatus("Cancelled");
                        break;
                    }

                    var isAccept = student.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase);
                    var isReject = student.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase);

                    if (isAccept && processAccepts)
                    {
                        await _automationService.ProcessStudentAcceptAsync(student);
                        await _automationService.NavigateToUclSelectAsync();
                    }
                    else if (isReject && processRejects)
                    {
                        await _automationService.ProcessStudentRejectAsync(student);
                        await _automationService.NavigateToUclSelectAsync();
                    }
                    else
                    {
                        student.Status = ProcessingStatus.Skipped;
                        processingWindow.LogMessage($"Skipped {student.StudentNo} (Decision: {student.Decision})");
                    }

                    RefreshStudentGrid();
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

            // Reset so the main Portico tab always runs against "Current Applications".
            _automationService.TargetAcademicYear = null;

            _isProcessing = false;
            _startButton.IsEnabled = true;
            _browseButton.IsEnabled = true;
            _loadSheetButton.IsEnabled = true;
            _stopButton.IsEnabled = false;

            processingWindow.LogMessage("Browser left open — use it freely; the next run will take it over automatically.");

            var successCount = _students.Count(s => s.Status == ProcessingStatus.Success);
            var failedCount = _students.Count(s => s.Status == ProcessingStatus.Failed);
            var skippedTotal = _students.Count(s => s.Status == ProcessingStatus.Skipped);
            var footerMsg = skippedTotal > 0
                ? $"Complete: {successCount} successful, {failedCount} failed, {skippedTotal} skipped"
                : $"Complete: {successCount} successful, {failedCount} failed";
            UpdateFooterStatus(footerMsg);
        }
    }

    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        _allStudents.Clear();
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

        _processRejectsCheckBox.IsChecked = true;
        _processAcceptsCheckBox.IsChecked = false;
        _mergeOverviewCheckBox.IsChecked = false;
        _debugModeCheckBox.IsChecked = false;

        UpdateFooterStatus("Waiting for file...");
        LogStatus("Past-applicant tab reset - ready to load new file");
    }

    private void PdfToolsButton_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var downloadPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                GetBatchFolderName());

            var window = new Views.PdfToolsWindow(_students.ToList(), downloadPath);
            window.Show();
            window.Activate();
        }
        catch (Exception ex)
        {
            LogStatus($"PDF Tools failed to open: {ex.Message}");
        }
    }

    // Prompts for a batch number before a Merge Overview run.
    // Returns (batchNum, folderName) or (0, null) if cancelled.
    private async Task<(int batchNum, string? folderName)> ShowBatchNumberDialog()
    {
        int resultNum = 0;
        string? resultFolder = null;

        var dialog = new Window
        {
            Title = "Merge Overview — Batch Setup (Past 2026/27)",
            Width = 480,
            Height = 230,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false
        };

        var programmeFolderName = GetBatchFolderName();

        var panel = new StackPanel { Margin = new Avalonia.Thickness(24), Spacing = 14 };

        panel.Children.Add(new TextBlock
        {
            Text = $"Files will download to Desktop/{programmeFolderName}/Batch N - ...\nEnter the batch number for this run:",
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.Parse("#2D3748"))
        });

        var inputRow = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 12 };
        inputRow.Children.Add(new TextBlock
        {
            Text = "Batch Number:",
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            Foreground = new SolidColorBrush(Color.Parse("#4A5568"))
        });
        var numBox = new TextBox { Width = 90 };
        inputRow.Children.Add(numBox);
        panel.Children.Add(inputRow);

        var previewText = new TextBlock
        {
            Text = "Folder: —",
            Foreground = new SolidColorBrush(Color.Parse("#4F46E5")),
            FontWeight = FontWeight.SemiBold,
            FontSize = 13
        };
        panel.Children.Add(previewText);

        numBox.TextChanged += (s, e) =>
        {
            if (int.TryParse(numBox.Text, out int n) && n > 0)
                previewText.Text = $"Folder: {programmeFolderName}/{ComputeBatchFolderName(n)}";
            else
                previewText.Text = "Folder: —";
        };

        var buttons = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
            Spacing = 10
        };

        var ok = new Button { Content = "Start Processing", Padding = new Avalonia.Thickness(16, 8) };
        var cancel = new Button { Content = "Cancel", Padding = new Avalonia.Thickness(12, 8) };

        ok.Click += (s, e) =>
        {
            if (int.TryParse(numBox.Text, out int n) && n > 0)
            {
                resultNum = n;
                resultFolder = ComputeBatchFolderName(n);
                dialog.Close();
            }
        };
        cancel.Click += (s, e) => dialog.Close();

        buttons.Children.Add(ok);
        buttons.Children.Add(cancel);
        panel.Children.Add(buttons);
        dialog.Content = panel;

        var owner = OwnerWindow;
        if (owner != null)
            await dialog.ShowDialog(owner);
        else
            dialog.Show();

        return (resultNum, resultFolder);
    }

    private string ComputeBatchFolderName(int batchNum)
    {
        var dates = _students
            .Where(s => s.ReceivedDate.HasValue)
            .Select(s => s.ReceivedDate!.Value.Date)
            .ToList();

        if (dates.Count == 0)
            return $"Batch {batchNum}";

        var start = dates.Min();
        var end = dates.Max();
        return start.Date == end.Date
            ? $"Batch {batchNum} - {start:MMM d}"
            : $"Batch {batchNum} - {start:MMM d} - {end:MMM d}";
    }

    private string GetBatchFolderName()
    {
        var programmes = _students
            .Select(s => s.Programme?.Trim().ToUpperInvariant())
            .Where(p => !string.IsNullOrEmpty(p))
            .Distinct()
            .OrderBy(p => p)
            .ToList();

        if (programmes.Count == 0 || programmes.Count > 3)
            return "LATEST_BATCH";

        var safe = programmes.Select(p => new string(p!.Where(char.IsLetterOrDigit).ToArray()));
        return string.Join("_", safe) + "_LATEST";
    }

    private void AutoRenameStudentPdf(StudentRecord student, string folderPath, Views.ProcessingWindow processingWindow)
    {
        try
        {
            var overviewFiles = Directory.GetFiles(folderPath, $"{student.StudentNo}-*.pdf", SearchOption.TopDirectoryOnly)
                .Concat(Directory.GetFiles(folderPath, $"{student.StudentNo}-*.PDF", SearchOption.TopDirectoryOnly))
                .ToList();

            if (overviewFiles.Count == 0) return;

            var originalPath = overviewFiles[0];
            var newFilename = _pdfRenameService.GenerateNewFilename(student);
            var newPath = Path.Combine(folderPath, newFilename);

            if (File.Exists(newPath)) return;
            File.Move(originalPath, newPath);
            processingWindow.LogMessage($"Renamed: {Path.GetFileName(originalPath)} → {newFilename}");
        }
        catch (Exception ex)
        {
            processingWindow.LogMessage($"Auto-rename failed for {student.StudentNo}: {ex.Message}");
        }
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

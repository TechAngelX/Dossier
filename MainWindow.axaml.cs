// MainWindow.axaml.cs

using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Playwrighter.Models;
using Playwrighter.Services;
using System.Collections.ObjectModel;

namespace Playwrighter;

public partial class MainWindow : Window
{
    private readonly IExcelService _excelService;
    private readonly IPorticoAutomationService _automationService;
    private readonly AppConfig _config;
    
    private string _currentFilePath = string.Empty;
    private List<StudentRecord> _allStudents = new();
    private ObservableCollection<StudentRecord> _students = new();
    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isProcessing;
    
    public MainWindow()
    {
        InitializeComponent();
        
        _excelService = new ExcelService();
        _automationService = new PorticoAutomationService();
        _config = new AppConfig();
        
        SetupEventHandlers();
        SetupDragDrop();
        
        _automationService.StatusUpdated += OnStatusUpdated;
        _automationService.StudentProcessed += OnStudentProcessed;
    }
    
    private void InitializeComponent()
    {
        Avalonia.Markup.Xaml.AvaloniaXamlLoader.Load(this);
        
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
        _startButton = this.FindControl<Button>("StartButton")!;
        _stopButton = this.FindControl<Button>("StopButton")!;
        _statusLog = this.FindControl<TextBox>("StatusLog")!;
        _clearLogButton = this.FindControl<Button>("ClearLogButton")!;
        _footerStatus = this.FindControl<TextBlock>("FooterStatus")!;
        _settingsButton = this.FindControl<Button>("SettingsButton")!;
        _exitButton = this.FindControl<Button>("ExitButton")!;
    }
    
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
    private Button _startButton = null!;
    private Button _stopButton = null!;
    private TextBox _statusLog = null!;
    private Button _clearLogButton = null!;
    private TextBlock _footerStatus = null!;
    private Button _settingsButton = null!;
    private Button _exitButton = null!;
    
    private void SetupEventHandlers()
    {
        _browseButton.Click += BrowseButton_Click;
        _loadSheetButton.Click += LoadSheetButton_Click;
        _startButton.Click += StartButton_Click;
        _stopButton.Click += StopButton_Click;
        _clearLogButton.Click += ClearLogButton_Click;
        _settingsButton.Click += SettingsButton_Click;
        _exitButton.Click += ExitButton_Click;
        _processAcceptsCheckBox.IsCheckedChanged += FilterCheckBox_Changed;
        _processRejectsCheckBox.IsCheckedChanged += FilterCheckBox_Changed;
    }
    
    private void FilterCheckBox_Changed(object? sender, RoutedEventArgs e)
    {
        FilterStudentList();
    }
    
    private void FilterStudentList()
    {
        if (_allStudents == null || _allStudents.Count == 0)
        {
            LogStatus("No students loaded to filter.");
            return;
        }
        
        var processAccepts = _processAcceptsCheckBox.IsChecked ?? false;
        var processRejects = _processRejectsCheckBox.IsChecked ?? false;
        
        var filtered = _allStudents.Where(s =>
        {
            var isAccept = s.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase);
            var isReject = s.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase);
            
            if (isAccept && processAccepts) return true;
            if (isReject && processRejects) return true;
            return false;
        }).ToList();
        
        _students = new ObservableCollection<StudentRecord>(filtered);
        _studentGrid.ItemsSource = null;
        _studentGrid.ItemsSource = _students;
        
        var acceptCount = filtered.Count(s => s.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase));
        var rejectCount = filtered.Count(s => s.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase));
        
        _studentCountText.Text = $"Showing: {filtered.Count} | Accepts: {acceptCount} | Rejects: {rejectCount}";
        UpdateFooterStatus($"Ready to process {filtered.Count} students");
        
        LogStatus($"Filtered: {filtered.Count} students to process (Accepts: {acceptCount}, Rejects: {rejectCount})");
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
        if (files != null)
        {
            var file = files.FirstOrDefault();
            if (file != null)
            {
                var path = file.Path.LocalPath;
                if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    await LoadExcelFileAsync(path);
                }
                else
                {
                    LogStatus("Please drop an Excel file (.xlsx or .xls)");
                }
            }
        }
    }
    
    private async void BrowseButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;
        
        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select Excel File",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Excel Files")
                {
                    Patterns = new[] { "*.xlsx", "*.xls" }
                }
            }
        });
        
        if (files.Count > 0)
        {
            await LoadExcelFileAsync(files[0].Path.LocalPath);
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
            {
                _sheetComboBox.SelectedItem = deptInTray;
            }
            else if (sheets.Count > 0)
            {
                _sheetComboBox.SelectedIndex = 0;
            }
            
            _sheetSelectionPanel.IsVisible = true;
            UpdateFooterStatus($"File loaded: {Path.GetFileName(filePath)}");
        }
        catch (Exception ex)
        {
            LogStatus($"Error loading file: {ex.Message}");
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
            
            var totalAccepts = _allStudents.Count(s => 
                s.Decision.Equals("Accept", StringComparison.OrdinalIgnoreCase));
            var totalRejects = _allStudents.Count(s => 
                s.Decision.Equals("Reject", StringComparison.OrdinalIgnoreCase));
            
            LogStatus($"Loaded {_allStudents.Count} students from '{selectedSheet}' (Accepts: {totalAccepts}, Rejects: {totalRejects})");
            
            _studentListPanel.IsVisible = true;
            _actionPanel.IsVisible = true;
            
            FilterStudentList();
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
        
        _isProcessing = true;
        _cancellationTokenSource = new CancellationTokenSource();
        
        _startButton.IsEnabled = false;
        _stopButton.IsEnabled = true;
        _browseButton.IsEnabled = false;
        _loadSheetButton.IsEnabled = false;
        
        try
        {
            LogStatus("Initialising browser automation...");
            await _automationService.InitialiseAsync(_config);
            
            LogStatus("Attempting login to Portico...");
            var loginSuccess = await _automationService.LoginAsync();
            
            if (!loginSuccess)
            {
                LogStatus("Login failed or timed out.");
                return;
            }
            
            await _automationService.NavigateToUclSelectAsync();
            
            var processAccepts = _processAcceptsCheckBox.IsChecked ?? true;
            var processRejects = _processRejectsCheckBox.IsChecked ?? false;
            
            foreach (var student in _students)
            {
                if (_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    LogStatus("Processing cancelled by user.");
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
                    LogStatus($"Skipped {student.StudentNo} (Decision: {student.Decision})");
                }
                
                RefreshStudentGrid();
            }
            
            LogStatus("Processing complete.");
        }
        catch (Exception ex)
        {
            LogStatus($"Error during processing: {ex.Message}");
            LogStatus("Browser left open for debugging. Click Stop to close it.");
            _stopButton.IsEnabled = true;
            return;
        }
        finally
        {
            _isProcessing = false;
            _startButton.IsEnabled = true;
            _browseButton.IsEnabled = true;
            _loadSheetButton.IsEnabled = true;
        }
        
        _stopButton.IsEnabled = false;
        await _automationService.CloseAsync();
        
        var successCount = _students.Count(s => s.Status == ProcessingStatus.Success);
        var failedCount = _students.Count(s => s.Status == ProcessingStatus.Failed);
        UpdateFooterStatus($"Complete: {successCount} successful, {failedCount} failed");
    }
    
    private async void StopButton_Click(object? sender, RoutedEventArgs e)
    {
        _cancellationTokenSource?.Cancel();
        LogStatus("Stopping and closing browser...");
        _stopButton.IsEnabled = false;
        await _automationService.CloseAsync();
        LogStatus("Browser closed.");
    }
    
    private void ClearLogButton_Click(object? sender, RoutedEventArgs e)
    {
        _statusLog.Text = string.Empty;
    }
    
    private async void SettingsButton_Click(object? sender, RoutedEventArgs e)
    {
        var settingsWindow = new SettingsWindow(_config);
        await settingsWindow.ShowDialog(this);
    }
    
    private async void ExitButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_isProcessing)
        {
            _cancellationTokenSource?.Cancel();
            await _automationService.CloseAsync();
        }
        Close();
    }
    
    private void OnStatusUpdated(object? sender, string message)
    {
        Dispatcher.UIThread.Post(() =>
        {
            LogStatus(message);
        });
    }
    
    private void OnStudentProcessed(object? sender, StudentRecord student)
    {
        Dispatcher.UIThread.Post(() =>
        {
            RefreshStudentGrid();
        });
    }
    
    private void LogStatus(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        _statusLog.Text += $"[{timestamp}] {message}\n";
        
        _statusLog.CaretIndex = _statusLog.Text?.Length ?? 0;
    }
    
    private void UpdateFooterStatus(string status)
    {
        _footerStatus.Text = status;
    }
    
    private void RefreshStudentGrid()
    {
        _studentGrid.ItemsSource = null;
        _studentGrid.ItemsSource = _students;
    }
    
    protected override async void OnClosing(WindowClosingEventArgs e)
    {
        if (_isProcessing)
        {
            _cancellationTokenSource?.Cancel();
            await _automationService.CloseAsync();
        }
        base.OnClosing(e);
    }
}

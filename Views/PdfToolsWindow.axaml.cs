// Views/PdfToolsWindow.axaml.cs

using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Dossier.Models;
using Dossier.Services;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;

namespace Dossier.Views;

public partial class PdfToolsWindow : Window
{
    private readonly IPdfRenameService _service = new PdfRenameService();
    private readonly IExcelService _excelService = new ExcelService();

    private List<StudentRecord> _students;
    private string _folderPath;
    private bool _hasRankingData;

    private readonly ObservableCollection<PdfRenamePreviewItem> _previewItems = new();
    private List<(string destPath, string originalPath)> _lastCompleted = new();

    // Controls
    private TextBlock _studentsLoadedText = null!;
    private TextBlock _folderPathText = null!;
    private TextBlock _pdfCountText = null!;
    private TextBlock _previewCountText = null!;
    private TextBlock _footerStatus = null!;
    private Button _loadSpreadsheetButton = null!;
    private Button _changeFolderButton = null!;
    private Button _previewButton = null!;
    private Button _renameAllButton = null!;
    private Button _undoButton = null!;
    private Button _appendRankingButton = null!;
    private Button _openFolderButton = null!;
    private Button _closeButton = null!;
    private DataGrid _previewGrid = null!;
    private TextBox _statusLog = null!;

    public PdfToolsWindow(List<StudentRecord>? students = null, string? defaultFolderPath = null)
    {
        _students = students ?? new List<StudentRecord>();

        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        _folderPath = defaultFolderPath ?? Path.Combine(desktop, "LATEST_BATCH");

        InitializeComponent();

        _previewGrid.ItemsSource = _previewItems;

        UpdateStudentsDisplay();
        RefreshFolderDisplay();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);

        _studentsLoadedText = this.FindControl<TextBlock>("StudentsLoadedText")!;
        _folderPathText = this.FindControl<TextBlock>("FolderPathText")!;
        _pdfCountText = this.FindControl<TextBlock>("PdfCountText")!;
        _previewCountText = this.FindControl<TextBlock>("PreviewCountText")!;
        _footerStatus = this.FindControl<TextBlock>("FooterStatus")!;
        _loadSpreadsheetButton = this.FindControl<Button>("LoadSpreadsheetButton")!;
        _changeFolderButton = this.FindControl<Button>("ChangeFolderButton")!;
        _previewButton = this.FindControl<Button>("PreviewButton")!;
        _renameAllButton = this.FindControl<Button>("RenameAllButton")!;
        _undoButton = this.FindControl<Button>("UndoButton")!;
        _appendRankingButton = this.FindControl<Button>("AppendRankingButton")!;
        _openFolderButton = this.FindControl<Button>("OpenFolderButton")!;
        _closeButton = this.FindControl<Button>("CloseButton")!;
        _previewGrid = this.FindControl<DataGrid>("PreviewGrid")!;
        _statusLog = this.FindControl<TextBox>("StatusLog")!;

        _loadSpreadsheetButton.Click += LoadSpreadsheetButton_Click;
        _changeFolderButton.Click += ChangeFolderButton_Click;
        _previewButton.Click += PreviewButton_Click;
        _renameAllButton.Click += RenameAllButton_Click;
        _undoButton.Click += UndoButton_Click;
        _appendRankingButton.Click += AppendRankingButton_Click;
        _openFolderButton.Click += (s, e) => { if (Directory.Exists(_folderPath)) OpenFolder(_folderPath); };
        _closeButton.Click += (s, e) => Close();
    }

    // ── Spreadsheet loading ────────────────────────────────────────────────

    private async void LoadSpreadsheetButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select Student Data Spreadsheet",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Spreadsheet") { Patterns = new[] { "*.xlsx", "*.xls", "*.csv" } }
            }
        });

        if (files.Count == 0) return;

        var path = files[0].Path.LocalPath;
        try
        {
            if (path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                _students = _excelService.LoadStudentsFromCsv(path);
            }
            else
            {
                var sheets = _excelService.GetSheetNames(path);
                var sheet = sheets.FirstOrDefault(s =>
                    s.Contains("Dept", StringComparison.OrdinalIgnoreCase) &&
                    s.Contains("tray", StringComparison.OrdinalIgnoreCase))
                    ?? sheets.FirstOrDefault()
                    ?? throw new Exception("No worksheets found.");
                _students = _excelService.LoadStudentsFromFile(path, sheet);
            }

            Log($"Loaded {_students.Count} students from {Path.GetFileName(path)}.");
            UpdateStudentsDisplay();
            ClearPreview();
            RefreshFolderDisplay();
        }
        catch (Exception ex)
        {
            Log($"ERROR loading spreadsheet: {ex.Message}");
        }
    }

    private void UpdateStudentsDisplay()
    {
        _hasRankingData = _students.Any(s => !string.IsNullOrWhiteSpace(s.ApplicationQualityRank));

        if (_students.Count == 0)
        {
            _studentsLoadedText.Text = "No spreadsheet loaded";
            _studentsLoadedText.Foreground = new SolidColorBrush(Color.Parse("#94A3B8"));
        }
        else
        {
            int rankCount = _students.Count(s => !string.IsNullOrWhiteSpace(s.ApplicationQualityRank));
            var dateRange = ComputeDateRange();
            var details = dateRange != null ? $"  •  {dateRange}" : "";
            var rankNote = rankCount > 0 ? $"  •  {rankCount} with ranking" : "";
            _studentsLoadedText.Text = $"{_students.Count} students{details}{rankNote}";
            _studentsLoadedText.Foreground = new SolidColorBrush(Color.Parse("#16A34A"));
        }
    }

    // ── Folder ─────────────────────────────────────────────────────────────

    private async void ChangeFolderButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Select folder containing PDF files",
            AllowMultiple = false
        });

        if (folders.Count == 0) return;

        _folderPath = folders[0].Path.LocalPath;
        ClearPreview();
        RefreshFolderDisplay();
        Log($"Folder: {_folderPath}");
    }

    private void RefreshFolderDisplay()
    {
        _folderPathText.Text = _folderPath;

        bool folderExists = Directory.Exists(_folderPath);
        int pdfCount = folderExists
            ? Directory.GetFiles(_folderPath, "*.pdf", SearchOption.TopDirectoryOnly).Length
            : 0;

        _pdfCountText.Text = folderExists
            ? $"{pdfCount} PDF file{(pdfCount == 1 ? "" : "s")} in folder"
            : "Folder does not exist yet";

        bool ready = folderExists && pdfCount > 0 && _students.Count > 0;
        _previewButton.IsEnabled = ready;
        _renameAllButton.IsEnabled = false;
        _appendRankingButton.IsEnabled = _hasRankingData && folderExists && pdfCount > 0;

        _footerStatus.Text = _students.Count == 0
            ? "Load a spreadsheet to begin."
            : !folderExists
                ? "Folder does not exist — run Merge Overview or change folder."
                : pdfCount == 0
                    ? "No PDFs found in folder."
                    : "Ready — click Preview or Rename All.";
    }

    // ── Preview ────────────────────────────────────────────────────────────

    private void PreviewButton_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            _previewItems.Clear();

            var items = _service.PreviewRenames(_folderPath, _students);
            foreach (var item in items)
                _previewItems.Add(item);

            int matchCount = items.Count(i => i.CanRename);
            int skipCount = items.Count - matchCount;

            _previewCountText.Text = $"{matchCount} to rename, {skipCount} skipped";
            _footerStatus.Text = $"{matchCount} files matched — click Rename All to proceed.";
            Log($"Preview: {matchCount} matches, {skipCount} no match.");

            _renameAllButton.IsEnabled = matchCount > 0;
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
        }
    }

    // ── Rename All (into batch folder) ─────────────────────────────────────

    private async void RenameAllButton_Click(object? sender, RoutedEventArgs e)
    {
        int matchCount = _previewItems.Count(i => i.CanRename);

        // Ask for batch number and confirm folder name
        var (batchNum, batchFolderName) = await ShowBatchNumberDialog();
        if (batchFolderName == null) return;

        bool confirmed = await ShowConfirmDialog("Rename All",
            $"Move {matchCount} files into:\n{Path.Combine(_folderPath, batchFolderName)}\n\nContinue?");
        if (!confirmed) return;

        try
        {
            _renameAllButton.IsEnabled = false;
            Log($"--- Renaming into '{batchFolderName}' ---");

            var (success, errors, completed) = _service.RenameIntoBatchFolder(
                _folderPath, _previewItems.ToList(), batchFolderName);

            _lastCompleted = completed;
            _undoButton.IsEnabled = completed.Count > 0;

            Log($"Done: {success} moved, {errors} errors.");
            Log($"Output: {Path.Combine(_folderPath, batchFolderName)}");
            _footerStatus.Text = $"Done: {success} renamed into '{batchFolderName}'.";

            if (success > 0)
            {
                bool open = await ShowConfirmDialog("Open Folder",
                    $"Open '{batchFolderName}'?");
                if (open) OpenFolder(Path.Combine(_folderPath, batchFolderName));
            }

            PreviewButton_Click(null, null!);
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
        }
    }

    // ── Undo ───────────────────────────────────────────────────────────────

    private async void UndoButton_Click(object? sender, RoutedEventArgs e)
    {
        bool confirmed = await ShowConfirmDialog("Undo",
            $"Move {_lastCompleted.Count} files back to their original locations?");
        if (!confirmed) return;

        int success = 0, errors = 0;
        foreach (var (destPath, originalPath) in _lastCompleted)
        {
            try
            {
                if (File.Exists(destPath) && !File.Exists(originalPath))
                {
                    File.Move(destPath, originalPath);
                    success++;
                }
            }
            catch { errors++; }
        }

        _lastCompleted.Clear();
        _undoButton.IsEnabled = false;
        Log($"Undo: {success} restored, {errors} errors.");
        _footerStatus.Text = $"Undo complete: {success} files restored.";

        PreviewButton_Click(null, null!);
    }

    // ── Append Ranking ─────────────────────────────────────────────────────

    private async void AppendRankingButton_Click(object? sender, RoutedEventArgs e)
    {
        int rankCount = _students.Count(s => !string.IsNullOrWhiteSpace(s.ApplicationQualityRank));
        int pdfCount = Directory.GetFiles(_folderPath, "*.pdf", SearchOption.TopDirectoryOnly).Length;
        string outputFolder = Path.Combine(_folderPath, "RankRenamed");

        bool confirmed = await ShowConfirmDialog("Append Ranking",
            $"Prefix filenames with ranking letter.\n\n" +
            $"Example:  A - b7 John Smith 12345678 H 2_1.pdf\n\n" +
            $"{rankCount} students have rankings  •  {pdfCount} PDFs in folder\n\n" +
            $"Output: {outputFolder}\nOriginals unchanged.");
        if (!confirmed) return;

        try
        {
            _appendRankingButton.IsEnabled = false;
            Log("--- Appending ranking prefixes ---");

            var (success, skipped, errors) = _service.AppendRanking(_folderPath, _students);

            Log($"Done: {success} copied, {skipped} skipped, {errors} errors.");
            if (success > 0) Log($"Output: {outputFolder}");
            _footerStatus.Text = $"Ranking: {success} copied, {skipped} skipped.";

            if (success > 0)
            {
                bool open = await ShowConfirmDialog("Open Folder", "Open the RankRenamed folder?");
                if (open) OpenFolder(outputFolder);
            }
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
            _appendRankingButton.IsEnabled = _hasRankingData;
        }
    }

    // ── Batch folder name helpers ──────────────────────────────────────────

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

    private string? ComputeDateRange()
    {
        var dates = _students
            .Where(s => s.ReceivedDate.HasValue)
            .Select(s => s.ReceivedDate!.Value.Date)
            .ToList();

        if (dates.Count == 0) return null;
        var start = dates.Min();
        var end = dates.Max();
        return start.Date == end.Date ? $"{start:MMM d}" : $"{start:MMM d} – {end:MMM d}";
    }

    // ── Dialogs ────────────────────────────────────────────────────────────

    private async Task<(int batchNum, string? folderName)> ShowBatchNumberDialog()
    {
        int resultNum = 0;
        string? resultFolder = null;

        var dialog = new Window
        {
            Title = "Batch Number",
            Width = 420,
            Height = 210,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false
        };

        var panel = new StackPanel { Margin = new Avalonia.Thickness(24), Spacing = 16 };

        panel.Children.Add(new TextBlock
        {
            Text = "Enter the batch number for this set of files:",
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.Parse("#2D3748"))
        });

        var inputRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };
        inputRow.Children.Add(new TextBlock
        {
            Text = "Batch Number:",
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = new SolidColorBrush(Color.Parse("#4A5568"))
        });

        var numBox = new TextBox { Width = 80, Text = "" };
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
                previewText.Text = $"Folder: {ComputeBatchFolderName(n)}";
            else
                previewText.Text = "Folder: —";
        };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 10
        };

        var ok = new Button { Content = "OK", Width = 80 };
        var cancel = new Button { Content = "Cancel", Width = 80 };

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

        await dialog.ShowDialog(this);
        return (resultNum, resultFolder);
    }

    private async Task<bool> ShowConfirmDialog(string title, string message)
    {
        bool result = false;
        var dialog = new Window
        {
            Title = title,
            Width = 460,
            Height = 230,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false
        };

        var panel = new StackPanel { Margin = new Avalonia.Thickness(24), Spacing = 20 };
        panel.Children.Add(new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap });

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            Spacing = 12
        };
        var yes = new Button { Content = "Yes", Width = 80 };
        var no = new Button { Content = "No", Width = 80 };
        yes.Click += (s, e) => { result = true; dialog.Close(); };
        no.Click += (s, e) => { result = false; dialog.Close(); };
        buttons.Children.Add(yes);
        buttons.Children.Add(no);
        panel.Children.Add(buttons);
        dialog.Content = panel;

        await dialog.ShowDialog(this);
        return result;
    }

    // ── Utilities ──────────────────────────────────────────────────────────

    private void ClearPreview()
    {
        _previewItems.Clear();
        _previewCountText.Text = "";
        _renameAllButton.IsEnabled = false;
        _undoButton.IsEnabled = false;
        _lastCompleted.Clear();
    }

    private void OpenFolder(string path)
    {
        try
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                System.Diagnostics.Process.Start("open", path);
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                System.Diagnostics.Process.Start("explorer.exe", path);
            else
                System.Diagnostics.Process.Start("xdg-open", path);
        }
        catch { }
    }

    private void Log(string message)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var ts = DateTime.Now.ToString("HH:mm:ss");
            _statusLog.Text += $"[{ts}] {message}\n";
            _statusLog.CaretIndex = _statusLog.Text?.Length ?? 0;
        });
    }
}

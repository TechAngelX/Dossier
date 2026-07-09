// Views/ProcessingWindow.axaml.cs

using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using Dossier.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;

namespace Dossier.Views;

public partial class ProcessingWindow : Window
{
    private ObservableCollection<ProcessingStudentViewModel> _students = new();
    private List<StudentRecord> _sourceRecords = new();
    private int _totalStudents;
    private int _processedCount;
    private System.Timers.Timer? _pulseTimer;
    private Stopwatch _elapsed = new();
    private System.Timers.Timer? _clockTimer;

    public ProcessingWindow()
    {
        InitializeComponent();
        StudentsList.ItemsSource = _students;

        CancelButton.Click += (s, e) => OnCancelRequested();
        CloseButton.Click += (s, e) => Close();
        OutputFailedButton.Click += OnOutputFailedClicked;
    }
    
    public event EventHandler? CancelRequested;

    private async void OnStudentNoClicked(object? sender, Avalonia.Input.PointerPressedEventArgs e)
    {
        if (sender is not TextBlock tb || tb.Tag is not string studentNo || string.IsNullOrEmpty(studentNo))
            return;

        var clipboard = Clipboard;
        if (clipboard is null)
            return;

        await clipboard.SetTextAsync(studentNo);

        // Brief visual confirmation
        var original = tb.Foreground;
        tb.Foreground = Avalonia.Media.Brushes.Green;
        await System.Threading.Tasks.Task.Delay(500);
        tb.Foreground = original;
    }

    private void OnOutputFailedClicked(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        var failedNumbers = _students
            .Where(s => s.StatusText == "Failed")
            .Select(s => s.StudentNo)
            .ToHashSet();

        if (failedNumbers.Count == 0)
            return;

        // Prefer full source records (Programme, ErrorMessage); fall back to the row view model.
        var rows = _sourceRecords
            .Where(r => failedNumbers.Contains(r.StudentNo))
            .Select(r => (r.StudentNo, r.Forename, r.Surname, r.Decision, r.Programme, r.ErrorMessage))
            .ToList();

        var covered = rows.Select(r => r.StudentNo).ToHashSet();
        foreach (var vm in _students.Where(s => s.StatusText == "Failed" && !covered.Contains(s.StudentNo)))
            rows.Add((vm.StudentNo, vm.Forename, vm.Surname, vm.Decision, "", ""));

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("StudentNo,Forename,Surname,Decision,Programme,ErrorMessage");
        foreach (var r in rows)
        {
            sb.AppendLine(string.Join(",", new[]
            {
                CsvEscape(r.StudentNo),
                CsvEscape(r.Forename),
                CsvEscape(r.Surname),
                CsvEscape(r.Decision),
                CsvEscape(r.Programme),
                CsvEscape(r.ErrorMessage)
            }));
        }

        try
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var fileName = $"Failed Records - {DateTime.Now:yyyy-MM-dd HHmm}.csv";
            var savePath = System.IO.Path.Combine(desktop, fileName);
            System.IO.File.WriteAllText(savePath, sb.ToString());

            LogMessage($"Saved {rows.Count} failed record(s) to {savePath}");
            FooterStatus.Text = $"Saved {rows.Count} failed record(s) to Desktop: {fileName}";
            RevealInFileManager(savePath);
        }
        catch (Exception ex)
        {
            LogMessage($"ERROR saving failed records: {ex.Message}");
            FooterStatus.Text = $"Could not save failed records: {ex.Message}";
        }
    }

    private static string CsvEscape(string? value)
    {
        var v = value ?? "";
        if (v.Contains(',') || v.Contains('"') || v.Contains('\n') || v.Contains('\r'))
            return "\"" + v.Replace("\"", "\"\"") + "\"";
        return v;
    }

    private static void RevealInFileManager(string path)
    {
        try
        {
            if (OperatingSystem.IsMacOS())
                Process.Start("open", new[] { "-R", path });
            else if (OperatingSystem.IsWindows())
                Process.Start("explorer.exe", $"/select,\"{path}\"");
            else if (OperatingSystem.IsLinux())
                Process.Start("xdg-open", System.IO.Path.GetDirectoryName(path) ?? path);
        }
        catch
        {
            // Non-fatal: the file is already saved.
        }
    }
    
    public void Initialize(List<StudentRecord> students)
    {
        _sourceRecords = students;
        _totalStudents = students.Count;
        _processedCount = 0;

        _students.Clear();
        
        foreach (var student in students)
        {
            _students.Add(new ProcessingStudentViewModel
            {
                StudentNo = student.StudentNo,
                Forename = student.Forename,
                Surname = student.Surname,
                Decision = student.Decision,
                StatusIcon = "⏳",
                StatusText = "Pending",
                StatusColor = "#A0AEC0"
            });
        }
        
        _elapsed.Restart();
        _clockTimer?.Stop();
        _clockTimer = new System.Timers.Timer(1000);
        _clockTimer.Elapsed += (s, e) => Dispatcher.UIThread.Post(UpdateFooterText);
        _clockTimer.Start();

        UpdateProgress();
    }

    public void ReorderStudents(List<StudentRecord> sorted)
    {
        Dispatcher.UIThread.Post(() =>
        {
            _students.Clear();
            foreach (var student in sorted)
            {
                _students.Add(new ProcessingStudentViewModel
                {
                    StudentNo = student.StudentNo,
                    Forename = student.Forename,
                    Surname = student.Surname,
                    Decision = student.Decision,
                    StatusIcon = student.Status == ProcessingStatus.Pending ? "\u23f3" : "\u2714",
                    StatusText = student.Status == ProcessingStatus.Pending ? "Pending" : student.Status.ToString(),
                    StatusColor = "#A0AEC0"
                });
            }
        });
    }

    public void UpdateStudentStatus(string studentNo, ProcessingStatus status, string? errorMessage = null)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var student = _students.FirstOrDefault(s => s.StudentNo == studentNo);
            if (student == null) return;
            
            switch (status)
            {
                case ProcessingStatus.Processing:
                    student.StatusIcon = "⚙";
                    student.StatusText = "Processing";
                    student.StatusColor = "#4299E1";
                    StartPulseAnimation(studentNo);
                    break;
                    
                case ProcessingStatus.Success:
                    student.StatusIcon = "✓";
                    student.StatusText = "Done";
                    student.StatusColor = "#48BB78";
                    StopPulseAnimation(studentNo);
                    _processedCount++;
                    break;
                    
                case ProcessingStatus.Failed:
                    student.StatusIcon = "✗";
                    student.StatusText = "Failed";
                    student.StatusColor = "#E53E3E";
                    StopPulseAnimation(studentNo);
                    _processedCount++;
                    break;

                case ProcessingStatus.Skipped:
                    student.StatusIcon = "⏭";
                    student.StatusText = "Skipped";
                    student.StatusColor = "#ED8936";
                    StopPulseAnimation(studentNo);
                    _processedCount++;
                    break;
            }
            
            UpdateProgress();
            ScrollToCurrentStudent(studentNo);
        });
    }
    
    private void StartPulseAnimation(string studentNo)
    {
        _pulseTimer?.Stop();
        _pulseTimer = new System.Timers.Timer(500);
        var pulse = true;
        
        _pulseTimer.Elapsed += (s, e) =>
        {
            Dispatcher.UIThread.Post(() =>
            {
                var student = _students.FirstOrDefault(s => s.StudentNo == studentNo);
                if (student != null && student.StatusText == "Processing")
                {
                    student.StatusColor = pulse ? "#4299E1" : "#63B3ED";
                    pulse = !pulse;
                }
            });
        };
        
        _pulseTimer.Start();
    }
    
    private void StopPulseAnimation(string studentNo)
    {
        _pulseTimer?.Stop();
        _pulseTimer = null;
    }
    
    public void LogMessage(string message)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            StatusLog.Text += $"[{timestamp}] {message}\n";
            StatusLog.CaretIndex = StatusLog.Text?.Length ?? 0;
        });
    }
    
    public void UpdateFooterStatus(string status)
    {
        Dispatcher.UIThread.Post(() =>
        {
            FooterStatus.Text = status;
        });
    }
    
    public void ProcessingComplete()
    {
        _clockTimer?.Stop();
        _elapsed.Stop();

        Dispatcher.UIThread.Post(() =>
        {
            var successCount = _students.Count(s => s.StatusText == "Done");
            var failedCount = _students.Count(s => s.StatusText == "Failed");
            var skippedCount = _students.Count(s => s.StatusText == "Skipped");
            var time = FormatElapsed(_elapsed.Elapsed);

            var summary = skippedCount > 0
                ? $"Complete: {successCount} successful, {failedCount} failed, {skippedCount} skipped"
                : $"Complete: {successCount} successful, {failedCount} failed";

            SubtitleText.Text = summary;
            FooterStatus.Text = $"Done in {time}  —  {summary}";

            CancelButton.IsVisible = false;
            CloseButton.IsVisible = true;
            OutputFailedButton.IsVisible = failedCount > 0;
        });
    }
    
    private void UpdateProgress()
    {
        var percentage = _totalStudents > 0 ? (_processedCount * 100.0 / _totalStudents) : 0;
        ProgressBar.Value = percentage;
        ProgressText.Text = $"{_processedCount} / {_totalStudents} ({percentage:F0}%)";
        UpdateFooterText();
    }

    private string FormatElapsed(TimeSpan ts)
    {
        return ts.TotalHours >= 1
            ? $"{(int)ts.TotalHours}h {ts.Minutes:D2}m {ts.Seconds:D2}s"
            : ts.TotalMinutes >= 1
                ? $"{(int)ts.TotalMinutes}m {ts.Seconds:D2}s"
                : $"{ts.Seconds}s";
    }

    private void UpdateFooterText()
    {
        var time = FormatElapsed(_elapsed.Elapsed);
        var remaining = _totalStudents - _processedCount;

        if (_processedCount > 0 && remaining > 0)
            FooterStatus.Text = $"Processed {_processedCount} of {_totalStudents}  —  {remaining} remaining  —  {time}";
        else if (_processedCount > 0 && remaining == 0)
            FooterStatus.Text = $"All {_totalStudents} records processed  —  {time} total";
        else
            FooterStatus.Text = $"Starting...  —  {time}";
    }
    
    private void ScrollToCurrentStudent(string studentNo)
    {
        var index = _students.ToList().FindIndex(s => s.StudentNo == studentNo);
        if (index >= 0)
        {
            var itemHeight = 65.0;
            var viewportHeight = StudentsScrollViewer.Viewport.Height;
            var scrollPosition = (index * itemHeight) - (viewportHeight / 2) + (itemHeight / 2);
            scrollPosition = Math.Max(0, scrollPosition);
            
            var currentOffset = StudentsScrollViewer.Offset.Y;
            var targetOffset = scrollPosition;
            var steps = 20;
            var stepSize = (targetOffset - currentOffset) / steps;
            
            var scrollTimer = new System.Timers.Timer(10);
            var currentStep = 0;
            
            scrollTimer.Elapsed += (s, e) =>
            {
                if (currentStep >= steps)
                {
                    scrollTimer.Stop();
                    scrollTimer.Dispose();
                    return;
                }
                
                Dispatcher.UIThread.Post(() =>
                {
                    currentOffset += stepSize;
                    StudentsScrollViewer.Offset = new Avalonia.Vector(0, currentOffset);
                    currentStep++;
                });
            };
            
            scrollTimer.Start();
        }
    }
    
    private void OnCancelRequested()
    {
        CancelRequested?.Invoke(this, EventArgs.Empty);
        Close();
    }
}

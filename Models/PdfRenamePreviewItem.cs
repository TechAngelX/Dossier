using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Dossier.Models;

public class PdfRenamePreviewItem : INotifyPropertyChanged
{
    private string _currentFilename = "";
    private string _newFilename = "";
    private string _status = "";
    private string _statusColor = "#64748B";

    public string OriginalPath { get; set; } = "";
    public bool CanRename { get; set; }

    public string CurrentFilename
    {
        get => _currentFilename;
        set { _currentFilename = value; OnPropertyChanged(); }
    }

    public string NewFilename
    {
        get => _newFilename;
        set { _newFilename = value; OnPropertyChanged(); }
    }

    public string Status
    {
        get => _status;
        set { _status = value; OnPropertyChanged(); }
    }

    public string StatusColor
    {
        get => _statusColor;
        set { _statusColor = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

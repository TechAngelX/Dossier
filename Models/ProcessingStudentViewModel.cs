using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Dossier.Models;

public class ProcessingStudentViewModel : INotifyPropertyChanged
{
    private string _studentNo = "";
    private string _forename = "";
    private string _surname = "";
    private string _decision = "";
    private string _statusIcon = "⏳";
    private string _statusText = "Pending";
    private string _statusColor = "#A0AEC0";
    
    public string StudentNo
    {
        get => _studentNo;
        set { _studentNo = value; OnPropertyChanged(); }
    }
    
    public string Forename
    {
        get => _forename;
        set { _forename = value; OnPropertyChanged(); }
    }
    
    public string Surname
    {
        get => _surname;
        set { _surname = value; OnPropertyChanged(); }
    }
    
    public string Decision
    {
        get => _decision;
        set { _decision = value; OnPropertyChanged(); }
    }
    
    public string StatusIcon
    {
        get => _statusIcon;
        set { _statusIcon = value; OnPropertyChanged(); }
    }
    
    public string StatusText
    {
        get => _statusText;
        set { _statusText = value; OnPropertyChanged(); }
    }
    
    public string StatusColor
    {
        get => _statusColor;
        set { _statusColor = value; OnPropertyChanged(); }
    }
    
    public event PropertyChangedEventHandler? PropertyChanged;
    
    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

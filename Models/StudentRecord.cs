// Models/StudentRecord.cs

namespace Dossier.Models;

public class StudentRecord
{
    public string StudentNo { get; set; } = string.Empty;
    public string Decision { get; set; } = string.Empty;
    public string Forename { get; set; } = "";
    public string Surname { get; set; } = "";
    public string Programme { get; set; } = string.Empty;
    public DateTime? ReceivedDate { get; set; }
    public DateTime? DueDate { get; set; }
    public DateTime? DateOfBirth { get; set; }
    public ProcessingStatus Status { get; set; } = ProcessingStatus.Pending;
    public string ErrorMessage { get; set; } = string.Empty;
    public string Name => $"{Forename} {Surname}".Trim();
    public string ReceivedDateDisplay => ReceivedDate?.ToString("dd/MM/yyyy") ?? "";
    public string DueDateDisplay => DueDate?.ToString("dd/MM/yyyy") ?? "";
    public string DateOfBirthDisplay => DateOfBirth?.ToString("dd/MM/yyyy") ?? "";
}

public enum ProcessingStatus
{
    Pending,
    Processing,
    Success,
    Failed,
    Skipped
}

public enum DecisionType
{
    Accept,
    Reject,
    Unknown
}

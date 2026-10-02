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

    // PDF rename fields (from PDFusion integration)
    public string Batch { get; set; } = string.Empty;
    public string FeeStatus { get; set; } = string.Empty;
    public string UKGrade { get; set; } = string.Empty;
    public string ApplicationQualityRank { get; set; } = string.Empty;

    // Personal Tutor Allocation (PAT tab) — the tutor code to assign in Portico,
    // read from the "2026-27 PAT (PIVOT)" column (e.g. "MHERB70").
    public string PersonalTutor { get; set; } = string.Empty;

    // PAT tab — the "PAT Required" column. An "N" here means skip this record.
    public string PatRequired { get; set; } = string.Empty;
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

// Models/StudentRecord.cs

namespace Playwrighter.Models;

public class StudentRecord
{
    public string StudentNo { get; set; } = string.Empty;
    public string Decision { get; set; } = string.Empty;
    public string Forename { get; set; } = ""; 
    public string Surname { get; set; } = "";  
    public string Programme { get; set; } = string.Empty;
    public ProcessingStatus Status { get; set; } = ProcessingStatus.Pending;
    public string ErrorMessage { get; set; } = string.Empty;
    public string Name => $"{Forename} {Surname}".Trim(); 

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

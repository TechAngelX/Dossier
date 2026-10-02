// Services/IPorticoAutomationService.cs

using Dossier.Models;

namespace Dossier.Services;

public interface IPorticoAutomationService
{
    event EventHandler<string>? StatusUpdated;
    event EventHandler<StudentRecord>? StudentProcessed;
    
    bool DebugMode { get; set; }
    bool IsInitialised { get; }

    // When set (e.g. "2026/27"), each student search first selects that academic year
    // in the Portico "Year" dropdown instead of the default "Current Applications".
    // Left null by the main Portico tab so its behaviour is unchanged.
    string? TargetAcademicYear { get; set; }
    
    Task InitialiseAsync(AppConfig config);
    Task<bool> LoginAsync();
    Task<bool> NavigateToUclSelectAsync();
    Task ProcessStudentAcceptAsync(StudentRecord student);
    Task ProcessStudentRejectAsync(StudentRecord student);
    Task ProcessStudentMergeOverviewAsync(StudentRecord student, string downloadPath);

    // PAT (Personal Tutor Allocations) flow.
    // NavigateToPersonalTutorAllocationsAsync opens the allocations page from the My Portico home;
    // ProcessStudentPatAsync enters one student's number + tutor code and clicks "Apply New Criteria".
    Task<bool> NavigateToPersonalTutorAllocationsAsync(bool unassignedOnly = false);
    Task ProcessStudentPatAsync(StudentRecord student);
    Task<string> DownloadDepartmentReportAsync(string fullProgrammeName, string downloadDir);
    Task<string> DownloadIndividualStudentOverviewCsvAsync(string studentNumber, string downloadDir);
    Task CloseAsync(bool handOffToUser = false);
}

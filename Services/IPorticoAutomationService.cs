// Services/IPorticoAutomationService.cs

using Playwrighter.Models;

namespace Playwrighter.Services;

public interface IPorticoAutomationService
{
    event EventHandler<string>? StatusUpdated;
    event EventHandler<StudentRecord>? StudentProcessed;
    
    bool DebugMode { get; set; }
    bool IsInitialised { get; }
    
    Task InitialiseAsync(AppConfig config);
    Task<bool> LoginAsync();
    Task<bool> NavigateToUclSelectAsync();
    Task ProcessStudentAcceptAsync(StudentRecord student);
    Task ProcessStudentRejectAsync(StudentRecord student);
    Task ProcessStudentMergeOverviewAsync(StudentRecord student, string downloadPath);
    Task CloseAsync();
}

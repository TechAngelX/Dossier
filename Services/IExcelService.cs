// Services/IExcelService.cs

using Playwrighter.Models;

namespace Playwrighter.Services;

public interface IExcelService
{
    List<StudentRecord> LoadStudentsFromFile(string filePath, string sheetName = "Dept In-tray");
    List<StudentRecord> LoadStudentsFromCsv(string filePath);
    List<string> GetSheetNames(string filePath);
}

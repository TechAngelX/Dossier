// Services/IExcelService.cs

using Dossier.Models;

namespace Dossier.Services;

public interface IExcelService
{
    List<StudentRecord> LoadStudentsFromFile(string filePath, string sheetName = "Dept In-tray");
    List<StudentRecord> LoadStudentsFromCsv(string filePath);
    List<string> GetSheetNames(string filePath);
}

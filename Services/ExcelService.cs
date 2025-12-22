// Services/ExcelService.cs

using OfficeOpenXml;
using Playwrighter.Models;

namespace Playwrighter.Services;

public class ExcelService : IExcelService
{
    public List<string> GetSheetNames(string filePath)
    {
        var sheetNames = new List<string>();
        using var package = new ExcelPackage(new FileInfo(filePath));
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            sheetNames.Add(worksheet.Name);
        }
        return sheetNames;
    }

    public List<StudentRecord> LoadStudentsFromFile(string filePath, string sheetName = "Dept In-tray")
    {
        var students = new List<StudentRecord>();
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[sheetName];
        if (worksheet == null)
        {
            worksheet = package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet == null) throw new InvalidOperationException("No worksheets found.");
        }
        
        int studentNoCol = -1;
        int decisionCol = -1;
        int nameCol = -1;
        int forenameCol = -1;
        int surnameCol = -1;
        int programmeCol = -1;
        
        int headerRow = 1;
        int colCount = worksheet.Dimension?.Columns ?? 0;
        
        Console.WriteLine("=== Excel Column Headers ===");
        for (int col = 1; col <= colCount; col++)
        {
            string header = worksheet.Cells[headerRow, col].Value?.ToString()?.Trim() ?? "";
            Console.WriteLine($"  Column {col}: '{header}'");
        }
        
        for (int col = 1; col <= colCount; col++)
        {
            string header = worksheet.Cells[headerRow, col].Value?.ToString()?.Trim().ToLowerInvariant() ?? "";
            
            if (string.IsNullOrEmpty(header)) continue;

            if (studentNoCol == -1 && (header == "studentno" || header == "student_no" || header == "student number" || header == "student id" || header == "id"))
                studentNoCol = col;
            
            if (decisionCol == -1 && (header == "decision" || header == "status" || header == "offer"))
                decisionCol = col;
            
            if (nameCol == -1 && (header == "name" || header == "applicant name" || header == "student name"))
                nameCol = col;
                
            if (forenameCol == -1 && (header == "forename" || header == "firstname" || header == "first name"))
                forenameCol = col;
            
            if (surnameCol == -1 && (header == "surname" || header == "lastname" || header == "last name"))
                surnameCol = col;

            if (programmeCol == -1 && header == "programme")
                programmeCol = col;
        }

        for (int col = 1; col <= colCount; col++)
        {
            string header = worksheet.Cells[headerRow, col].Value?.ToString()?.Trim().ToLowerInvariant() ?? "";
            if (string.IsNullOrEmpty(header)) continue;

            if (studentNoCol == -1 && header.Contains("student") && header.Contains("no")) studentNoCol = col;
            if (decisionCol == -1 && header.Contains("decision")) decisionCol = col;
            
            if (programmeCol == -1 && (header == "prog" || header == "progcode" || header == "prog code" || header == "progshort" || header == "route"))
                programmeCol = col;
            
            if (forenameCol == -1 && header.Contains("forename")) forenameCol = col;
            if (surnameCol == -1 && header.Contains("surname")) surnameCol = col;
        }
        
        Console.WriteLine($"=== Detected Columns ===");
        Console.WriteLine($"  StudentNo column: {studentNoCol}");
        Console.WriteLine($"  Decision column: {decisionCol}");
        
        if (nameCol > 0) Console.WriteLine($"  Name column: {nameCol}");
        if (forenameCol > 0) Console.WriteLine($"  Forename column: {forenameCol}");
        if (surnameCol > 0) Console.WriteLine($"  Surname column: {surnameCol}");
        
        Console.WriteLine($"  Programme column: {programmeCol}");
        
        if (studentNoCol == -1) throw new InvalidOperationException("Could not find StudentNo column.");
        if (decisionCol == -1) throw new InvalidOperationException("Could not find Decision column.");
        if (programmeCol == -1)
        {
            Console.WriteLine("WARNING: Programme column not found! Row matching may fail.");
        }

        int rowCount = worksheet.Dimension?.Rows ?? 0;
        for (int row = headerRow + 1; row <= rowCount; row++)
        {
            string? studentNo = worksheet.Cells[row, studentNoCol].Value?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(studentNo)) continue;

            string programmeValue = programmeCol > 0 ? worksheet.Cells[row, programmeCol].Value?.ToString()?.Trim() ?? "" : "";
            
            string forename = "";
            string surname = "";
            
            if (nameCol > 0)
            {
                var fullName = worksheet.Cells[row, nameCol].Value?.ToString()?.Trim() ?? "";
                var nameParts = fullName.Split(' ', 2);
                forename = nameParts.Length > 0 ? nameParts[0] : "";
                surname = nameParts.Length > 1 ? nameParts[1] : "";
            }
            else
            {
                forename = forenameCol > 0 ? worksheet.Cells[row, forenameCol].Value?.ToString()?.Trim() ?? "" : "";
                surname = surnameCol > 0 ? worksheet.Cells[row, surnameCol].Value?.ToString()?.Trim() ?? "" : "";
            }

            var record = new StudentRecord
            {
                StudentNo = studentNo,
                Decision = worksheet.Cells[row, decisionCol].Value?.ToString()?.Trim() ?? "",
                Forename = forename,
                Surname = surname,
                Programme = programmeValue,
                Status = ProcessingStatus.Pending
            };

            if (students.Count < 5)
            {
                Console.WriteLine($"  Loaded: StudentNo={record.StudentNo}, Decision={record.Decision}, Name={record.Forename} {record.Surname}, Programme='{record.Programme}'");
            }

            students.Add(record);
        }
        
        Console.WriteLine($"=== Total loaded: {students.Count} students ===");
        return students;
    }
}

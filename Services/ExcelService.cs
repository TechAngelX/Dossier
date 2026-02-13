// Services/ExcelService.cs

using OfficeOpenXml;
using Dossier.Models;

namespace Dossier.Services;

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
        if (decisionCol == -1)
        {
            Console.WriteLine("WARNING: Decision column not found. Accept/Reject processing won't work, but Merge Overview will.");
        }
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
                Decision = decisionCol > 0 ? worksheet.Cells[row, decisionCol].Value?.ToString()?.Trim() ?? "" : "",
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

    public List<StudentRecord> LoadStudentsFromCsv(string filePath)
    {
        var students = new List<StudentRecord>();
        var lines = File.ReadAllLines(filePath);

        if (lines.Length == 0) throw new InvalidOperationException("CSV file is empty.");

        // Parse header row - detect delimiter (comma or tab)
        var delimiter = lines[0].Contains('\t') ? '\t' : ',';
        var headers = ParseCsvLine(lines[0], delimiter);

        Console.WriteLine("=== CSV Column Headers ===");
        for (int col = 0; col < headers.Length; col++)
        {
            Console.WriteLine($"  Column {col}: '{headers[col]}'");
        }

        int studentNoCol = -1;
        int decisionCol = -1;
        int nameCol = -1;
        int forenameCol = -1;
        int surnameCol = -1;
        int programmeCol = -1;

        // Pass 1: exact match
        for (int col = 0; col < headers.Length; col++)
        {
            string header = headers[col].Trim().ToLowerInvariant();
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

        // Pass 2: fuzzy match
        for (int col = 0; col < headers.Length; col++)
        {
            string header = headers[col].Trim().ToLowerInvariant();
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
        if (nameCol >= 0) Console.WriteLine($"  Name column: {nameCol}");
        if (forenameCol >= 0) Console.WriteLine($"  Forename column: {forenameCol}");
        if (surnameCol >= 0) Console.WriteLine($"  Surname column: {surnameCol}");
        Console.WriteLine($"  Programme column: {programmeCol}");

        if (studentNoCol == -1) throw new InvalidOperationException("Could not find StudentNo column in CSV.");
        if (decisionCol == -1)
        {
            Console.WriteLine("WARNING: Decision column not found in CSV. Accept/Reject processing won't work, but Merge Overview will.");
        }
        if (programmeCol == -1)
        {
            Console.WriteLine("WARNING: Programme column not found! Row matching may fail.");
        }

        // Parse data rows
        for (int row = 1; row < lines.Length; row++)
        {
            if (string.IsNullOrWhiteSpace(lines[row])) continue;

            var fields = ParseCsvLine(lines[row], delimiter);

            string studentNo = studentNoCol < fields.Length ? fields[studentNoCol].Trim() : "";
            if (string.IsNullOrWhiteSpace(studentNo)) continue;

            string programmeValue = programmeCol >= 0 && programmeCol < fields.Length
                ? fields[programmeCol].Trim() : "";

            string forename = "";
            string surname = "";

            if (nameCol >= 0 && nameCol < fields.Length)
            {
                var fullName = fields[nameCol].Trim();
                var nameParts = fullName.Split(' ', 2);
                forename = nameParts.Length > 0 ? nameParts[0] : "";
                surname = nameParts.Length > 1 ? nameParts[1] : "";
            }
            else
            {
                forename = forenameCol >= 0 && forenameCol < fields.Length ? fields[forenameCol].Trim() : "";
                surname = surnameCol >= 0 && surnameCol < fields.Length ? fields[surnameCol].Trim() : "";
            }

            var record = new StudentRecord
            {
                StudentNo = studentNo,
                Decision = decisionCol >= 0 && decisionCol < fields.Length ? fields[decisionCol].Trim() : "",
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

    private static string[] ParseCsvLine(string line, char delimiter)
    {
        var fields = new List<string>();
        bool inQuotes = false;
        var current = new System.Text.StringBuilder();

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    current.Append(c);
                }
            }
            else
            {
                if (c == '"')
                {
                    inQuotes = true;
                }
                else if (c == delimiter)
                {
                    fields.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
        }

        fields.Add(current.ToString());
        return fields.ToArray();
    }
}

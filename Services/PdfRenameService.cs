using System.Globalization;
using Dossier.Models;

namespace Dossier.Services;

public class PdfRenameService : IPdfRenameService
{
    public List<PdfRenamePreviewItem> PreviewRenames(string folderPath, IEnumerable<StudentRecord> students)
    {
        var items = new List<PdfRenamePreviewItem>();
        var lookup = students
            .GroupBy(s => s.StudentNo, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var pdfFiles = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly)
                                .OrderBy(f => f);

        foreach (var filePath in pdfFiles)
        {
            string filename = Path.GetFileName(filePath);
            string studentNo = ExtractStudentNumber(filename);

            if (string.IsNullOrEmpty(studentNo))
            {
                items.Add(new PdfRenamePreviewItem
                {
                    OriginalPath = filePath,
                    CurrentFilename = filename,
                    NewFilename = "(no change)",
                    Status = "No student number found",
                    StatusColor = "#94A3B8",
                    CanRename = false
                });
                continue;
            }

            if (lookup.TryGetValue(studentNo, out var student))
            {
                items.Add(new PdfRenamePreviewItem
                {
                    OriginalPath = filePath,
                    CurrentFilename = filename,
                    NewFilename = GenerateNewFilename(student),
                    Status = $"{student.Forename} {student.Surname}".Trim(),
                    StatusColor = "#059669",
                    CanRename = true
                });
            }
            else
            {
                items.Add(new PdfRenamePreviewItem
                {
                    OriginalPath = filePath,
                    CurrentFilename = filename,
                    NewFilename = "(no change)",
                    Status = $"No record for {studentNo}",
                    StatusColor = "#DC2626",
                    CanRename = false
                });
            }
        }

        return items;
    }

    public (int success, int errors, List<(string newPath, string originalPath)> completed) RenameInPlace(
        string folderPath, IReadOnlyList<PdfRenamePreviewItem> items)
    {
        int success = 0, errors = 0;
        var completed = new List<(string newPath, string originalPath)>();

        foreach (var item in items.Where(i => i.CanRename))
        {
            var newPath = Path.Combine(folderPath, item.NewFilename);
            if (string.Equals(item.OriginalPath, newPath, StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                if (File.Exists(newPath)) { errors++; continue; }
                File.Move(item.OriginalPath, newPath);
                completed.Add((newPath, item.OriginalPath));
                success++;
            }
            catch
            {
                errors++;
            }
        }

        return (success, errors, completed);
    }

    public (int success, int errors, List<(string destPath, string originalPath)> completed) RenameIntoBatchFolder(
        string folderPath, IReadOnlyList<PdfRenamePreviewItem> items, string batchFolderName)
    {
        var batchFolder = Path.Combine(folderPath, batchFolderName);
        Directory.CreateDirectory(batchFolder);

        int success = 0, errors = 0;
        var completed = new List<(string destPath, string originalPath)>();

        foreach (var item in items.Where(i => i.CanRename))
        {
            var destPath = Path.Combine(batchFolder, item.NewFilename);
            try
            {
                if (File.Exists(destPath)) { errors++; continue; }
                File.Move(item.OriginalPath, destPath);
                completed.Add((destPath, item.OriginalPath));
                success++;
            }
            catch
            {
                errors++;
            }
        }

        return (success, errors, completed);
    }

    public (int success, int skipped, int errors) AppendRanking(
        string folderPath, IEnumerable<StudentRecord> students)
    {
        var rankLookup = students
            .Where(s => !string.IsNullOrWhiteSpace(s.ApplicationQualityRank))
            .GroupBy(s => s.StudentNo, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().ApplicationQualityRank.Trim(), StringComparer.OrdinalIgnoreCase);

        var outputFolder = Path.Combine(folderPath, "RankRenamed");
        Directory.CreateDirectory(outputFolder);

        int success = 0, skipped = 0, errors = 0;
        var pdfFiles = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly);

        foreach (var filePath in pdfFiles)
        {
            string filename = Path.GetFileName(filePath);
            string studentNo = ExtractStudentNumber(filename);

            if (string.IsNullOrEmpty(studentNo) || !rankLookup.TryGetValue(studentNo, out var ranking))
            {
                skipped++;
                continue;
            }

            // Skip already-prefixed files: single letter + " - " not starting with "b"
            if (filename.Length > 4 && filename[1] == ' ' && filename[2] == '-' && filename[3] == ' '
                && char.IsLetter(filename[0]) && !filename.StartsWith("b", StringComparison.OrdinalIgnoreCase))
            {
                skipped++;
                continue;
            }

            string newFilename = $"{ranking} - {filename}";
            string newPath = Path.Combine(outputFolder, newFilename);

            try
            {
                if (File.Exists(newPath)) { skipped++; continue; }
                File.Copy(filePath, newPath);
                success++;
            }
            catch
            {
                errors++;
            }
        }

        return (success, skipped, errors);
    }

    public string GenerateNewFilename(StudentRecord student)
    {
        string batchNum = ExtractBatchNumber(student.Batch);
        string forename = ToProperCase(student.Forename);
        string surname = ToProperCase(student.Surname);
        string feeCode = DetermineFeeStatusCode(student.FeeStatus);
        string gradeCode = FormatUKGrade(student.UKGrade);

        string filename = $"b{batchNum} {forename} {surname} {student.StudentNo} {feeCode} {gradeCode}.pdf";

        foreach (char c in Path.GetInvalidFileNameChars())
            filename = filename.Replace(c, '_');

        return filename;
    }

    private static string ExtractStudentNumber(string filename)
    {
        // Format 1: XXXXXXXX-XX-XX-OVERVIEW.PDF (digits before first dash)
        var dashParts = filename.Split('-');
        if (dashParts.Length >= 1)
        {
            string firstPart = dashParts[0].Trim();
            if (firstPart.All(char.IsDigit) && firstPart.Length >= 7 && firstPart.Length <= 10)
                return firstPart;
        }

        // Format 2: b1 John Smith 26049530 H 2_1.pdf (longest 7-10 digit space-delimited token)
        string? best = null;
        foreach (var part in filename.Split(' '))
        {
            string clean = part.Replace(".pdf", "", StringComparison.OrdinalIgnoreCase);
            if (clean.Length >= 7 && clean.Length <= 10 && clean.All(char.IsDigit))
            {
                if (best == null || clean.Length > best.Length)
                    best = clean;
            }
        }
        return best ?? "";
    }

    private static string ExtractBatchNumber(string batch)
    {
        if (string.IsNullOrWhiteSpace(batch)) return "0";
        var digits = new string(batch.Where(char.IsDigit).ToArray());
        return string.IsNullOrEmpty(digits) ? "0" : digits;
    }

    private static string ToProperCase(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text.ToLower());
    }

    private static string DetermineFeeStatusCode(string feeStatus)
    {
        if (string.IsNullOrWhiteSpace(feeStatus)) return "?";
        string lower = feeStatus.ToLower();
        if (lower.Contains("european") || lower.Contains("overseas")) return "OS";
        if (lower.Contains("home")) return "H";
        return "?";
    }

    private static string FormatUKGrade(string ukGrade)
    {
        if (string.IsNullOrWhiteSpace(ukGrade)) return "XX";
        string lower = ukGrade.Trim().ToLower();
        if (lower.Contains("2.1") || lower.Contains("2:1")) return "2_1";
        if (lower.Contains("2.2") || lower.Contains("2:2")) return "2_2";
        if (lower.Contains("1st") || lower == "1" || lower.Contains("first")) return "1";
        if (lower.Contains("3rd") || lower == "3" || lower.Contains("third")) return "3";
        if (lower.Contains("master")) return "Masters";
        return "XX";
    }
}

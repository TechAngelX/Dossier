using Xunit;
using Dossier.Models;
using Dossier.Services;
using OfficeOpenXml;

namespace Dossier.Tests.Services;

public class ExcelServiceExcelTests : IDisposable
{
    private readonly ExcelService _service = new();
    private readonly List<string> _tempFiles = new();

    public ExcelServiceExcelTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    private string CreateTempXlsx(Action<ExcelPackage> configure)
    {
        var path = Path.Combine(Path.GetTempPath(), $"dossier_test_{Guid.NewGuid()}.xlsx");
        using var package = new ExcelPackage();
        configure(package);
        package.SaveAs(new FileInfo(path));
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }

    [Fact]
    public void StandardXlsx_ParsesAllColumns()
    {
        var path = CreateTempXlsx(pkg =>
        {
            var ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "StudentNo";
            ws.Cells[1, 2].Value = "Decision";
            ws.Cells[1, 3].Value = "Forename";
            ws.Cells[1, 4].Value = "Surname";
            ws.Cells[1, 5].Value = "Programme";

            ws.Cells[2, 1].Value = "12345";
            ws.Cells[2, 2].Value = "Accept";
            ws.Cells[2, 3].Value = "John";
            ws.Cells[2, 4].Value = "Smith";
            ws.Cells[2, 5].Value = "ML";

            ws.Cells[3, 1].Value = "67890";
            ws.Cells[3, 2].Value = "Reject";
            ws.Cells[3, 3].Value = "Jane";
            ws.Cells[3, 4].Value = "Doe";
            ws.Cells[3, 5].Value = "CS";
        });

        var students = _service.LoadStudentsFromFile(path, "Sheet1");

        Assert.Equal(2, students.Count);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("Accept", students[0].Decision);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Smith", students[0].Surname);
        Assert.Equal("ML", students[0].Programme);

        Assert.Equal("67890", students[1].StudentNo);
        Assert.Equal("Reject", students[1].Decision);
    }

    [Fact]
    public void NamedSheet_NotFound_FallsBackToFirst()
    {
        var path = CreateTempXlsx(pkg =>
        {
            var ws = pkg.Workbook.Worksheets.Add("MyData");
            ws.Cells[1, 1].Value = "StudentNo";
            ws.Cells[2, 1].Value = "12345";
        });

        // Request "Dept In-tray" which doesn't exist — should fall back to "MyData"
        var students = _service.LoadStudentsFromFile(path, "Dept In-tray");

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
    }

    [Fact]
    public void FuzzyColumnHeaders_DetectedCorrectly()
    {
        var path = CreateTempXlsx(pkg =>
        {
            var ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Student Number";
            ws.Cells[1, 2].Value = "Status";
            ws.Cells[1, 3].Value = "First Name";
            ws.Cells[1, 4].Value = "Last Name";

            ws.Cells[2, 1].Value = "12345";
            ws.Cells[2, 2].Value = "Accept";
            ws.Cells[2, 3].Value = "John";
            ws.Cells[2, 4].Value = "Smith";
        });

        var students = _service.LoadStudentsFromFile(path, "Sheet1");

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("Accept", students[0].Decision);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Smith", students[0].Surname);
    }

    [Fact]
    public void MissingStudentNoColumn_ThrowsInvalidOperation()
    {
        var path = CreateTempXlsx(pkg =>
        {
            var ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Name";
            ws.Cells[1, 2].Value = "Decision";
            ws.Cells[2, 1].Value = "John Smith";
            ws.Cells[2, 2].Value = "Accept";
        });

        Assert.Throws<InvalidOperationException>(() =>
            _service.LoadStudentsFromFile(path, "Sheet1"));
    }

    [Fact]
    public void MissingDecisionColumn_LoadsSuccessfully()
    {
        var path = CreateTempXlsx(pkg =>
        {
            var ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "StudentNo";
            ws.Cells[1, 2].Value = "Forename";
            ws.Cells[2, 1].Value = "12345";
            ws.Cells[2, 2].Value = "John";
        });

        var students = _service.LoadStudentsFromFile(path, "Sheet1");

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("", students[0].Decision);
    }

    [Fact]
    public void GetSheetNames_ReturnsAllSheets()
    {
        var path = CreateTempXlsx(pkg =>
        {
            pkg.Workbook.Worksheets.Add("Sheet1");
            pkg.Workbook.Worksheets.Add("Dept In-tray");
            pkg.Workbook.Worksheets.Add("Summary");
        });

        var names = _service.GetSheetNames(path);

        Assert.Equal(3, names.Count);
        Assert.Contains("Sheet1", names);
        Assert.Contains("Dept In-tray", names);
        Assert.Contains("Summary", names);
    }

    [Fact]
    public void NameColumn_SplitsIntoForenameAndSurname()
    {
        var path = CreateTempXlsx(pkg =>
        {
            var ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "StudentNo";
            ws.Cells[1, 2].Value = "Name";
            ws.Cells[2, 1].Value = "12345";
            ws.Cells[2, 2].Value = "John Michael Smith";
        });

        var students = _service.LoadStudentsFromFile(path, "Sheet1");

        Assert.Single(students);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Michael Smith", students[0].Surname);
    }

    [Fact]
    public void EmptyRows_AreSkipped()
    {
        var path = CreateTempXlsx(pkg =>
        {
            var ws = pkg.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "StudentNo";
            ws.Cells[1, 2].Value = "Decision";
            ws.Cells[2, 1].Value = "12345";
            ws.Cells[2, 2].Value = "Accept";
            // Row 3 is empty
            ws.Cells[4, 1].Value = "67890";
            ws.Cells[4, 2].Value = "Reject";
        });

        var students = _service.LoadStudentsFromFile(path, "Sheet1");

        Assert.Equal(2, students.Count);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("67890", students[1].StudentNo);
    }
}

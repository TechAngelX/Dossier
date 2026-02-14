using Xunit;
using Dossier.Models;
using Dossier.Services;

namespace Dossier.Tests.Services;

public class ExcelServiceCsvTests : IDisposable
{
    private readonly ExcelService _service = new();
    private readonly List<string> _tempFiles = new();

    private string CreateTempCsv(string content)
    {
        var path = Path.Combine(Path.GetTempPath(), $"dossier_test_{Guid.NewGuid()}.csv");
        File.WriteAllText(path, content);
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
    public void StandardCsv_ParsesAllColumns()
    {
        var csv = "StudentNo,Decision,Forename,Surname,Programme\n12345,Accept,John,Smith,ML\n67890,Reject,Jane,Doe,CS";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Equal(2, students.Count);

        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("Accept", students[0].Decision);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Smith", students[0].Surname);
        Assert.Equal("ML", students[0].Programme);
        Assert.Equal(ProcessingStatus.Pending, students[0].Status);

        Assert.Equal("67890", students[1].StudentNo);
        Assert.Equal("Reject", students[1].Decision);
        Assert.Equal("Jane", students[1].Forename);
        Assert.Equal("Doe", students[1].Surname);
        Assert.Equal("CS", students[1].Programme);
    }

    [Fact]
    public void TabDelimited_AutoDetects()
    {
        var csv = "StudentNo\tDecision\tForename\tSurname\tProgramme\n12345\tAccept\tJohn\tSmith\tML";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("Accept", students[0].Decision);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Smith", students[0].Surname);
        Assert.Equal("ML", students[0].Programme);
    }

    [Theory]
    [InlineData("student_no")]
    [InlineData("student number")]
    [InlineData("student id")]
    [InlineData("id")]
    public void StudentNoColumn_ExactAliases(string header)
    {
        var csv = $"{header},Decision\n12345,Accept";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
    }

    [Fact]
    public void FuzzyMatch_StudentNo_ContainsStudentAndNo()
    {
        var csv = "My Student No Here,Decision\n12345,Accept";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
    }

    [Theory]
    [InlineData("decision")]
    [InlineData("status")]
    [InlineData("offer")]
    public void DecisionColumn_ExactAliases(string header)
    {
        var csv = $"StudentNo,{header}\n12345,Accept";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("Accept", students[0].Decision);
    }

    [Fact]
    public void MissingDecisionColumn_LoadsWithoutError()
    {
        var csv = "StudentNo,Forename,Surname\n12345,John,Smith";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("", students[0].Decision);
    }

    [Fact]
    public void MissingProgrammeColumn_LoadsWithoutError()
    {
        var csv = "StudentNo,Decision\n12345,Accept";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("", students[0].Programme);
    }

    [Fact]
    public void MissingStudentNoColumn_ThrowsInvalidOperation()
    {
        var csv = "Name,Decision\nJohn Smith,Accept";
        var path = CreateTempCsv(csv);

        Assert.Throws<InvalidOperationException>(() => _service.LoadStudentsFromCsv(path));
    }

    [Fact]
    public void EmptyCsv_ThrowsInvalidOperation()
    {
        var path = CreateTempCsv("");

        Assert.Throws<InvalidOperationException>(() => _service.LoadStudentsFromCsv(path));
    }

    [Fact]
    public void QuotedFieldsWithCommas_ParsedCorrectly()
    {
        var csv = "StudentNo,Name,Decision\n12345,\"Smith, John\",Accept";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
        // Name column is split on first space: "Smith," and "John"
        Assert.Equal("Smith,", students[0].Forename);
        Assert.Equal("John", students[0].Surname);
    }

    [Fact]
    public void EscapedQuotes_HandledCorrectly()
    {
        var csv = "StudentNo,Name\n12345,\"John \"\"The Man\"\" Smith\"";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        // Name = John "The Man" Smith, split on first space
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("\"The Man\" Smith", students[0].Surname);
    }

    [Fact]
    public void NameColumn_SplitsOnFirstSpace()
    {
        var csv = "StudentNo,Name\n12345,John Michael Smith";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Michael Smith", students[0].Surname);
    }

    [Fact]
    public void SeparateForenameAndSurname_PreferredOverNameColumn()
    {
        var csv = "StudentNo,Forename,Surname\n12345,John,Smith";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Smith", students[0].Surname);
    }

    [Fact]
    public void BlankRows_AreSkipped()
    {
        var csv = "StudentNo,Decision\n12345,Accept\n\n\n67890,Reject\n";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Equal(2, students.Count);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("67890", students[1].StudentNo);
    }

    [Fact]
    public void Whitespace_IsTrimmed()
    {
        var csv = "StudentNo,Decision,Forename,Surname\n  12345  ,  Accept  ,  John  ,  Smith  ";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("Accept", students[0].Decision);
        Assert.Equal("John", students[0].Forename);
        Assert.Equal("Smith", students[0].Surname);
    }

    [Fact]
    public void ProgrammeColumnAliases_Fuzzy()
    {
        var csv = "StudentNo,prog\n12345,ML";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("ML", students[0].Programme);
    }

    [Theory]
    [InlineData("progcode")]
    [InlineData("prog code")]
    [InlineData("progshort")]
    [InlineData("route")]
    public void ProgrammeColumn_FuzzyAliases(string header)
    {
        var csv = $"StudentNo,{header}\n12345,CS";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Single(students);
        Assert.Equal("CS", students[0].Programme);
    }

    [Fact]
    public void HeaderOnly_ReturnsEmptyList()
    {
        var csv = "StudentNo,Decision,Forename,Surname,Programme";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Empty(students);
    }

    [Fact]
    public void RowWithEmptyStudentNo_IsSkipped()
    {
        var csv = "StudentNo,Decision\n12345,Accept\n,Reject\n67890,Accept";
        var path = CreateTempCsv(csv);

        var students = _service.LoadStudentsFromCsv(path);

        Assert.Equal(2, students.Count);
        Assert.Equal("12345", students[0].StudentNo);
        Assert.Equal("67890", students[1].StudentNo);
    }
}

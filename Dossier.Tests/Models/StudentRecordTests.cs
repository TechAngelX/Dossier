using Xunit;
using Dossier.Models;

namespace Dossier.Tests.Models;

public class StudentRecordTests
{
    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var record = new StudentRecord();

        Assert.Equal(string.Empty, record.StudentNo);
        Assert.Equal(string.Empty, record.Decision);
        Assert.Equal("", record.Forename);
        Assert.Equal("", record.Surname);
        Assert.Equal(string.Empty, record.Programme);
        Assert.Equal(ProcessingStatus.Pending, record.Status);
        Assert.Equal(string.Empty, record.ErrorMessage);
    }

    [Fact]
    public void Name_WithBothNames_ReturnsCombined()
    {
        var record = new StudentRecord { Forename = "John", Surname = "Smith" };
        Assert.Equal("John Smith", record.Name);
    }

    [Fact]
    public void Name_WithForenameOnly_ReturnsForename()
    {
        var record = new StudentRecord { Forename = "John" };
        Assert.Equal("John", record.Name);
    }

    [Fact]
    public void Name_WithSurnameOnly_ReturnsSurname()
    {
        var record = new StudentRecord { Surname = "Smith" };
        Assert.Equal("Smith", record.Name);
    }

    [Fact]
    public void Name_WithBothEmpty_ReturnsEmpty()
    {
        var record = new StudentRecord();
        Assert.Equal("", record.Name);
    }

    [Fact]
    public void Name_TrimsOuterWhitespace()
    {
        // Name = $"{Forename} {Surname}".Trim() — only trims leading/trailing
        var record = new StudentRecord { Forename = "  John", Surname = "Smith  " };
        Assert.Equal("John Smith", record.Name);
    }

    [Theory]
    [InlineData(ProcessingStatus.Pending)]
    [InlineData(ProcessingStatus.Processing)]
    [InlineData(ProcessingStatus.Success)]
    [InlineData(ProcessingStatus.Failed)]
    [InlineData(ProcessingStatus.Skipped)]
    public void Status_CanBeSetToAllValues(ProcessingStatus status)
    {
        var record = new StudentRecord { Status = status };
        Assert.Equal(status, record.Status);
    }

    [Theory]
    [InlineData(DecisionType.Accept)]
    [InlineData(DecisionType.Reject)]
    [InlineData(DecisionType.Unknown)]
    public void DecisionType_HasExpectedValues(DecisionType decision)
    {
        Assert.True(Enum.IsDefined(typeof(DecisionType), decision));
    }
}

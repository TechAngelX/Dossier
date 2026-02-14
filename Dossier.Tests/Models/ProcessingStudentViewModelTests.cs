using Xunit;
using System.ComponentModel;
using Dossier.Models;

namespace Dossier.Tests.Models;

public class ProcessingStudentViewModelTests
{
    [Fact]
    public void DefaultValues_AreCorrect()
    {
        var vm = new ProcessingStudentViewModel();

        Assert.Equal("", vm.StudentNo);
        Assert.Equal("", vm.Forename);
        Assert.Equal("", vm.Surname);
        Assert.Equal("", vm.Decision);
        Assert.Equal("\u23f3", vm.StatusIcon); // ⏳
        Assert.Equal("Pending", vm.StatusText);
        Assert.Equal("#A0AEC0", vm.StatusColor);
    }

    [Theory]
    [InlineData(nameof(ProcessingStudentViewModel.StudentNo), "12345")]
    [InlineData(nameof(ProcessingStudentViewModel.Forename), "John")]
    [InlineData(nameof(ProcessingStudentViewModel.Surname), "Smith")]
    [InlineData(nameof(ProcessingStudentViewModel.Decision), "Accept")]
    [InlineData(nameof(ProcessingStudentViewModel.StatusIcon), "\u2713")] // ✓
    [InlineData(nameof(ProcessingStudentViewModel.StatusText), "Done")]
    [InlineData(nameof(ProcessingStudentViewModel.StatusColor), "#48BB78")]
    public void PropertyChanged_FiresForEachProperty(string propertyName, string newValue)
    {
        var vm = new ProcessingStudentViewModel();
        string? changedProperty = null;

        vm.PropertyChanged += (_, e) => changedProperty = e.PropertyName;

        var prop = typeof(ProcessingStudentViewModel).GetProperty(propertyName)!;
        prop.SetValue(vm, newValue);

        Assert.Equal(propertyName, changedProperty);
    }

    [Fact]
    public void MultiplePropertyChanges_FireMultipleEvents()
    {
        var vm = new ProcessingStudentViewModel();
        var changedProperties = new List<string?>();

        vm.PropertyChanged += (_, e) => changedProperties.Add(e.PropertyName);

        vm.StudentNo = "12345";
        vm.Forename = "John";
        vm.StatusText = "Processing";

        Assert.Equal(3, changedProperties.Count);
        Assert.Contains(nameof(ProcessingStudentViewModel.StudentNo), changedProperties);
        Assert.Contains(nameof(ProcessingStudentViewModel.Forename), changedProperties);
        Assert.Contains(nameof(ProcessingStudentViewModel.StatusText), changedProperties);
    }

    [Fact]
    public void ImplementsINotifyPropertyChanged()
    {
        var vm = new ProcessingStudentViewModel();
        Assert.IsAssignableFrom<INotifyPropertyChanged>(vm);
    }

    [Fact]
    public void SetProperty_UpdatesValue()
    {
        var vm = new ProcessingStudentViewModel
        {
            StudentNo = "99999",
            Forename = "Jane",
            Surname = "Doe",
            Decision = "Reject",
            StatusIcon = "\u2717", // ✗
            StatusText = "Failed",
            StatusColor = "#E53E3E"
        };

        Assert.Equal("99999", vm.StudentNo);
        Assert.Equal("Jane", vm.Forename);
        Assert.Equal("Doe", vm.Surname);
        Assert.Equal("Reject", vm.Decision);
        Assert.Equal("\u2717", vm.StatusIcon);
        Assert.Equal("Failed", vm.StatusText);
        Assert.Equal("#E53E3E", vm.StatusColor);
    }
}

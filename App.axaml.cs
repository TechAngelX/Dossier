// App.axaml.cs

using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using OfficeOpenXml;

namespace Dossier;

public partial class App : Application
{
    public override void Initialize()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new MainWindow();
        }

        base.OnFrameworkInitializationCompleted();
    }
}

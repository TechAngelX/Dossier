// SettingsWindow.axaml.cs

using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Dossier.Models;

namespace Dossier;

public partial class SettingsWindow : Window
{
    private readonly AppConfig _config;
    
    public SettingsWindow(AppConfig config)
    {
        InitializeComponent();
        _config = config;
        LoadSettings();
        SetupEventHandlers();
    }
    
    private void InitializeComponent()
    {
        Avalonia.Markup.Xaml.AvaloniaXamlLoader.Load(this);
        
        _headlessModeCheckBox = this.FindControl<CheckBox>("HeadlessModeCheckBox")!;
        _useExistingSsoCheckBox = this.FindControl<CheckBox>("UseExistingSsoCheckBox")!;
        _edgeUserDataDirTextBox = this.FindControl<TextBox>("EdgeUserDataDirTextBox")!;
        _browseDirButton = this.FindControl<Button>("BrowseDirButton")!;
        _actionDelayNumeric = this.FindControl<NumericUpDown>("ActionDelayNumeric")!;
        _porticoUrlTextBox = this.FindControl<TextBox>("PorticoUrlTextBox")!;
        _cancelButton = this.FindControl<Button>("CancelButton")!;
        _saveButton = this.FindControl<Button>("SaveButton")!;
    }
    
    private CheckBox _headlessModeCheckBox = null!;
    private CheckBox _useExistingSsoCheckBox = null!;
    private TextBox _edgeUserDataDirTextBox = null!;
    private Button _browseDirButton = null!;
    private NumericUpDown _actionDelayNumeric = null!;
    private TextBox _porticoUrlTextBox = null!;
    private Button _cancelButton = null!;
    private Button _saveButton = null!;
    
    private void SetupEventHandlers()
    {
        _browseDirButton.Click += BrowseDirButton_Click;
        _cancelButton.Click += CancelButton_Click;
        _saveButton.Click += SaveButton_Click;
    }
    
    private void LoadSettings()
    {
        _headlessModeCheckBox.IsChecked = _config.HeadlessMode;
        _useExistingSsoCheckBox.IsChecked = _config.UseExistingSsoSession;
        _edgeUserDataDirTextBox.Text = _config.EdgeUserDataDir;
        _actionDelayNumeric.Value = _config.ActionDelayMs;
        _porticoUrlTextBox.Text = _config.PorticoUrl;
    }
    
    private async void BrowseDirButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;
        
        var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Select Edge User Data Directory",
            AllowMultiple = false
        });
        
        if (folders.Count > 0)
        {
            _edgeUserDataDirTextBox.Text = folders[0].Path.LocalPath;
        }
    }
    
    private void CancelButton_Click(object? sender, RoutedEventArgs e)
    {
        Close();
    }
    
    private void SaveButton_Click(object? sender, RoutedEventArgs e)
    {
        _config.HeadlessMode = _headlessModeCheckBox.IsChecked ?? false;
        _config.UseExistingSsoSession = _useExistingSsoCheckBox.IsChecked ?? true;
        _config.EdgeUserDataDir = _edgeUserDataDirTextBox.Text ?? string.Empty;
        _config.ActionDelayMs = (int)(_actionDelayNumeric.Value ?? 500);
        _config.PorticoUrl = _porticoUrlTextBox.Text ?? "https://evision.ucl.ac.uk/urd/sits.urd/run/siw_lgn";
        
        Close();
    }
}

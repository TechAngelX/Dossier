// Models/AppConfig.cs

namespace Dossier.Models;

public class AppConfig
{
    public string PorticoUrl { get; set; } = "https://evision.ucl.ac.uk/urd/sits.urd/run/siw_lgn";
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public bool UseExistingSsoSession { get; set; } = true;
    public string EdgeUserDataDir { get; set; } = string.Empty;
    public int ActionDelayMs { get; set; } = 500;
    public bool HeadlessMode { get; set; } = false;
}

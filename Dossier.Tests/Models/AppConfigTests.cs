using Xunit;
using Dossier.Models;

namespace Dossier.Tests.Models;

public class AppConfigTests
{
    [Fact]
    public void DefaultPorticoUrl_IsCorrect()
    {
        var config = new AppConfig();
        Assert.Equal("https://evision.ucl.ac.uk/urd/sits.urd/run/siw_lgn", config.PorticoUrl);
    }

    [Fact]
    public void DefaultActionDelayMs_Is500()
    {
        var config = new AppConfig();
        Assert.Equal(500, config.ActionDelayMs);
    }

    [Fact]
    public void DefaultHeadlessMode_IsFalse()
    {
        var config = new AppConfig();
        Assert.False(config.HeadlessMode);
    }

    [Fact]
    public void DefaultUseExistingSsoSession_IsTrue()
    {
        var config = new AppConfig();
        Assert.True(config.UseExistingSsoSession);
    }

    [Fact]
    public void DefaultStrings_AreEmpty()
    {
        var config = new AppConfig();
        Assert.Equal(string.Empty, config.Username);
        Assert.Equal(string.Empty, config.Password);
        Assert.Equal(string.Empty, config.EdgeUserDataDir);
    }

    [Fact]
    public void Properties_CanBeSet()
    {
        var config = new AppConfig
        {
            PorticoUrl = "https://example.com",
            Username = "testuser",
            Password = "testpass",
            UseExistingSsoSession = false,
            EdgeUserDataDir = "/tmp/edge",
            ActionDelayMs = 1000,
            HeadlessMode = true
        };

        Assert.Equal("https://example.com", config.PorticoUrl);
        Assert.Equal("testuser", config.Username);
        Assert.Equal("testpass", config.Password);
        Assert.False(config.UseExistingSsoSession);
        Assert.Equal("/tmp/edge", config.EdgeUserDataDir);
        Assert.Equal(1000, config.ActionDelayMs);
        Assert.True(config.HeadlessMode);
    }
}

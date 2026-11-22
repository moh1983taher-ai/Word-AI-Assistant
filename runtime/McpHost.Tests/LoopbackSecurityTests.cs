using ResearchTools.McpHost.Security;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class LoopbackSecurityTests
{
    [Theory]
    [InlineData("https://moh1983taher-ai.github.io")]
    [InlineData("https://moh1983taher-ai.github.io/")]
    [InlineData("https://localhost:3000")]
    public void AllowsConfiguredBrowserOrigins(string origin)
    {
        Assert.True(LoopbackSecurity.IsAllowedOrigin(origin));
    }

    [Theory]
    [InlineData("https://moh1983taher-ai.github.io.attacker.example")]
    [InlineData("http://moh1983taher-ai.github.io")]
    [InlineData("https://attacker.example")]
    [InlineData("")]
    public void RejectsUntrustedOrigins(string origin)
    {
        Assert.False(LoopbackSecurity.IsAllowedOrigin(origin));
    }

    [Fact]
    public void AllowsLocalMcpClientsWithoutBrowserOrigin()
    {
        Assert.True(LoopbackSecurity.IsAllowedMcpOrigin(null));
        Assert.True(LoopbackSecurity.IsAllowedMcpOrigin(""));
        Assert.True(LoopbackSecurity.IsAllowedMcpOrigin("http://127.0.0.1:18081"));
        Assert.True(LoopbackSecurity.IsAllowedMcpOrigin("http://localhost:18081"));
        Assert.True(LoopbackSecurity.IsAllowedMcpOrigin("https://moh1983taher-ai.github.io"));
    }

    [Theory]
    [InlineData("https://attacker.example")]
    [InlineData("http://127.0.0.1:18082")]
    [InlineData("http://127.0.0.1.attacker.example:18081")]
    [InlineData("http://localhost:8080")]
    public void RejectsUntrustedMcpOrigins(string origin)
    {
        Assert.False(LoopbackSecurity.IsAllowedMcpOrigin(origin));
    }

    [Fact]
    public void SupportsExplicitAlternateMcpLoopbackPort()
    {
        Assert.True(LoopbackSecurity.IsAllowedMcpOrigin("http://127.0.0.1:18082", 18082));
        Assert.False(LoopbackSecurity.IsAllowedMcpOrigin("http://127.0.0.1:18081", 18082));
    }

    [Theory]
    [InlineData("127.0.0.1:18081")]
    [InlineData("localhost:18081")]
    public void AllowsOnlyConfiguredDefaultLoopbackHosts(string host)
    {
        Assert.True(LoopbackSecurity.IsAllowedHost(host));
    }

    [Theory]
    [InlineData("0.0.0.0:18081")]
    [InlineData("192.168.1.10:18081")]
    [InlineData("attacker.example:18081")]
    [InlineData("127.0.0.1:8080")]
    public void RejectsNonLoopbackOrUnexpectedHosts(string host)
    {
        Assert.False(LoopbackSecurity.IsAllowedHost(host));
    }

    [Fact]
    public void AllowsAnExplicitAlternateLoopbackPortWithoutOpeningNetworkInterfaces()
    {
        Assert.True(LoopbackSecurity.IsAllowedHost("127.0.0.1:18082", 18082));
        Assert.True(LoopbackSecurity.IsAllowedHost("localhost:18082", 18082));
        Assert.False(LoopbackSecurity.IsAllowedHost("192.168.1.10:18082", 18082));
        Assert.False(LoopbackSecurity.IsAllowedHost("127.0.0.1:18081", 18082));
    }

    [Fact]
    public void RejectsInvalidAlternatePorts()
    {
        Assert.False(LoopbackSecurity.IsAllowedHost("127.0.0.1:80", 80));
        Assert.False(LoopbackSecurity.IsAllowedHost("127.0.0.1:70000", 70000));
    }

    [Fact]
    public void ResearchRequiresTheExpectedOriginAndClientMarker()
    {
        Assert.True(LoopbackSecurity.IsAuthorizedResearchClient(
            "https://moh1983taher-ai.github.io", "word-addin"));
        Assert.False(LoopbackSecurity.IsAuthorizedResearchClient(
            "https://attacker.example", "word-addin"));
        Assert.False(LoopbackSecurity.IsAuthorizedResearchClient(
            "https://moh1983taher-ai.github.io", "other-client"));
    }
}

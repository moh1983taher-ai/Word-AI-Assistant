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

    [Theory]
    [InlineData("127.0.0.1:18081")]
    [InlineData("localhost:18081")]
    public void AllowsOnlyConfiguredLoopbackHosts(string host)
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

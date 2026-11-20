using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class DuckAiEndpointSelectionTests
{
    [Fact]
    public void PrefersDirectBackendBeforeLocalAgentProxy()
    {
        Assert.Equal("http://127.0.0.1:18080", DuckAiClient.CandidateBaseAddresses[0]);
        Assert.Equal("http://127.0.0.1:8080", DuckAiClient.CandidateBaseAddresses[1]);
    }
}

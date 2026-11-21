using System.Text.Json;
using System.Text.Json.Serialization;

namespace ResearchTools.McpHost.Infrastructure;

public static class JsonDefaults
{
    public static readonly JsonSerializerOptions Web = Create();

    private static JsonSerializerOptions Create()
    {
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = false,
            MaxDepth = 64
        };
        options.Converters.Add(new JsonStringEnumConverter());
        return options;
    }
}

namespace ResearchTools.McpHost.Services;

public sealed record ResearchQuery(string Query, string Purpose);

public sealed record ResearchPlan(
    string Intent,
    string Kind,
    IReadOnlyList<ResearchQuery> Queries,
    IReadOnlyList<string> Coverage);

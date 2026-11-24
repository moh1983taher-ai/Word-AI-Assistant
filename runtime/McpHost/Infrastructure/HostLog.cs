using System.Text;

namespace ResearchTools.McpHost.Infrastructure;

/// <summary>
/// Small rolling file logger. User prompts, tool arguments, response bodies and raw exception
/// messages are deliberately excluded from persistent diagnostics.
/// </summary>
public sealed class HostLog
{
    private const long MaxBytes = 2 * 1024 * 1024;
    private readonly object _sync = new();
    private readonly string _path;

    public HostLog()
    {
        var directory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ResearchTools", "DuckAI");
        Directory.CreateDirectory(directory);

        // Preview runs use a private log so their diagnostics cannot be confused
        // with concurrent requests handled by the permanent 18081 service.
        var previewLogId = Environment.GetEnvironmentVariable("RESEARCHTOOLS_MCP_PREVIEW_LOG_ID");
        var isValidPreviewId = !string.IsNullOrWhiteSpace(previewLogId) &&
            previewLogId.All(char.IsAsciiLetterOrDigit);
        var fileName = isValidPreviewId
            ? $"DuckAI-MCP-Host-preview-{previewLogId}.log"
            : "DuckAI-MCP-Host.log";
        _path = Path.Combine(directory, fileName);
    }

    public void Info(string message) => Write("INFO", message);
    public void Warning(string message) => Write("WARN", message);
    public void Error(string message) => Write("ERROR", message);
    public void Error(string operation, Exception exception) =>
        Write("ERROR", $"{operation}: {exception.GetType().Name}");

    private void Write(string level, string message)
    {
        try
        {
            lock (_sync)
            {
                var info = new FileInfo(_path);
                if (info.Exists && info.Length >= MaxBytes)
                {
                    var previous = _path + ".old";
                    if (File.Exists(previous)) File.Delete(previous);
                    File.Move(_path, previous);
                }

                var line = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz} [{level}] {message}{Environment.NewLine}";
                File.AppendAllText(_path, line, new UTF8Encoding(false));
            }
        }
        catch
        {
            // Diagnostics must never interrupt request processing.
        }
    }
}

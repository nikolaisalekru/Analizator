using Analizator.Core;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;

namespace Analizator;

internal static class MonitoringWorker
{
    private const string WorkerArgument = "--monitoring-worker";
    private const string ProgressPrefix = "[MONITORING_PROGRESS]";
    private const string ResultPrefix = "[MONITORING_RESULT]";
    private const string ErrorPrefix = "[MONITORING_ERROR]";

    public static bool IsWorkerInvocation(string[] args) =>
        args.Length == 2 && string.Equals(args[0], WorkerArgument, StringComparison.Ordinal);

    public static int Run(string[] args)
    {
        try
        {
            var request = DeserializePayload<MonitoringReportRequest>(args[1]);
            var progress = new InlineProgress<MonitoringReportProgress>(item =>
                Console.WriteLine(ProgressPrefix + JsonSerializer.Serialize(item)));
            var result = new MonitoringReportService().Run(request, progress);
            Console.WriteLine(ResultPrefix + JsonSerializer.Serialize(result));
            return 0;
        }
        catch (Exception exception)
        {
            Console.WriteLine(ErrorPrefix + JsonSerializer.Serialize(exception.Message));
            return 1;
        }
    }

    public static async Task<MonitoringReportResult> RunIsolatedAsync(
        MonitoringReportRequest request,
        IProgress<MonitoringReportProgress> progress,
        CancellationToken cancellationToken)
    {
        var startInfo = CreateStartInfo();
        startInfo.ArgumentList.Add(WorkerArgument);
        startInfo.ArgumentList.Add(SerializePayload(request));
        startInfo.Environment["DOTNET_EnableDiagnostics_Debugger"] = "0";
        startInfo.Environment["DOTNET_MODIFIABLE_ASSEMBLIES"] = "none";

        using var process = new Process { StartInfo = startInfo };
        if (!process.Start())
            throw new InvalidOperationException("Не удалось запустить процесс мониторинга КТК.");

        var stderrTask = process.StandardError.ReadToEndAsync();
        MonitoringReportResult? result = null;
        string? workerError = null;
        try
        {
            while (await process.StandardOutput.ReadLineAsync(cancellationToken) is { } line)
            {
                if (line.StartsWith(ProgressPrefix, StringComparison.Ordinal))
                {
                    var item = JsonSerializer.Deserialize<MonitoringReportProgress>(line[ProgressPrefix.Length..]);
                    if (item is not null)
                        progress.Report(item);
                }
                else if (line.StartsWith(ResultPrefix, StringComparison.Ordinal))
                    result = JsonSerializer.Deserialize<MonitoringReportResult>(line[ResultPrefix.Length..]);
                else if (line.StartsWith(ErrorPrefix, StringComparison.Ordinal))
                    workerError = JsonSerializer.Deserialize<string>(line[ErrorPrefix.Length..]);
            }

            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            TryStop(process);
            throw;
        }

        var stderr = await stderrTask;
        if (process.ExitCode != 0)
            throw new InvalidOperationException(!string.IsNullOrWhiteSpace(workerError)
                ? workerError
                : string.IsNullOrWhiteSpace(stderr)
                    ? $"Процесс мониторинга завершился с кодом {process.ExitCode}."
                    : stderr.Trim());
        return result ?? throw new InvalidOperationException(
            "Процесс мониторинга завершился без итоговых данных.");
    }

    private static ProcessStartInfo CreateStartInfo()
    {
        var processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Не удалось определить путь к программе.");
        var startInfo = new ProcessStartInfo
        {
            FileName = processPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()?.Location
                ?? throw new InvalidOperationException("Не удалось определить сборку программы."));
        return startInfo;
    }

    private static string SerializePayload<T>(T value) =>
        Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value)));

    private static T DeserializePayload<T>(string payload) =>
        JsonSerializer.Deserialize<T>(Encoding.UTF8.GetString(Convert.FromBase64String(payload)))
        ?? throw new InvalidOperationException("Не удалось прочитать параметры мониторинга.");

    private static void TryStop(Process process)
    {
        try
        {
            if (!process.HasExited)
                process.Kill(entireProcessTree: true);
        }
        catch (InvalidOperationException)
        {
        }
    }

    private sealed class InlineProgress<T>(Action<T> handler) : IProgress<T>
    {
        public void Report(T value) => handler(value);
    }
}

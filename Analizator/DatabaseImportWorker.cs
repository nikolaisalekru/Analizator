using Analizator.Core;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;

namespace Analizator;

internal sealed record DatabaseImportWorkerRequest(
    string DatabasePath,
    string[] Files,
    Dictionary<string, List<string>> Installations);

internal static class DatabaseImportWorker
{
    private const string WorkerArgument = "--database-import-worker";
    private const string ProgressPrefix = "[DATABASE_PROGRESS]";
    private const string ResultPrefix = "[DATABASE_RESULT]";
    private const string ErrorPrefix = "[DATABASE_ERROR]";

    public static bool IsWorkerInvocation(string[] args) =>
        args.Length == 2 && string.Equals(args[0], WorkerArgument, StringComparison.Ordinal);

    public static int Run(string[] args)
    {
        try
        {
            var request = DeserializePayload<DatabaseImportWorkerRequest>(args[1]);
            var progress = new InlineProgress<DatabaseImportProgress>(item =>
                Console.WriteLine(ProgressPrefix + JsonSerializer.Serialize(item)));
            var result = new TrainingDatabaseService(request.DatabasePath)
                .ImportFiles(request.Files, request.Installations, progress);
            Console.WriteLine(ResultPrefix + JsonSerializer.Serialize(result));
            return 0;
        }
        catch (Exception exception)
        {
            Console.WriteLine(ErrorPrefix + JsonSerializer.Serialize(exception.Message));
            return 1;
        }
    }

    public static async Task<DatabaseImportResult> RunIsolatedAsync(
        DatabaseImportWorkerRequest request,
        IProgress<DatabaseImportProgress> progress,
        CancellationToken cancellationToken)
    {
        var startInfo = CreateStartInfo();
        startInfo.ArgumentList.Add(WorkerArgument);
        startInfo.ArgumentList.Add(SerializePayload(request));
        startInfo.Environment["DOTNET_EnableDiagnostics_Debugger"] = "0";
        startInfo.Environment["DOTNET_MODIFIABLE_ASSEMBLIES"] = "none";

        using var process = new Process { StartInfo = startInfo };
        if (!process.Start())
            throw new InvalidOperationException("Не удалось запустить процесс импорта в базу данных.");

        var stderrTask = process.StandardError.ReadToEndAsync();
        DatabaseImportResult? result = null;
        string? workerError = null;
        try
        {
            while (await process.StandardOutput.ReadLineAsync(cancellationToken) is { } line)
            {
                if (line.StartsWith(ProgressPrefix, StringComparison.Ordinal))
                {
                    var item = JsonSerializer.Deserialize<DatabaseImportProgress>(line[ProgressPrefix.Length..]);
                    if (item is not null)
                        progress.Report(item);
                }
                else if (line.StartsWith(ResultPrefix, StringComparison.Ordinal))
                {
                    result = JsonSerializer.Deserialize<DatabaseImportResult>(line[ResultPrefix.Length..]);
                }
                else if (line.StartsWith(ErrorPrefix, StringComparison.Ordinal))
                {
                    workerError = JsonSerializer.Deserialize<string>(line[ErrorPrefix.Length..]);
                }
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
        {
            throw new InvalidOperationException(
                !string.IsNullOrWhiteSpace(workerError)
                    ? workerError
                    : string.IsNullOrWhiteSpace(stderr)
                        ? $"Процесс импорта завершился с кодом {process.ExitCode}."
                        : stderr.Trim());
        }

        return result ?? throw new InvalidOperationException(
            "Процесс импорта завершился без итоговых данных.");
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
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet",
                StringComparison.OrdinalIgnoreCase))
        {
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()?.Location
                ?? throw new InvalidOperationException("Не удалось определить сборку программы."));
        }
        return startInfo;
    }

    private static string SerializePayload<T>(T value) =>
        Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value)));

    private static T DeserializePayload<T>(string payload) =>
        JsonSerializer.Deserialize<T>(Encoding.UTF8.GetString(Convert.FromBase64String(payload)))
        ?? throw new InvalidOperationException("Не удалось прочитать параметры импорта.");

    private static void TryStop(Process process)
    {
        try
        {
            if (!process.HasExited)
                process.Kill(entireProcessTree: true);
        }
        catch (InvalidOperationException)
        {
            // The worker completed between the checks.
        }
    }

    private sealed class InlineProgress<T>(Action<T> handler) : IProgress<T>
    {
        public void Report(T value) => handler(value);
    }
}

using System.Text;
using ClosedXML.Excel;

namespace Analizator.Core;

public sealed class AnalyzerService
{
    public AnalyzerResult Run(
        AnalyzerRequest request,
        IProgress<AnalyzerProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ValidateRequest(request);
        Directory.CreateDirectory(request.OutputDirectory);
        var extension = string.Equals(Path.GetExtension(request.AnalysisFile), ".xlsm", StringComparison.OrdinalIgnoreCase)
            ? ".xlsm"
            : ".xlsx";
        var outputFile = Path.Combine(request.OutputDirectory, "Анализ_прохождения_результат" + extension);
        EnsureOutputFileIsAvailable(outputFile);
        var dataRoot = Directory.GetParent(request.ConfigurationDirectory)?.FullName ?? request.OutputDirectory;
        var logsDirectory = Path.Combine(dataRoot, "logs");
        Directory.CreateDirectory(logsDirectory);
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
        var logFile = Path.Combine(logsDirectory, timestamp + ".txt");

        using var writer = new StreamWriter(logFile, false, new UTF8Encoding(false)) { AutoFlush = true };
        void Log(string message)
        {
            writer.WriteLine(message);
            writer.Flush();
        }

        try
        {
            progress?.Report(new AnalyzerProgress(3, "Проверка входных файлов"));
            Log("АНАЛИЗАТОР ПРОХОЖДЕНИЯ КТК — C# CORE");
            Log($"Начало: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            Log($"Файл анализа: {request.AnalysisFile}");
            Log($"Файл выгрузки: {request.ExportFile}");
            cancellationToken.ThrowIfCancellationRequested();

            if (request.CreateBackup)
            {
                var backupDirectory = Path.Combine(request.OutputDirectory, "backups");
                Directory.CreateDirectory(backupDirectory);
                var backupPath = Path.Combine(backupDirectory, $"{timestamp} {Path.GetFileName(request.AnalysisFile)}");
                File.Copy(request.AnalysisFile, backupPath, true);
                Log($"Резервная копия: {backupPath}");
            }

            progress?.Report(new AnalyzerProgress(10, "Загрузка конфигурации"));
            var configuration = ConfigurationService.Load(request.ConfigurationDirectory);
            Log($"Загружено установок: {configuration.Installations.Count}");

            var month = TextNormalization.MonthFromFileName(request.ExportFile)
                        ?? throw new InvalidOperationException(
                            $"Не удалось определить месяц по имени файла выгрузки: {Path.GetFileName(request.ExportFile)}");
            Log($"Месяц выгрузки: {TextNormalization.MonthNames[month]} (№{month})");

            progress?.Report(new AnalyzerProgress(18, "Чтение выгрузки КТК"));
            var export = ExportWorkbookReader.Read(
                request.ExportFile,
                configuration.Installations,
                configuration.Settings.ModeFilter,
                cancellationToken,
                Log);
            Log($"Принято отчётов: {export.TotalReports}");

            progress?.Report(new AnalyzerProgress(28, "Открытие книги анализа"));
            using var workbook = new XLWorkbook(request.AnalysisFile);
            var statistics = new ProcessingStatistics { Reports = export.TotalReports };

            progress?.Report(new AnalyzerProgress(40, "Расчёт и запись результатов"));
            var processing = AnalysisWorkbookProcessor.Process(
                workbook,
                export,
                configuration,
                month,
                request.AutoAddNewPeople,
                statistics,
                cancellationToken,
                Log);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new AnalyzerProgress(82, "Формирование служебных листов"));
            ReportingService.FinalizeWorkbook(
                workbook,
                processing,
                export,
                configuration.Settings,
                configuration.Installations,
                statistics,
                progress,
                cancellationToken,
                Log);

            progress?.Report(new AnalyzerProgress(94, "Сохранение результирующей книги"));
            try
            {
                workbook.SaveAs(outputFile);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                throw new AnalysisFileAccessException(outputFile, exception);
            }

            Log($"Результат: {outputFile}");
            Log($"Сопоставлено сотрудников: {statistics.MatchedPeople}");
            Log($"Новых сотрудников: {statistics.NewPeople}; добавлено: {statistics.AddedPeople}");
            Log($"Отсутствовали в выгрузке: {statistics.AbsentPeople}");
            Log($"Завершено: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            progress?.Report(new AnalyzerProgress(100, "Обработка успешно завершена"));

            return new AnalyzerResult
            {
                OutputFile = outputFile,
                Reports = statistics.Reports,
                MatchedPeople = statistics.MatchedPeople,
                NewPeople = statistics.NewPeople,
                AddedPeople = statistics.AddedPeople,
                AbsentPeople = statistics.AbsentPeople,
                ExactMatches = statistics.ExactMatches,
                KeyMatches = statistics.KeyMatches,
                FuzzyMatches = statistics.FuzzyMatches,
                UnmatchedInstallations = processing.UnmatchedInstallations,
                LogFile = logFile
            };
        }
        catch (Exception exception)
        {
            Log("");
            Log("КРИТИЧЕСКАЯ ОШИБКА");
            Log(exception.ToString());
            throw;
        }
    }

    private static void EnsureOutputFileIsAvailable(string outputFile)
    {
        if (!File.Exists(outputFile))
            return;

        try
        {
            using var _ = new FileStream(
                outputFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new AnalysisFileAccessException(outputFile, exception);
        }
    }

    private static void ValidateRequest(AnalyzerRequest request)
    {
        if (!File.Exists(request.AnalysisFile))
            throw new FileNotFoundException("Файл анализа не найден.", request.AnalysisFile);
        if (!File.Exists(request.ExportFile))
            throw new FileNotFoundException("Файл выгрузки не найден.", request.ExportFile);
        var analysisExtension = Path.GetExtension(request.AnalysisFile);
        var exportExtension = Path.GetExtension(request.ExportFile);
        if (analysisExtension is not ".xlsx" and not ".xlsm" && analysisExtension is not ".XLSX" and not ".XLSM")
            throw new InvalidDataException("Файл анализа должен иметь формат XLSX или XLSM.");
        if (exportExtension is not ".xlsx" and not ".xlsm" && exportExtension is not ".XLSX" and not ".XLSM")
            throw new InvalidDataException("Файл выгрузки должен иметь формат XLSX или XLSM.");
        if (string.Equals(Path.GetFullPath(request.AnalysisFile), Path.GetFullPath(request.ExportFile), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Файл анализа и файл выгрузки не могут совпадать.");
    }
}

using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Analizator.Core;
using ClosedXML.Excel;

var failures = new List<string>();
Run("Нормализация", () =>
{
    Equal("ф и о по отчету", TextNormalization.NormalizeHeader("  Ф.И.О.\nпо-отчёту  "));
    Equal(8, TextNormalization.MonthFromFileName("Выгрузка_КТК_август_2026.xlsx"));
    Equal(92.5, TextNormalization.SafeDouble("92,5%"));
});
Run("Сопоставление", () =>
{
    var installations = new Dictionary<string, List<string>> { ["ГФУ-2"] = ["ГФУ", "ГФУ 2"] };
    Equal("ГФУ-2", MatchingService.MatchInstallation("ГФУ", installations).Name);
    var exact = MatchingService.MatchPerson("Иванов Иван", ["Иванов  Иван"], "ГФУ-2",
        new Dictionary<string, Dictionary<string, string>>(), installations);
    Equal(MatchKind.Exact, exact.Kind);
});

var root = Path.Combine(Path.GetTempPath(), "Analizator-CSharp-tests-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    Run("Сохранение настройки порогов", () =>
    {
        var config = Path.Combine(root, "threshold-save-config");
        Directory.CreateDirectory(config);
        var settingsPath = Path.Combine(config, "settings.json");
        File.WriteAllText(settingsPath, "{\"custom_setting\":\"preserve-me\"}");
        ConfigurationService.SavePassingThresholds(config, new PassingThresholdSettings
        {
            UsePerInstallation = true,
            GlobalPercent = 70d,
            Installations = new Dictionary<string, double>
            {
                ["ГФУ-2"] = 55d,
                ["ВБ"] = 65d
            }
        });

        var loaded = ConfigurationService.Load(config).Settings.PassingThreshold;
        True(loaded.UsePerInstallation, "Индивидуальный режим не сохранился");
        Equal(70d, loaded.GlobalPercent);
        Equal(55d, loaded.GetForInstallation("ГФУ-2"));
        Equal(65d, loaded.GetForInstallation("ВБ"));
        using var document = JsonDocument.Parse(File.ReadAllText(settingsPath));
        Equal("preserve-me", document.RootElement.GetProperty("custom_setting").GetString());
    });

    Run("Сквозная обработка XLSX", () =>
    {
        var analysis = Path.Combine(root, "Анализ_прохождения 2026_08.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_2026.xlsx");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export,
        [
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 80d),
            ("04.08.2026", "Сценарий А", "Иванов Иван Иванович", 95d),
            ("05.08.2026", "Сценарий Б", "Иванов Иван Иванович", 75d)
        ]);
        var sourceHash = Hash(analysis);
        var service = new AnalyzerService();
        var progressUpdates = new List<AnalyzerProgress>();
        var result = service.Run(new AnalyzerRequest(
                analysis, export, Path.Combine(root, "output"), Path.Combine(root, "config"), false, false),
            new InlineProgress<AnalyzerProgress>(progressUpdates.Add));
        Equal(sourceHash, Hash(analysis));
        Equal(3, result.Reports);
        Equal(1, result.MatchedPeople);
        True(progressUpdates.Any(item => item.Percent is >= 83 and <= 93),
            "Нет промежуточного прогресса при формировании служебных листов");
        int firstConditionalFormats;
        using (var workbook = new XLWorkbook(result.OutputFile))
        {
            var sheet = workbook.Worksheet("ГФУ-2");
            Equal(95d, sheet.Cell("C3").GetDouble());
            Equal(75d, sheet.Cell("D3").GetDouble());
            Equal("ДА", sheet.Cell("H3").GetString());
            True(workbook.TryGetWorksheet("Новые сотрудники", out _), "Нет листа новых сотрудников");
            True(workbook.TryGetWorksheet("Отчёт обработки", out _), "Нет отчёта обработки");
            firstConditionalFormats = sheet.ConditionalFormats.Count();
        }
        var second = service.Run(new AnalyzerRequest(
            result.OutputFile, export, Path.Combine(root, "second-output"), Path.Combine(root, "config"), false, false));
        using var secondWorkbook = new XLWorkbook(second.OutputFile);
        var secondSheet = secondWorkbook.Worksheet("ГФУ-2");
        Equal(firstConditionalFormats, secondSheet.ConditionalFormats.Count());
        Equal(95d, secondSheet.Cell("C3").GetDouble());
    });

    Run("Создание отсутствующего месяца", () =>
    {
        var analysis = Path.Combine(root, "Анализ_июль.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август.xlsx");
        WriteAnalysis(analysis, "Июль", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 90d)]);
        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "month-output"), Path.Combine(root, "config"), false, false));
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        var blocks = AnalysisWorkbookProcessor.FindMonthBlocks(sheet);
        True(blocks.Any(block => block.Month == 7), "Пропал исходный месяц");
        True(blocks.Any(block => block.Month == 8), "Новый месяц не создан");
        var august = blocks.First(block => block.Month == 8);
        Equal(90d, sheet.Cell(3, august.ScenarioColumns[1]).GetDouble());
    });

    Run("Автодобавление сотрудника", () =>
    {
        var analysis = Path.Combine(root, "Анализ_add.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_add.xlsx");
        WriteAnalysis(analysis, "Август", ["Алексеев Алексей Алексеевич", "Яковлев Яков Яковлевич"]);
        WriteExport(export,
        [
            ("03.08.2026", "Сценарий А", "Алексеев Алексей Алексеевич", 90d),
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 91d),
            ("03.08.2026", "Сценарий Б", "Иванов Иван Иванович", 82d),
            ("03.08.2026", "Сценарий А", "Яковлев Яков Яковлевич", 80d)
        ]);
        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "add-output"), Path.Combine(root, "config"), false, true));
        Equal(1, result.NewPeople);
        Equal(1, result.AddedPeople);
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        Equal("Иванов Иван Иванович", sheet.Cell("A4").GetString());
        Equal(91d, sheet.Cell("C4").GetDouble());
        Equal(82d, sheet.Cell("D4").GetDouble());
    });

    Run("Единый порог прохождения", () =>
    {
        var analysis = Path.Combine(root, "Анализ_threshold_global.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_threshold_global.xlsx");
        var config = Path.Combine(root, "threshold-global-config");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 60d)]);
        WriteThresholdSettings(config, false, 75d, new Dictionary<string, double>());

        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "threshold-global-output"), config, false, false));
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        Equal(XLColor.FromHtml("#FFF2CC").Color.ToArgb() & 0xFFFFFF,
            sheet.Cell("C3").Style.Fill.BackgroundColor.Color.ToArgb() & 0xFFFFFF);
        True(sheet.Cell("C4").FormulaA1.Contains(">=75", StringComparison.Ordinal),
            "Итоговая формула не использует единый порог 75%");
    });

    Run("Индивидуальный порог установки", () =>
    {
        var analysis = Path.Combine(root, "Анализ_threshold_installation.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_threshold_installation.xlsx");
        var config = Path.Combine(root, "threshold-installation-config");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 60d)]);
        WriteThresholdSettings(config, true, 90d,
            new Dictionary<string, double> { ["ГФУ-2"] = 50d });

        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "threshold-installation-output"), config, false, false));
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        True(sheet.Cell("C4").FormulaA1.Contains(">=50", StringComparison.Ordinal),
            "Итоговая формула не использует индивидуальный порог 50%");
        True(!sheet.Cell("C4").FormulaA1.Contains(">=90", StringComparison.Ordinal),
            "Для установки ошибочно применён общий порог");
    });

    Run("Понятная ошибка занятого файла результата", () =>
    {
        var analysis = Path.Combine(root, "Анализ_locked.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_locked.xlsx");
        var output = Path.Combine(root, "locked-output");
        Directory.CreateDirectory(output);
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 90d)]);
        var resultFile = Path.Combine(output, "Анализ_прохождения_результат.xlsx");
        File.WriteAllBytes(resultFile, [0]);

        using var lockStream = new FileStream(
            resultFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        try
        {
            new AnalyzerService().Run(new AnalyzerRequest(
                analysis, export, output, Path.Combine(root, "config"), false, false));
            throw new InvalidOperationException("Ожидалась ошибка занятого файла");
        }
        catch (AnalysisFileAccessException exception)
        {
            True(exception.Message.Contains("Закройте файл", StringComparison.Ordinal),
                "Ошибка не содержит понятной инструкции пользователю");
            Equal(resultFile, exception.FilePath);
        }
    });

    Run("Изолированный процесс анализа", () =>
    {
        var analysis = Path.Combine(root, "Анализ_worker.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_worker.xlsx");
        var output = Path.Combine(root, "worker-output");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 90d)]);

        var request = new AnalyzerRequest(
            analysis, export, output, Path.Combine(root, "config"), false, false);
        var payload = Convert.ToBase64String(
            Encoding.UTF8.GetBytes(JsonSerializer.Serialize(request)));
        var executable = Path.Combine(AppContext.BaseDirectory, "Analizator.exe");
        True(File.Exists(executable), "Не найден исполняемый файл изолированного процесса");

        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        startInfo.ArgumentList.Add("--analysis-worker");
        startInfo.ArgumentList.Add(payload);
        startInfo.Environment["DOTNET_EnableDiagnostics_Debugger"] = "0";
        startInfo.Environment["DOTNET_MODIFIABLE_ASSEMBLIES"] = "none";

        using var process = new Process { StartInfo = startInfo };
        True(process.Start(), "Не удалось запустить изолированный процесс");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("Изолированный процесс не завершился за 30 секунд");
        }

        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        Equal(0, process.ExitCode);
        True(string.IsNullOrWhiteSpace(stderr), "Ошибка изолированного процесса: " + stderr);
        True(stdout.Contains("\"Percent\":100", StringComparison.Ordinal),
            "Изолированный процесс не сообщил о 100% прогресса");
        var resultFile = Path.Combine(output, "Анализ_прохождения_результат.xlsx");
        True(File.Exists(resultFile),
            "Изолированный процесс не создал результат");

        using var lockStream = new FileStream(
            resultFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var failingProcess = new Process { StartInfo = startInfo };
        True(failingProcess.Start(), "Не удалось повторно запустить изолированный процесс");
        var failingStdoutTask = failingProcess.StandardOutput.ReadToEndAsync();
        var failingStderrTask = failingProcess.StandardError.ReadToEndAsync();
        if (!failingProcess.WaitForExit(30_000))
        {
            failingProcess.Kill(entireProcessTree: true);
            throw new TimeoutException("Проверка занятого файла не завершилась за 30 секунд");
        }

        var failingStdout = failingStdoutTask.GetAwaiter().GetResult();
        var failingStderr = failingStderrTask.GetAwaiter().GetResult();
        Equal(1, failingProcess.ExitCode);
        True(string.IsNullOrWhiteSpace(failingStderr),
            "Техническая ошибка попала в поток диагностики: " + failingStderr);
        var errorLine = failingStdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
            .SingleOrDefault(line => line.StartsWith("[ANALIZATOR_ERROR]", StringComparison.Ordinal));
        True(errorLine is not null, "Изолированный процесс не передал понятную ошибку интерфейсу");
        var message = JsonSerializer.Deserialize<string>(errorLine!["[ANALIZATOR_ERROR]".Length..]);
        True(message?.Contains("Закройте файл", StringComparison.Ordinal) == true,
            "Ошибка изолированного процесса не содержит инструкции закрыть файл");
        True(!failingStdout.Contains("System.IO.IOException", StringComparison.Ordinal),
            "Пользователю передаётся технический стек вызовов");
    });
}
finally
{
    try { Directory.Delete(root, true); } catch { }
}

if (failures.Count > 0)
{
    Console.Error.WriteLine($"FAILED ({failures.Count})");
    foreach (var failure in failures)
        Console.Error.WriteLine(failure);
    return 1;
}
Console.WriteLine("OK: все C#-тесты пройдены");
return 0;

void Run(string name, Action test)
{
    try
    {
        test();
        Console.WriteLine("PASS: " + name);
    }
    catch (Exception exception)
    {
        failures.Add($"{name}: {exception.Message}\n{exception}");
        Console.WriteLine("FAIL: " + name);
    }
}

static void WriteAnalysis(string path, string month, IReadOnlyList<string> people)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.AddWorksheet("ГФУ-2");
    sheet.Cell("C1").Value = month;
    sheet.Range("C1:H1").Merge();
    string[] headers = ["ФИО", "Должность", "1", "2", "3", "4", "5", "Посещение КТК"];
    for (var column = 1; column <= headers.Length; column++)
    {
        sheet.Cell(2, column).Value = headers[column - 1];
        sheet.Cell(2, column).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAF7");
    }
    var row = 3;
    foreach (var person in people)
    {
        sheet.Cell(row, 1).Value = person;
        sheet.Cell(row, 2).Value = "Оператор";
        for (var column = 3; column <= 7; column++)
            sheet.Cell(row, column).Value = 0;
        sheet.Cell(row, 8).Value = "Нет";
        row++;
    }
    sheet.Cell(row, 1).Value = "Итого";
    sheet.Cell(row, 3).FormulaA1 = $"COUNT(C3:C{row - 1})";
    sheet.SheetView.FreezeRows(2);
    sheet.SheetView.FreezeColumns(2);
    sheet.Range(2, 1, row - 1, 8).SetAutoFilter();
    workbook.SaveAs(path);
}

static void WriteExport(string path, IReadOnlyList<(string Date, string Scenario, string Person, double Percent)> rows)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.AddWorksheet("ГФУ-2");
    sheet.Cell(1, 1).Value = "Тестовая выгрузка";
    string[] headers = ["Дата", "Тема", "Ф.И.О. по отчёту", "Процент", "Режим"];
    for (var column = 1; column <= headers.Length; column++)
        sheet.Cell(2, column).Value = headers[column - 1];
    var row = 3;
    foreach (var item in rows)
    {
        sheet.Cell(row, 1).Value = item.Date;
        sheet.Cell(row, 2).Value = item.Scenario;
        sheet.Cell(row, 3).Value = item.Person;
        sheet.Cell(row, 4).Value = item.Percent;
        sheet.Cell(row, 5).Value = "Случайный выбор";
        row++;
    }
    workbook.SaveAs(path);
}

static void WriteThresholdSettings(
    string directory,
    bool usePerInstallation,
    double global,
    IReadOnlyDictionary<string, double> installations)
{
    Directory.CreateDirectory(directory);
    File.WriteAllText(Path.Combine(directory, "settings.json"), JsonSerializer.Serialize(new
    {
        passing_threshold = new
        {
            use_per_installation = usePerInstallation,
            global,
            installations
        }
    }));
}

static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"Ожидалось: {expected}; получено: {actual}");
}

static void True(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

file sealed class InlineProgress<T>(Action<T> handler) : IProgress<T>
{
    public void Report(T value) => handler(value);
}

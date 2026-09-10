using ClosedXML.Excel;
using System.Globalization;

namespace Analizator.Core;

public sealed class StatisticsReportService
{
    private const string FontName = "Arial";
    private static readonly XLColor Navy = XLColor.FromHtml("#1F4E78");
    private static readonly XLColor Blue = XLColor.FromHtml("#5B9BD5");
    private static readonly XLColor LightBlue = XLColor.FromHtml("#D9EAF7");
    private static readonly XLColor LightGray = XLColor.FromHtml("#F3F6FA");
    private static readonly XLColor Muted = XLColor.FromHtml("#66758A");
    private static readonly XLColor Green = XLColor.FromHtml("#E2F0D9");
    private static readonly XLColor Amber = XLColor.FromHtml("#FFF2CC");
    private static readonly XLColor Red = XLColor.FromHtml("#FCE4D6");

    public StatisticsReportResult Run(
        StatisticsReportRequest request,
        IProgress<StatisticsReportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ValidateRequest(request);
        progress?.Report(new StatisticsReportProgress(3, "Чтение записей базы данных"));

        var selectedInstallations = request.Installations
            .Select(TextNormalization.NormalizeText)
            .Where(value => value.Length > 0)
            .ToHashSet(StringComparer.Ordinal);
        var database = new TrainingDatabaseService(request.DatabasePath);
        var scopedRows = database.SearchAttempts(
                request.PeriodStart.Date,
                request.PeriodEnd.Date,
                null,
                int.MaxValue)
            .Items
            .Where(item => selectedInstallations.Contains(
                TextNormalization.NormalizeText(item.Installation)))
            .ToArray();

        cancellationToken.ThrowIfCancellationRequested();
        var exactDuplicates = scopedRows.Where(IsExactDuplicate).ToArray();
        var canonicalRows = scopedRows.Where(item => !IsExactDuplicate(item)).ToArray();
        var possibleDuplicates = canonicalRows.Where(IsPossibleDuplicate).ToArray();
        var attempts = request.IncludePossibleDuplicates
            ? canonicalRows
            : canonicalRows.Where(item => !IsPossibleDuplicate(item)).ToArray();
        if (attempts.Length == 0)
            throw new InvalidOperationException(
                "За выбранный период и по выбранным установкам нет записей для статистики.");

        progress?.Report(new StatisticsReportProgress(12, "Расчёт показателей сотрудников"));
        var people = BuildPeople(attempts, request, cancellationToken);
        progress?.Report(new StatisticsReportProgress(22, "Расчёт показателей установок"));
        var installations = BuildInstallations(attempts, request, cancellationToken);
        progress?.Report(new StatisticsReportProgress(31, "Расчёт показателей сценариев"));
        var scenarios = BuildScenarios(attempts, request, cancellationToken);
        progress?.Report(new StatisticsReportProgress(40, "Расчёт динамики по месяцам"));
        var months = BuildMonths(attempts, request);
        var modes = BuildModes(attempts, request);
        var quality = BuildQuality(
            attempts, scopedRows, exactDuplicates.Length, possibleDuplicates.Length, request);

        var outputFile = Path.GetFullPath(request.OutputFile);
        Directory.CreateDirectory(Path.GetDirectoryName(outputFile)
            ?? throw new InvalidOperationException("Не удалось определить папку результата."));

        try
        {
            progress?.Report(new StatisticsReportProgress(48, "Создание сводного листа"));
            using var workbook = new XLWorkbook();
            workbook.Properties.Title = "Подробная статистика прохождения КТК";
            workbook.Properties.Subject =
                $"Фактические данные за {request.PeriodStart:dd.MM.yyyy} — {request.PeriodEnd:dd.MM.yyyy}";
            workbook.Properties.Author = "Анализатор КТК";

            WriteSummarySheet(workbook, attempts, people, installations, scenarios,
                months, modes, exactDuplicates.Length, possibleDuplicates.Length, request);
            progress?.Report(new StatisticsReportProgress(56, "Таблица сотрудников"));
            WritePeopleSheet(workbook, people, request);
            progress?.Report(new StatisticsReportProgress(63, "Таблица установок"));
            WriteInstallationsSheet(workbook, installations, request);
            progress?.Report(new StatisticsReportProgress(70, "Таблица сценариев"));
            WriteScenariosSheet(workbook, scenarios, request);
            progress?.Report(new StatisticsReportProgress(76, "Динамика и режимы"));
            WriteMonthsSheet(workbook, months, request);
            WriteModesSheet(workbook, modes, request);
            progress?.Report(new StatisticsReportProgress(82, "Контроль качества данных"));
            WriteQualitySheet(workbook, quality, request);
            progress?.Report(new StatisticsReportProgress(87, "Добавление исходных записей"));
            WriteSourceSheet(workbook, attempts, request, cancellationToken);
            WriteMethodologySheet(workbook, request);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new StatisticsReportProgress(93, "Сохранение книги Excel"));
            workbook.SaveAs(outputFile);
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new StatisticsReportProgress(97, "Добавление диаграмм"));
            StatisticsChartWriter.AddCharts(
                outputFile,
                months.Select(item => MonthLabel(item.Month)).ToArray(),
                months.Select(item => (double)item.Attempts).ToArray(),
                months.Select(item => (double)item.Passed).ToArray(),
                installations.Select(item => item.Installation).ToArray(),
                installations.Select(item => (double)item.Attempts).ToArray());
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new AnalysisFileAccessException(outputFile, exception);
        }

        progress?.Report(new StatisticsReportProgress(100, "Отчёт сформирован"));
        return new StatisticsReportResult
        {
            OutputFile = outputFile,
            Attempts = attempts.Length,
            People = people.Count,
            Installations = installations.Count,
            Scenarios = scenarios.Count,
            SuccessfulAttempts = attempts.Count(item => IsPassed(item, request.PassingThreshold)),
            ExactDuplicatesExcluded = exactDuplicates.Length,
            PossibleDuplicatesIncluded = request.IncludePossibleDuplicates
                ? possibleDuplicates.Length
                : 0
        };
    }

    private static List<PersonStatisticsRow> BuildPeople(
        IReadOnlyList<DatabaseAttempt> attempts,
        StatisticsReportRequest request,
        CancellationToken cancellationToken)
    {
        var result = new List<PersonStatisticsRow>();
        foreach (var group in attempts.GroupBy(item => PersonKey(item.Installation, item.PersonName)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rows = group.OrderBy(item => item.AttemptedAt).ThenBy(item => item.Id).ToArray();
            var first = rows[0];
            var last = rows[^1];
            var threshold = request.PassingThreshold.GetForInstallation(first.Installation);
            var passed = rows.Count(item => item.Percent >= threshold);
            var scenarioGroups = rows.GroupBy(item => TextNormalization.NormalizeText(item.ScenarioName))
                .Select(items => new
                {
                    Name = items.First().ScenarioName,
                    Average = items.Average(item => item.Percent),
                    Passed = items.Any(item => item.Percent >= threshold)
                })
                .OrderByDescending(item => item.Average)
                .ToArray();
            var activeDates = rows.Select(item => item.AttemptedAt.Date).Distinct().Order().ToArray();
            var gaps = activeDates.Zip(activeDates.Skip(1),
                (left, right) => (right - left).TotalDays).ToArray();
            var durations = rows.Select(item => ParseDurationMinutes(item.Duration))
                .Where(item => item.HasValue).Select(item => item!.Value).ToArray();
            var actions = rows.Where(item => item.Actions.HasValue)
                .Select(item => (double)item.Actions!.Value).ToArray();
            var attemptsToFirstPass = Array.FindIndex(rows, item => item.Percent >= threshold);
            var currentSuccessStreak = rows.Reverse().TakeWhile(item => item.Percent >= threshold).Count();
            var daysSinceLast = Math.Max(0, (request.PeriodEnd.Date - last.AttemptedAt.Date).Days);
            var passRate = passed / (double)rows.Length;
            var deviation = StandardDeviation(rows.Select(item => item.Percent));
            var status = PersonStatus(daysSinceLast, passRate, deviation,
                rows.Length, last.Percent, threshold);

            result.Add(new PersonStatisticsRow(
                first.Installation,
                first.PersonName,
                rows.Length,
                activeDates.Length,
                rows.Select(item => new DateTime(item.AttemptedAt.Year, item.AttemptedAt.Month, 1))
                    .Distinct().Count(),
                scenarioGroups.Length,
                scenarioGroups.Count(item => item.Passed),
                passed,
                passRate,
                rows.Average(item => item.Percent),
                Median(rows.Select(item => item.Percent)),
                rows.Max(item => item.Percent),
                rows.Min(item => item.Percent),
                first.AttemptedAt,
                last.AttemptedAt,
                daysSinceLast,
                gaps.Length == 0 ? null : gaps.Average(),
                last.Percent - first.Percent,
                deviation,
                attemptsToFirstPass < 0 ? null : attemptsToFirstPass + 1,
                currentSuccessStreak,
                durations.Length == 0 ? null : durations.Average(),
                actions.Length == 0 ? null : actions.Average(),
                rows.Count(IsPossibleDuplicate),
                scenarioGroups.First().Name,
                scenarioGroups.Last().Name,
                status));
        }
        return result.OrderBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.PersonName, StringComparer.CurrentCultureIgnoreCase).ToList();
    }

    private static List<StatisticsInstallationRow> BuildInstallations(
        IReadOnlyList<DatabaseAttempt> attempts,
        StatisticsReportRequest request,
        CancellationToken cancellationToken)
    {
        var cutoff = request.PeriodEnd.Date.AddDays(-29);
        var result = new List<StatisticsInstallationRow>();
        foreach (var group in attempts.GroupBy(item => item.Installation,
                     StringComparer.CurrentCultureIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rows = group.ToArray();
            var threshold = request.PassingThreshold.GetForInstallation(group.Key);
            var people = rows.GroupBy(item => TextNormalization.NormalizePersonName(item.PersonName)).ToArray();
            var passed = rows.Count(item => item.Percent >= threshold);
            var durations = rows.Select(item => ParseDurationMinutes(item.Duration))
                .Where(item => item.HasValue).Select(item => item!.Value).ToArray();
            var actions = rows.Where(item => item.Actions.HasValue)
                .Select(item => (double)item.Actions!.Value).ToArray();
            var problemScenario = rows.GroupBy(item => TextNormalization.NormalizeText(item.ScenarioName))
                .Select(items => new
                {
                    Name = items.First().ScenarioName,
                    PassRate = items.Count(item => item.Percent >= threshold) / (double)items.Count(),
                    Average = items.Average(item => item.Percent)
                })
                .OrderBy(item => item.PassRate).ThenBy(item => item.Average).First().Name;

            result.Add(new StatisticsInstallationRow(
                group.Key,
                people.Length,
                people.Count(person => person.Any(item => item.AttemptedAt.Date >= cutoff)),
                rows.Length,
                rows.Select(item => TextNormalization.NormalizeText(item.ScenarioName)).Distinct().Count(),
                passed,
                passed / (double)rows.Length,
                rows.Average(item => item.Percent),
                Median(rows.Select(item => item.Percent)),
                rows.Max(item => item.Percent),
                rows.Min(item => item.Percent),
                rows.Length / (double)people.Length,
                threshold,
                durations.Length == 0 ? null : durations.Average(),
                actions.Length == 0 ? null : actions.Average(),
                rows.Count(IsPossibleDuplicate),
                rows.Min(item => item.AttemptedAt),
                rows.Max(item => item.AttemptedAt),
                problemScenario));
        }
        return result.OrderByDescending(item => item.Attempts)
            .ThenBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase).ToList();
    }

    private static List<ScenarioStatisticsRow> BuildScenarios(
        IReadOnlyList<DatabaseAttempt> attempts,
        StatisticsReportRequest request,
        CancellationToken cancellationToken)
    {
        var result = new List<ScenarioStatisticsRow>();
        foreach (var group in attempts.GroupBy(item =>
                     TextNormalization.NormalizeText(item.Installation) + "\u001F" +
                     TextNormalization.NormalizeText(item.ScenarioName)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rows = group.ToArray();
            var installation = rows[0].Installation;
            var threshold = request.PassingThreshold.GetForInstallation(installation);
            var passed = rows.Count(item => item.Percent >= threshold);
            var participants = rows.Select(item => TextNormalization.NormalizePersonName(item.PersonName))
                .Distinct().Count();
            var durations = rows.Select(item => ParseDurationMinutes(item.Duration))
                .Where(item => item.HasValue).Select(item => item!.Value).ToArray();
            var actions = rows.Where(item => item.Actions.HasValue)
                .Select(item => (double)item.Actions!.Value).ToArray();
            var passRate = passed / (double)rows.Length;
            result.Add(new ScenarioStatisticsRow(
                installation,
                rows[0].ScenarioName,
                participants,
                rows.Length,
                passed,
                passRate,
                rows.Average(item => item.Percent),
                Median(rows.Select(item => item.Percent)),
                rows.Max(item => item.Percent),
                rows.Min(item => item.Percent),
                rows.Length - participants,
                threshold,
                durations.Length == 0 ? null : durations.Average(),
                actions.Length == 0 ? null : actions.Average(),
                rows.Count(IsPossibleDuplicate),
                rows.Min(item => item.AttemptedAt),
                rows.Max(item => item.AttemptedAt),
                ScenarioStatus(rows.Length, passRate)));
        }
        return result.OrderBy(item => item.PassRate)
            .ThenByDescending(item => item.Attempts)
            .ThenBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.Scenario, StringComparer.CurrentCultureIgnoreCase).ToList();
    }

    private static List<MonthStatisticsRow> BuildMonths(
        IReadOnlyList<DatabaseAttempt> attempts,
        StatisticsReportRequest request)
    {
        var grouped = attempts.GroupBy(item => new DateTime(item.AttemptedAt.Year,
                item.AttemptedAt.Month, 1))
            .ToDictionary(group => group.Key, group => group.ToArray());
        var result = new List<MonthStatisticsRow>();
        for (var month = new DateTime(request.PeriodStart.Year, request.PeriodStart.Month, 1);
             month <= request.PeriodEnd.Date;
             month = month.AddMonths(1))
        {
            var rows = grouped.GetValueOrDefault(month) ?? [];
            var passed = rows.Count(item => IsPassed(item, request.PassingThreshold));
            result.Add(new MonthStatisticsRow(
                month,
                rows.Length,
                rows.Select(item => PersonKey(item.Installation, item.PersonName)).Distinct().Count(),
                rows.Select(item => TextNormalization.NormalizeText(item.Installation)).Distinct().Count(),
                rows.Select(item => TextNormalization.NormalizeText(item.Installation) + "\u001F" +
                                    TextNormalization.NormalizeText(item.ScenarioName)).Distinct().Count(),
                passed,
                rows.Length - passed,
                rows.Length == 0 ? null : passed / (double)rows.Length,
                rows.Length == 0 ? null : rows.Average(item => item.Percent),
                rows.Length == 0 ? null : Median(rows.Select(item => item.Percent)),
                rows.Count(IsPossibleDuplicate)));
        }
        return result;
    }

    private static List<ModeStatisticsRow> BuildModes(
        IReadOnlyList<DatabaseAttempt> attempts,
        StatisticsReportRequest request) =>
        attempts.GroupBy(item => string.IsNullOrWhiteSpace(item.Mode) ? "Не указан" : item.Mode.Trim(),
                StringComparer.CurrentCultureIgnoreCase)
            .Select(group =>
            {
                var rows = group.ToArray();
                var passed = rows.Count(item => IsPassed(item, request.PassingThreshold));
                return new ModeStatisticsRow(
                    group.Key,
                    rows.Length,
                    rows.Select(item => PersonKey(item.Installation, item.PersonName)).Distinct().Count(),
                    rows.Select(item => TextNormalization.NormalizeText(item.Installation)).Distinct().Count(),
                    passed,
                    passed / (double)rows.Length,
                    rows.Average(item => item.Percent));
            })
            .OrderByDescending(item => item.Attempts).ToList();

    private static List<QualityStatisticsRow> BuildQuality(
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<DatabaseAttempt> scopedRows,
        int exactDuplicates,
        int possibleDuplicates,
        StatisticsReportRequest request)
    {
        var canonical = scopedRows.Where(item => !IsExactDuplicate(item)).ToArray();
        return
        [
            new("Записей в выбранном диапазоне до обработки повторов", scopedRows.Count,
                "Все строки выбранных установок, включая точные повторы"),
            new("Точных повторов исключено", exactDuplicates,
                "Точные повторы не участвуют ни в одном показателе"),
            new("Возможных повторов найдено", possibleDuplicates,
                request.IncludePossibleDuplicates
                    ? "Включены в показатели"
                    : "Исключены из показателей по выбранной настройке"),
            new("Записей без режима", attempts.Count(item => string.IsNullOrWhiteSpace(item.Mode)),
                "Чаще встречается в старых Excel-выгрузках"),
            new("Записей без длительности", attempts.Count(item => string.IsNullOrWhiteSpace(item.Duration)),
                "Поле недоступно в части источников"),
            new("Длительность не удалось распознать", attempts.Count(item =>
                    !string.IsNullOrWhiteSpace(item.Duration) && !ParseDurationMinutes(item.Duration).HasValue),
                "Такие значения не входят в среднее время"),
            new("Записей без количества действий", attempts.Count(item => !item.Actions.HasValue),
                "Такие значения не входят в среднее количество действий"),
            new("Записей без проекта", attempts.Count(item => string.IsNullOrWhiteSpace(item.Project)),
                "Поле необязательно для текущего формата"),
            new("Записей без роли", attempts.Count(item => string.IsNullOrWhiteSpace(item.Role)),
                "Поле необязательно для текущего формата"),
            new("Записей из JSON", attempts.Count(item =>
                    Path.GetExtension(item.SourceFile).Equals(".json", StringComparison.OrdinalIgnoreCase)),
                "JSON обычно содержит наиболее подробные данные"),
            new("Записей из Excel", attempts.Count(item =>
                    !Path.GetExtension(item.SourceFile).Equals(".json", StringComparison.OrdinalIgnoreCase)),
                "Текущий и старый форматы Excel"),
            new("Результатов 100%", attempts.Count(item => Math.Abs(item.Percent - 100d) < 0.0001),
                "Количество полностью выполненных прохождений"),
            new("Результатов 0%", attempts.Count(item => Math.Abs(item.Percent) < 0.0001),
                "Требуют внимания при проверке исходных данных"),
            new("Исходных файлов", canonical.Select(item => item.SourceFile)
                    .Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                "Уникальные имена источников в выбранном диапазоне")
        ];
    }

    private static void WriteSummarySheet(
        XLWorkbook workbook,
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<PersonStatisticsRow> people,
        IReadOnlyList<StatisticsInstallationRow> installations,
        IReadOnlyList<ScenarioStatisticsRow> scenarios,
        IReadOnlyList<MonthStatisticsRow> months,
        IReadOnlyList<ModeStatisticsRow> modes,
        int exactDuplicates,
        int possibleDuplicates,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Сводка");
        PrepareSheet(sheet, "Подробная статистика прохождения КТК", Context(request), 20);
        var passed = attempts.Count(item => IsPassed(item, request.PassingThreshold));
        var cutoff = request.PeriodEnd.Date.AddDays(-29);
        var activeLast30 = attempts.Where(item => item.AttemptedAt.Date >= cutoff)
            .Select(item => PersonKey(item.Installation, item.PersonName)).Distinct().Count();
        var kpis = new (string Label, object Value, string Format)[]
        {
            ("Попытки", attempts.Count, "#,##0"),
            ("Сотрудники", people.Count, "#,##0"),
            ("Установки", installations.Count, "#,##0"),
            ("Сценарии", scenarios.Count, "#,##0"),
            ("Успешность", passed / (double)attempts.Count, "0.0%"),
            ("Средний результат", attempts.Average(item => item.Percent), "0.0\"%\""),
            ("Медианный результат", Median(attempts.Select(item => item.Percent)), "0.0\"%\""),
            ("Активны последние 30 дней", activeLast30, "#,##0"),
            ("Точных повторов исключено", exactDuplicates, "#,##0"),
            ("Возможных повторов в базе", possibleDuplicates, "#,##0")
        };
        for (var index = 0; index < kpis.Length; index++)
        {
            var column = index + 1;
            sheet.Cell(6, column).Value = kpis[index].Label;
            sheet.Cell(7, column).Value = XLCellValue.FromObject(kpis[index].Value);
            sheet.Cell(7, column).Style.NumberFormat.Format = kpis[index].Format;
            sheet.Range(6, column, 8, column).Style.Fill.BackgroundColor =
                index % 2 == 0 ? LightBlue : LightGray;
            sheet.Range(6, column, 8, column).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Range(6, column, 8, column).Style.Border.OutsideBorderColor = XLColor.FromHtml("#D6DEE8");
            sheet.Cell(6, column).Style.Font.Bold = true;
            sheet.Cell(6, column).Style.Font.FontSize = 9;
            sheet.Cell(6, column).Style.Alignment.WrapText = true;
            sheet.Cell(7, column).Style.Font.Bold = true;
            sheet.Cell(7, column).Style.Font.FontSize = 15;
            sheet.Cell(7, column).Style.Font.FontColor = Navy;
            sheet.Cell(7, column).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Column(column).Width = index is 7 or 8 or 9 ? 22 : 17;
        }
        sheet.Row(6).Height = 31;
        sheet.Row(7).Height = 27;
        sheet.Row(8).Height = 7;

        sheet.Cell(10, 1).Value = "Последние месяцы";
        StyleSectionTitle(sheet.Cell(10, 1));
        string[] headers =
            ["Месяц", "Попытки", "Сотрудники", "Успешно", "Неуспешно", "Успешность", "Средний результат"];
        WriteHeaders(sheet, 11, headers);
        var recentMonths = months.TakeLast(12).ToArray();
        for (var index = 0; index < recentMonths.Length; index++)
        {
            var row = 12 + index;
            var item = recentMonths[index];
            sheet.Cell(row, 1).Value = MonthLabel(item.Month);
            sheet.Cell(row, 2).Value = item.Attempts;
            sheet.Cell(row, 3).Value = item.ActivePeople;
            sheet.Cell(row, 4).Value = item.Passed;
            sheet.Cell(row, 5).Value = item.Failed;
            WriteNullable(sheet.Cell(row, 6), item.PassRate);
            WriteNullable(sheet.Cell(row, 7), item.AveragePercent);
        }
        var recentLastRow = 11 + recentMonths.Length;
        StyleTable(sheet, 11, recentLastRow, headers.Length, "SummaryMonths", 0);
        sheet.Range(12, 6, recentLastRow, 6).Style.NumberFormat.Format = "0.0%";
        sheet.Range(12, 7, recentLastRow, 7).Style.NumberFormat.Format = "0.0\"%\"";
        ApplyRateFormatting(sheet.Range(12, 6, recentLastRow, 6));
        SetWidths(sheet, [18, 12, 14, 12, 13, 14, 19]);

        var modeStart = recentLastRow + 3;
        sheet.Cell(modeStart, 1).Value = "Режимы прохождения";
        StyleSectionTitle(sheet.Cell(modeStart, 1));
        string[] modeHeaders = ["Режим", "Попытки", "Сотрудники", "Успешность", "Средний результат"];
        WriteHeaders(sheet, modeStart + 1, modeHeaders);
        for (var index = 0; index < modes.Count; index++)
        {
            var row = modeStart + 2 + index;
            var item = modes[index];
            sheet.Cell(row, 1).Value = item.Mode;
            sheet.Cell(row, 2).Value = item.Attempts;
            sheet.Cell(row, 3).Value = item.People;
            sheet.Cell(row, 4).Value = item.PassRate;
            sheet.Cell(row, 5).Value = item.AveragePercent;
        }
        var modeLastRow = modeStart + 1 + modes.Count;
        StyleTable(sheet, modeStart + 1, modeLastRow, modeHeaders.Length, "SummaryModes", 0);
        sheet.Range(modeStart + 2, 4, modeLastRow, 4).Style.NumberFormat.Format = "0.0%";
        sheet.Range(modeStart + 2, 5, modeLastRow, 5).Style.NumberFormat.Format = "0.0\"%\"";
        ApplyRateFormatting(sheet.Range(modeStart + 2, 4, modeLastRow, 4));

        sheet.Column(11).Width = 24;
        for (var column = 12; column <= 20; column++)
            sheet.Column(column).Width = 11;
        sheet.SheetView.FreezeRows(4);
    }

    private static void WritePeopleSheet(
        XLWorkbook workbook,
        IReadOnlyList<PersonStatisticsRow> rows,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Сотрудники");
        PrepareSheet(sheet, "Статистика по сотрудникам", Context(request), 27);
        string[] headers =
        [
            "Установка", "ФИО", "Попытки", "Активных дней", "Активных месяцев",
            "Сценарии", "Пройдено сценариев", "Успешных попыток", "Успешность",
            "Средний результат", "Медиана", "Лучший", "Худший", "Первая попытка",
            "Последняя попытка", "Дней с последней", "Средний перерыв, дней",
            "Изменение результата", "Стандартное отклонение", "Попыток до первого успеха",
            "Текущая серия успехов", "Среднее время, мин", "Среднее действий",
            "Возможные повторы", "Сильный сценарий", "Слабый сценарий", "Статус"
        ];
        WriteHeaders(sheet, 6, headers);
        for (var index = 0; index < rows.Count; index++)
        {
            var row = 7 + index;
            var item = rows[index];
            object?[] values =
            [
                item.Installation, item.PersonName, item.Attempts, item.ActiveDays, item.ActiveMonths,
                item.Scenarios, item.PassedScenarios, item.PassedAttempts, item.PassRate,
                item.AveragePercent, item.MedianPercent, item.BestPercent, item.WorstPercent,
                item.FirstAttempt, item.LastAttempt, item.DaysSinceLast, item.AverageGapDays,
                item.Improvement, item.StandardDeviation, item.AttemptsToFirstPass,
                item.CurrentSuccessStreak, item.AverageDurationMinutes, item.AverageActions,
                item.PossibleDuplicates, item.StrongestScenario, item.WeakestScenario, item.Status
            ];
            WriteRow(sheet, row, values);
            StyleStatus(sheet.Cell(row, 27), item.Status);
        }
        var lastRow = 6 + rows.Count;
        StyleTable(sheet, 6, lastRow, headers.Length, "PeopleStatistics", 2);
        sheet.Range(7, 9, lastRow, 9).Style.NumberFormat.Format = "0.0%";
        sheet.Range(7, 10, lastRow, 13).Style.NumberFormat.Format = "0.0\"%\"";
        sheet.Range(7, 14, lastRow, 15).Style.DateFormat.Format = "dd.MM.yyyy HH:mm";
        sheet.Range(7, 17, lastRow, 19).Style.NumberFormat.Format = "0.0";
        sheet.Range(7, 22, lastRow, 23).Style.NumberFormat.Format = "0.0";
        ApplyRateFormatting(sheet.Range(7, 9, lastRow, 9));
        SetWidths(sheet,
            [18, 31, 10, 13, 15, 11, 17, 17, 13, 18, 11, 11, 11, 18, 18, 16,
             19, 20, 21, 21, 19, 18, 17, 18, 42, 42, 27]);
    }

    private static void WriteInstallationsSheet(
        XLWorkbook workbook,
        IReadOnlyList<StatisticsInstallationRow> rows,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Установки");
        PrepareSheet(sheet, "Статистика по установкам", Context(request), 19);
        string[] headers =
        [
            "Установка", "Сотрудники", "Активны за 30 дней", "Попытки", "Сценарии",
            "Успешных", "Успешность", "Средний результат", "Медиана", "Лучший", "Худший",
            "Попыток на сотрудника", "Порог", "Среднее время, мин", "Среднее действий",
            "Возможные повторы", "Первая запись", "Последняя запись", "Наиболее сложный сценарий"
        ];
        WriteHeaders(sheet, 6, headers);
        for (var index = 0; index < rows.Count; index++)
        {
            var item = rows[index];
            WriteRow(sheet, 7 + index,
            [
                item.Installation, item.People, item.ActiveLast30Days, item.Attempts, item.Scenarios,
                item.Passed, item.PassRate, item.AveragePercent, item.MedianPercent, item.BestPercent,
                item.WorstPercent, item.AttemptsPerPerson, item.Threshold, item.AverageDurationMinutes,
                item.AverageActions, item.PossibleDuplicates, item.FirstAttempt, item.LastAttempt,
                item.ProblemScenario
            ]);
        }
        var lastRow = 6 + rows.Count;
        StyleTable(sheet, 6, lastRow, headers.Length, "InstallationStatistics", 1);
        sheet.Range(7, 7, lastRow, 7).Style.NumberFormat.Format = "0.0%";
        sheet.Range(7, 8, lastRow, 13).Style.NumberFormat.Format = "0.0\"%\"";
        sheet.Range(7, 12, lastRow, 12).Style.NumberFormat.Format = "0.0";
        sheet.Range(7, 14, lastRow, 15).Style.NumberFormat.Format = "0.0";
        sheet.Range(7, 17, lastRow, 18).Style.DateFormat.Format = "dd.MM.yyyy HH:mm";
        ApplyRateFormatting(sheet.Range(7, 7, lastRow, 7));
        SetWidths(sheet,
            [21, 14, 20, 12, 12, 13, 14, 19, 12, 12, 12, 21, 12, 20, 18, 19,
             18, 18, 47]);
    }

    private static void WriteScenariosSheet(
        XLWorkbook workbook,
        IReadOnlyList<ScenarioStatisticsRow> rows,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Сценарии");
        PrepareSheet(sheet, "Статистика по сценариям", Context(request), 18);
        string[] headers =
        [
            "Установка", "Сценарий", "Сотрудники", "Попытки", "Успешных", "Успешность",
            "Средний результат", "Медиана", "Лучший", "Худший", "Повторные попытки", "Порог",
            "Среднее время, мин", "Среднее действий", "Возможные повторы", "Первая запись",
            "Последняя запись", "Статус"
        ];
        WriteHeaders(sheet, 6, headers);
        for (var index = 0; index < rows.Count; index++)
        {
            var item = rows[index];
            WriteRow(sheet, 7 + index,
            [
                item.Installation, item.Scenario, item.People, item.Attempts, item.Passed,
                item.PassRate, item.AveragePercent, item.MedianPercent, item.BestPercent,
                item.WorstPercent, item.RepeatAttempts, item.Threshold, item.AverageDurationMinutes,
                item.AverageActions, item.PossibleDuplicates, item.FirstAttempt, item.LastAttempt,
                item.Status
            ]);
            StyleStatus(sheet.Cell(7 + index, 18), item.Status);
        }
        var lastRow = 6 + rows.Count;
        StyleTable(sheet, 6, lastRow, headers.Length, "ScenarioStatistics", 2);
        sheet.Range(7, 6, lastRow, 6).Style.NumberFormat.Format = "0.0%";
        sheet.Range(7, 7, lastRow, 12).Style.NumberFormat.Format = "0.0\"%\"";
        sheet.Range(7, 13, lastRow, 14).Style.NumberFormat.Format = "0.0";
        sheet.Range(7, 16, lastRow, 17).Style.DateFormat.Format = "dd.MM.yyyy HH:mm";
        ApplyRateFormatting(sheet.Range(7, 6, lastRow, 6));
        SetWidths(sheet,
            [20, 48, 13, 11, 12, 14, 19, 12, 12, 12, 19, 12, 20, 18, 19, 18, 18, 23]);
    }

    private static void WriteMonthsSheet(
        XLWorkbook workbook,
        IReadOnlyList<MonthStatisticsRow> rows,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Динамика");
        PrepareSheet(sheet, "Динамика по месяцам", Context(request), 11);
        string[] headers =
        [
            "Месяц", "Попытки", "Сотрудники", "Установки", "Сценарии", "Успешных",
            "Неуспешных", "Успешность", "Средний результат", "Медиана", "Возможные повторы"
        ];
        WriteHeaders(sheet, 6, headers);
        for (var index = 0; index < rows.Count; index++)
        {
            var item = rows[index];
            WriteRow(sheet, 7 + index,
            [
                MonthLabel(item.Month), item.Attempts, item.ActivePeople, item.Installations,
                item.Scenarios, item.Passed, item.Failed, item.PassRate, item.AveragePercent,
                item.MedianPercent, item.PossibleDuplicates
            ]);
        }
        var lastRow = 6 + rows.Count;
        StyleTable(sheet, 6, lastRow, headers.Length, "MonthlyStatistics", 1);
        sheet.Range(7, 8, lastRow, 8).Style.NumberFormat.Format = "0.0%";
        sheet.Range(7, 9, lastRow, 10).Style.NumberFormat.Format = "0.0\"%\"";
        ApplyRateFormatting(sheet.Range(7, 8, lastRow, 8));
        SetWidths(sheet, [18, 12, 14, 13, 13, 13, 15, 14, 19, 12, 19]);
    }

    private static void WriteModesSheet(
        XLWorkbook workbook,
        IReadOnlyList<ModeStatisticsRow> rows,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Режимы");
        PrepareSheet(sheet, "Статистика по режимам прохождения", Context(request), 7);
        string[] headers =
            ["Режим", "Попытки", "Сотрудники", "Установки", "Успешных", "Успешность", "Средний результат"];
        WriteHeaders(sheet, 6, headers);
        for (var index = 0; index < rows.Count; index++)
        {
            var item = rows[index];
            WriteRow(sheet, 7 + index,
                [item.Mode, item.Attempts, item.People, item.Installations, item.Passed,
                 item.PassRate, item.AveragePercent]);
        }
        var lastRow = 6 + rows.Count;
        StyleTable(sheet, 6, lastRow, headers.Length, "ModeStatistics", 1);
        sheet.Range(7, 6, lastRow, 6).Style.NumberFormat.Format = "0.0%";
        sheet.Range(7, 7, lastRow, 7).Style.NumberFormat.Format = "0.0\"%\"";
        ApplyRateFormatting(sheet.Range(7, 6, lastRow, 6));
        SetWidths(sheet, [28, 13, 14, 14, 13, 14, 19]);
    }

    private static void WriteQualitySheet(
        XLWorkbook workbook,
        IReadOnlyList<QualityStatisticsRow> rows,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Качество данных");
        PrepareSheet(sheet, "Качество исходных данных", Context(request), 3);
        WriteHeaders(sheet, 6, ["Показатель", "Количество", "Пояснение"]);
        for (var index = 0; index < rows.Count; index++)
        {
            WriteRow(sheet, 7 + index,
                [rows[index].Metric, rows[index].Value, rows[index].Note]);
            if (rows[index].Value > 0 && rows[index].Metric.Contains("не удалось", StringComparison.OrdinalIgnoreCase))
                sheet.Cell(7 + index, 2).Style.Fill.BackgroundColor = Red;
        }
        var lastRow = 6 + rows.Count;
        StyleTable(sheet, 6, lastRow, 3, "DataQuality", 1);
        SetWidths(sheet, [52, 16, 72]);
        sheet.Range(7, 3, lastRow, 3).Style.Alignment.WrapText = true;
    }

    private static void WriteSourceSheet(
        XLWorkbook workbook,
        IReadOnlyList<DatabaseAttempt> attempts,
        StatisticsReportRequest request,
        CancellationToken cancellationToken)
    {
        var sheet = workbook.AddWorksheet("Исходные данные");
        PrepareSheet(sheet, "Исходные записи отчёта", Context(request), 20);
        string[] headers =
        [
            "ID", "Установка", "ФИО", "Дата", "Время", "Сценарий", "Процент", "Порог",
            "Пройдено", "Режим", "Длительность", "Длительность, мин", "Действий", "Проект",
            "Роль", "Зачтено", "Источник", "Лист источника", "Строка", "Возможный повтор"
        ];
        WriteHeaders(sheet, 6, headers);
        for (var index = 0; index < attempts.Count; index++)
        {
            if ((index & 255) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            var item = attempts[index];
            var threshold = request.PassingThreshold.GetForInstallation(item.Installation);
            WriteRow(sheet, 7 + index,
            [
                item.Id, item.Installation, item.PersonName, item.AttemptedAt.Date,
                item.AttemptedAt.TimeOfDay, item.ScenarioName, item.Percent, threshold,
                item.Percent >= threshold ? "Да" : "Нет", item.Mode, item.Duration,
                ParseDurationMinutes(item.Duration), item.Actions, item.Project, item.Role,
                item.Credited, item.SourceFile, item.SourceSheet, item.SourceRow,
                IsPossibleDuplicate(item) ? "Да" : "Нет"
            ]);
        }
        var lastRow = 6 + attempts.Count;
        StyleTable(sheet, 6, lastRow, headers.Length, "StatisticsSourceData", 3);
        sheet.Range(7, 4, lastRow, 4).Style.DateFormat.Format = "dd.MM.yyyy";
        sheet.Range(7, 5, lastRow, 5).Style.NumberFormat.Format = "hh:mm:ss";
        sheet.Range(7, 7, lastRow, 8).Style.NumberFormat.Format = "0.0\"%\"";
        sheet.Range(7, 12, lastRow, 12).Style.NumberFormat.Format = "0.0";
        SetWidths(sheet,
            [10, 20, 31, 13, 12, 48, 13, 12, 13, 23, 16, 19, 12, 22, 20, 13, 31, 23, 11, 20]);
    }

    private static void WriteMethodologySheet(
        XLWorkbook workbook,
        StatisticsReportRequest request)
    {
        var sheet = workbook.AddWorksheet("Методика");
        PrepareSheet(sheet, "Методика расчёта", Context(request), 6);
        sheet.Cell(6, 1).Value = "Правило";
        sheet.Cell(6, 2).Value = "Описание";
        var rules = new (string, string)[]
        {
            ("Успешная попытка", "Результат равен порогу установки или превышает его."),
            ("Успешность", "Количество успешных попыток, делённое на общее количество учитываемых попыток."),
            ("Медиана", "Центральное значение упорядоченного набора результатов."),
            ("Стандартное отклонение", "Разброс результатов сотрудника. Рассчитывается по всей выбранной совокупности."),
            ("Изменение результата", "Результат последней попытки минус результат первой попытки в выбранном периоде."),
            ("Средний перерыв", "Среднее число дней между различными активными датами сотрудника."),
            ("Повторные попытки", "Количество попыток сценария сверх одной попытки на каждого сотрудника."),
            ("Активность за 30 дней", "Наличие хотя бы одной записи за последние 30 календарных дней выбранного периода."),
            ("Точные повторы", "Всегда исключаются из статистики."),
            ("Возможные повторы", request.IncludePossibleDuplicates
                ? "Включены в расчёты по выбранной настройке."
                : "Исключены из расчётов по выбранной настройке."),
            ("Нет активности 30+ дней", "Последняя запись сотрудника находится не менее чем за 30 дней до конца отчётного периода."),
            ("Низкая успешность", "Менее половины попыток сотрудника достигли действующего порога."),
            ("Высокий разброс", "Не менее трёх попыток и стандартное отклонение результата от 20 процентных пунктов."),
            ("Ограничение", "Отчёт описывает только сотрудников, встречающихся в базе за выбранный период. Он не заменяет штатный список.")
        };
        for (var index = 0; index < rules.Length; index++)
            WriteRow(sheet, 7 + index, [rules[index].Item1, rules[index].Item2]);
        StyleTable(sheet, 6, 6 + rules.Length, 2, "StatisticsMethodology", 0);
        sheet.Range(7, 2, 6 + rules.Length, 2).Style.Alignment.WrapText = true;
        sheet.Column(1).Width = 32;
        sheet.Column(2).Width = 95;

        sheet.Cell(6, 4).Value = "Установка";
        sheet.Cell(6, 5).Value = "Порог";
        var installations = request.Installations.OrderBy(item => item,
            StringComparer.CurrentCultureIgnoreCase).ToArray();
        for (var index = 0; index < installations.Length; index++)
        {
            sheet.Cell(7 + index, 4).Value = installations[index];
            sheet.Cell(7 + index, 5).Value = request.PassingThreshold.GetForInstallation(installations[index]);
        }
        StyleTable(sheet, 6, 6 + installations.Length, 5, "StatisticsThresholds", 0, 4);
        sheet.Range(7, 5, 6 + installations.Length, 5).Style.NumberFormat.Format = "0.0\"%\"";
        sheet.Column(4).Width = 28;
        sheet.Column(5).Width = 14;
    }

    private static void PrepareSheet(
        IXLWorksheet sheet,
        string title,
        string context,
        int lastColumn)
    {
        sheet.ShowGridLines = false;
        sheet.Cell(2, 1).Value = title;
        sheet.Cell(2, 1).Style.Font.FontName = FontName;
        sheet.Cell(2, 1).Style.Font.FontSize = 16;
        sheet.Cell(2, 1).Style.Font.Bold = true;
        sheet.Cell(2, 1).Style.Font.FontColor = Navy;
        sheet.Cell(3, 1).Value = context;
        sheet.Cell(3, 1).Style.Font.FontName = FontName;
        sheet.Cell(3, 1).Style.Font.FontSize = 10;
        sheet.Cell(3, 1).Style.Font.Italic = true;
        sheet.Cell(3, 1).Style.Font.FontColor = Muted;
        sheet.Range(4, 1, 4, lastColumn).Style.Fill.BackgroundColor = Navy;
        sheet.Row(4).Height = 2.5;
        sheet.Row(5).Height = 9;
        sheet.Range(1, 1, 5, lastColumn).Style.Font.FontName = FontName;
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        sheet.PageSetup.Margins.Left = 0.25;
        sheet.PageSetup.Margins.Right = 0.25;
        sheet.PageSetup.Margins.Top = 0.4;
        sheet.PageSetup.Margins.Bottom = 0.4;
    }

    private static void WriteHeaders(IXLWorksheet sheet, int row, IReadOnlyList<string> headers)
    {
        for (var column = 1; column <= headers.Count; column++)
            sheet.Cell(row, column).Value = headers[column - 1];
    }

    private static void WriteRow(IXLWorksheet sheet, int row, IReadOnlyList<object?> values)
    {
        for (var column = 1; column <= values.Count; column++)
        {
            var value = values[column - 1];
            if (value is null)
                continue;
            sheet.Cell(row, column).Value = XLCellValue.FromObject(value);
        }
    }

    private static void StyleTable(
        IXLWorksheet sheet,
        int headerRow,
        int lastRow,
        int lastColumn,
        string tableName,
        int freezeColumns,
        int firstColumn = 1)
    {
        var range = sheet.Range(headerRow, firstColumn, lastRow, lastColumn);
        var table = range.CreateTable(tableName);
        table.Theme = XLTableTheme.TableStyleMedium2;
        table.ShowAutoFilter = true;
        range.Style.Font.FontName = FontName;
        range.Style.Font.FontSize = 10;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        var header = sheet.Range(headerRow, firstColumn, headerRow, lastColumn);
        header.Style.Fill.BackgroundColor = Navy;
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Font.Bold = true;
        header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        header.Style.Alignment.WrapText = true;
        sheet.Row(headerRow).Height = 32;
        sheet.SheetView.FreezeRows(headerRow);
        if (freezeColumns > 0)
            sheet.SheetView.FreezeColumns(freezeColumns);
    }

    private static void StyleSectionTitle(IXLCell cell)
    {
        cell.Style.Font.FontName = FontName;
        cell.Style.Font.FontSize = 12;
        cell.Style.Font.Bold = true;
        cell.Style.Font.FontColor = Navy;
    }

    private static void SetWidths(IXLWorksheet sheet, IReadOnlyList<double> widths)
    {
        for (var index = 0; index < widths.Count; index++)
            sheet.Column(index + 1).Width = widths[index];
    }

    private static void ApplyRateFormatting(IXLRange range)
    {
        foreach (var cell in range.Cells())
        {
            if (cell.DataType != XLDataType.Number)
                continue;

            var value = cell.GetDouble();
            cell.Style.Fill.BackgroundColor = value switch
            {
                < 0.5 => Red,
                < 0.8 => Amber,
                _ => Green
            };
        }
    }

    private static void StyleStatus(IXLCell cell, string status)
    {
        cell.Style.Fill.BackgroundColor = status switch
        {
            "Без замечаний" or "Высокая успешность" => Green,
            "Недостаточно данных" or "Средняя успешность" or "Высокий разброс" => Amber,
            _ => Red
        };
    }

    private static void WriteNullable(IXLCell cell, double? value)
    {
        if (value.HasValue)
            cell.Value = value.Value;
    }

    private static string Context(StatisticsReportRequest request) =>
        $"Период: {request.PeriodStart:dd.MM.yyyy} — {request.PeriodEnd:dd.MM.yyyy}. " +
        $"Установок: {request.Installations.Length}. Возможные повторы: " +
        (request.IncludePossibleDuplicates ? "включены" : "исключены") + ".";

    private static string MonthLabel(DateTime month) =>
        month.ToString("MMMM yyyy", CultureInfo.GetCultureInfo("ru-RU"));

    private static bool IsPassed(DatabaseAttempt attempt, PassingThresholdSettings settings) =>
        attempt.Percent >= settings.GetForInstallation(attempt.Installation);

    private static bool IsExactDuplicate(DatabaseAttempt attempt) =>
        attempt.DuplicateOfAttemptId.HasValue ||
        string.Equals(attempt.DuplicateKind, "Точный", StringComparison.OrdinalIgnoreCase);

    private static bool IsPossibleDuplicate(DatabaseAttempt attempt) =>
        !IsExactDuplicate(attempt) &&
        (attempt.PossibleDuplicate ||
         string.Equals(attempt.DuplicateKind, "Возможный", StringComparison.OrdinalIgnoreCase));

    private static string PersonKey(string installation, string personName) =>
        TextNormalization.NormalizeText(installation) + "\u001F" +
        TextNormalization.NormalizePersonName(personName);

    private static double Median(IEnumerable<double> values)
    {
        var ordered = values.Order().ToArray();
        if (ordered.Length == 0)
            throw new InvalidOperationException("Невозможно рассчитать медиану пустого набора.");
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2d
            : ordered[middle];
    }

    private static double StandardDeviation(IEnumerable<double> values)
    {
        var items = values.ToArray();
        if (items.Length <= 1)
            return 0;
        var average = items.Average();
        return Math.Sqrt(items.Average(item => Math.Pow(item - average, 2)));
    }

    private static double? ParseDurationMinutes(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (TimeSpan.TryParse(value.Trim(), CultureInfo.InvariantCulture, out var duration) ||
            TimeSpan.TryParse(value.Trim(), CultureInfo.CurrentCulture, out duration))
            return duration.TotalMinutes;
        return null;
    }

    private static string PersonStatus(
        int daysSinceLast,
        double passRate,
        double deviation,
        int attempts,
        double lastPercent,
        double threshold)
    {
        if (daysSinceLast >= 30)
            return "Нет активности 30+ дней";
        if (passRate < 0.5)
            return "Успешность ниже 50%";
        if (attempts >= 3 && deviation >= 20)
            return "Высокий разброс";
        if (lastPercent < threshold)
            return "Последняя попытка ниже порога";
        return "Без замечаний";
    }

    private static string ScenarioStatus(int attempts, double passRate)
    {
        if (attempts < 3)
            return "Недостаточно данных";
        if (passRate < 0.5)
            return "Низкая успешность";
        if (passRate < 0.8)
            return "Средняя успешность";
        return "Высокая успешность";
    }

    private static void ValidateRequest(StatisticsReportRequest request)
    {
        if (!File.Exists(request.DatabasePath))
            throw new FileNotFoundException("Файл базы данных не найден.", request.DatabasePath);
        if (string.IsNullOrWhiteSpace(request.OutputFile))
            throw new ArgumentException("Не указан файл результата.", nameof(request));
        if (request.PeriodEnd.Date < request.PeriodStart.Date)
            throw new ArgumentException("Дата окончания периода не может быть раньше даты начала.");
        if (request.Installations.Length == 0)
            throw new ArgumentException("Выберите хотя бы одну установку.");
    }

    private sealed record PersonStatisticsRow(
        string Installation,
        string PersonName,
        int Attempts,
        int ActiveDays,
        int ActiveMonths,
        int Scenarios,
        int PassedScenarios,
        int PassedAttempts,
        double PassRate,
        double AveragePercent,
        double MedianPercent,
        double BestPercent,
        double WorstPercent,
        DateTime FirstAttempt,
        DateTime LastAttempt,
        int DaysSinceLast,
        double? AverageGapDays,
        double Improvement,
        double StandardDeviation,
        int? AttemptsToFirstPass,
        int CurrentSuccessStreak,
        double? AverageDurationMinutes,
        double? AverageActions,
        int PossibleDuplicates,
        string StrongestScenario,
        string WeakestScenario,
        string Status);

    private sealed record StatisticsInstallationRow(
        string Installation,
        int People,
        int ActiveLast30Days,
        int Attempts,
        int Scenarios,
        int Passed,
        double PassRate,
        double AveragePercent,
        double MedianPercent,
        double BestPercent,
        double WorstPercent,
        double AttemptsPerPerson,
        double Threshold,
        double? AverageDurationMinutes,
        double? AverageActions,
        int PossibleDuplicates,
        DateTime FirstAttempt,
        DateTime LastAttempt,
        string ProblemScenario);

    private sealed record ScenarioStatisticsRow(
        string Installation,
        string Scenario,
        int People,
        int Attempts,
        int Passed,
        double PassRate,
        double AveragePercent,
        double MedianPercent,
        double BestPercent,
        double WorstPercent,
        int RepeatAttempts,
        double Threshold,
        double? AverageDurationMinutes,
        double? AverageActions,
        int PossibleDuplicates,
        DateTime FirstAttempt,
        DateTime LastAttempt,
        string Status);

    private sealed record MonthStatisticsRow(
        DateTime Month,
        int Attempts,
        int ActivePeople,
        int Installations,
        int Scenarios,
        int Passed,
        int Failed,
        double? PassRate,
        double? AveragePercent,
        double? MedianPercent,
        int PossibleDuplicates);

    private sealed record ModeStatisticsRow(
        string Mode,
        int Attempts,
        int People,
        int Installations,
        int Passed,
        double PassRate,
        double AveragePercent);

    private sealed record QualityStatisticsRow(string Metric, int Value, string Note);
}

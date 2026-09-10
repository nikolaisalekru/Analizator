using ClosedXML.Excel;
using System.Globalization;

namespace Analizator.Core;

public sealed class MonitoringReportService
{
    private const string FontName = "Arial";
    private static readonly XLColor Navy = XLColor.FromHtml("#17365D");
    private static readonly XLColor Blue = XLColor.FromHtml("#4472C4");
    private static readonly XLColor LightBlue = XLColor.FromHtml("#D9EAF7");
    private static readonly XLColor LightGray = XLColor.FromHtml("#F3F6FA");
    private static readonly XLColor Muted = XLColor.FromHtml("#66758A");
    private static readonly CultureInfo RussianCulture = CultureInfo.GetCultureInfo("ru-RU");

    public MonitoringReportResult Run(
        MonitoringReportRequest request,
        IProgress<MonitoringReportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ValidateRequest(request);
        progress?.Report(new MonitoringReportProgress(3, "Чтение записей базы данных"));

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
                "За выбранный период и по выбранным установкам нет записей для мониторинга.");

        progress?.Report(new MonitoringReportProgress(16, "Расчёт помесячных показателей"));
        var months = attempts
            .Select(item => new MonthKey(item.AttemptedAt.Year, item.AttemptedAt.Month))
            .Distinct()
            .OrderBy(item => item.Year)
            .ThenBy(item => item.Month)
            .ToArray();
        var quarters = Enumerable.Range(1, 4)
            .Select(quarter =>
            {
                var quarterStart = new DateTime(request.PeriodStart.Year, (quarter - 1) * 3 + 1, 1);
                var quarterEnd = quarterStart.AddMonths(3).AddDays(-1);
                return new QuarterScope(
                    new QuarterKey(request.PeriodStart.Year, quarter),
                    request.PeriodStart.Date <= quarterStart && request.PeriodEnd.Date >= quarterEnd);
            })
            .ToArray();
        var completeQuarters = quarters.Where(item => item.IsComplete).ToArray();

        progress?.Report(new MonitoringReportProgress(30, "Расчёт квартальных показателей"));
        var installationOrder = attempts
            .Select(item => item.Installation)
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();

        var reportOutput = Path.GetFullPath(request.ReportOutputFile);
        var formOutput = Path.GetFullPath(request.FormOutputFile);
        Directory.CreateDirectory(Path.GetDirectoryName(reportOutput)
            ?? throw new InvalidOperationException("Не удалось определить папку подробного отчёта."));
        Directory.CreateDirectory(Path.GetDirectoryName(formOutput)
            ?? throw new InvalidOperationException("Не удалось определить папку формы сбора."));

        try
        {
            progress?.Report(new MonitoringReportProgress(42, "Формирование подробного отчёта"));
            WriteReport(
                reportOutput,
                attempts,
                months,
                completeQuarters,
                installationOrder,
                request,
                exactDuplicates.Length,
                possibleDuplicates.Length,
                cancellationToken);

            progress?.Report(new MonitoringReportProgress(86, "Заполнение формы сбора"));
            var quarterResults = WriteCollectionForm(
                request.FormTemplateFile,
                formOutput,
                attempts,
                quarters,
                request.PeriodStart.Year);

            progress?.Report(new MonitoringReportProgress(100, "Мониторинг сформирован"));
            var periodMetrics = CalculateMetrics(attempts);
            return new MonitoringReportResult
            {
                ReportOutputFile = reportOutput,
                FormOutputFile = formOutput,
                Attempts = attempts.Length,
                People = periodMetrics.Students,
                Installations = installationOrder.Length,
                IndividualScenarios = periodMetrics.IndividualScenarios,
                ExactDuplicatesExcluded = exactDuplicates.Length,
                PossibleDuplicatesIncluded = request.IncludePossibleDuplicates
                    ? possibleDuplicates.Length
                    : 0,
                Quarters = quarterResults
            };
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            var file = File.Exists(request.FormTemplateFile)
                ? formOutput
                : request.FormTemplateFile;
            throw new AnalysisFileAccessException(file, exception);
        }
    }

    private static void WriteReport(
        string outputFile,
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<MonthKey> months,
        IReadOnlyList<QuarterScope> quarters,
        IReadOnlyList<string> installationOrder,
        MonitoringReportRequest request,
        int exactDuplicates,
        int possibleDuplicates,
        CancellationToken cancellationToken)
    {
        using var workbook = new XLWorkbook();
        workbook.Properties.Title = "Мониторинг прохождения КТК";
        workbook.Properties.Subject =
            $"Данные за {request.PeriodStart:dd.MM.yyyy} — {request.PeriodEnd:dd.MM.yyyy}";
        workbook.Properties.Author = "Анализатор КТК";

        WriteSummarySheet(workbook, attempts, months, quarters, installationOrder, request, cancellationToken);
        WriteEmployeesSheet(workbook, attempts, months, quarters, request, cancellationToken);
        WriteDetailsSheet(workbook, attempts, cancellationToken);
        WriteDiagnosticsSheet(workbook, exactDuplicates, possibleDuplicates, request, quarters);
        workbook.CalculateMode = XLCalculateMode.Auto;
        workbook.SaveAs(outputFile);
    }

    private static void WriteSummarySheet(
        XLWorkbook workbook,
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<MonthKey> months,
        IReadOnlyList<QuarterScope> quarters,
        IReadOnlyList<string> installationOrder,
        MonitoringReportRequest request,
        CancellationToken cancellationToken)
    {
        var sheet = workbook.AddWorksheet("Статистика");
        sheet.ShowGridLines = false;
        var lastColumn = installationOrder.Count + 2;
        sheet.Cell(1, 1).Value = "МОНИТОРИНГ КТК";
        sheet.Range(1, 1, 1, lastColumn).Merge();
        sheet.Cell(1, 1).Style.Font.FontName = FontName;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontColor = Navy;
        sheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        sheet.Row(1).Height = 28;

        sheet.Cell(2, 1).Value = "Показатель";
        sheet.Cell(2, 2).Value = "ВСЕ УСТАНОВКИ";
        for (var index = 0; index < installationOrder.Count; index++)
            sheet.Cell(2, index + 3).Value = installationOrder[index];
        StyleHeader(sheet.Range(2, 1, 2, lastColumn));

        var row = 3;
        row = WriteSection(sheet, row, lastColumn, "ПОМЕСЯЧНАЯ СТАТИСТИКА");
        foreach (var month in months)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var monthRows = attempts.Where(item =>
                item.AttemptedAt.Year == month.Year && item.AttemptedAt.Month == month.Month).ToArray();
            row = WriteMetricBlock(
                sheet,
                row,
                lastColumn,
                MonthLabel(month),
                monthRows,
                installationOrder);
        }

        if (quarters.Count > 0)
        {
            row = WriteSection(sheet, row, lastColumn, "КВАРТАЛЬНАЯ СТАТИСТИКА");
            foreach (var quarter in quarters)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var quarterRows = attempts.Where(item =>
                    item.AttemptedAt.Year == quarter.Key.Year &&
                    QuarterOf(item.AttemptedAt.Month) == quarter.Key.Quarter).ToArray();
                row = WriteMetricBlock(
                    sheet,
                    row,
                    lastColumn,
                    QuarterLabel(quarter.Key),
                    quarterRows,
                    installationOrder);
            }
        }

        row = WriteSection(sheet, row, lastColumn,
            $"ИТОГ ЗА {request.PeriodStart:dd.MM.yyyy} — {request.PeriodEnd:dd.MM.yyyy}");
        WriteMetricRows(sheet, row, attempts, installationOrder, useSigmaLabels: true);

        sheet.Column(1).Width = 39;
        sheet.Column(2).Width = 19;
        for (var column = 3; column <= lastColumn; column++)
            sheet.Column(column).Width = 18;
        sheet.Range(1, 1, sheet.LastRowUsed()!.RowNumber(), lastColumn).Style.Font.FontName = FontName;
        sheet.Range(3, 2, sheet.LastRowUsed()!.RowNumber(), lastColumn)
            .Style.NumberFormat.Format = "#,##0";
        sheet.SheetView.FreezeRows(2);
        sheet.SheetView.FreezeColumns(2);
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        sheet.PageSetup.FitToPages(1, 0);
    }

    private static int WriteMetricBlock(
        IXLWorksheet sheet,
        int row,
        int lastColumn,
        string title,
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<string> installationOrder)
    {
        sheet.Cell(row, 1).Value = title;
        var titleRange = sheet.Range(row, 1, row, lastColumn);
        titleRange.Style.Fill.BackgroundColor = LightBlue;
        titleRange.Style.Font.FontColor = Navy;
        titleRange.Style.Font.Bold = true;
        titleRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        titleRange.Style.Border.BottomBorderColor = Blue;
        row++;
        WriteMetricRows(sheet, row, attempts, installationOrder, useSigmaLabels: false);
        return row + 5;
    }

    private static void WriteMetricRows(
        IXLWorksheet sheet,
        int startRow,
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<string> installationOrder,
        bool useSigmaLabels)
    {
        var labels = useSigmaLabels
            ? new[]
            {
                "Σ уникальных сотрудников",
                "Σ занятий (отчётов)",
                "Σ уникальных сценариев",
                "Σ индивидуальных сценариев"
            }
            : new[]
            {
                "Число студентов",
                "Число занятий (отчётов)",
                "Число уникальных сценариев",
                "Число индивидуальных сценариев"
            };
        var global = CalculateMetrics(attempts);
        var globalValues = MetricValues(global);
        for (var offset = 0; offset < labels.Length; offset++)
        {
            sheet.Cell(startRow + offset, 1).Value = labels[offset];
            sheet.Cell(startRow + offset, 2).Value = globalValues[offset];
            sheet.Cell(startRow + offset, 2).Style.Font.Bold = true;
        }

        for (var index = 0; index < installationOrder.Count; index++)
        {
            var installation = installationOrder[index];
            var installationRows = attempts.Where(item => string.Equals(
                item.Installation,
                installation,
                StringComparison.CurrentCultureIgnoreCase)).ToArray();
            var values = MetricValues(CalculateMetrics(installationRows));
            for (var offset = 0; offset < values.Length; offset++)
                sheet.Cell(startRow + offset, index + 3).Value = values[offset];
        }

        var body = sheet.Range(startRow, 1, startRow + 3, installationOrder.Count + 2);
        body.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
        body.Style.Border.BottomBorderColor = XLColor.FromHtml("#D9E2F3");
        for (var offset = 0; offset < 4; offset++)
            if (offset % 2 == 1)
                sheet.Range(startRow + offset, 1, startRow + offset, installationOrder.Count + 2)
                    .Style.Fill.BackgroundColor = LightGray;
    }

    private static int WriteSection(IXLWorksheet sheet, int row, int lastColumn, string title)
    {
        sheet.Cell(row, 1).Value = title;
        var range = sheet.Range(row, 1, row, lastColumn);
        range.Style.Fill.BackgroundColor = Navy;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Font.Bold = true;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        sheet.Row(row).Height = 22;
        return row + 1;
    }

    private static void WriteEmployeesSheet(
        XLWorkbook workbook,
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<MonthKey> months,
        IReadOnlyList<QuarterScope> quarters,
        MonitoringReportRequest request,
        CancellationToken cancellationToken)
    {
        var sheet = workbook.AddWorksheet("Сотрудники");
        sheet.ShowGridLines = false;
        var headers = new List<string> { "Установка", "ФИО" };
        foreach (var month in months)
        {
            headers.Add($"{MonthLabel(month)} · занятия");
            headers.Add($"{MonthLabel(month)} · инд. сценарии");
        }
        foreach (var quarter in quarters)
        {
            headers.Add($"{QuarterLabel(quarter.Key)} · занятия");
            headers.Add($"{QuarterLabel(quarter.Key)} · инд. сценарии");
        }
        headers.Add($"{request.PeriodStart:dd.MM.yyyy}–{request.PeriodEnd:dd.MM.yyyy} · занятия");
        headers.Add($"{request.PeriodStart:dd.MM.yyyy}–{request.PeriodEnd:dd.MM.yyyy} · инд. сценарии");
        WriteHeaders(sheet, 1, headers);

        var employeeGroups = attempts
            .GroupBy(item => InstallationKey(item.Installation) + "\u001F" + PersonKey(item.PersonName))
            .Select(group => group.OrderBy(item => item.AttemptedAt).ThenBy(item => item.Id).ToArray())
            .OrderBy(group => group[0].Installation, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(group => group[0].PersonName, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
        var row = 2;
        foreach (var group in employeeGroups)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheet.Cell(row, 1).Value = group[0].Installation;
            sheet.Cell(row, 2).Value = group[0].PersonName;
            var column = 3;
            foreach (var month in months)
            {
                var periodRows = group.Where(item =>
                    item.AttemptedAt.Year == month.Year && item.AttemptedAt.Month == month.Month).ToArray();
                sheet.Cell(row, column++).Value = periodRows.Length;
                sheet.Cell(row, column++).Value = CountIndividualScenarios(periodRows);
            }
            foreach (var quarter in quarters)
            {
                var periodRows = group.Where(item =>
                    item.AttemptedAt.Year == quarter.Key.Year &&
                    QuarterOf(item.AttemptedAt.Month) == quarter.Key.Quarter).ToArray();
                sheet.Cell(row, column++).Value = periodRows.Length;
                sheet.Cell(row, column++).Value = CountIndividualScenarios(periodRows);
            }
            sheet.Cell(row, column++).Value = group.Length;
            sheet.Cell(row, column).Value = CountIndividualScenarios(group);
            row++;
        }

        StyleDataTable(sheet, 1, row - 1, headers.Count, "MonitoringEmployees", 2);
        sheet.Column(1).Width = 23;
        sheet.Column(2).Width = 35;
        for (var column = 3; column <= headers.Count; column++)
            sheet.Column(column).Width = 18;
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
    }

    private static void WriteDetailsSheet(
        XLWorkbook workbook,
        IReadOnlyList<DatabaseAttempt> attempts,
        CancellationToken cancellationToken)
    {
        var sheet = workbook.AddWorksheet("Детализация");
        sheet.ShowGridLines = false;
        string[] headers =
            ["Установка", "ФИО", "Месяц", "Дата", "Тема", "Режим", "Файл", "Лист"];
        WriteHeaders(sheet, 1, headers);
        var ordered = attempts
            .OrderBy(item => item.AttemptedAt)
            .ThenBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.PersonName, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.Id)
            .ToArray();
        for (var index = 0; index < ordered.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var item = ordered[index];
            var row = index + 2;
            sheet.Cell(row, 1).Value = item.Installation;
            sheet.Cell(row, 2).Value = item.PersonName;
            sheet.Cell(row, 3).Value = MonthLabel(new MonthKey(item.AttemptedAt.Year, item.AttemptedAt.Month));
            sheet.Cell(row, 4).Value = item.AttemptedAt.Date;
            sheet.Cell(row, 5).Value = item.ScenarioName;
            sheet.Cell(row, 6).Value = item.Mode ?? "";
            sheet.Cell(row, 7).Value = item.SourceFile;
            sheet.Cell(row, 8).Value = item.SourceSheet;
        }
        StyleDataTable(sheet, 1, ordered.Length + 1, headers.Length, "MonitoringDetails", 2);
        sheet.Range(2, 4, ordered.Length + 1, 4).Style.DateFormat.Format = "dd.MM.yyyy";
        SetWidths(sheet, [23, 34, 17, 14, 56, 24, 42, 26]);
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
    }

    private static void WriteDiagnosticsSheet(
        XLWorkbook workbook,
        int exactDuplicates,
        int possibleDuplicates,
        MonitoringReportRequest request,
        IReadOnlyList<QuarterScope> quarters)
    {
        var sheet = workbook.AddWorksheet("Диагностика");
        sheet.ShowGridLines = false;
        string[] headers = ["Тип", "Файл", "Лист", "Причина"];
        WriteHeaders(sheet, 1, headers);
        var rows = new List<object?[]>
        {
            new object?[] { "Источник", "Локальная база данных", "", "В отчёт включены записи выбранного периода и установок." },
            new object?[] { "Точные повторы", "", "", $"Исключено записей: {exactDuplicates:N0}." },
            new object?[] { "Возможные повторы", "", "", request.IncludePossibleDuplicates
                ? $"Включено записей: {possibleDuplicates:N0}."
                : $"Исключено записей: {possibleDuplicates:N0}." },
            new object?[] { "Методика", "", "", "Индивидуальный сценарий — уникальная тройка: технический объект + сотрудник + сценарий." },
            new object?[] { "Групповое прохождение", "", "", "Каждый участник учитывается отдельно, если он представлен отдельной записью в исходных данных." }
        };
        foreach (var quarter in quarters.Where(item => !item.IsComplete))
            rows.Add(new object?[] {
                "Неполный квартал",
                "",
                "",
                $"{QuarterLabel(quarter.Key)} не записан в форму: выбранный период не покрывает квартал полностью."
            });
        for (var index = 0; index < rows.Count; index++)
            WriteRow(sheet, index + 2, rows[index]);
        StyleDataTable(sheet, 1, rows.Count + 1, headers.Length, "MonitoringDiagnostics", 0);
        SetWidths(sheet, [24, 30, 22, 95]);
        sheet.Range(2, 4, rows.Count + 1, 4).Style.Alignment.WrapText = true;
    }

    private static IReadOnlyList<MonitoringQuarterResult> WriteCollectionForm(
        string templateFile,
        string outputFile,
        IReadOnlyList<DatabaseAttempt> attempts,
        IReadOnlyList<QuarterScope> quarters,
        int year)
    {
        using var templateStream = new FileStream(
            templateFile,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        using var workbook = new XLWorkbook(templateStream);
        var sheet = workbook.Worksheets.FirstOrDefault(item =>
                        string.Equals(item.Name, "Форма сбора", StringComparison.CurrentCultureIgnoreCase))
                    ?? throw new InvalidOperationException(
                        "В шаблоне отсутствует лист «Форма сбора».");
        var headerRow = FindQuarterHeaderRow(sheet);
        var indicatorRow = FindIndicatorRow(sheet);
        var quarterColumns = new Dictionary<int, int>();
        for (var column = 1; column <= sheet.LastColumnUsed()!.ColumnNumber(); column++)
        {
            var normalized = TextNormalization.NormalizeText(sheet.Cell(headerRow, column).GetString());
            if (normalized.Length == 2 && normalized[0] == 'q' &&
                int.TryParse(normalized[1..], out var quarter) && quarter is >= 1 and <= 4)
            {
                quarterColumns[quarter] = column;
            }
        }
        if (quarterColumns.Count != 4)
            throw new InvalidOperationException(
                "В форме сбора не найдены все квартальные колонки Q1–Q4.");

        foreach (var column in quarterColumns.Values)
            sheet.Cell(indicatorRow, column).Clear(XLClearOptions.Contents);

        var results = new List<MonitoringQuarterResult>();
        for (var quarter = 1; quarter <= 4; quarter++)
        {
            var quarterScope = quarters.FirstOrDefault(item =>
                item.Key.Year == year && item.Key.Quarter == quarter);
            var quarterRows = attempts.Where(item =>
                item.AttemptedAt.Year == year && QuarterOf(item.AttemptedAt.Month) == quarter).ToArray();
            var individualScenarios = CountIndividualScenarios(quarterRows);
            var write = quarterScope?.IsComplete == true;
            if (write)
            {
                var cell = sheet.Cell(indicatorRow, quarterColumns[quarter]);
                cell.Value = individualScenarios;
                cell.Style.NumberFormat.Format = "#,##0";
            }
            results.Add(new MonitoringQuarterResult(year, quarter, individualScenarios, write));
        }

        workbook.CalculateMode = XLCalculateMode.Auto;
        workbook.SaveAs(outputFile);
        return results;
    }

    private static int FindQuarterHeaderRow(IXLWorksheet sheet)
    {
        var lastRow = Math.Min(sheet.LastRowUsed()?.RowNumber() ?? 0, 100);
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (var row = 1; row <= lastRow; row++)
        {
            var values = Enumerable.Range(1, lastColumn)
                .Select(column => TextNormalization.NormalizeText(sheet.Cell(row, column).GetString()))
                .ToHashSet(StringComparer.Ordinal);
            if (values.Contains("q1") && values.Contains("q2") &&
                values.Contains("q3") && values.Contains("q4"))
                return row;
        }
        throw new InvalidOperationException("В форме сбора не найдена строка заголовков Q1–Q4.");
    }

    private static int FindIndicatorRow(IXLWorksheet sheet)
    {
        const string expected = "кол во прохождения сценариев";
        foreach (var row in sheet.RowsUsed())
        {
            foreach (var cell in row.CellsUsed())
            {
                if (TextNormalization.NormalizeText(cell.GetString()) == expected)
                    return row.RowNumber();
            }
        }
        throw new InvalidOperationException(
            "В форме сбора не найдена строка «Кол-во прохождения сценариев».");
    }

    private static MetricSet CalculateMetrics(IReadOnlyList<DatabaseAttempt> attempts) => new(
        attempts.Select(item => PersonKey(item.PersonName)).Distinct(StringComparer.Ordinal).Count(),
        attempts.Count,
        attempts.Select(item => InstallationScenarioKey(item.Installation, item.ScenarioName))
            .Distinct(StringComparer.Ordinal).Count(),
        CountIndividualScenarios(attempts));

    private static int CountIndividualScenarios(IEnumerable<DatabaseAttempt> attempts) => attempts
        .Select(item => IndividualScenarioKey(item.Installation, item.PersonName, item.ScenarioName))
        .Distinct(StringComparer.Ordinal)
        .Count();

    private static int[] MetricValues(MetricSet value) =>
        [value.Students, value.Lessons, value.UniqueScenarios, value.IndividualScenarios];

    private static string IndividualScenarioKey(string installation, string person, string scenario) =>
        InstallationKey(installation) + "\u001F" + PersonKey(person) + "\u001F" + ScenarioKey(scenario);

    private static string InstallationScenarioKey(string installation, string scenario) =>
        InstallationKey(installation) + "\u001F" + ScenarioKey(scenario);

    private static string InstallationKey(string value) => TextNormalization.NormalizeText(value);
    private static string PersonKey(string value) => TextNormalization.NormalizePersonName(value);
    private static string ScenarioKey(string value) => TextNormalization.NormalizeText(value);
    private static int QuarterOf(int month) => (month - 1) / 3 + 1;

    private static string MonthLabel(MonthKey month) =>
        new DateTime(month.Year, month.Month, 1).ToString("MMMM yyyy", RussianCulture);

    private static string QuarterLabel(QuarterKey quarter)
    {
        var firstMonth = (quarter.Quarter - 1) * 3 + 1;
        var lastMonth = firstMonth + 2;
        return $"{RomanQuarter(quarter.Quarter)} квартал {quarter.Year} ({firstMonth:00}–{lastMonth:00})";
    }

    private static string RomanQuarter(int quarter) => quarter switch
    {
        1 => "I",
        2 => "II",
        3 => "III",
        4 => "IV",
        _ => quarter.ToString(CultureInfo.InvariantCulture)
    };

    private static bool IsExactDuplicate(DatabaseAttempt attempt) =>
        attempt.DuplicateOfAttemptId.HasValue ||
        string.Equals(attempt.DuplicateKind, "Точный", StringComparison.OrdinalIgnoreCase);

    private static bool IsPossibleDuplicate(DatabaseAttempt attempt) =>
        !IsExactDuplicate(attempt) &&
        (string.Equals(attempt.DuplicateKind, "Возможный", StringComparison.OrdinalIgnoreCase) ||
         attempt.PossibleDuplicate);

    private static void ValidateRequest(MonitoringReportRequest request)
    {
        if (!File.Exists(request.DatabasePath))
            throw new FileNotFoundException("Файл локальной базы данных не найден.", request.DatabasePath);
        if (!File.Exists(request.FormTemplateFile))
            throw new FileNotFoundException("Шаблон формы сбора не найден.", request.FormTemplateFile);
        if (request.PeriodEnd.Date < request.PeriodStart.Date)
            throw new ArgumentException("Дата окончания периода не может быть раньше даты начала.");
        if (request.PeriodStart.Year != request.PeriodEnd.Year)
            throw new ArgumentException(
                "Форма сбора заполняется за один календарный год. Выберите даты одного года.");
        if (request.Installations.Length == 0)
            throw new ArgumentException("Выберите хотя бы одну установку.");
        if (Path.GetFullPath(request.ReportOutputFile).Equals(
                Path.GetFullPath(request.FormOutputFile),
                StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Подробный отчёт и форма сбора должны сохраняться в разные файлы.");
    }

    private static void WriteHeaders(IXLWorksheet sheet, int row, IReadOnlyList<string> headers)
    {
        for (var column = 1; column <= headers.Count; column++)
            sheet.Cell(row, column).Value = headers[column - 1];
    }

    private static void WriteRow(IXLWorksheet sheet, int row, IReadOnlyList<object?> values)
    {
        for (var column = 1; column <= values.Count; column++)
            if (values[column - 1] is { } value)
                sheet.Cell(row, column).Value = XLCellValue.FromObject(value);
    }

    private static void StyleHeader(IXLRange range)
    {
        range.Style.Fill.BackgroundColor = Navy;
        range.Style.Font.FontName = FontName;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Font.Bold = true;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        range.Style.Alignment.WrapText = true;
        range.FirstCell().Worksheet.Row(range.FirstRow().RowNumber()).Height = 32;
    }

    private static void StyleDataTable(
        IXLWorksheet sheet,
        int headerRow,
        int lastRow,
        int lastColumn,
        string tableName,
        int freezeColumns)
    {
        var range = sheet.Range(headerRow, 1, lastRow, lastColumn);
        var table = range.CreateTable(tableName);
        table.Theme = XLTableTheme.TableStyleMedium2;
        table.ShowAutoFilter = true;
        range.Style.Font.FontName = FontName;
        range.Style.Font.FontSize = 10;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        StyleHeader(sheet.Range(headerRow, 1, headerRow, lastColumn));
        sheet.SheetView.FreezeRows(headerRow);
        if (freezeColumns > 0)
            sheet.SheetView.FreezeColumns(freezeColumns);
    }

    private static void SetWidths(IXLWorksheet sheet, IReadOnlyList<double> widths)
    {
        for (var index = 0; index < widths.Count; index++)
            sheet.Column(index + 1).Width = widths[index];
    }

    private sealed record MetricSet(
        int Students,
        int Lessons,
        int UniqueScenarios,
        int IndividualScenarios);

    private sealed record MonthKey(int Year, int Month);
    private sealed record QuarterKey(int Year, int Quarter);
    private sealed record QuarterScope(QuarterKey Key, bool IsComplete);
}

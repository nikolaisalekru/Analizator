using System.Diagnostics;
using ClosedXML.Excel;

namespace Analizator.Core;

public static class ReportingService
{
    private static readonly string[] NewPeopleHeaders =
    [
        "Установка", "ФИО", "Должность", "1", "2", "3", "4", "5",
        "Посещение КТК", "Автодобавление", "Добавлен в анализ", "Добавлен в график", "Строка в графике"
    ];

    public static void FinalizeWorkbook(
        XLWorkbook workbook,
        WorkbookProcessingResult processing,
        ExportReadResult export,
        AnalyzerSettings settings,
        IReadOnlyDictionary<string, List<string>> installations,
        ProcessingStatistics statistics,
        IProgress<AnalyzerProgress>? progress = null,
        CancellationToken cancellationToken = default,
        Action<string>? log = null)
    {
        var timer = Stopwatch.StartNew();
        progress?.Report(new AnalyzerProgress(83, "Обновление сводного листа"));
        FillSummary(workbook, processing.Sheets, installations);
        log?.Invoke($"Сводный лист обновлён за {timer.Elapsed.TotalSeconds:F1} с");

        cancellationToken.ThrowIfCancellationRequested();
        timer.Restart();
        progress?.Report(new AnalyzerProgress(85, "Подготовка графика посещаемости"));
        FillAttendanceChart(workbook, processing.Sheets, processing.NewPeople, progress, cancellationToken);
        log?.Invoke($"График посещаемости сформирован за {timer.Elapsed.TotalSeconds:F1} с");

        cancellationToken.ThrowIfCancellationRequested();
        timer.Restart();
        progress?.Report(new AnalyzerProgress(92, "Формирование списка новых сотрудников"));
        CreateNewPeopleSheet(workbook, processing.NewPeople, settings.PassingThreshold);
        log?.Invoke($"Лист новых сотрудников сформирован за {timer.Elapsed.TotalSeconds:F1} с");

        cancellationToken.ThrowIfCancellationRequested();
        timer.Restart();
        progress?.Report(new AnalyzerProgress(93, "Формирование отчёта обработки"));
        CreateProcessingReport(workbook, export, processing, settings, statistics);
        log?.Invoke($"Отчёт обработки сформирован за {timer.Elapsed.TotalSeconds:F1} с");
        workbook.CalculateMode = XLCalculateMode.Auto;
    }

    private static void CreateNewPeopleSheet(
        XLWorkbook workbook,
        IReadOnlyList<NewPerson> people,
        PassingThresholdSettings passingThreshold)
    {
        DeleteSheetIfExists(workbook, "Новые сотрудники");
        var sheet = workbook.AddWorksheet("Новые сотрудники");
        for (var column = 1; column <= NewPeopleHeaders.Length; column++)
        {
            var cell = sheet.Cell(1, column);
            cell.Value = NewPeopleHeaders[column - 1];
            ApplyHeader(cell, "#D9EAF7");
        }

        var row = 2;
        foreach (var person in people
                     .OrderBy(item => TextNormalization.NormalizeText(item.Installation), StringComparer.Ordinal)
                     .ThenBy(item => TextNormalization.NormalizePersonName(item.Name), StringComparer.Ordinal))
        {
            var threshold = passingThreshold.GetForInstallation(person.Installation);
            sheet.Cell(row, 1).Value = person.Installation;
            sheet.Cell(row, 2).Value = person.Name;
            sheet.Cell(row, 3).Value = person.Position;
            for (var index = 0; index < 5; index++)
            {
                var cell = sheet.Cell(row, 4 + index);
                var value = index < person.Scenarios.Count ? Math.Round(person.Scenarios[index].BestPercent, 2) : 0d;
                cell.Value = value;
                if (index >= person.Scenarios.Count)
                    cell.Style.Fill.BackgroundColor = LegacyColor("FFC7CE");
                else if (value < threshold)
                    cell.Style.Fill.BackgroundColor = LegacyColor("FFF2CC");
            }
            WriteBoolean(sheet.Cell(row, 9), person.Visited);
            WriteBoolean(sheet.Cell(row, 10), person.AutoAdd);
            WriteBoolean(sheet.Cell(row, 11), person.AddedToAnalysis);
            WriteBoolean(sheet.Cell(row, 12), person.AddedToAttendanceChart);
            if (person.AttendanceChartRow.HasValue)
                sheet.Cell(row, 13).Value = person.AttendanceChartRow.Value;
            sheet.Range(row, 4, row, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Row(row).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            row++;
        }

        double[] widths = [20, 35, 35, 12, 12, 12, 12, 12, 20, 20, 20, 22, 20];
        for (var column = 1; column <= widths.Length; column++)
            sheet.Column(column).Width = widths[column - 1];
        sheet.SheetView.FreezeRows(1);
        sheet.Range(1, 1, Math.Max(1, row - 1), 13).SetAutoFilter();
    }

    private static void CreateProcessingReport(
        XLWorkbook workbook,
        ExportReadResult export,
        WorkbookProcessingResult processing,
        AnalyzerSettings settings,
        ProcessingStatistics statistics)
    {
        DeleteSheetIfExists(workbook, "Отчёт обработки");
        var sheet = workbook.AddWorksheet("Отчёт обработки");
        var row = 1;
        sheet.Cell(row++, 1).Value = "ОТЧЁТ ОБРАБОТКИ";
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 14;
        row++;

        Section(sheet, ref row, "ЛИСТЫ ВЫГРУЗКИ");
        HeaderRow(sheet, row++, "Лист", "Установка", "Количество отчётов", "Статус");
        if (export.InstallationStatistics.Count == 0)
            StatusRow(sheet, row++, "#FFF2CC", "Нет данных");
        foreach (var item in export.InstallationStatistics)
            StatusRow(sheet, row++, item.Status == "OK" ? "#D9EAD3" : "#FFF2CC",
                item.Sheet, item.Installation ?? "", item.Reports, item.Status);

        row += 2;
        Section(sheet, ref row, "РЕЖИМЫ ПРОХОЖДЕНИЯ");
        StatusRow(sheet, row++, settings.ModeFilter.Enabled ? "#D9EAD3" : "#FFF2CC",
            "Учёт режима", settings.ModeFilter.Enabled ? "ВКЛЮЧЕН" : "ВЫКЛЮЧЕН");
        StatusRow(sheet, row++, null, "Учитываемые режимы",
            settings.ModeFilter.Modes.Count == 0 ? "Не указаны" : string.Join(", ", settings.ModeFilter.Modes));
        row += 2;
        Subsection(sheet, ref row, "Количество отчётов по режимам");
        HeaderRow(sheet, row++, "Режим", "Количество");
        foreach (var mode in new[] { "Случайный выбор", "Ежемесячный тренинг", "Все сценарии", "Экзамен" })
            StatusRow(sheet, row++, null, mode, export.ModeCounts.GetValueOrDefault(mode));
        row += 2;
        Section(sheet, ref row, "ЛИСТЫ БЕЗ СТОЛБЦА РЕЖИМА");
        if (export.SheetsWithoutMode.Count == 0)
            StatusRow(sheet, row++, "#D9EAD3", "Нет листов без режима");
        else
        {
            HeaderRow(sheet, row++, "Лист", "Установка", "Статус");
            foreach (var _ in export.SheetsWithoutMode)
                StatusRow(sheet, row++, "#FFF2CC", null, null, "Режим отсутствует");
        }
        row += 2;
        StatusRow(sheet, row++, null, "Отчётов отфильтровано по режиму", export.FilteredReports);

        row += 2;
        Section(sheet, ref row, "НЕСОПОСТАВЛЕННЫЕ УСТАНОВКИ");
        if (processing.UnmatchedInstallations.Count == 0)
            StatusRow(sheet, row++, "#D9EAD3", "Нет");
        else
            foreach (var item in processing.UnmatchedInstallations)
                StatusRow(sheet, row++, "#FFF2CC", item);

        row += 2;
        Section(sheet, ref row, "НОВЫЕ СОТРУДНИКИ");
        if (processing.NewPeople.Count == 0)
            StatusRow(sheet, row++, "#D9EAD3", "Новых сотрудников не найдено");
        else
        {
            HeaderRow(sheet, row++, "Установка", "ФИО", "Добавлен в анализ", "Автодобавление", "Добавлен в график", "Строка в графике");
            foreach (var person in processing.NewPeople)
                StatusRow(sheet, row++, person.AddedToAnalysis ? "#D9EAD3" : "#FFF2CC",
                    person.Installation, person.Name, YesNo(person.AddedToAnalysis), YesNo(person.AutoAdd),
                    YesNo(person.AddedToAttendanceChart), person.AttendanceChartRow?.ToString() ?? "");
        }

        row += 2;
        Section(sheet, ref row, "FUZZY-СОВПАДЕНИЯ ФИО");
        if (statistics.FuzzyDetails.Count == 0)
            StatusRow(sheet, row++, "#D9EAD3", "Fuzzy-совпадений нет");
        else
        {
            HeaderRow(sheet, row++, "Установка", "ФИО из выгрузки", "ФИО в анализе", "Оценка");
            foreach (var item in statistics.FuzzyDetails)
                StatusRow(sheet, row++, "#FFF2CC", item.Installation, item.ExportName, item.AnalysisName, item.Score);
        }

        row += 2;
        Section(sheet, ref row, "СОВПАДЕНИЯ ПО КЛЮЧАМ ФИО");
        if (statistics.KeyDetails.Count == 0)
            StatusRow(sheet, row++, "#D9EAD3", "Совпадений по ключам нет");
        else
        {
            HeaderRow(sheet, row++, "Установка", "ФИО из выгрузки", "ФИО в анализе");
            foreach (var item in statistics.KeyDetails)
                StatusRow(sheet, row++, null, item.Installation, item.ExportName, item.AnalysisName);
        }

        sheet.Column(1).Width = 32;
        sheet.Column(2).Width = 35;
        sheet.Column(3).Width = 25;
        sheet.Column(4).Width = 30;
        sheet.Column(5).Width = 25;
        sheet.Column(6).Width = 20;
        sheet.SheetView.FreezeRows(3);
    }

    private static void FillSummary(
        XLWorkbook workbook,
        IReadOnlyList<AnalysisSheet> sheets,
        IReadOnlyDictionary<string, List<string>> installations)
    {
        var summary = workbook.Worksheets.FirstOrDefault(sheet =>
            TextNormalization.NormalizeText(sheet.Name) == "сводная");
        if (summary is null)
            return;

        var lastRow = summary.LastRowUsed()?.RowNumber() ?? 0;
        var lastColumn = summary.LastColumnUsed()?.ColumnNumber() ?? 0;
        int? headerRow = null;
        var passedColumns = new List<int>();
        for (var row = 1; row <= Math.Min(15, lastRow); row++)
        {
            var current = new List<int>();
            for (var column = 1; column <= lastColumn; column++)
            {
                var header = TextNormalization.NormalizeHeader(summary.Cell(row, column).GetString());
                if (header is "кол во чел прошедших 5 сценариев" or "количество человек прошедших 5 сценариев")
                    current.Add(column);
            }
            if (current.Count == 0)
                continue;
            headerRow = row;
            passedColumns = current.Order().ToList();
            break;
        }
        if (!headerRow.HasValue || passedColumns.Count == 0)
            return;

        int? installationColumn = null;
        var bestCount = 0;
        for (var column = 1; column < passedColumns[0]; column++)
        {
            var count = Enumerable.Range(headerRow.Value + 1, Math.Max(0, lastRow - headerRow.Value))
                .Count(row => !summary.Cell(row, column).IsEmpty());
            if (count <= bestCount)
                continue;
            bestCount = count;
            installationColumn = column;
        }
        if (!installationColumn.HasValue)
            return;

        var sheetMap = sheets
            .GroupBy(sheet => sheet.Installation, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        for (var row = headerRow.Value + 1; row <= lastRow; row++)
        {
            var installationName = summary.Cell(row, installationColumn.Value).GetString().Trim();
            if (installationName.Length == 0)
                continue;
            var canonical = MatchingService.MatchInstallation(installationName, installations).Name;
            if (canonical is null || !sheetMap.TryGetValue(canonical, out var source) || source.People.Count == 0)
                continue;
            var scenarioBlocks = FindScenarioBlocks(source.Worksheet);
            var formulaRow = source.People.Max(person => person.RowNumber) + 1;
            for (var monthIndex = 0; monthIndex < passedColumns.Count && monthIndex < scenarioBlocks.Count; monthIndex++)
            {
                var passedColumn = passedColumns[monthIndex];
                var countCell = $"{XLHelper.GetColumnLetterFromNumber(passedColumn - 1)}{row}";
                var block = scenarioBlocks[monthIndex];
                var sourceName = source.Worksheet.Name.Replace("'", "''");
                var sourceSum = $"SUM('{sourceName}'!" +
                                $"{XLHelper.GetColumnLetterFromNumber(block.StartColumn)}{formulaRow}:" +
                                $"{XLHelper.GetColumnLetterFromNumber(block.EndColumn)}{formulaRow})";
                summary.Cell(row, passedColumn).FormulaA1 =
                    $"IF({countCell}=\"\",{sourceSum},IF({countCell}<{sourceSum},{countCell},{sourceSum}))";
            }
        }
    }

    private static void FillAttendanceChart(
        XLWorkbook workbook,
        IReadOnlyList<AnalysisSheet> sheets,
        IReadOnlyList<NewPerson> newPeople,
        IProgress<AnalyzerProgress>? progress,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var chart = workbook.Worksheets.FirstOrDefault(sheet =>
            TextNormalization.NormalizeText(sheet.Name) == "график посещаемости");
        if (chart is null)
        {
            foreach (var person in newPeople)
            {
                person.AddedToAttendanceChart = false;
                person.AttendanceChartRow = null;
            }
            return;
        }

        var installationRanges = chart.MergedRanges
            .Where(range => range.RangeAddress.FirstAddress.ColumnNumber == 1 &&
                            range.RangeAddress.LastAddress.ColumnNumber == 1 &&
                            range.RangeAddress.FirstAddress.RowNumber >= 3)
            .OrderBy(range => range.RangeAddress.FirstAddress.RowNumber)
            .ToArray();
        if (installationRanges.Length < 2)
            return;
        var templateStart = installationRanges[1].RangeAddress.FirstAddress.RowNumber;
        var templateEnd = installationRanges[1].RangeAddress.LastAddress.RowNumber;
        var templateMiddle = Math.Min(templateStart + 1, templateEnd);
        var template = new AttendanceTemplate(
            CaptureRow(chart, templateStart, 1, 19),
            CaptureRow(chart, templateMiddle, 1, 19),
            CaptureRow(chart, templateEnd, 1, 19),
            [
                CaptureRow(chart, templateStart, 17, 19),
                CaptureRow(chart, templateStart + 1, 17, 19),
                CaptureRow(chart, templateStart + 2, 17, 19)
            ],
            CaptureRow(chart, templateStart + 3, 16, 19));

        var groups = sheets
            .Select(sheet => new AttendanceGroup(sheet, FindScenarioBlocks(sheet.Worksheet)))
            .Where(group => group.ScenarioBlocks.Count >= 12 && group.Sheet.People.Count > 0)
            .ToList();
        if (groups.Count == 0)
            return;

        progress?.Report(new AnalyzerProgress(86, "Очистка прежнего графика посещаемости"));
        foreach (var range in chart.MergedRanges
                     .Where(range => range.RangeAddress.FirstAddress.RowNumber >= 3)
                     .ToArray())
        {
            cancellationToken.ThrowIfCancellationRequested();
            range.Unmerge();
        }
        var oldLastRow = chart.LastRowUsed()?.RowNumber() ?? 2;
        if (oldLastRow >= 3)
        {
            var oldLastColumn = Math.Max(19, chart.LastColumnUsed()?.ColumnNumber() ?? 19);
            // Preserve column/row template formatting while removing the old data.
            // Reapplying formats cell-by-cell outside the active 19-column chart is
            // both unnecessary and considerably slower in ClosedXML.
            chart.Range(3, 1, oldLastRow, oldLastColumn).Clear(XLClearOptions.Contents);
        }

        var chartRows = new Dictionary<(string Installation, string Person), int>();
        var currentRow = 3;
        for (var groupIndex = 0; groupIndex < groups.Count; groupIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var group = groups[groupIndex];
            var percent = 87 + (int)Math.Floor(5d * groupIndex / groups.Count);
            progress?.Report(new AnalyzerProgress(
                percent,
                $"График посещаемости: {groupIndex + 1} из {groups.Count} установок"));
            var groupStart = currentRow;
            var groupEnd = groupStart + group.Sheet.People.Count - 1;
            for (var offset = 0; offset < group.Sheet.People.Count; offset++)
            {
                if ((offset & 31) == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                var person = group.Sheet.People[offset];
                var targetRow = groupStart + offset;
                var rowTemplate = targetRow == groupStart
                    ? template.Start
                    : targetRow == groupEnd ? template.End : template.Middle;
                ApplyRow(chart, targetRow, rowTemplate, 1, 15, true);
                ApplyRow(chart, targetRow, template.SideBlank, 16, 19, false);
                chart.Cell(targetRow, 2).Value = person.Name;
                for (var monthIndex = 0; monthIndex < 12; monthIndex++)
                {
                    var block = group.ScenarioBlocks[monthIndex];
                    var sourceName = group.Sheet.Worksheet.Name.Replace("'", "''");
                    var attendanceColumn = XLHelper.GetColumnLetterFromNumber(block.AttendanceColumn);
                    chart.Cell(targetRow, 3 + monthIndex).FormulaA1 =
                        $"IF('{sourceName}'!${attendanceColumn}{person.RowNumber}=\"Да\",1,0)";
                }
                chart.Cell(targetRow, 15).FormulaA1 =
                    $"IFERROR(ROUND(SUM(C{targetRow}:N{targetRow})/($S${groupStart}),1),\"\")";
                chartRows[(TextNormalization.NormalizeText(group.Sheet.Installation),
                    TextNormalization.NormalizePersonName(person.Name))] = targetRow;
            }

            chart.Cell(groupStart, 1).Value = group.Sheet.Installation;
            for (var offset = 0; offset < template.Coefficient.Length && groupStart + offset <= groupEnd; offset++)
                ApplyRow(chart, groupStart + offset, template.Coefficient[offset], 17, 19, false);
            chart.Cell(groupStart, 17).Value = "Количество месяцев:";
            var parts = Enumerable.Range(3, 12).Select(column =>
            {
                var letter = XLHelper.GetColumnLetterFromNumber(column);
                return $"IF(COUNTIF({letter}{groupStart}:{letter}{groupEnd},\">0\")>0,1,0)";
            });
            chart.Cell(groupStart, 19).FormulaA1 = "SUM(" + string.Join(",", parts) + ")";
            if (groupEnd > groupStart)
                chart.Range(groupStart, 1, groupEnd, 1).Merge();
            chart.Range(groupStart, 17, groupStart, 18).Merge();
            var boxEnd = Math.Min(groupStart + 2, groupEnd);
            if (boxEnd >= groupStart + 1)
                chart.Range(groupStart + 1, 17, boxEnd, 19).Merge();
            currentRow = groupEnd + 1;
        }

        foreach (var person in newPeople)
        {
            var key = (TextNormalization.NormalizeText(person.Installation), TextNormalization.NormalizePersonName(person.Name));
            var found = chartRows.TryGetValue(key, out var chartRow);
            person.AddedToAttendanceChart = person.AddedToAnalysis && found;
            person.AttendanceChartRow = person.AddedToAttendanceChart ? chartRow : null;
        }

        var year = FindYear(workbook, chart);
        if (year is not null)
            chart.Cell("A1").Value = $"Статистика за {year} год";
        chart.SheetView.FreezeRows(2);
        chart.SheetView.FreezeColumns(2);
    }

    private static List<ScenarioBlock> FindScenarioBlocks(IXLWorksheet sheet)
    {
        var blocks = new List<ScenarioBlock>();
        var used = new HashSet<int>();
        var maxRow = Math.Min(10, sheet.LastRowUsed()?.RowNumber() ?? 0);
        var maxColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (var row = 1; row <= maxRow; row++)
        for (var column = 1; column <= maxColumn - 5; column++)
        {
            var complete = true;
            for (var offset = 0; offset < 5; offset++)
                complete &= sheet.Cell(row, column + offset).GetString().Trim() == (offset + 1).ToString();
            if (!complete || TextNormalization.NormalizeHeader(sheet.Cell(row, column + 5).GetString()) != "посещение ктк" || !used.Add(column))
                continue;
            blocks.Add(new ScenarioBlock(column, column + 4, column + 5));
            column += 5;
        }
        return blocks.OrderBy(block => block.StartColumn).ToList();
    }

    private static RowTemplate CaptureRow(IXLWorksheet sheet, int row, int firstColumn, int lastColumn)
    {
        var styles = new Dictionary<int, IXLStyle>();
        for (var column = firstColumn; column <= lastColumn; column++)
            styles[column] = sheet.Cell(row, column).Style;
        return new RowTemplate(sheet.Row(row).Height, styles);
    }

    private static void ApplyRow(
        IXLWorksheet sheet,
        int row,
        RowTemplate template,
        int firstColumn,
        int lastColumn,
        bool applyHeight)
    {
        if (applyHeight)
            sheet.Row(row).Height = template.Height;
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            sheet.Cell(row, column).Style = template.Styles[column];
            sheet.Cell(row, column).Clear(XLClearOptions.Contents);
        }
    }

    private static string? FindYear(XLWorkbook workbook, IXLWorksheet chart)
    {
        var candidates = new List<string>();
        var summary = workbook.Worksheets.FirstOrDefault(sheet => TextNormalization.NormalizeText(sheet.Name) == "сводная");
        if (summary is not null)
            for (var row = 1; row <= Math.Min(5, summary.LastRowUsed()?.RowNumber() ?? 0); row++)
            for (var column = 1; column <= Math.Min(8, summary.LastColumnUsed()?.ColumnNumber() ?? 0); column++)
                candidates.Add(summary.Cell(row, column).GetString());
        candidates.Add(chart.Cell("A1").GetString());
        foreach (var candidate in candidates)
        {
            var match = System.Text.RegularExpressions.Regex.Match(candidate, @"\b(20\d{2})\b");
            if (match.Success)
                return match.Groups[1].Value;
        }
        return null;
    }

    private sealed record ScenarioBlock(int StartColumn, int EndColumn, int AttendanceColumn);
    private sealed record AttendanceGroup(AnalysisSheet Sheet, List<ScenarioBlock> ScenarioBlocks);
    private sealed record RowTemplate(double Height, Dictionary<int, IXLStyle> Styles);
    private sealed record AttendanceTemplate(
        RowTemplate Start,
        RowTemplate Middle,
        RowTemplate End,
        RowTemplate[] Coefficient,
        RowTemplate SideBlank);

    private static void DeleteSheetIfExists(XLWorkbook workbook, string name)
    {
        var existing = workbook.Worksheets.FirstOrDefault(sheet =>
            string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));
        existing?.Delete();
    }

    private static void ApplyHeader(IXLCell cell, string color, bool center = true)
    {
        cell.Style.Fill.BackgroundColor = LegacyColor(color.TrimStart('#'));
        cell.Style.Font.Bold = true;
        if (center)
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        cell.Style.Alignment.WrapText = true;
        ApplyBorder(cell);
    }

    private static void HeaderRow(IXLWorksheet sheet, int row, params string[] values)
    {
        for (var column = 1; column <= values.Length; column++)
        {
            sheet.Cell(row, column).Value = values[column - 1];
            ApplyHeader(sheet.Cell(row, column), "#D9EAF7", false);
        }
    }

    private static void Section(IXLWorksheet sheet, ref int row, string title)
    {
        var cell = sheet.Cell(row++, 1);
        cell.Value = title;
        cell.Style.Font.Bold = true;
        cell.Style.Fill.BackgroundColor = LegacyColor("B4C7E7");
    }

    private static void Subsection(IXLWorksheet sheet, ref int row, string title)
    {
        sheet.Cell(row++, 1).Value = title;
        sheet.Cell(row - 1, 1).Style.Font.Bold = true;
    }

    private static void StatusRow(IXLWorksheet sheet, int row, string? color, params object?[] values)
    {
        for (var column = 1; column <= values.Length; column++)
        {
            var cell = sheet.Cell(row, column);
            cell.Value = XLCellValue.FromObject(values[column - 1]);
            if (color is not null)
                cell.Style.Fill.BackgroundColor = LegacyColor(color.TrimStart('#'));
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            ApplyBorder(cell);
        }
    }

    private static void ApplyBorder(IXLCell cell)
    {
        cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        cell.Style.Border.LeftBorderColor = LegacyColor("B7B7B7");
        cell.Style.Border.RightBorderColor = LegacyColor("B7B7B7");
        cell.Style.Border.TopBorderColor = LegacyColor("B7B7B7");
        cell.Style.Border.BottomBorderColor = LegacyColor("B7B7B7");
    }

    private static void WriteBoolean(IXLCell cell, bool value)
    {
        cell.Value = YesNo(value).ToUpperInvariant();
        cell.Style.Fill.BackgroundColor = LegacyColor(value ? "C6EFCE" : "EDE4D3");
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    private static string YesNo(bool value) => value ? "Да" : "Нет";

    private static XLColor LegacyColor(string rgb) => XLColor.FromArgb(
        0,
        Convert.ToByte(rgb[..2], 16),
        Convert.ToByte(rgb.Substring(2, 2), 16),
        Convert.ToByte(rgb.Substring(4, 2), 16));
}

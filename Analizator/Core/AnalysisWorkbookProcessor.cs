using ClosedXML.Excel;
using System.Globalization;

namespace Analizator.Core;

public static class AnalysisWorkbookProcessor
{
    private static readonly HashSet<string> ServiceSheets =
    [
        "новые сотрудники", "отчет обработки", "отчет обработки", "график посещаемости"
    ];

    private static readonly HashSet<string> IgnoredPersonRows =
    [
        "фио", "итого", "итого", "всего", "всего сотрудников", "список сотрудников",
        "количество", "количество сотрудников"
    ];

    public static WorkbookProcessingResult Process(
        XLWorkbook workbook,
        ExportReadResult export,
        ConfigurationBundle configuration,
        int month,
        bool autoAddOverride,
        ProcessingStatistics statistics,
        CancellationToken cancellationToken,
        Action<string>? log = null)
    {
        var sheets = DiscoverSheets(workbook, configuration.Installations, log);
        var sheetsByInstallation = sheets
            .GroupBy(sheet => sheet.Installation, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var unmatched = sheets.Count == 0 ? new List<string>() : new List<string>();
        foreach (var installation in export.Data.Keys)
            if (!sheetsByInstallation.ContainsKey(installation))
                unmatched.Add(installation);

        var newPeople = new List<NewPerson>();
        foreach (var sheet in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var passingThreshold = configuration.Settings.PassingThreshold.GetForInstallation(sheet.Installation);
            log?.Invoke($"[ОБРАБОТКА ЛИСТА] {sheet.Title} → {sheet.Installation}");
            var monthBlock = EnsureMonthBlock(sheet.Worksheet, month, log);
            if (monthBlock is null)
            {
                unmatched.Add(sheet.Title);
                continue;
            }

            export.Data.TryGetValue(sheet.Installation, out var exportPeople);
            exportPeople ??= new Dictionary<string, Dictionary<string, ScenarioAggregate>>(StringComparer.OrdinalIgnoreCase);
            var matchedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var exportPerson in exportPeople)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var match = MatchingService.MatchPerson(
                    exportPerson.Key,
                    sheet.People.Select(person => person.Name),
                    sheet.Installation,
                    configuration.FioKeys,
                    configuration.Installations);
                if (!match.IsMatch || match.Name is null)
                {
                    var normalized = TextNormalization.NormalizePersonName(exportPerson.Key);
                    if (!configuration.ExcludedPeople.Contains(normalized))
                    {
                        newPeople.Add(new NewPerson
                        {
                            Installation = sheet.Installation,
                            Name = exportPerson.Key,
                            Scenarios = SelectBestScenarios(exportPerson.Value),
                            AutoAdd = autoAddOverride && configuration.Settings.AutoAddNewPeople.IsEnabled(sheet.Installation)
                        });
                    }
                    continue;
                }

                var personRow = sheet.People.FirstOrDefault(person =>
                    TextNormalization.NormalizePersonName(person.Name) == TextNormalization.NormalizePersonName(match.Name));
                if (personRow is null)
                    continue;
                matchedNames.Add(TextNormalization.NormalizePersonName(personRow.Name));
                WritePersonResults(
                    sheet.Worksheet,
                    personRow.RowNumber,
                    monthBlock,
                    SelectBestScenarios(exportPerson.Value),
                    passingThreshold);
                statistics.MatchedPeople++;
                switch (match.Kind)
                {
                    case MatchKind.Exact:
                        statistics.ExactMatches++;
                        break;
                    case MatchKind.Key:
                        statistics.KeyMatches++;
                        statistics.KeyDetails.Add((sheet.Installation, exportPerson.Key, personRow.Name));
                        break;
                    case MatchKind.Fuzzy:
                        statistics.FuzzyMatches++;
                        statistics.FuzzyDetails.Add((sheet.Installation, exportPerson.Key, personRow.Name, match.Score));
                        break;
                }
            }

            foreach (var person in sheet.People)
            {
                if (matchedNames.Contains(TextNormalization.NormalizePersonName(person.Name)))
                    continue;
                MarkAbsent(sheet.Worksheet, person.RowNumber, sheet.FioColumn, monthBlock);
                statistics.AbsentPeople++;
            }
        }

        statistics.NewPeople = newPeople.Count;
        foreach (var person in newPeople.Where(person => person.AutoAdd))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!sheetsByInstallation.TryGetValue(person.Installation, out var sheet))
                continue;
            var passingThreshold = configuration.Settings.PassingThreshold.GetForInstallation(person.Installation);
            if (InsertNewPerson(sheet, person, month, passingThreshold, log))
            {
                person.AddedToAnalysis = true;
                statistics.AddedPeople++;
            }
        }

        foreach (var sheet in sheets)
        {
            var people = ReadPeople(sheet.Worksheet, sheet.HeaderRow, sheet.FioColumn, sheet.PositionColumn);
            var blocks = FindMonthBlocks(sheet.Worksheet);
            var passingThreshold = configuration.Settings.PassingThreshold.GetForInstallation(sheet.Installation);
            RefreshMonthFormulas(sheet.Worksheet, people, blocks, passingThreshold);
            ApplyBlankConditionalFormatting(sheet.Worksheet, people, blocks);
        }

        return new WorkbookProcessingResult(sheets, newPeople, unmatched.Distinct(StringComparer.OrdinalIgnoreCase).ToList());
    }

    public static List<AnalysisSheet> DiscoverSheets(
        XLWorkbook workbook,
        IReadOnlyDictionary<string, List<string>> installations,
        Action<string>? log = null)
    {
        var result = new List<AnalysisSheet>();
        foreach (var worksheet in workbook.Worksheets)
        {
            if (ServiceSheets.Contains(TextNormalization.NormalizeText(worksheet.Name)))
                continue;
            var header = FindAnalysisHeader(worksheet);
            if (header.Row is null || header.FioColumn is null)
                continue;
            var people = ReadPeople(worksheet, header.Row.Value, header.FioColumn.Value, header.PositionColumn);
            if (people.Count == 0)
                continue;
            var match = MatchingService.MatchInstallation(worksheet.Name, installations);
            if (match.Name is null)
            {
                log?.Invoke($"[ПРЕДУПРЕЖДЕНИЕ] Лист {worksheet.Name} не сопоставлен с установкой");
                continue;
            }
            var sheet = new AnalysisSheet
            {
                Title = worksheet.Name,
                Installation = match.Name,
                Worksheet = worksheet,
                HeaderRow = header.Row.Value,
                FioColumn = header.FioColumn.Value,
                PositionColumn = header.PositionColumn
            };
            sheet.People.AddRange(people);
            sheet.MonthBlocks.AddRange(FindMonthBlocks(worksheet));
            result.Add(sheet);
        }
        return result;
    }

    private static (int? Row, int? FioColumn, int? PositionColumn) FindAnalysisHeader(IXLWorksheet sheet)
    {
        var maxRow = Math.Min(50, sheet.LastRowUsed()?.RowNumber() ?? 0);
        var maxColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (var row = 1; row <= maxRow; row++)
        {
            int? fio = null;
            int? position = null;
            for (var column = 1; column <= maxColumn; column++)
            {
                var value = TextNormalization.NormalizeHeader(sheet.Cell(row, column).GetString());
                if (value is "фио" or "ф и о" or "фио сотрудника" or "ф и о сотрудника")
                    fio = column;
                else if (value is "должность" or "должность сотрудника")
                    position = column;
            }
            if (fio.HasValue)
                return (row, fio, position);
        }
        return (null, null, null);
    }

    private static List<PersonRow> ReadPeople(IXLWorksheet sheet, int headerRow, int fioColumn, int? positionColumn)
    {
        var result = new List<PersonRow>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;
        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var name = sheet.Cell(row, fioColumn).GetString().Trim();
            var normalized = TextNormalization.NormalizePersonName(name);
            if (normalized.Length < 3 || IgnoredPersonRows.Contains(normalized) || !seen.Add(normalized))
                continue;
            var position = positionColumn.HasValue ? sheet.Cell(row, positionColumn.Value).GetString().Trim() : "";
            result.Add(new PersonRow(name, position, row));
        }
        return result;
    }

    public static List<MonthBlock> FindMonthBlocks(IXLWorksheet sheet)
    {
        var maxRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var maxColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var headers = new List<(int Month, string Name, int Row, int Column)>();
        foreach (var month in TextNormalization.MonthNames)
        for (var row = 1; row <= Math.Min(10, maxRow); row++)
        for (var column = 1; column <= maxColumn; column++)
        {
            var text = TextNormalization.NormalizeText(sheet.Cell(row, column).GetString());
            if (text == TextNormalization.NormalizeText(month.Value))
                headers.Add((month.Key, month.Value, row, column));
        }

        var result = new List<MonthBlock>();
        foreach (var group in headers.GroupBy(header => header.Row))
        {
            var ordered = group.OrderBy(header => header.Column).ToArray();
            for (var index = 0; index < ordered.Length; index++)
            {
                var item = ordered[index];
                var endColumn = index + 1 < ordered.Length ? ordered[index + 1].Column - 1 : maxColumn;
                for (var candidateRow = item.Row + 1; candidateRow <= Math.Min(item.Row + 5, maxRow); candidateRow++)
                {
                    var scenarios = new Dictionary<int, int>();
                    int? attendance = null;
                    for (var column = item.Column; column <= endColumn; column++)
                    {
                        var value = sheet.Cell(candidateRow, column).GetString().Trim();
                        if (int.TryParse(value, out var number) && number is >= 1 and <= 5)
                            scenarios.TryAdd(number, column);
                        else if (TextNormalization.NormalizeHeader(value) == "посещение ктк")
                            attendance = column;
                    }
                    if (scenarios.Count == 5)
                    {
                        result.Add(new MonthBlock(item.Month, item.Name, item.Row, item.Column, endColumn, scenarios, attendance));
                        break;
                    }
                }
            }
        }
        return result.OrderBy(block => block.HeaderRow).ThenBy(block => block.StartColumn).ToList();
    }

    public static MonthBlock? EnsureMonthBlock(IXLWorksheet sheet, int month, Action<string>? log = null)
    {
        var blocks = FindMonthBlocks(sheet);
        var existing = blocks.FirstOrDefault(block => block.Month == month);
        if (existing is not null)
            return existing;
        if (month is < 1 or > 12 || blocks.Count == 0)
            return null;

        var template = blocks.Where(block => block.Month < month).OrderByDescending(block => block.Month).FirstOrDefault()
                       ?? blocks.OrderBy(block => block.Month).First();
        var width = template.EndColumn - template.StartColumn + 1;
        var later = blocks.Where(block => block.Month > month).OrderBy(block => block.Month).FirstOrDefault();
        var insertColumn = month > template.Month
            ? later?.StartColumn ?? template.EndColumn + 1
            : template.StartColumn;
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? template.EndColumn;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? template.HeaderRow + 5;

        var cells = new List<CellTemplate>();
        for (var row = 1; row <= lastRow; row++)
        for (var offset = 0; offset < width; offset++)
        {
            var cell = sheet.Cell(row, template.StartColumn + offset);
            cells.Add(new CellTemplate(row, offset, cell.Value, cell.HasFormula ? cell.FormulaR1C1 : null, cell.Style));
        }
        var widths = Enumerable.Range(0, width).Select(offset => sheet.Column(template.StartColumn + offset).Width).ToArray();
        var merges = sheet.MergedRanges
            .Where(range => range.RangeAddress.FirstAddress.ColumnNumber >= template.StartColumn &&
                            range.RangeAddress.LastAddress.ColumnNumber <= template.EndColumn &&
                            range.RangeAddress.FirstAddress.RowNumber >= template.HeaderRow &&
                            range.RangeAddress.LastAddress.RowNumber <= template.HeaderRow + 4)
            .Select(range => new MergeTemplate(
                range.RangeAddress.FirstAddress.RowNumber,
                range.RangeAddress.LastAddress.RowNumber,
                range.RangeAddress.FirstAddress.ColumnNumber - template.StartColumn,
                range.RangeAddress.LastAddress.ColumnNumber - template.StartColumn))
            .ToArray();

        if (insertColumn <= lastColumn)
            sheet.Column(insertColumn).InsertColumnsBefore(width);

        foreach (var item in cells)
        {
            var target = sheet.Cell(item.Row, insertColumn + item.Offset);
            target.Style = item.Style;
            if (item.FormulaR1C1 is not null)
                target.FormulaR1C1 = item.FormulaR1C1;
            else
                target.Value = item.Value;
        }
        for (var offset = 0; offset < width; offset++)
            sheet.Column(insertColumn + offset).Width = widths[offset];
        foreach (var merge in merges)
        {
            var target = sheet.Range(merge.FirstRow, insertColumn + merge.FirstColumnOffset,
                merge.LastRow, insertColumn + merge.LastColumnOffset);
            if (!target.IsMerged())
                target.Merge();
        }

        var oldName = TextNormalization.NormalizeText(template.MonthName);
        var replaced = false;
        for (var row = template.HeaderRow; row <= template.HeaderRow + 4 && !replaced; row++)
        for (var offset = 0; offset < width; offset++)
        {
            var cell = WritableCell(sheet.Cell(row, insertColumn + offset));
            if (TextNormalization.NormalizeText(cell.GetString()) != oldName)
                continue;
            cell.Value = TextNormalization.MonthNames[month];
            replaced = true;
            break;
        }
        if (!replaced)
            WritableCell(sheet.Cell(template.HeaderRow, insertColumn)).Value = TextNormalization.MonthNames[month];

        var created = FindMonthBlocks(sheet).FirstOrDefault(block => block.Month == month);
        if (created is not null)
            log?.Invoke($"[МЕСЯЦ] Добавлен блок {created.MonthName} на лист {sheet.Name}");
        return created;
    }

    private static List<ScenarioAggregate> SelectBestScenarios(Dictionary<string, ScenarioAggregate> scenarios) =>
        scenarios.Values
            .OrderByDescending(item => item.BestPercent)
            .ThenBy(item => TextNormalization.NormalizeText(item.DisplayName), StringComparer.Ordinal)
            .Take(5)
            .Select(item => new ScenarioAggregate
            {
                DisplayName = item.DisplayName,
                BestPercent = item.BestPercent,
                Attempts = item.Attempts
            })
            .ToList();

    private static void WritePersonResults(
        IXLWorksheet sheet,
        int row,
        MonthBlock block,
        IReadOnlyList<ScenarioAggregate> scenarios,
        double passingThreshold)
    {
        for (var number = 1; number <= 5; number++)
        {
            if (!block.ScenarioColumns.TryGetValue(number, out var column))
                continue;
            var cell = WritableCell(sheet.Cell(row, column));
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            if (number <= scenarios.Count)
            {
                var value = Math.Round(scenarios[number - 1].BestPercent, 2);
                cell.Value = value;
                cell.Style.Fill.BackgroundColor = value < passingThreshold
                    ? LegacyColor("FFF2CC")
                    : XLColor.NoColor;
            }
            else
            {
                cell.Value = 0d;
                cell.Style.Fill.BackgroundColor = LegacyColor("FFC7CE");
            }
        }
        if (block.AttendanceColumn.HasValue)
        {
            var attendance = WritableCell(sheet.Cell(row, block.AttendanceColumn.Value));
            attendance.Value = scenarios.Count > 0 ? "ДА" : "НЕТ";
            attendance.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            attendance.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            attendance.Style.Fill.BackgroundColor = LegacyColor(scenarios.Count > 0 ? "C6EFCE" : "EDE4D3");
        }
    }

    private static void MarkAbsent(IXLWorksheet sheet, int row, int fioColumn, MonthBlock block)
    {
        var fio = WritableCell(sheet.Cell(row, fioColumn));
        fio.Style.Fill.BackgroundColor = LegacyColor("EDE4D3");
        fio.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        fio.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        var processed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in block.ScenarioColumns.OrderBy(pair => pair.Key).Select(pair => pair.Value))
        {
            var cell = WritableCell(sheet.Cell(row, column));
            if (!processed.Add(cell.Address.ToString() ?? $"{row}:{column}"))
                continue;
            if (!cell.IsEmpty())
                continue;
            cell.Value = 0d;
            cell.Style.Fill.BackgroundColor = LegacyColor("FFC7CE");
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        if (block.AttendanceColumn.HasValue)
        {
            var attendance = WritableCell(sheet.Cell(row, block.AttendanceColumn.Value));
            if (attendance.IsEmpty())
            {
                attendance.Value = "НЕТ";
                attendance.Style.Fill.BackgroundColor = LegacyColor("EDE4D3");
                attendance.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
        }
    }

    private static bool InsertNewPerson(
        AnalysisSheet sheet,
        NewPerson person,
        int month,
        double passingThreshold,
        Action<string>? log)
    {
        var currentPeople = ReadPeople(sheet.Worksheet, sheet.HeaderRow, sheet.FioColumn, sheet.PositionColumn);
        var normalized = TextNormalization.NormalizePersonName(person.Name);
        if (currentPeople.Any(item => TextNormalization.NormalizePersonName(item.Name) == normalized))
            return false;
        var block = EnsureMonthBlock(sheet.Worksheet, month, log);
        if (block is null)
            return false;

        var next = currentPeople.FirstOrDefault(item =>
            string.Compare(TextNormalization.NormalizePersonName(item.Name), normalized, StringComparison.Ordinal) > 0);
        var insertRow = next?.RowNumber ?? (currentPeople.Count > 0 ? currentPeople.Max(item => item.RowNumber) + 1 : sheet.HeaderRow + 1);
        var styleSourceBefore = next?.RowNumber ?? (currentPeople.Count > 0 ? currentPeople.Max(item => item.RowNumber) : sheet.HeaderRow);
        sheet.Worksheet.Row(insertRow).InsertRowsAbove(1);
        var styleSourceAfter = styleSourceBefore >= insertRow ? styleSourceBefore + 1 : styleSourceBefore;
        var maxColumn = sheet.Worksheet.LastColumnUsed()?.ColumnNumber() ?? block.EndColumn;
        for (var column = 1; column <= maxColumn; column++)
        {
            var source = sheet.Worksheet.Cell(styleSourceAfter, column);
            var target = sheet.Worksheet.Cell(insertRow, column);
            target.Style = source.Style;
            target.Clear(XLClearOptions.Contents);
        }
        sheet.Worksheet.Row(insertRow).Height = sheet.Worksheet.Row(styleSourceAfter).Height;
        sheet.Worksheet.Cell(insertRow, sheet.FioColumn).Value = person.Name;
        if (sheet.PositionColumn.HasValue)
            sheet.Worksheet.Cell(insertRow, sheet.PositionColumn.Value).Value = person.Position;
        block = FindMonthBlocks(sheet.Worksheet).First(item => item.Month == month);
        WritePersonResults(sheet.Worksheet, insertRow, block, person.Scenarios, passingThreshold);
        var updatedPeople = ReadPeople(sheet.Worksheet, sheet.HeaderRow, sheet.FioColumn, sheet.PositionColumn);
        sheet.People.Clear();
        sheet.People.AddRange(updatedPeople);
        RefreshMonthFormulas(sheet.Worksheet, updatedPeople, FindMonthBlocks(sheet.Worksheet), passingThreshold);
        log?.Invoke($"[ДОБАВЛЕН] {sheet.Title}: {person.Name}");
        return true;
    }

    private static void RefreshMonthFormulas(
        IXLWorksheet sheet,
        IReadOnlyList<PersonRow> people,
        IReadOnlyList<MonthBlock> blocks,
        double passingThreshold)
    {
        if (people.Count == 0)
            return;
        var firstRow = people.Min(person => person.RowNumber);
        var lastRow = people.Max(person => person.RowNumber);
        var formulaRow = lastRow + 1;
        var thresholdText = passingThreshold.ToString("0.##", CultureInfo.InvariantCulture);
        foreach (var block in blocks)
        {
            if (block.ScenarioColumns.Count != 5 || !block.AttendanceColumn.HasValue)
                continue;
            var conditions = block.ScenarioColumns.OrderBy(pair => pair.Key)
                .Select(pair => $"({XLHelper.GetColumnLetterFromNumber(pair.Value)}{firstRow}:" +
                                $"{XLHelper.GetColumnLetterFromNumber(pair.Value)}{lastRow}>={thresholdText})");
            var firstScenarioColumn = block.ScenarioColumns.OrderBy(pair => pair.Key).First().Value;
            sheet.Cell(formulaRow, firstScenarioColumn).FormulaA1 = "SUMPRODUCT(" + string.Join("*", conditions) + ")";
            var attendanceLetter = XLHelper.GetColumnLetterFromNumber(block.AttendanceColumn.Value);
            sheet.Cell(formulaRow, block.AttendanceColumn.Value).FormulaA1 =
                $"COUNTIF({attendanceLetter}{firstRow}:{attendanceLetter}{lastRow},\"Да\")";
        }
    }

    private static void ApplyBlankConditionalFormatting(IXLWorksheet sheet, IReadOnlyList<PersonRow> people, IReadOnlyList<MonthBlock> blocks)
    {
        if (people.Count == 0)
            return;
        var first = people.Min(person => person.RowNumber);
        var last = people.Max(person => person.RowNumber);
        foreach (var block in blocks)
        foreach (var column in block.ScenarioColumns.Values.Append(block.AttendanceColumn ?? 0).Where(column => column > 0).Distinct())
        {
            var range = sheet.Range(first, column, last, column);
            range.AddConditionalFormat().WhenIsBlank().Fill.SetBackgroundColor(LegacyColor("F2F2F2"));
        }
    }

    private static IXLCell WritableCell(IXLCell cell) => cell.IsMerged() ? cell.MergedRange().FirstCell() : cell;

    private static XLColor LegacyColor(string rgb) => XLColor.FromArgb(
        0,
        Convert.ToByte(rgb[..2], 16),
        Convert.ToByte(rgb.Substring(2, 2), 16),
        Convert.ToByte(rgb.Substring(4, 2), 16));

    private sealed record CellTemplate(int Row, int Offset, XLCellValue Value, string? FormulaR1C1, IXLStyle Style);
    private sealed record MergeTemplate(int FirstRow, int LastRow, int FirstColumnOffset, int LastColumnOffset);
}

public sealed record WorkbookProcessingResult(
    IReadOnlyList<AnalysisSheet> Sheets,
    IReadOnlyList<NewPerson> NewPeople,
    IReadOnlyList<string> UnmatchedInstallations);

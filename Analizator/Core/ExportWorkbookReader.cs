using ClosedXML.Excel;

namespace Analizator.Core;

public static class ExportWorkbookReader
{
    private static readonly string[] FieldOrder = ["date", "topic", "person", "percent", "mode", "project"];
    private static readonly HashSet<string> RequiredFields = ["date", "topic", "person", "percent"];

    private static readonly Dictionary<string, HashSet<string>> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["date"] = ["дата", "дата прохождения", "дата отчета", "дата выполнения", "дата тренинга"],
        ["topic"] = ["тема", "сценарий", "название сценария", "наименование сценария", "название темы", "наименование темы"],
        ["person"] = ["фио", "ф и о", "фио сотрудника", "ф и о сотрудника", "сотрудник", "сотрудник фио", "фио по отчету", "фио по отчете", "фио по результату", "пользователь", "студент", "фио студента", "студент фио"],
        ["percent"] = ["%", "процент", "процент выполнения", "процент выполнения сценария", "результат", "результат %", "результат процент", "результат выполнения", "процент результата"],
        ["mode"] = ["режим", "режим прохождения", "режим обучения", "тип прохождения", "вид прохождения"],
        ["project"] = ["проект", "название проекта", "наименование проекта"]
    };

    public static ExportReadResult Read(
        string path,
        IReadOnlyDictionary<string, List<string>> installations,
        ModeFilterSettings modeFilter,
        CancellationToken cancellationToken,
        Action<string>? log = null)
    {
        var result = new ExportReadResult();
        var dates = new List<DateTime>();
        using var workbook = new XLWorkbook(path);
        foreach (var sheet in workbook.Worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var header = FindHeader(sheet);
            if (header.Row is null)
            {
                result.InstallationStatistics.Add(new InstallationStatistic(sheet.Name, null, 0, "заголовок не найден"));
                log?.Invoke($"[ОШИБКА] {sheet.Name}: заголовок выгрузки не найден");
                continue;
            }

            var installationMatch = MatchingService.MatchInstallation(sheet.Name, installations);
            if (installationMatch.Name is null)
            {
                result.InstallationStatistics.Add(new InstallationStatistic(sheet.Name, null, 0, "установка не сопоставлена", installationMatch.Score, installationMatch.Type));
                log?.Invoke($"[ОШИБКА] {sheet.Name}: установка не сопоставлена");
                continue;
            }

            if (header.Columns.ContainsKey("mode"))
                result.SheetsWithMode.Add(sheet.Name);
            else
                result.SheetsWithoutMode.Add(sheet.Name);

            var localReports = 0;
            var localFiltered = 0;
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? header.Row.Value;
            for (var row = header.Row.Value + 1; row <= lastRow; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var dateValue = CellValue(sheet.Cell(row, header.Columns["date"]));
                var topicValue = CellValue(sheet.Cell(row, header.Columns["topic"]));
                var personValue = CellValue(sheet.Cell(row, header.Columns["person"]));
                var percentValue = CellValue(sheet.Cell(row, header.Columns["percent"]));
                var modeValue = header.Columns.TryGetValue("mode", out var modeColumn)
                    ? CellValue(sheet.Cell(row, modeColumn))
                    : null;

                if (dateValue is null && topicValue is null && personValue is null)
                    continue;
                var person = personValue?.ToString()?.Trim();
                var scenario = topicValue?.ToString()?.Trim();
                if (string.IsNullOrEmpty(person) || string.IsNullOrEmpty(scenario))
                    continue;
                var date = TextNormalization.ParseDate(dateValue);
                if (date.HasValue)
                    dates.Add(date.Value);
                var percent = TextNormalization.SafeDouble(percentValue);
                if (!percent.HasValue)
                    continue;

                var normalizedMode = TextNormalization.NormalizeMode(modeValue);
                if (modeFilter.Enabled && !TextNormalization.ModeIsAllowed(modeValue, modeFilter))
                {
                    localFiltered++;
                    result.FilteredReports++;
                    if (normalizedMode.Length == 0)
                        result.ReportsWithoutMode++;
                    else
                        CountMode(result, modeValue, false);
                    continue;
                }
                if (normalizedMode.Length > 0)
                    CountMode(result, modeValue, true);

                localReports++;
                result.TotalReports++;
                if (!result.Data.TryGetValue(installationMatch.Name, out var people))
                    result.Data[installationMatch.Name] = people = new(StringComparer.OrdinalIgnoreCase);
                if (!people.TryGetValue(person, out var scenarios))
                    people[person] = scenarios = new(StringComparer.OrdinalIgnoreCase);
                var key = TextNormalization.NormalizeText(scenario);
                if (!scenarios.TryGetValue(key, out var aggregate))
                    scenarios[key] = new ScenarioAggregate
                    {
                        DisplayName = scenario,
                        BestPercent = percent.Value,
                        Attempts = 1
                    };
                else
                {
                    aggregate.Attempts++;
                    aggregate.BestPercent = Math.Max(aggregate.BestPercent, percent.Value);
                }
            }

            result.InstallationStatistics.Add(new InstallationStatistic(
                sheet.Name, installationMatch.Name, localReports, "OK", installationMatch.Score, installationMatch.Type));
            log?.Invoke($"{sheet.Name} → {installationMatch.Name}: {localReports} отчётов" +
                        (localFiltered > 0 ? $"; отфильтровано: {localFiltered}" : ""));
        }

        result.Month = dates.Count == 0
            ? TextNormalization.MonthFromFileName(path)
            : dates.GroupBy(date => date.Month).OrderByDescending(group => group.Count()).First().Key;
        return result;
    }

    private static void CountMode(ExportReadResult result, object? modeValue, bool accepted)
    {
        var known = TextNormalization.KnownMode(modeValue);
        if (known is not null)
        {
            result.ModeCounts.TryGetValue(known, out var count);
            result.ModeCounts[known] = count + 1;
            return;
        }
        var normalized = TextNormalization.NormalizeMode(modeValue);
        if (normalized.Length == 0)
            return;
        result.UnknownModes.TryGetValue(normalized, out var unknownCount);
        result.UnknownModes[normalized] = unknownCount + 1;
    }

    private static object? CellValue(IXLCell cell)
    {
        if (cell.IsEmpty())
            return null;
        return cell.DataType switch
        {
            XLDataType.DateTime => cell.GetDateTime(),
            XLDataType.Number => cell.GetDouble(),
            XLDataType.Boolean => cell.GetBoolean(),
            XLDataType.TimeSpan => cell.GetTimeSpan(),
            _ => cell.GetString()
        };
    }

    private static (int? Row, Dictionary<string, int> Columns) FindHeader(IXLWorksheet sheet)
    {
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        var lastRow = Math.Min(50, sheet.LastRowUsed()?.RowNumber() ?? 0);
        int? bestRow = null;
        var bestColumns = new Dictionary<string, int>();
        var bestScore = int.MinValue;
        for (var row = 1; row <= lastRow; row++)
        {
            var columns = new Dictionary<string, int>();
            for (var column = 1; column <= lastColumn; column++)
            {
                var normalized = TextNormalization.NormalizeHeader(sheet.Cell(row, column).GetString());
                var field = FieldOrder.FirstOrDefault(candidate =>
                    !columns.ContainsKey(candidate) && MatchesHeader(normalized, candidate));
                if (field is not null)
                    columns[field] = column;
            }
            if (!RequiredFields.All(columns.ContainsKey))
                continue;
            var score = 100 + (columns.ContainsKey("mode") ? 10 : 0) + (columns.ContainsKey("project") ? 5 : 0);
            var requiredColumns = RequiredFields.Select(field => columns[field]).ToArray();
            if (requiredColumns.Max() - requiredColumns.Min() + 1 <= 10)
                score += 5;
            if (score > bestScore)
            {
                bestScore = score;
                bestRow = row;
                bestColumns = columns;
            }
        }
        return (bestRow, bestColumns);
    }

    private static bool MatchesHeader(string value, string field)
    {
        if (value.Length == 0)
            return false;
        if (Aliases[field].Contains(value))
            return true;
        return field switch
        {
            "date" => value.StartsWith("дата ", StringComparison.Ordinal),
            "topic" => new[] { "тема ", "название темы", "наименование темы", "название сценария", "наименование сценария", "сценарий " }.Any(value.StartsWith),
            "person" => value.StartsWith("фио ") || value.StartsWith("ф и о ") || value.StartsWith("студент ") ||
                        (value.Contains("фио") && new[] { "сотрудник", "отчет", "пользователь", "студент" }.Any(value.Contains)),
            "percent" => value.Contains("процент") || value.Contains("результат") && value.Contains('%'),
            "mode" => value.StartsWith("режим") || value.EndsWith("прохождения"),
            "project" => value.StartsWith("проект"),
            _ => false
        };
    }
}

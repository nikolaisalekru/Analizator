using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;

namespace Analizator.Core;

public enum TrainingExportFormat
{
    Current,
    Legacy,
    Mixed,
    Json
}

public sealed record JsonAttemptMetadata(
    string? Guid,
    string? MachineName,
    string? PlantName,
    string? StudentSerial,
    DateTime? EntryDate,
    DateTime? SessionStartDate,
    DateTime? SessionEndDate,
    bool? IsClosed,
    string? FilePath,
    int? Resume,
    bool? Result,
    bool? IsValid,
    int? ReportType,
    bool? HasReportFile,
    string? ReportFile,
    string? TrainingTimesJson,
    string? ExerciseLogsJson,
    int? TrainingTimesCount,
    int? ExerciseLogsCount);

public sealed record TrainingAttemptRecord(
    string Installation,
    string PersonName,
    string ScenarioName,
    DateTime AttemptedAt,
    double Percent,
    string? Mode,
    string? Project,
    string? Role,
    string? Credited,
    string? Duration,
    int? Actions,
    string SourceSheet,
    int SourceRow,
    JsonAttemptMetadata? JsonMetadata = null);

public sealed record ExportImportError(string Sheet, int? Row, string Message);

public sealed class ParsedTrainingExport
{
    public required TrainingExportFormat Format { get; init; }
    public List<TrainingAttemptRecord> Attempts { get; } = [];
    public List<ExportImportError> Errors { get; } = [];
    public DateTime? PeriodStart => Attempts.Count == 0 ? null : Attempts.Min(item => item.AttemptedAt);
    public DateTime? PeriodEnd => Attempts.Count == 0 ? null : Attempts.Max(item => item.AttemptedAt);
}

/// <summary>
/// Reads both generations of KTK Excel exports into lossless row-level records.
/// </summary>
public static class TrainingExportParser
{
    public const int Version = 3;

    private static readonly string[] FieldOrder =
    [
        "date", "topic", "person", "percent", "mode", "project",
        "duration", "actions", "role", "credited"
    ];

    private static readonly HashSet<string> RequiredFields =
        ["date", "topic", "person", "percent"];

    private static readonly Dictionary<string, HashSet<string>> Aliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["date"] = ["дата", "дата прохождения", "дата отчета", "дата выполнения", "дата тренинга"],
            ["topic"] = ["тема", "сценарий", "название сценария", "наименование сценария", "название темы", "наименование темы"],
            ["person"] = ["фио", "ф и о", "фио сотрудника", "ф и о сотрудника", "сотрудник", "сотрудник фио", "фио по отчету", "фио по отчете", "фио по результату", "пользователь", "студент", "фио студента", "студент фио"],
            ["percent"] = ["%", "процент", "процент выполнения", "процент выполнения сценария", "результат", "результат %", "результат процент", "результат выполнения", "процент результата"],
            ["mode"] = ["режим", "режим прохождения", "режим обучения", "тип прохождения", "вид прохождения"],
            ["project"] = ["проект", "название проекта", "наименование проекта"],
            ["duration"] = ["время", "длительность", "время прохождения"],
            ["actions"] = ["действий", "количество действий", "число действий"],
            ["role"] = ["роль", "роль студента"],
            ["credited"] = ["зачтено", "зачет", "статус зачета"]
        };

    public static ParsedTrainingExport Parse(
        string path,
        IReadOnlyDictionary<string, List<string>> installations,
        CancellationToken cancellationToken = default,
        Action<int, string>? progress = null)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Файл выгрузки не найден.", path);

        var attempts = new List<TrainingAttemptRecord>();
        var errors = new List<ExportImportError>();
        var sawCurrent = false;
        var sawLegacy = false;

        using var workbook = OpenWorkbook(path, progress);
        var sheets = workbook.Worksheets.ToArray();
        for (var sheetIndex = 0; sheetIndex < sheets.Length; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = sheets[sheetIndex];
            progress?.Invoke(
                sheets.Length == 0 ? 100 : (int)Math.Round(sheetIndex * 100d / sheets.Length),
                $"Чтение листа «{sheet.Name}»");

            var header = FindHeader(sheet);
            if (header.Row is null)
            {
                errors.Add(new ExportImportError(sheet.Name, null, "Не найдена строка заголовков выгрузки."));
                continue;
            }

            if (header.Columns.ContainsKey("mode"))
                sawCurrent = true;
            if (header.Columns.ContainsKey("project") || header.Columns.ContainsKey("role") ||
                header.Columns.ContainsKey("credited"))
                sawLegacy = true;

            var installationMatch = MatchingService.MatchInstallation(sheet.Name, installations);
            var installation = installationMatch.Name;
            if (installation is null)
            {
                errors.Add(new ExportImportError(sheet.Name, null,
                    "Установка не найдена в настройках. Лист пропущен; " +
                    "добавьте его название как синоним нужной установки."));
                continue;
            }

            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? header.Row.Value;
            for (var row = header.Row.Value + 1; row <= lastRow; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var dateValue = ReadValue(sheet.Cell(row, header.Columns["date"]));
                var topicValue = ReadValue(sheet.Cell(row, header.Columns["topic"]));
                var personValue = ReadValue(sheet.Cell(row, header.Columns["person"]));
                var percentCell = sheet.Cell(row, header.Columns["percent"]);

                if (dateValue is null && topicValue is null && personValue is null && percentCell.IsEmpty())
                    continue;

                var person = personValue?.ToString()?.Trim();
                var scenario = topicValue?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(person) || string.IsNullOrWhiteSpace(scenario))
                {
                    errors.Add(new ExportImportError(sheet.Name, row, "Не заполнено ФИО или название сценария."));
                    continue;
                }

                var attemptedAt = ParseDateTime(dateValue);
                if (!attemptedAt.HasValue)
                {
                    errors.Add(new ExportImportError(sheet.Name, row, "Не удалось прочитать дату прохождения."));
                    continue;
                }

                var percent = ReadPercent(percentCell);
                if (!percent.HasValue || percent.Value is < 0 or > 100)
                {
                    errors.Add(new ExportImportError(sheet.Name, row, "Процент должен быть числом от 0 до 100."));
                    continue;
                }

                attempts.Add(new TrainingAttemptRecord(
                    installation,
                    person,
                    scenario,
                    attemptedAt.Value,
                    percent.Value,
                    ReadOptionalText(sheet, row, header.Columns, "mode"),
                    ReadOptionalText(sheet, row, header.Columns, "project"),
                    ReadOptionalText(sheet, row, header.Columns, "role"),
                    ReadOptionalText(sheet, row, header.Columns, "credited"),
                    ReadOptionalText(sheet, row, header.Columns, "duration"),
                    ReadOptionalInt(sheet, row, header.Columns, "actions"),
                    sheet.Name,
                    row));
            }
        }

        progress?.Invoke(100, "Чтение выгрузки завершено");
        var parsed = new ParsedTrainingExport
        {
            Format = sawCurrent && sawLegacy
                ? TrainingExportFormat.Mixed
                : sawLegacy && !sawCurrent
                    ? TrainingExportFormat.Legacy
                    : TrainingExportFormat.Current
        };
        parsed.Attempts.AddRange(attempts);
        parsed.Errors.AddRange(errors);
        return parsed;
    }

    private static string? ReadOptionalText(
        IXLWorksheet sheet,
        int row,
        IReadOnlyDictionary<string, int> columns,
        string field)
    {
        if (!columns.TryGetValue(field, out var column))
            return null;
        var text = ReadValue(sheet.Cell(row, column))?.ToString()?.Trim();
        return string.IsNullOrEmpty(text) ? null : text;
    }

    private static XLWorkbook OpenWorkbook(string path, Action<int, string>? progress)
    {
        try
        {
            return new XLWorkbook(path);
        }
        catch (Exception exception) when (
            string.Equals(exception.GetType().FullName,
                "ClosedXML.Excel.IO.PartStructureException", StringComparison.Ordinal))
        {
            // Some exports contain malformed pivot-cache records. The pivot tables are only
            // auxiliary Excel objects and are not used by the importer, so remove them from
            // an in-memory copy and leave the user's source workbook untouched.
            progress?.Invoke(1, "Восстановление структуры книги во временной копии");
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var repairedCopy = new MemoryStream();
            source.CopyTo(repairedCopy);
            repairedCopy.Position = 0;

            using (var document = SpreadsheetDocument.Open(repairedCopy, true))
            {
                var workbookPart = document.WorkbookPart
                    ?? throw new InvalidDataException("В книге Excel отсутствует основная часть workbook.");
                foreach (var worksheetPart in workbookPart.WorksheetParts)
                    foreach (var pivotTablePart in worksheetPart.PivotTableParts.ToArray())
                        worksheetPart.DeletePart(pivotTablePart);
                foreach (var cachePart in workbookPart.PivotTableCacheDefinitionParts.ToArray())
                    workbookPart.DeletePart(cachePart);
                workbookPart.Workbook.PivotCaches?.Remove();
                workbookPart.Workbook.Save();
            }

            repairedCopy.Position = 0;
            return new XLWorkbook(repairedCopy);
        }
    }

    private static int? ReadOptionalInt(
        IXLWorksheet sheet,
        int row,
        IReadOnlyDictionary<string, int> columns,
        string field)
    {
        if (!columns.TryGetValue(field, out var column))
            return null;
        var number = TextNormalization.SafeDouble(ReadValue(sheet.Cell(row, column)));
        return number.HasValue ? (int)Math.Round(number.Value) : null;
    }

    private static double? ReadPercent(IXLCell cell)
    {
        var value = TextNormalization.SafeDouble(ReadValue(cell));
        if (!value.HasValue)
            return null;
        var format = cell.Style.NumberFormat.Format ?? "";
        return format.Contains('%') && value.Value is >= 0 and <= 1
            ? value.Value * 100d
            : value.Value;
    }

    private static object? ReadValue(IXLCell cell)
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

    private static DateTime? ParseDateTime(object? value)
    {
        if (value is DateTime dateTime)
            return dateTime;
        if (value is double serial && serial is > 0 and < 2958466)
            return DateTime.FromOADate(serial);

        var text = value?.ToString()?.Trim();
        if (string.IsNullOrEmpty(text))
            return null;
        string[] formats =
        [
            "dd.MM.yyyy", "dd.MM.yyyy H:mm", "dd.MM.yyyy HH:mm",
            "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy HH:mm:ss",
            "yyyy-MM-dd", "yyyy-MM-dd H:mm:ss", "yyyy-MM-dd HH:mm:ss",
            "M/d/yyyy", "M/d/yyyy H:mm:ss"
        ];
        if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out dateTime))
            return dateTime;
        return DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"),
            DateTimeStyles.AllowWhiteSpaces, out dateTime)
            ? dateTime
            : null;
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
            var score = 100 + (columns.ContainsKey("mode") ? 10 : 0) +
                        (columns.ContainsKey("project") ? 5 : 0);
            var requiredColumns = RequiredFields.Select(field => columns[field]).ToArray();
            if (requiredColumns.Max() - requiredColumns.Min() + 1 <= 10)
                score += 5;
            if (score <= bestScore)
                continue;
            bestScore = score;
            bestRow = row;
            bestColumns = columns;
        }

        return (bestRow, bestColumns);
    }

    private static bool MatchesHeader(string value, string field)
    {
        if (value.Length == 0 || !Aliases.TryGetValue(field, out var aliases))
            return false;
        if (aliases.Contains(value))
            return true;
        return field switch
        {
            "date" => value.StartsWith("дата ", StringComparison.Ordinal),
            "topic" => new[] { "тема ", "название темы", "наименование темы", "название сценария", "наименование сценария", "сценарий " }.Any(value.StartsWith),
            "person" => value.StartsWith("фио ") || value.StartsWith("ф и о ") || value.StartsWith("студент ") ||
                        value.Contains("фио") && new[] { "сотрудник", "отчет", "пользователь", "студент" }.Any(value.Contains),
            "percent" => value.Contains("процент") || value.Contains("результат") && value.Contains('%'),
            "mode" => value.StartsWith("режим") || value.EndsWith("прохождения"),
            "project" => value.StartsWith("проект"),
            "duration" => value.StartsWith("время") || value.Contains("длительност"),
            "actions" => value.Contains("действ"),
            "role" => value.StartsWith("роль"),
            "credited" => value.StartsWith("зачт"),
            _ => false
        };
    }
}

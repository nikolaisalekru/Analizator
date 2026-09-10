using ClosedXML.Excel;
using System.Globalization;

namespace Analizator.Core;

public sealed class DatabaseSelectionExportService
{
    public static string SuggestedFileName(DateTime periodStart, DateTime periodEnd)
    {
        var start = periodStart.Date;
        var end = periodEnd.Date;
        if (start.Day == 1 && end == start.AddMonths(1).AddDays(-1))
        {
            var month = TextNormalization.MonthNames[start.Month].ToLowerInvariant();
            return $"Выгрузка_КТК_МНПЗ_{month}_{start:yyyy}.xlsx";
        }

        return $"Выгрузка_КТК_МНПЗ_{start:yyyy-MM-dd}_{end:yyyy-MM-dd}.xlsx";
    }

    public DatabaseSelectionExportResult Export(
        TrainingDatabaseService database,
        DatabaseSelectionExportRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(database);
        ArgumentNullException.ThrowIfNull(request);
        ValidateRequest(request);

        var selectedInstallations = request.Installations
            .Select(TextNormalization.NormalizeText)
            .ToHashSet(StringComparer.Ordinal);
        var selectedPeople = request.People
            .Select(item => item.NormalizedValue)
            .ToHashSet(StringComparer.Ordinal);
        var filterPeople = selectedPeople.Count > 0;
        var attempts = database.GetAttempts(request.PeriodStart, request.PeriodEnd)
            .Where(item => selectedInstallations.Contains(
                TextNormalization.NormalizeText(item.Installation)))
            .Where(item => !filterPeople ||
                selectedPeople.Contains(PersonKey(item.Installation, item.DisplayPersonName)))
            .OrderBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.AttemptedAt)
            .ThenBy(item => item.DisplayPersonName, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.ScenarioName, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();

        if (attempts.Length == 0)
            throw new InvalidOperationException(
                filterPeople
                    ? "За выбранный период у отмеченных сотрудников нет записей для выгрузки."
                    : "За выбранный период у отмеченных установок нет записей для выгрузки.");

        var outputFile = Path.GetFullPath(request.OutputFile);
        var outputDirectory = Path.GetDirectoryName(outputFile)
            ?? throw new InvalidOperationException("Не удалось определить папку результата.");
        Directory.CreateDirectory(outputDirectory);

        try
        {
            using var workbook = new XLWorkbook();
            var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var installationGroup in attempts.GroupBy(item => item.Installation))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var sheetName = UniqueWorksheetName(installationGroup.Key, usedSheetNames);
                var sheet = workbook.AddWorksheet(sheetName);
                WriteInstallationSheet(
                    sheet,
                    installationGroup.Key,
                    installationGroup.ToArray(),
                    request,
                    cancellationToken);
            }

            cancellationToken.ThrowIfCancellationRequested();
            workbook.SaveAs(outputFile);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new AnalysisFileAccessException(outputFile, exception);
        }

        return new DatabaseSelectionExportResult(
            outputFile,
            attempts.Length,
            attempts.Select(item => PersonKey(item.Installation, item.DisplayPersonName)).Distinct().Count(),
            attempts.Select(item => TextNormalization.NormalizeText(item.Installation)).Distinct().Count());
    }

    private static void ValidateRequest(DatabaseSelectionExportRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.OutputFile))
            throw new ArgumentException("Не указан файл результата.", nameof(request));
        if (request.PeriodEnd.Date < request.PeriodStart.Date)
            throw new ArgumentException("Дата окончания периода не может быть раньше даты начала.");
        if (!request.Installations.Any(item => !string.IsNullOrWhiteSpace(item)))
            throw new ArgumentException("Выберите хотя бы одну установку.");
        if (request.Columns.Count == 0)
            throw new ArgumentException("Выберите столбцы для выгрузки.");
        if (request.Columns.Distinct().Count() != request.Columns.Count)
            throw new ArgumentException("Список столбцов содержит повторы.");
    }

    private static void WriteInstallationSheet(
        IXLWorksheet sheet,
        string installation,
        IReadOnlyList<DatabaseAttempt> attempts,
        DatabaseSelectionExportRequest request,
        CancellationToken cancellationToken)
    {
        const int headerRow = 14;
        const int firstDataRow = headerRow + 1;
        const int firstDataColumn = 2;
        var lastDataColumn = firstDataColumn + request.Columns.Count - 1;
        var layoutLastColumn = Math.Max(8, lastDataColumn);

        sheet.Cell(1, 1).Value = "Генератор Отчетов";
        sheet.Range(1, 1, 1, layoutLastColumn).Merge();
        sheet.Cell(2, 1).Value = $"Выгрузка КТК · {installation}";
        sheet.Range(2, 1, 2, layoutLastColumn).Merge();
        sheet.Cell(4, 1).Value =
            $"Период: {request.PeriodStart:dd.MM.yyyy} — {request.PeriodEnd:dd.MM.yyyy}";
        sheet.Range(4, 1, 4, layoutLastColumn).Merge();

        sheet.Cell(5, 1).Value = "Параметры:";
        sheet.Range(5, 1, 5, layoutLastColumn).Merge();
        sheet.Cell(6, 1).Value = "Отчеты:";
        sheet.Cell(6, 2).Value = attempts.Count;
        sheet.Cell(7, 1).Value = "Сценарии:";
        sheet.Cell(7, 2).Value = attempts
            .Select(item => TextNormalization.NormalizeText(item.ScenarioName)).Distinct().Count();
        sheet.Cell(8, 1).Value = "Студенты:";
        sheet.Cell(8, 2).Value = attempts
            .Select(item => TextNormalization.NormalizePersonName(item.DisplayPersonName)).Distinct().Count();
        sheet.Cell(9, 1).Value = "Проект:";
        sheet.Cell(9, 2).Value = ProjectSummary(attempts);
        sheet.Cell(10, 1).Value = "Тип записи:";
        sheet.Cell(10, 2).Value = request.People.Count == 0
            ? "Все"
            : "Выбранные сотрудники";

        sheet.Range(11, 1, 11, layoutLastColumn).Merge();
        sheet.Range(12, 1, 12, layoutLastColumn).Merge();
        var dateColumnIndex = request.Columns
            .Select((column, index) => (column, index))
            .Where(item => item.column == DatabaseExportColumn.Date)
            .Select(item => item.index)
            .DefaultIfEmpty(-1)
            .First();
        if (dateColumnIndex >= 0)
        {
            var dateColumn = firstDataColumn + dateColumnIndex;
            sheet.Cell(13, dateColumn).Value = "Общие данные";
            if (dateColumn > firstDataColumn)
            {
                sheet.Cell(13, firstDataColumn).Value = "Отчеты";
                if (dateColumn - firstDataColumn > 1)
                    sheet.Range(13, firstDataColumn, 13, dateColumn - 1).Merge();
            }
            if (dateColumn < lastDataColumn)
            {
                sheet.Cell(13, dateColumn + 1).Value = "Отчеты";
                if (lastDataColumn - dateColumn > 1)
                    sheet.Range(13, dateColumn + 1, 13, lastDataColumn).Merge();
            }
        }
        else
        {
            sheet.Cell(13, firstDataColumn).Value = "Отчеты";
            if (lastDataColumn > firstDataColumn)
                sheet.Range(13, firstDataColumn, 13, lastDataColumn).Merge();
        }

        for (var index = 0; index < request.Columns.Count; index++)
        {
            var column = firstDataColumn + index;
            sheet.Cell(headerRow, column).Value = Header(request.Columns[index]);
            sheet.Column(column).Width = ColumnWidth(request.Columns[index]);
        }

        for (var index = 0; index < attempts.Count; index++)
        {
            if ((index & 255) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            var row = firstDataRow + index;
            for (var columnIndex = 0; columnIndex < request.Columns.Count; columnIndex++)
                WriteValue(sheet.Cell(row, firstDataColumn + columnIndex),
                    request.Columns[columnIndex], attempts[index]);
        }

        var lastDataRow = firstDataRow + attempts.Count - 1;
        StyleSheet(sheet, headerRow, firstDataRow, lastDataRow, firstDataColumn,
            lastDataColumn, layoutLastColumn, request.Columns);

    }

    private static void StyleSheet(
        IXLWorksheet sheet,
        int headerRow,
        int firstDataRow,
        int lastDataRow,
        int firstDataColumn,
        int lastDataColumn,
        int layoutLastColumn,
        IReadOnlyList<DatabaseExportColumn> columns)
    {
        var title = sheet.Range(1, 1, 1, layoutLastColumn);
        title.Style.Font.FontName = "Tahoma";
        title.Style.Font.FontSize = 20;
        title.Style.Font.Bold = true;
        title.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        title.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        sheet.Row(1).Height = 33;

        var subtitle = sheet.Range(2, 1, 2, layoutLastColumn);
        subtitle.Style.Font.FontName = "Tahoma";
        subtitle.Style.Font.FontSize = 10;
        subtitle.Style.Font.FontColor = XLColor.FromHtml("#6E7C91");
        subtitle.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet.Row(2).Height = 20;

        var period = sheet.Range(4, 1, 4, layoutLastColumn);
        period.Style.Font.FontName = "Tahoma";
        period.Style.Font.FontSize = 9;
        period.Style.Font.Italic = true;
        period.Style.Font.FontColor = XLColor.FromHtml("#44546A");

        var parameters = sheet.Range(5, 1, 10, layoutLastColumn);
        parameters.Style.Font.FontName = "Tahoma";
        parameters.Style.Font.FontSize = 8;
        parameters.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        parameters.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        parameters.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        sheet.Range(5, 1, 10, 1).Style.Font.Bold = true;

        var groupHeader = sheet.Range(13, firstDataColumn, 13, lastDataColumn);
        var tableHeader = sheet.Range(headerRow, firstDataColumn, headerRow, lastDataColumn);
        foreach (var range in new[] { groupHeader, tableHeader })
        {
            range.Style.Fill.BackgroundColor = XLColor.FromHtml("#DCDCDC");
            range.Style.Font.FontName = "Times New Roman";
            range.Style.Font.FontSize = 9;
            range.Style.Font.Bold = true;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Alignment.WrapText = true;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        sheet.Row(13).Height = 21;
        sheet.Row(headerRow).Height = 29;

        var dataRange = sheet.Range(firstDataRow, firstDataColumn, lastDataRow, lastDataColumn);
        dataRange.Style.Font.FontName = "Times New Roman";
        dataRange.Style.Font.FontSize = 10;
        dataRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

        for (var index = 0; index < columns.Count; index++)
        {
            var columnNumber = firstDataColumn + index;
            var columnRange = sheet.Range(firstDataRow, columnNumber, lastDataRow, columnNumber);
            columnRange.Style.Alignment.Horizontal = columns[index] is DatabaseExportColumn.Scenario
                or DatabaseExportColumn.Person or DatabaseExportColumn.Project
                or DatabaseExportColumn.Role or DatabaseExportColumn.Source
                ? XLAlignmentHorizontalValues.Left
                : XLAlignmentHorizontalValues.Center;

            if (columns[index] == DatabaseExportColumn.Date)
                columnRange.Style.DateFormat.Format = "dd.MM.yyyy";
            if (columns[index] == DatabaseExportColumn.Duration)
                columnRange.Style.DateFormat.Format = "[h]:mm:ss";
            if (columns[index] == DatabaseExportColumn.Percent)
            {
                columnRange.Style.NumberFormat.Format = "0.00";
                for (var row = firstDataRow; row <= lastDataRow; row++)
                {
                    var value = sheet.Cell(row, columnNumber).GetDouble();
                    sheet.Cell(row, columnNumber).Style.Fill.BackgroundColor = value >= 100
                        ? XLColor.FromHtml("#C6E0B4")
                        : XLColor.FromHtml("#FFB3B3");
                }
            }
        }

        sheet.Column(1).Width = 9.14;
        sheet.Range(headerRow, firstDataColumn, lastDataRow, lastDataColumn).SetAutoFilter();
        sheet.SheetView.FreezeRows(headerRow);
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        sheet.PageSetup.FitToPages(1, 0);
        sheet.PageSetup.Margins.Left = 0.25;
        sheet.PageSetup.Margins.Right = 0.25;
        sheet.PageSetup.Margins.Top = 0.4;
        sheet.PageSetup.Margins.Bottom = 0.4;
    }

    private static void WriteValue(
        IXLCell cell,
        DatabaseExportColumn column,
        DatabaseAttempt attempt)
    {
        switch (column)
        {
            case DatabaseExportColumn.Date:
                cell.Value = attempt.AttemptedAt;
                break;
            case DatabaseExportColumn.Scenario:
                cell.Value = attempt.ScenarioName;
                break;
            case DatabaseExportColumn.Person:
                cell.Value = attempt.DisplayPersonName;
                break;
            case DatabaseExportColumn.Duration:
                if (TimeSpan.TryParse(attempt.Duration, CultureInfo.InvariantCulture,
                        out var duration))
                    cell.Value = duration;
                else
                    cell.Value = attempt.Duration ?? "";
                break;
            case DatabaseExportColumn.Actions:
                if (attempt.Actions.HasValue)
                    cell.Value = attempt.Actions.Value;
                break;
            case DatabaseExportColumn.Percent:
                cell.Value = attempt.Percent;
                break;
            case DatabaseExportColumn.Mode:
                cell.Value = attempt.Mode ?? "";
                break;
            case DatabaseExportColumn.Project:
                cell.Value = attempt.Project ?? "";
                break;
            case DatabaseExportColumn.Role:
                cell.Value = attempt.Role ?? "";
                break;
            case DatabaseExportColumn.Credited:
                cell.Value = attempt.Credited ?? "";
                break;
            case DatabaseExportColumn.Source:
                cell.Value = attempt.SourceFile;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(column), column, null);
        }
    }

    private static string Header(DatabaseExportColumn column) => column switch
    {
        DatabaseExportColumn.Date => "Дата",
        DatabaseExportColumn.Scenario => "Тема",
        DatabaseExportColumn.Person => "ФИО по отчёту",
        DatabaseExportColumn.Duration => "Время",
        DatabaseExportColumn.Actions => "Действий",
        DatabaseExportColumn.Percent => "Процент",
        DatabaseExportColumn.Mode => "Режим",
        DatabaseExportColumn.Project => "Проект",
        DatabaseExportColumn.Role => "Роль",
        DatabaseExportColumn.Credited => "Зачтено",
        DatabaseExportColumn.Source => "Источник",
        _ => throw new ArgumentOutOfRangeException(nameof(column), column, null)
    };

    private static double ColumnWidth(DatabaseExportColumn column) => column switch
    {
        DatabaseExportColumn.Date => 14.7,
        DatabaseExportColumn.Scenario => 54.7,
        DatabaseExportColumn.Person => 30.7,
        DatabaseExportColumn.Duration => 10.5,
        DatabaseExportColumn.Actions => 13.7,
        DatabaseExportColumn.Percent => 12.7,
        DatabaseExportColumn.Mode => 23.7,
        DatabaseExportColumn.Project => 24,
        DatabaseExportColumn.Role => 20,
        DatabaseExportColumn.Credited => 13,
        DatabaseExportColumn.Source => 30,
        _ => 14
    };

    private static string ProjectSummary(IReadOnlyList<DatabaseAttempt> attempts)
    {
        var projects = attempts.Select(item => item.Project?.Trim())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .Take(3)
            .ToArray();
        return projects.Length switch
        {
            0 => "Все",
            1 => projects[0]!,
            2 => string.Join(", ", projects),
            _ => "Несколько"
        };
    }

    private static string UniqueWorksheetName(string value, ISet<string> used)
    {
        var invalid = new HashSet<char>(['[', ']', ':', '*', '?', '/', '\\']);
        var baseName = new string(value.Trim().Select(character =>
            invalid.Contains(character) ? '_' : character).ToArray());
        if (string.IsNullOrWhiteSpace(baseName))
            baseName = "Установка";
        if (baseName.Length > 31)
            baseName = baseName[..31];

        var candidate = baseName;
        for (var number = 2; !used.Add(candidate); number++)
        {
            var suffix = $" ({number})";
            candidate = baseName[..Math.Min(baseName.Length, 31 - suffix.Length)] + suffix;
        }
        return candidate;
    }

    private static string PersonKey(string installation, string personName) =>
        TextNormalization.NormalizeText(installation) + "\u001F" +
        TextNormalization.NormalizePersonName(personName);
}

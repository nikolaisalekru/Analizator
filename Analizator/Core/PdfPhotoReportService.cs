using ClosedXML.Excel;
using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Analizator.Core;

public sealed partial class PdfPhotoReportService
{
    private const string FontName = "Arial";
    private static readonly XLColor Navy = XLColor.FromHtml("#17365D");
    private static readonly XLColor Blue = XLColor.FromHtml("#4472C4");
    private static readonly XLColor LightBlue = XLColor.FromHtml("#D9EAF7");
    private static readonly XLColor LightGreen = XLColor.FromHtml("#E2F0D9");
    private static readonly XLColor LightRed = XLColor.FromHtml("#FCE4D6");
    private static readonly XLColor LightAmber = XLColor.FromHtml("#FFF2CC");
    private static readonly CultureInfo RussianCulture = CultureInfo.GetCultureInfo("ru-RU");

    public PdfPhotoReportResult Run(
        PdfPhotoReportRequest request,
        IProgress<PdfPhotoProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        Validate(request);
        progress?.Report(new PdfPhotoProgress(2, "Поиск PDF-файлов"));
        var enumerationNotes = new ConcurrentBag<string>();
        var files = EnumeratePdfFiles(request.SourceFolder, enumerationNotes)
            .OrderBy(path => path, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
        if (files.Length == 0)
            throw new InvalidOperationException("В выбранной папке и её подпапках PDF-файлы не найдены.");

        var entries = new ConcurrentBag<PdfPhotoEntry>();
        var completed = 0;
        Parallel.ForEach(files, new ParallelOptions
        {
            CancellationToken = cancellationToken,
            MaxDegreeOfParallelism = Math.Clamp(Environment.ProcessorCount / 2, 1, 6)
        }, file =>
        {
            var entry = ReadPdf(file, request.SourceFolder, request.MinimumPhotoWidth, request.MinimumPhotoHeight);
            entries.Add(entry);
            var done = Interlocked.Increment(ref completed);
            var percent = 5 + (int)Math.Round(done * 78d / files.Length);
            progress?.Report(new PdfPhotoProgress(percent, "Проверка фотографий и реквизитов", Path.GetFileName(file)));
        });

        cancellationToken.ThrowIfCancellationRequested();
        var ordered = entries
            .OrderBy(item => item.AttemptedAt ?? DateTime.MaxValue)
            .ThenBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.PersonName, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.PdfPath, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
        var filtered = ordered.Where(item =>
                (!request.PeriodStart.HasValue || !item.AttemptedAt.HasValue || item.AttemptedAt.Value.Date >= request.PeriodStart.Value.Date) &&
                (!request.PeriodEnd.HasValue || !item.AttemptedAt.HasValue || item.AttemptedAt.Value.Date <= request.PeriodEnd.Value.Date))
            .ToArray();
        if (filtered.Length == 0)
            throw new InvalidOperationException("В выбранном периоде PDF-отчёты не найдены.");

        progress?.Report(new PdfPhotoProgress(86, "Формирование Excel-таблицы"));
        WriteWorkbook(request, filtered, enumerationNotes.OrderBy(x => x).ToArray());
        progress?.Report(new PdfPhotoProgress(100, "Проверка PDF завершена"));
        return new PdfPhotoReportResult
        {
            OutputFile = Path.GetFullPath(request.OutputFile),
            PdfFiles = filtered.Length,
            WithPhoto = filtered.Count(item => item.HasPhoto),
            WithoutPhoto = filtered.Count(item => !item.HasPhoto && !IsError(item)),
            ReadErrors = filtered.Count(IsError),
            MissingPerson = filtered.Count(item => item.PersonName.Length == 0),
            MissingScenario = filtered.Count(item => item.Scenario.Length == 0),
            Installations = filtered.Select(item => item.Installation)
                .Distinct(StringComparer.CurrentCultureIgnoreCase).Count()
        };
    }

    private static PdfPhotoEntry ReadPdf(string pdfPath, string sourceRoot, int minWidth, int minHeight)
    {
        var installation = InferInstallation(sourceRoot, pdfPath);
        var reportFolder = InferReportFolder(sourceRoot, pdfPath);
        try
        {
            using var stream = new FileStream(pdfPath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            using var document = PdfDocument.Open(memory.ToArray());
            var text = new StringBuilder();
            var matchingImages = 0;
            foreach (var page in document.GetPages())
            {
                text.AppendLine(ContentOrderTextExtractor.GetText(page));
                matchingImages += page.GetImages().Count(image =>
                    image.WidthInSamples >= minWidth && image.HeightInSamples >= minHeight);
            }

            var content = text.ToString();
            var person = ExtractPerson(content);
            var personFromFile = false;
            if (person.Length == 0)
            {
                person = PersonFromFileName(pdfPath);
                personFromFile = person.Length > 0;
            }
            var scenario = ExtractLabeledValue(content,
                "(?:Название\\s+)?сценар(?:ий|ия)", "Тема", "Упражнение", "Тренинг");
            var mode = ExtractLabeledValue(content, "Режим(?:\\s+прохождения)?");
            var attemptedAt = ExtractDateFromText(content) ?? ExtractDateFromPath(sourceRoot, pdfPath);
            var missing = new List<string>();
            if (person.Length == 0) missing.Add("ФИО не найдено");
            if (scenario.Length == 0) missing.Add("сценарий не найден");
            if (!attemptedAt.HasValue) missing.Add("дата не найдена");
            if (personFromFile) missing.Add("ФИО взято из имени файла");
            return new PdfPhotoEntry(
                installation,
                person,
                attemptedAt,
                scenario,
                mode,
                matchingImages > 0,
                matchingImages,
                document.NumberOfPages,
                reportFolder,
                Path.GetFileName(pdfPath),
                Path.GetFullPath(pdfPath),
                missing.Count == 0 ? "Распознано" : string.Join("; ", missing));
        }
        catch (Exception exception)
        {
            return new PdfPhotoEntry(
                installation, "", ExtractDateFromPath(sourceRoot, pdfPath), "", "", false, 0, 0,
                reportFolder, Path.GetFileName(pdfPath), Path.GetFullPath(pdfPath),
                "Ошибка чтения: " + FirstLine(exception.Message));
        }
    }

    private static IEnumerable<string> EnumeratePdfFiles(string root, ConcurrentBag<string> notes)
    {
        var pending = new Stack<string>();
        pending.Push(Path.GetFullPath(root));
        while (pending.Count > 0)
        {
            var directory = pending.Pop();
            string[] files;
            try
            {
                files = Directory.GetFiles(directory, "*.pdf", SearchOption.TopDirectoryOnly);
            }
            catch (Exception exception) when (exception is UnauthorizedAccessException or IOException)
            {
                notes.Add($"Не прочитана папка: {directory}. {FirstLine(exception.Message)}");
                continue;
            }
            foreach (var file in files)
                yield return file;

            string[] directories;
            try
            {
                directories = Directory.GetDirectories(directory);
            }
            catch (Exception exception) when (exception is UnauthorizedAccessException or IOException)
            {
                notes.Add($"Не прочитаны подпапки: {directory}. {FirstLine(exception.Message)}");
                continue;
            }
            foreach (var child in directories)
            {
                try
                {
                    if ((File.GetAttributes(child) & FileAttributes.ReparsePoint) == 0)
                        pending.Push(child);
                }
                catch (Exception exception) when (exception is UnauthorizedAccessException or IOException)
                {
                    notes.Add($"Пропущена подпапка: {child}. {FirstLine(exception.Message)}");
                }
            }
        }
    }

    private static void WriteWorkbook(
        PdfPhotoReportRequest request,
        IReadOnlyList<PdfPhotoEntry> entries,
        IReadOnlyList<string> enumerationNotes)
    {
        var output = Path.GetFullPath(request.OutputFile);
        Directory.CreateDirectory(Path.GetDirectoryName(output)
            ?? throw new InvalidOperationException("Не удалось определить папку результата."));
        try
        {
            using var workbook = new XLWorkbook();
            workbook.Properties.Title = "Контроль фотографий в PDF-отчётах КТК";
            workbook.Properties.Author = "Анализатор КТК";
            WriteSummary(workbook, request, entries, enumerationNotes);
            WriteDetails(workbook, request, entries);
            workbook.CalculateMode = XLCalculateMode.Auto;
            workbook.SaveAs(output);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new AnalysisFileAccessException(output, exception);
        }
    }

    private static void WriteSummary(
        XLWorkbook workbook,
        PdfPhotoReportRequest request,
        IReadOnlyList<PdfPhotoEntry> entries,
        IReadOnlyList<string> enumerationNotes)
    {
        var sheet = workbook.AddWorksheet("Сводка");
        sheet.ShowGridLines = false;
        sheet.Cell("A2").Value = "КОНТРОЛЬ ФОТОГРАФИЙ В PDF";
        sheet.Cell("A2").Style.Font.FontName = FontName;
        sheet.Cell("A2").Style.Font.FontSize = 16;
        sheet.Cell("A2").Style.Font.Bold = true;
        sheet.Cell("A2").Style.Font.FontColor = Navy;
        sheet.Range("A2:H2").Merge();
        sheet.Range("A3:H3").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        sheet.Range("A3:H3").Style.Border.BottomBorderColor = Blue;
        sheet.Cell("A4").Value = "Папка";
        sheet.Cell("B4").Value = Path.GetFullPath(request.SourceFolder);
        sheet.Range("B4:H4").Merge();
        sheet.Cell("A5").Value = "Порог изображения";
        sheet.Cell("B5").Value = $"не менее {request.MinimumPhotoWidth} × {request.MinimumPhotoHeight} px";

        string[] cardLabels = ["Всего PDF", "С фотографией", "Без фотографии", "Ошибки чтения"];
        int[] cardValues =
        [
            entries.Count,
            entries.Count(item => item.HasPhoto),
            entries.Count(item => !item.HasPhoto && !IsError(item)),
            entries.Count(IsError)
        ];
        for (var index = 0; index < cardLabels.Length; index++)
        {
            var column = index * 2 + 1;
            sheet.Cell(7, column).Value = cardLabels[index];
            sheet.Cell(8, column).Value = cardValues[index];
            sheet.Range(7, column, 7, column + 1).Merge();
            sheet.Range(8, column, 8, column + 1).Merge();
            var card = sheet.Range(7, column, 8, column + 1);
            card.Style.Fill.BackgroundColor = index switch
            {
                1 => LightGreen,
                2 => LightRed,
                3 => LightAmber,
                _ => LightBlue
            };
            card.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            card.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            card.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            card.Style.Border.OutsideBorderColor = XLColor.FromHtml("#D4DFEE");
            sheet.Cell(7, column).Style.Font.Bold = true;
            sheet.Cell(7, column).Style.Font.FontColor = Navy;
            sheet.Cell(8, column).Style.Font.Bold = true;
            sheet.Cell(8, column).Style.Font.FontSize = 16;
            sheet.Cell(8, column).Style.Font.FontColor = Navy;
        }

        string[] headers = ["Установка", "PDF", "С фото", "Без фото", "Ошибки", "Доля с фото"];
        for (var column = 1; column <= headers.Length; column++)
            sheet.Cell(11, column).Value = headers[column - 1];
        var groups = entries.GroupBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.CurrentCultureIgnoreCase).ToArray();
        for (var index = 0; index < groups.Length; index++)
        {
            var group = groups[index].ToArray();
            var row = index + 12;
            sheet.Cell(row, 1).Value = groups[index].Key;
            sheet.Cell(row, 2).Value = group.Length;
            sheet.Cell(row, 3).Value = group.Count(item => item.HasPhoto);
            sheet.Cell(row, 4).Value = group.Count(item => !item.HasPhoto && !IsError(item));
            sheet.Cell(row, 5).Value = group.Count(IsError);
            sheet.Cell(row, 6).Value = group.Length == 0 ? 0d : group.Count(item => item.HasPhoto) / (double)group.Length;
        }
        var lastRow = 11 + groups.Length;
        var table = sheet.Range(11, 1, lastRow, headers.Length).CreateTable("PdfPhotoSummary");
        table.Theme = XLTableTheme.TableStyleMedium2;
        sheet.Range(12, 6, lastRow, 6).Style.NumberFormat.Format = "0.0%";
        StyleHeader(sheet.Range(11, 1, 11, headers.Length));
        sheet.SheetView.FreezeRows(11);
        SetWidths(sheet, [28, 13, 13, 15, 12, 16, 14, 14]);
        if (enumerationNotes.Count > 0)
        {
            sheet.Cell("H11").Value = "Примечания";
            sheet.Cell("H11").Style.Font.Bold = true;
            sheet.Cell("H11").Style.Font.FontColor = Navy;
            for (var index = 0; index < enumerationNotes.Count; index++)
                sheet.Cell(index + 12, 8).Value = enumerationNotes[index];
            sheet.Column(8).Width = 72;
            sheet.Range(12, 8, 11 + enumerationNotes.Count, 8).Style.Alignment.WrapText = true;
        }
        sheet.RangeUsed()!.Style.Font.FontName = FontName;
    }

    private static void WriteDetails(
        XLWorkbook workbook,
        PdfPhotoReportRequest request,
        IReadOnlyList<PdfPhotoEntry> entries)
    {
        var sheet = workbook.AddWorksheet("Отчёты PDF");
        sheet.ShowGridLines = false;
        sheet.Cell("A2").Value = "ОТЧЁТЫ PDF";
        sheet.Cell("A2").Style.Font.FontName = FontName;
        sheet.Cell("A2").Style.Font.FontSize = 16;
        sheet.Cell("A2").Style.Font.Bold = true;
        sheet.Cell("A2").Style.Font.FontColor = Navy;
        sheet.Cell("A3").Value = "Фотография — растровое изображение размером не менее " +
                                 $"{request.MinimumPhotoWidth} × {request.MinimumPhotoHeight} px";
        sheet.Cell("A3").Style.Font.Italic = true;
        sheet.Cell("A3").Style.Font.FontColor = XLColor.FromHtml("#66758A");
        string[] headers =
        [
            "№", "Установка", "ФИО", "Дата", "Сценарий", "Режим", "Фото в PDF",
            "Изображений", "Страниц", "Папка отчёта", "Имя файла", "Ссылка", "Примечание"
        ];
        for (var column = 1; column <= headers.Length; column++)
            sheet.Cell(5, column).Value = headers[column - 1];
        for (var index = 0; index < entries.Count; index++)
        {
            var item = entries[index];
            var row = index + 6;
            sheet.Cell(row, 1).Value = index + 1;
            sheet.Cell(row, 2).Value = item.Installation;
            sheet.Cell(row, 3).Value = item.PersonName;
            if (item.AttemptedAt.HasValue)
            {
                sheet.Cell(row, 4).Value = item.AttemptedAt.Value;
                sheet.Cell(row, 4).Style.DateFormat.Format = item.AttemptedAt.Value.TimeOfDay == TimeSpan.Zero
                    ? "dd.MM.yyyy"
                    : "dd.MM.yyyy HH:mm:ss";
            }
            sheet.Cell(row, 5).Value = item.Scenario;
            sheet.Cell(row, 6).Value = item.Mode;
            sheet.Cell(row, 7).Value = IsError(item) ? "Ошибка" : item.HasPhoto ? "Есть" : "Нет";
            sheet.Cell(row, 8).Value = item.MatchingImages;
            sheet.Cell(row, 9).Value = item.Pages;
            sheet.Cell(row, 10).Value = item.ReportFolder;
            sheet.Cell(row, 11).Value = item.FileName;
            sheet.Cell(row, 12).Value = "Открыть PDF";
            sheet.Cell(row, 12).SetHyperlink(new XLHyperlink(item.PdfPath));
            sheet.Cell(row, 13).Value = item.RecognitionNote;
            sheet.Cell(row, 7).Style.Fill.BackgroundColor = IsError(item)
                ? LightAmber
                : item.HasPhoto ? LightGreen : LightRed;
            sheet.Cell(row, 7).Style.Font.Bold = true;
        }
        var lastRow = entries.Count + 5;
        var table = sheet.Range(5, 1, lastRow, headers.Length).CreateTable("PdfPhotoDetails");
        table.Theme = XLTableTheme.TableStyleMedium2;
        StyleHeader(sheet.Range(5, 1, 5, headers.Length));
        sheet.Range(6, 5, lastRow, 6).Style.Alignment.WrapText = true;
        sheet.Range(6, 13, lastRow, 13).Style.Alignment.WrapText = true;
        sheet.RangeUsed()!.Style.Font.FontName = FontName;
        sheet.SheetView.FreezeRows(5);
        sheet.SheetView.FreezeColumns(3);
        SetWidths(sheet, [7, 25, 34, 20, 52, 22, 15, 14, 11, 25, 42, 17, 42]);
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
    }

    private static string InferInstallation(string root, string file)
    {
        var relativeDirectory = Path.GetDirectoryName(Path.GetRelativePath(root, file)) ?? "";
        var parts = relativeDirectory.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .Where(part => part.Length > 0).ToArray();
        var datedIndex = Array.FindIndex(parts, LooksLikeReportFolder);
        if (datedIndex > 0)
            return parts[0];
        if (datedIndex == 0 || parts.Length == 0)
            return Path.GetFileName(Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar));
        return parts[0];
    }

    private static string InferReportFolder(string root, string file)
    {
        var directory = new DirectoryInfo(Path.GetDirectoryName(file)!);
        var rootFull = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar);
        while (directory is not null && directory.FullName.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase))
        {
            if (LooksLikeReportFolder(directory.Name))
                return directory.Name;
            if (directory.FullName.Equals(rootFull, StringComparison.OrdinalIgnoreCase))
                break;
            directory = directory.Parent;
        }
        return Path.GetFileName(Path.GetDirectoryName(file)) ?? "";
    }

    private static DateTime? ExtractDateFromPath(string root, string file)
    {
        var relative = Path.GetRelativePath(root, file);
        foreach (var part in relative.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Reverse())
            if (TryParseDate(part, out var value))
                return value;
        return null;
    }

    private static DateTime? ExtractDateFromText(string text)
    {
        var match = DateLabelRegex().Match(text);
        if (match.Success && TryParseDate(match.Groups["value"].Value, out var value))
            return value;
        return null;
    }

    private static bool TryParseDate(string value, out DateTime result)
    {
        var match = DateRegex().Match(value);
        if (match.Success)
        {
            var normalized = $"{match.Groups["y"].Value}-{match.Groups["m"].Value}-{match.Groups["d"].Value}";
            if (match.Groups["h"].Success)
                normalized += $" {match.Groups["h"].Value}:{match.Groups["min"].Value}:{match.Groups["s"].Value}";
            string[] formats = ["yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss"];
            if (DateTime.TryParseExact(normalized, formats, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out result))
                return true;
        }
        return DateTime.TryParse(value, RussianCulture, DateTimeStyles.AllowWhiteSpaces, out result);
    }

    private static string ExtractPerson(string text)
    {
        var direct = ExtractLabeledValue(text,
            "Ф\\.?\\s*И\\.?\\s*О\\.?(?:\\s+по\\s+отч[её]ту)?",
            "Студент", "Сотрудник", "Обучаемый");
        if (LooksLikePerson(direct))
            return direct;
        var lastName = ExtractLabeledValue(text, "Фамилия");
        var firstName = ExtractLabeledValue(text, "Имя");
        var middleName = ExtractLabeledValue(text, "Отчество");
        return NormalizeValue(string.Join(" ", new[] { lastName, firstName, middleName }
            .Where(value => value.Length > 0)));
    }

    private static string ExtractLabeledValue(string text, params string[] labels)
    {
        var pattern = $@"(?im)^\s*(?:{string.Join("|", labels)})\s*[:\-]?\s*(?<value>[^\r\n]{{1,180}})\s*$";
        var match = Regex.Match(text, pattern, RegexOptions.CultureInvariant);
        if (!match.Success)
            return "";
        var value = NormalizeValue(match.Groups["value"].Value);
        var doubleSpace = Regex.Match(value, @"\s{3,}");
        return doubleSpace.Success ? value[..doubleSpace.Index].Trim() : value;
    }

    private static string PersonFromFileName(string path)
    {
        var value = Path.GetFileNameWithoutExtension(path);
        value = Regex.Replace(value, "протокол", "", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        value = Regex.Replace(value, @"^\d{4}[-_]\d{2}[-_]\d{2}[ _-]*", "");
        value = Regex.Replace(value, @"[()]", " ").Replace('_', ' ');
        value = NormalizeValue(value);
        return LooksLikePerson(value) ? value : "";
    }

    private static bool LooksLikePerson(string value)
    {
        if (value.Length is < 5 or > 120 || value.Any(char.IsDigit))
            return false;
        var parts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length is >= 2 and <= 4 && parts.All(part => part.Any(char.IsLetter));
    }

    private static string NormalizeValue(string value) =>
        Regex.Replace(value.Replace('\u00A0', ' '), @"\s+", " ").Trim(' ', ':', '-', '—', ';');

    private static bool LooksLikeReportFolder(string name) => DateRegex().IsMatch(name);
    private static bool IsError(PdfPhotoEntry item) => item.RecognitionNote.StartsWith("Ошибка чтения", StringComparison.Ordinal);
    private static string FirstLine(string value) => value.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? value;

    private static void Validate(PdfPhotoReportRequest request)
    {
        if (!Directory.Exists(request.SourceFolder))
            throw new DirectoryNotFoundException($"Папка PDF-отчётов не найдена:\n{request.SourceFolder}");
        if (request.PeriodStart.HasValue && request.PeriodEnd.HasValue && request.PeriodEnd.Value.Date < request.PeriodStart.Value.Date)
            throw new ArgumentException("Дата окончания периода не может быть раньше даты начала.");
        if (request.MinimumPhotoWidth < 1 || request.MinimumPhotoHeight < 1)
            throw new ArgumentException("Минимальный размер фотографии должен быть положительным.");
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

    private static void SetWidths(IXLWorksheet sheet, IReadOnlyList<double> widths)
    {
        for (var index = 0; index < widths.Count; index++)
            sheet.Column(index + 1).Width = widths[index];
    }

    [GeneratedRegex(@"(?<y>20\d{2})[_.-](?<m>\d{2})[_.-](?<d>\d{2})(?:[ T_]+(?<h>\d{2})[-_:](?<min>\d{2})[-_:](?<s>\d{2}))?", RegexOptions.CultureInvariant)]
    private static partial Regex DateRegex();

    [GeneratedRegex(@"(?im)(?:Дата(?:\s+(?:прохождения|отч[её]та|тренинга))?)\s*[:\-]?\s*(?<value>(?:\d{2}\.\d{2}\.\d{4}|20\d{2}[_.-]\d{2}[_.-]\d{2})(?:\s+\d{2}[:_-]\d{2}(?:[:_-]\d{2})?)?)", RegexOptions.CultureInvariant)]
    private static partial Regex DateLabelRegex();
}

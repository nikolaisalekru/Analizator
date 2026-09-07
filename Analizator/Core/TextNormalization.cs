using System.Globalization;
using System.Text.RegularExpressions;

namespace Analizator.Core;

public static partial class TextNormalization
{
    private static readonly HashSet<string> AvailableModes = new(StringComparer.OrdinalIgnoreCase)
    {
        "Случайный выбор", "Ежемесячный тренинг", "Все сценарии", "Экзамен"
    };

    public static string NormalizeText(object? value)
    {
        if (value is null)
            return "";
        var text = value.ToString()!.Trim().ToLowerInvariant().Replace('ё', 'е');
        text = InvalidCharacters().Replace(text, " ");
        return Whitespace().Replace(text, " ").Trim();
    }

    public static string NormalizePersonName(object? value) => NormalizeText(value);
    public static string NormalizeHeader(object? value) => NormalizeText(value);
    public static string NormalizeMode(object? value) => NormalizeText(value);

    public static bool ModeIsAllowed(object? value, ModeFilterSettings settings)
    {
        if (!settings.Enabled)
            return true;
        var normalized = NormalizeMode(value);
        return normalized.Length > 0 && settings.Modes.Any(mode => NormalizeMode(mode) == normalized);
    }

    public static string? KnownMode(object? value)
    {
        var normalized = NormalizeMode(value);
        return AvailableModes.FirstOrDefault(mode => NormalizeMode(mode) == normalized);
    }

    public static double? SafeDouble(object? value)
    {
        if (value is null)
            return null;
        if (value is double number)
            return number;
        if (value is float single)
            return single;
        if (value is decimal decimalValue)
            return (double)decimalValue;
        if (value is int integer)
            return integer;
        var text = value.ToString()?.Trim().Replace("%", "").Replace(',', '.');
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) ? result : null;
    }

    public static DateTime? ParseDate(object? value)
    {
        if (value is DateTime date)
            return date.Date;
        if (value is double serial && serial is > 0 and < 2958466)
            return DateTime.FromOADate(serial).Date;
        var text = value?.ToString()?.Trim();
        if (string.IsNullOrEmpty(text))
            return null;
        string[] formats =
        [
            "dd.MM.yyyy", "dd.MM.yyyy H:mm", "dd.MM.yyyy H:mm:ss",
            "yyyy-MM-dd", "yyyy-MM-dd H:mm:ss", "M/d/yyyy", "M/d/yyyy H:mm:ss"
        ];
        return DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out date) ? date.Date : null;
    }

    public static int? MonthFromFileName(string path)
    {
        var name = NormalizeText(Path.GetFileNameWithoutExtension(path));
        var names = MonthNames;
        foreach (var pair in names)
            if (Regex.IsMatch(name, $@"(^|\s){Regex.Escape(NormalizeText(pair.Value))}($|\s)"))
                return pair.Key;
        return null;
    }

    public static IReadOnlyDictionary<int, string> MonthNames { get; } = new Dictionary<int, string>
    {
        [1] = "Январь", [2] = "Февраль", [3] = "Март", [4] = "Апрель",
        [5] = "Май", [6] = "Июнь", [7] = "Июль", [8] = "Август",
        [9] = "Сентябрь", [10] = "Октябрь", [11] = "Ноябрь", [12] = "Декабрь"
    };

    [GeneratedRegex(@"[^a-zа-я0-9]+", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex InvalidCharacters();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Whitespace();
}

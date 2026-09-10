using System.Globalization;
using System.Text.Json;

namespace Analizator.Core;

public static class JsonTrainingExportParser
{
    public const int Version = 3;

    public static ParsedTrainingExport Parse(
        string path,
        IReadOnlyDictionary<string, List<string>> installations,
        CancellationToken cancellationToken = default,
        Action<int, string>? progress = null)
    {
        progress?.Invoke(5, "Чтение JSON");
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var document = JsonDocument.Parse(stream, new JsonDocumentOptions
        {
            AllowTrailingCommas = true,
            CommentHandling = JsonCommentHandling.Skip
        });

        var sessions = document.RootElement.ValueKind switch
        {
            JsonValueKind.Array => document.RootElement.EnumerateArray().ToArray(),
            JsonValueKind.Object => [document.RootElement],
            _ => throw new InvalidDataException("Корневой элемент JSON должен быть массивом или объектом.")
        };
        var parsed = new ParsedTrainingExport { Format = TrainingExportFormat.Json };
        var sourceRow = 0;
        var unmatchedInstallations = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

        for (var sessionIndex = 0; sessionIndex < sessions.Length; sessionIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var session = sessions[sessionIndex];
            if (session.ValueKind != JsonValueKind.Object)
            {
                parsed.Errors.Add(new ExportImportError("JSON", sessionIndex + 1,
                    "Элемент сеанса должен быть объектом."));
                continue;
            }

            var sessionGuid = ReadString(session, "Guid");
            var machineName = ReadString(session, "MachineName");
            var sessionPlantName = ReadString(session, "PlantName");
            var studentSerial = ReadString(session, "StudentSerial");
            var sessionStudentName = ReadString(session, "StudentName");
            var entryDate = ReadDateTime(session, "EntryDate");
            var sessionStart = ReadDateTime(session, "StartDate");
            var sessionEnd = ReadDateTime(session, "EndDate");
            var isClosed = ReadBoolean(session, "IsClosed");
            var trainingTimesJson = ReadRawJson(session, "TrainingTimes");
            var exerciseLogsJson = ReadRawJson(session, "ExerciseLogs");
            var trainingModes = ReadTrainingModes(session);

            if (!TryGetProperty(session, "ExerciseReports", out var reports) ||
                reports.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
                continue;
            if (reports.ValueKind != JsonValueKind.Array)
            {
                parsed.Errors.Add(new ExportImportError("JSON", sessionIndex + 1,
                    "Поле ExerciseReports должно быть массивом."));
                continue;
            }

            foreach (var report in reports.EnumerateArray())
            {
                cancellationToken.ThrowIfCancellationRequested();
                sourceRow++;
                if (report.ValueKind != JsonValueKind.Object)
                {
                    parsed.Errors.Add(new ExportImportError("JSON", sourceRow,
                        "Элемент ExerciseReports должен быть объектом."));
                    continue;
                }

                var rawPlantName = ReadString(report, "PlantName") ?? sessionPlantName;
                var installation = MatchingService.MatchInstallation(rawPlantName, installations).Name;
                if (installation is null && !string.IsNullOrWhiteSpace(rawPlantName))
                {
                    if (unmatchedInstallations.Add(rawPlantName.Trim()))
                    {
                        parsed.Errors.Add(new ExportImportError(rawPlantName, sourceRow,
                            "Установка не найдена в настройках. Записи этой установки пропущены; " +
                            "добавьте PlantName как синоним нужной установки."));
                    }
                    continue;
                }
                var personName = ReadString(report, "StudentName") ?? sessionStudentName;
                var scenarioName = ReadString(report, "ExerciseName");
                var attemptedAt = ReadDateTime(report, "Date");
                var percent = ReadDouble(report, "Percent");
                if (string.IsNullOrWhiteSpace(installation) ||
                    string.IsNullOrWhiteSpace(personName) ||
                    string.IsNullOrWhiteSpace(scenarioName) ||
                    !attemptedAt.HasValue || !percent.HasValue || percent.Value is < 0 or > 100)
                {
                    parsed.Errors.Add(new ExportImportError(rawPlantName ?? "JSON", sourceRow,
                        "Не заполнены установка, ФИО, сценарий, дата или корректный процент."));
                    continue;
                }

                var role = ReadString(report, "Role");
                var result = ReadBoolean(report, "Result");
                var mode = ResolveMode(trainingModes, attemptedAt.Value, scenarioName);
                var metadata = new JsonAttemptMetadata(
                    sessionGuid,
                    machineName,
                    rawPlantName,
                    studentSerial,
                    entryDate,
                    sessionStart,
                    sessionEnd,
                    isClosed,
                    ReadString(report, "FilePath"),
                    ReadInt32(report, "Resume"),
                    result,
                    ReadBoolean(report, "IsValid"),
                    ReadInt32(report, "ReportType"),
                    ReadBoolean(report, "HasReportFile"),
                    ReadRawValue(report, "ReportFile"),
                    trainingTimesJson,
                    exerciseLogsJson,
                    ReadArrayCount(session, "TrainingTimes"),
                    ReadArrayCount(session, "ExerciseLogs"));
                parsed.Attempts.Add(new TrainingAttemptRecord(
                    installation,
                    personName,
                    scenarioName,
                    attemptedAt.Value,
                    percent.Value,
                    mode,
                    null,
                    role,
                    result.HasValue ? result.Value ? "Да" : "Нет" : null,
                    ReadString(report, "TimeSpent"),
                    ReadInt32(report, "Steps"),
                    rawPlantName ?? "JSON",
                    sourceRow,
                    metadata));
            }

            var percentComplete = sessions.Length == 0
                ? 90
                : 10 + (int)Math.Round((sessionIndex + 1) * 80d / sessions.Length);
            progress?.Invoke(percentComplete, "Чтение отчётов JSON");
        }

        progress?.Invoke(100, "Чтение JSON завершено");
        return parsed;
    }

    public static string? ResolveMode(
        string? trainingTimesJson,
        DateTime attemptedAt,
        string? scenarioName = null)
    {
        if (string.IsNullOrWhiteSpace(trainingTimesJson))
            return null;

        try
        {
            using var document = JsonDocument.Parse(trainingTimesJson, new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip
            });
            return ResolveMode(ReadTrainingModes(document.RootElement), attemptedAt, scenarioName);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static string? ResolveMode(
        IReadOnlyList<TrainingModeEntry> entries,
        DateTime attemptedAt,
        string? scenarioName)
    {
        if (entries.Count == 0)
            return null;

        var distinctModes = entries
            .Select(item => item.Mode)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (distinctModes.Length == 1)
            return distinctModes[0];

        var byTime = entries
            .Where(item => item.StartDate.HasValue && item.EndDate.HasValue &&
                           attemptedAt >= item.StartDate.Value.AddSeconds(-1) &&
                           attemptedAt <= item.EndDate.Value.AddSeconds(1))
            .ToArray();
        var resolved = UniqueMode(byTime, scenarioName);
        if (resolved is not null)
            return resolved;

        var normalizedScenario = TextNormalization.NormalizeText(scenarioName);
        if (normalizedScenario.Length == 0)
            return null;
        var byScenario = entries
            .Where(item => TextNormalization.NormalizeText(item.Exercises)
                .Contains(normalizedScenario, StringComparison.Ordinal))
            .ToArray();
        return UniqueMode(byScenario, scenarioName);
    }

    private static string? UniqueMode(
        IReadOnlyList<TrainingModeEntry> entries,
        string? scenarioName)
    {
        if (entries.Count == 0)
            return null;

        var normalizedScenario = TextNormalization.NormalizeText(scenarioName);
        if (normalizedScenario.Length > 0)
        {
            var scenarioMatches = entries
                .Where(item => TextNormalization.NormalizeText(item.Exercises)
                    .Contains(normalizedScenario, StringComparison.Ordinal))
                .Select(item => item.Mode)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (scenarioMatches.Length == 1)
                return scenarioMatches[0];
        }

        var modes = entries
            .Select(item => item.Mode)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return modes.Length == 1 ? modes[0] : null;
    }

    private static List<TrainingModeEntry> ReadTrainingModes(JsonElement container)
    {
        var value = container;
        if (container.ValueKind == JsonValueKind.Object)
        {
            if (!TryGetProperty(container, "TrainingTimes", out value))
                return [];
        }
        if (value.ValueKind != JsonValueKind.Array)
            return [];

        var result = new List<TrainingModeEntry>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object || ReadMode(item) is not { } mode)
                continue;
            result.Add(new TrainingModeEntry(
                mode,
                ReadDateTime(item, "StartDate"),
                ReadDateTime(item, "EndDate"),
                ReadString(item, "Exercises")));
        }
        return result;
    }

    private static string? ReadMode(JsonElement trainingTime)
    {
        if (!TryGetProperty(trainingTime, "Mode", out var value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var numericMode))
            return ModeName(numericMode);

        var text = NullIfWhiteSpace(value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : value.ToString());
        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out numericMode))
            return ModeName(numericMode);
        return TextNormalization.KnownMode(text) ?? text;
    }

    private static string ModeName(int mode) => mode switch
    {
        0 => "Все сценарии",
        1 => "Ежемесячный тренинг",
        2 => "Случайный выбор",
        _ => $"Код режима {mode}"
    };

    private static bool TryGetProperty(JsonElement element, string name, out JsonElement value)
    {
        if (element.TryGetProperty(name, out value))
            return true;
        foreach (var property in element.EnumerateObject())
        {
            if (!string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                continue;
            value = property.Value;
            return true;
        }
        value = default;
        return false;
    }

    private static string? ReadString(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        return value.ValueKind == JsonValueKind.String
            ? NullIfWhiteSpace(value.GetString())
            : NullIfWhiteSpace(value.ToString());
    }

    private static string? ReadRawValue(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        return value.ValueKind == JsonValueKind.String ? value.GetString() : value.GetRawText();
    }

    private static string? ReadRawJson(JsonElement element, string name) =>
        TryGetProperty(element, name, out var value) &&
        value.ValueKind is not JsonValueKind.Null and not JsonValueKind.Undefined
            ? value.GetRawText()
            : null;

    private static DateTime? ReadDateTime(JsonElement element, string name)
    {
        var text = ReadString(element, name);
        if (string.IsNullOrWhiteSpace(text))
            return null;
        return DateTime.TryParse(text, CultureInfo.InvariantCulture,
                   DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind, out var invariant)
            ? invariant
            : DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"),
                DateTimeStyles.AllowWhiteSpaces, out var russian)
                ? russian
                : null;
    }

    private static double? ReadDouble(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value))
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number))
            return number;
        var text = ReadString(element, name);
        return double.TryParse(text?.Replace(',', '.'), NumberStyles.Float,
            CultureInfo.InvariantCulture, out number) ? number : null;
    }

    private static int? ReadInt32(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value))
            return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
            return number;
        return int.TryParse(ReadString(element, name), NumberStyles.Integer,
            CultureInfo.InvariantCulture, out number) ? number : null;
    }

    private static bool? ReadBoolean(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value))
            return null;
        if (value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return value.GetBoolean();
        return bool.TryParse(ReadString(element, name), out var result) ? result : null;
    }

    private static int? ReadArrayCount(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value) || value.ValueKind != JsonValueKind.Array)
            return null;
        return value.GetArrayLength();
    }

    private static string? NullIfWhiteSpace(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private sealed record TrainingModeEntry(
        string Mode,
        DateTime? StartDate,
        DateTime? EndDate,
        string? Exercises);
}

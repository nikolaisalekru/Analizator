using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Analizator.Core;

public sealed class ConfigurationBundle
{
    public required Dictionary<string, List<string>> Installations { get; init; }
    public required AnalyzerSettings Settings { get; init; }
    public required Dictionary<string, Dictionary<string, string>> FioKeys { get; init; }
    public required HashSet<string> ExcludedPeople { get; init; }
}

public static class ConfigurationService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public static ConfigurationBundle Load(string directory)
    {
        Directory.CreateDirectory(directory);
        var installationsPath = Path.Combine(directory, "installations.json");
        var settingsPath = Path.Combine(directory, "settings.json");
        var keysPath = Path.Combine(directory, "fio_keys.json");
        var excludedPath = Path.Combine(directory, "excluded_people.json");

        EnsureJson(installationsPath, DefaultInstallations());
        EnsureJson(settingsPath, DefaultSettingsDocument());
        EnsureJson(keysPath, new Dictionary<string, object>());
        EnsureJson(excludedPath, new { people = Array.Empty<string>() });

        return new ConfigurationBundle
        {
            Installations = LoadInstallations(installationsPath),
            Settings = LoadSettings(settingsPath),
            FioKeys = LoadFioKeys(keysPath),
            ExcludedPeople = LoadExcludedPeople(excludedPath)
        };
    }

    public static void SavePassingThresholds(string directory, PassingThresholdSettings settings)
    {
        ValidateThreshold(settings.GlobalPercent, "Общий порог");
        foreach (var item in settings.Installations)
            ValidateThreshold(item.Value, $"Порог для установки «{item.Key}»");

        Directory.CreateDirectory(directory);
        var settingsPath = Path.Combine(directory, "settings.json");
        EnsureJson(settingsPath, DefaultSettingsDocument());
        var root = JsonNode.Parse(File.ReadAllText(settingsPath)) as JsonObject
                   ?? throw new InvalidDataException("settings.json должен содержать JSON-объект.");
        var installationValues = new JsonObject();
        foreach (var item in settings.Installations.OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase))
            installationValues[item.Key] = item.Value;
        root["passing_threshold"] = new JsonObject
        {
            ["use_per_installation"] = settings.UsePerInstallation,
            ["global"] = settings.GlobalPercent,
            ["installations"] = installationValues
        };
        File.WriteAllText(settingsPath, root.ToJsonString(JsonOptions));
    }

    private static void EnsureJson<T>(string path, T value)
    {
        if (!File.Exists(path))
            File.WriteAllText(path, JsonSerializer.Serialize(value, JsonOptions));
    }

    private static JsonDocument ReadDocument(string path)
    {
        try
        {
            return JsonDocument.Parse(File.ReadAllText(path));
        }
        catch (JsonException exception)
        {
            throw new InvalidDataException(
                $"Ошибка в {Path.GetFileName(path)} (строка {exception.LineNumber + 1}, " +
                $"позиция {exception.BytePositionInLine + 1}): {exception.Message}", exception);
        }
    }

    private static Dictionary<string, List<string>> LoadInstallations(string path)
    {
        using var document = ReadDocument(path);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("installations.json должен содержать объект с установками.");
        if (root.TryGetProperty("installations", out var wrapped))
            root = wrapped;
        if (root.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("Поле installations должно содержать объект.");

        var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in root.EnumerateObject())
        {
            var canonical = property.Name.Trim();
            if (canonical.Length == 0)
                continue;
            var variants = new List<string> { canonical };
            if (property.Value.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in property.Value.EnumerateArray())
                    if (item.ValueKind == JsonValueKind.String)
                        AddUnique(variants, item.GetString());
            }
            else if (property.Value.ValueKind == JsonValueKind.String)
                AddUnique(variants, property.Value.GetString());
            result[canonical] = variants;
        }
        return result;
    }

    private static AnalyzerSettings LoadSettings(string path)
    {
        using var document = ReadDocument(path);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("settings.json должен содержать JSON-объект.");

        var modeFilter = new ModeFilterSettings();
        if (root.TryGetProperty("mode_filter", out var mode) && mode.ValueKind == JsonValueKind.Object)
        {
            modeFilter = new ModeFilterSettings
            {
                // Compatibility with the legacy loader: it retained the
                // configured list, but did not expose mode_filter.enabled to the pipeline.
                Enabled = false,
                Modes = []
            };
        }

        var autoAdd = new AutoAddSettings();
        if (root.TryGetProperty("auto_add_new_people", out var auto) && auto.ValueKind == JsonValueKind.Object)
        {
            var installations = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            if (auto.TryGetProperty("installations", out var map) && map.ValueKind == JsonValueKind.Object)
                foreach (var property in map.EnumerateObject())
                    installations[property.Name.Trim()] = property.Value.ValueKind == JsonValueKind.True;
            autoAdd = new AutoAddSettings
            {
                Default = !auto.TryGetProperty("default", out var defaultValue) || defaultValue.ValueKind != JsonValueKind.False,
                Installations = installations
            };
        }

        var passingThreshold = new PassingThresholdSettings();
        if (root.TryGetProperty("passing_threshold", out var threshold) &&
            threshold.ValueKind == JsonValueKind.Object)
        {
            var global = PassingThresholdSettings.DefaultPercent;
            if (threshold.TryGetProperty("global", out var globalValue))
                global = ReadThreshold(globalValue, "Общий порог прохождения");

            var installationValues = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            if (threshold.TryGetProperty("installations", out var map) && map.ValueKind == JsonValueKind.Object)
                foreach (var property in map.EnumerateObject())
                {
                    var installation = property.Name.Trim();
                    if (installation.Length > 0)
                        installationValues[installation] = ReadThreshold(
                            property.Value, $"Порог для установки «{installation}»");
                }

            passingThreshold = new PassingThresholdSettings
            {
                UsePerInstallation = threshold.TryGetProperty("use_per_installation", out var perInstallation) &&
                                     perInstallation.ValueKind == JsonValueKind.True,
                GlobalPercent = global,
                Installations = installationValues
            };
        }

        var parallel = !root.TryGetProperty("parallel_processing", out var parallelValue) ||
                       parallelValue.ValueKind != JsonValueKind.False;
        var workers = root.TryGetProperty("max_workers", out var workersValue) && workersValue.TryGetInt32(out var parsed)
            ? Math.Max(1, parsed)
            : 5;
        return new AnalyzerSettings
        {
            ModeFilter = modeFilter,
            AutoAddNewPeople = autoAdd,
            PassingThreshold = passingThreshold,
            ParallelProcessing = parallel,
            MaxWorkers = workers
        };
    }

    private static double ReadThreshold(JsonElement value, string fieldName)
    {
        if (!value.TryGetDouble(out var threshold))
            throw new InvalidDataException($"{fieldName} должен быть числом от 0 до 100.");
        ValidateThreshold(threshold, fieldName);
        return threshold;
    }

    private static void ValidateThreshold(double threshold, string fieldName)
    {
        if (double.IsNaN(threshold) || double.IsInfinity(threshold) || threshold is < 0 or > 100)
            throw new InvalidDataException($"{fieldName} должен быть числом от 0 до 100.");
    }

    private static Dictionary<string, Dictionary<string, string>> LoadFioKeys(string path)
    {
        using var document = ReadDocument(path);
        if (document.RootElement.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("fio_keys.json должен содержать JSON-объект.");
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var section in document.RootElement.EnumerateObject())
        {
            if (section.Value.ValueKind != JsonValueKind.Object)
                continue;
            var installation = TextNormalization.NormalizeText(section.Name);
            if (installation.Length == 0)
                continue;
            var keys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in section.Value.EnumerateObject())
            {
                if (item.Value.ValueKind != JsonValueKind.String)
                    continue;
                var key = TextNormalization.NormalizePersonName(item.Name);
                var target = item.Value.GetString()?.Trim();
                if (key.Length > 0 && !string.IsNullOrWhiteSpace(target))
                    keys.TryAdd(key, target);
            }
            result[installation] = keys;
        }
        return result;
    }

    private static HashSet<string> LoadExcludedPeople(string path)
    {
        using var document = ReadDocument(path);
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (document.RootElement.ValueKind != JsonValueKind.Object ||
            !document.RootElement.TryGetProperty("people", out var people) || people.ValueKind != JsonValueKind.Array)
            return result;
        foreach (var item in people.EnumerateArray())
            if (item.ValueKind == JsonValueKind.String)
            {
                var name = TextNormalization.NormalizePersonName(item.GetString());
                if (name.Length > 0)
                    result.Add(name);
            }
        return result;
    }

    private static void AddUnique(List<string> values, string? value)
    {
        value = value?.Trim();
        if (!string.IsNullOrEmpty(value) && !values.Contains(value, StringComparer.OrdinalIgnoreCase))
            values.Add(value);
    }

    private static object DefaultSettingsDocument() => new
    {
        mode_filter = new
        {
            enabled = true,
            modes = new[] { "Случайный выбор", "Все сценарии", "Ежемесячный тренинг" }
        },
        auto_add_new_people = new { @default = true, installations = new Dictionary<string, bool>() },
        passing_threshold = new
        {
            use_per_installation = false,
            global = PassingThresholdSettings.DefaultPercent,
            installations = new Dictionary<string, double>()
        },
        parallel_processing = true,
        max_workers = 5
    };

    private static Dictionary<string, string[]> DefaultInstallations() => new(StringComparer.OrdinalIgnoreCase)
    {
        ["ЭЛОУ-АВТ-6"] = ["ЭЛОУ АВТ-6", "ЭЛОУ-АВТ 6", "ЭЛОУАВТ-6", "АВТ-6", "АВТ"],
        ["ГФУ-2"] = ["ГФУ 2", "ГФУ2", "ГФУ"],
        ["ВБ"] = ["АТ-ВБ", "Висбрекинг"],
        ["УПБ"] = ["Битум", "Битумка"],
        ["ЛЧ-35/11-1000"] = ["ЛЧ 35/11-1000", "35 11 1000", "3511", "миллионка"],
        ["Л-24_5"] = ["Л245", "Л-24 5", "Л-24/5", "Л-24-5"],
        ["УИЛН+КЦА"] = ["УИЛН КЦА", "УИЛН"],
        ["ЛЧ-24-2000"] = ["ЛЧ 24-2000", "ЛЧ 24 2000"],
        ["МТБЭ"] = ["MTBE"],
        ["ГОБКК"] = ["Гобка", "Кат. крекинг", "Каталитический крекинг"],
        ["УПС"] = ["Сера"],
        ["УПВ"] = ["Водородка", "Водород"],
        ["ФУ и УСУГиП"] = ["ФУ УСУГиП", "ФУиУСУГиП"],
        ["Г-43-107"] = ["Г 43 107", "Г-43/107", "Кат крекинг машинисты"],
        ["КУПН"] = [],
        ["УПА"] = ["УПА с ХЖА", "Азот", "Азотка"],
        ["Котельные"] = ["Котельная", "Котлы"],
        ["ЦВК"] = [], ["ХВП"] = [], ["ГРС"] = [],
        ["АСБиКТ"] = ["АСБИКТ", "АСБ и КТ"],
        ["СНМиВГ"] = ["СНМИВГ", "СНМ и ВГ"],
        ["БРиДТ"] = ["БРИДТ", "БР и ДТ"],
        ["РХиПНО"] = ["РХИПНО", "РХ и ПНО"],
        ["ЭНСН"] = [], ["КОС"] = [],
        ["БОВ-8а"] = ["БОВ-8А", "БОВ 8а", "БОВ 8А"],
        ["TАМЕ"] = ["TAME", "ТАМЕ", "ТАМЭ", "КТД-200_720", "200720"],
        ["РНК ЦУП"] = ["РНК-ЦУП", "РНКЦУП", "ЦУП", "ИТР ЦУП", "ИТР"]
    };
}

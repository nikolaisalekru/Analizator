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

    public static EditableConfigurationData LoadEditable(string directory)
    {
        var configuration = Load(directory);
        return new EditableConfigurationData
        {
            Installations = LoadEditableInstallations(Path.Combine(directory, "installations.json")),
            FioKeys = LoadEditableFioKeys(Path.Combine(directory, "fio_keys.json")),
            ExcludedPeople = LoadEditableExcludedPeople(Path.Combine(directory, "excluded_people.json")),
            Settings = configuration.Settings
        };
    }

    public static void SaveInstallations(
        string directory,
        IReadOnlyDictionary<string, IReadOnlyCollection<string>> installations)
    {
        if (installations.Count == 0)
            throw new ArgumentException("Должна остаться хотя бы одна установка.");

        var root = new JsonObject();
        foreach (var item in installations.OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase))
        {
            var canonical = item.Key.Trim();
            if (canonical.Length == 0)
                throw new ArgumentException("Название установки не может быть пустым.");
            if (root.ContainsKey(canonical))
                throw new ArgumentException($"Установка «{canonical}» указана несколько раз.");

            var aliases = new List<string> { canonical };
            foreach (var alias in item.Value)
                AddUnique(aliases, alias);
            root[canonical] = new JsonArray(aliases.Select(value => JsonValue.Create(value)).ToArray());
        }

        WriteConfigurationFile(directory, "installations.json", root);
    }

    public static void SaveFioKeys(
        string directory,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> sections)
    {
        var root = new JsonObject();
        foreach (var section in sections.OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase))
        {
            var sectionName = section.Key.Trim();
            if (sectionName.Length == 0)
                throw new ArgumentException("Название раздела не может быть пустым.");

            var values = new JsonObject();
            foreach (var item in section.Value.OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase))
            {
                var abbreviation = item.Key.Trim();
                var fullName = item.Value.Trim();
                if (abbreviation.Length == 0 || fullName.Length == 0)
                    throw new ArgumentException("Сокращение и полное ФИО должны быть заполнены.");
                values[abbreviation] = fullName;
            }
            root[sectionName] = values;
        }

        WriteConfigurationFile(directory, "fio_keys.json", root);
    }

    public static void SaveExcludedPeople(string directory, IEnumerable<string> people)
    {
        var values = people
            .Select(value => value.Trim())
            .Where(value => value.Length > 0)
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase)
            .Select(value => JsonValue.Create(value))
            .ToArray();
        WriteConfigurationFile(directory, "excluded_people.json", new JsonObject
        {
            ["people"] = new JsonArray(values)
        });
    }

    public static void SaveAnalysisSettings(
        string directory,
        bool parallelProcessing,
        int maximumWorkers,
        bool autoAddNewPeople)
    {
        if (maximumWorkers is < 1 or > 64)
            throw new ArgumentException("Количество параллельных обработчиков должно быть от 1 до 64.");

        var root = ReadSettingsObject(directory);
        root["parallel_processing"] = parallelProcessing;
        root["max_workers"] = maximumWorkers;
        var autoAdd = root["auto_add_new_people"] as JsonObject ?? new JsonObject();
        autoAdd["default"] = autoAddNewPeople;
        autoAdd["installations"] ??= new JsonObject();
        root["auto_add_new_people"] = autoAdd;
        WriteSettingsObject(directory, root);
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

    public static void SaveDatabaseBackupSettings(string directory, DatabaseBackupSettings settings)
    {
        ValidateDatabaseBackupSettings(settings);
        var root = ReadSettingsObject(directory);
        root["database_backups"] = new JsonObject
        {
            ["automatic_enabled"] = settings.AutomaticBackupsEnabled,
            ["interval_days"] = settings.IntervalDays,
            ["maximum_files"] = settings.MaximumBackupFiles,
            ["before_changes"] = settings.BackupBeforeChanges,
            ["confirm_row_deletion"] = settings.ConfirmRowDeletion,
            ["confirm_source_deletion"] = settings.ConfirmSourceDeletion
        };
        WriteSettingsObject(directory, root);
    }

    public static void SaveDatabaseEditorSettings(string directory, DatabaseEditorSettings settings)
    {
        var root = ReadSettingsObject(directory);
        root["database_editor"] = new JsonObject
        {
            ["hidden_columns"] = new JsonArray(settings.HiddenColumns
                .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase)
                .Select(value => JsonValue.Create(value))
                .ToArray())
        };
        WriteSettingsObject(directory, root);
    }

    private static JsonObject ReadSettingsObject(string directory)
    {
        Directory.CreateDirectory(directory);
        var settingsPath = Path.Combine(directory, "settings.json");
        EnsureJson(settingsPath, DefaultSettingsDocument());
        return JsonNode.Parse(File.ReadAllText(settingsPath)) as JsonObject
               ?? throw new InvalidDataException("settings.json должен содержать JSON-объект.");
    }

    private static void WriteSettingsObject(string directory, JsonObject root) =>
        File.WriteAllText(Path.Combine(directory, "settings.json"), root.ToJsonString(JsonOptions));

    private static void WriteConfigurationFile(string directory, string fileName, JsonNode root)
    {
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, fileName), root.ToJsonString(JsonOptions));
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

    private static Dictionary<string, List<string>> LoadEditableInstallations(string path)
    {
        using var document = ReadDocument(path);
        var root = document.RootElement;
        if (root.TryGetProperty("installations", out var wrapped))
            root = wrapped;
        if (root.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("installations.json должен содержать объект с установками.");

        var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in root.EnumerateObject())
        {
            var canonical = property.Name.Trim();
            if (canonical.Length == 0)
                continue;
            var aliases = new List<string>();
            if (property.Value.ValueKind == JsonValueKind.Array)
                foreach (var item in property.Value.EnumerateArray())
                    if (item.ValueKind == JsonValueKind.String &&
                        !string.Equals(item.GetString()?.Trim(), canonical, StringComparison.OrdinalIgnoreCase))
                        AddUnique(aliases, item.GetString());
            else if (property.Value.ValueKind == JsonValueKind.String &&
                     !string.Equals(property.Value.GetString()?.Trim(), canonical, StringComparison.OrdinalIgnoreCase))
                AddUnique(aliases, property.Value.GetString());
            result[canonical] = aliases;
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

        var databaseBackups = new DatabaseBackupSettings();
        if (root.TryGetProperty("database_backups", out var backups) &&
            backups.ValueKind == JsonValueKind.Object)
        {
            databaseBackups = new DatabaseBackupSettings
            {
                AutomaticBackupsEnabled = ReadBoolean(backups, "automatic_enabled", true),
                IntervalDays = ReadInteger(backups, "interval_days", 15, 1, 3650),
                MaximumBackupFiles = ReadInteger(backups, "maximum_files", 2, 1, 100),
                BackupBeforeChanges = ReadBoolean(backups, "before_changes", true),
                ConfirmRowDeletion = ReadBoolean(backups, "confirm_row_deletion", true),
                ConfirmSourceDeletion = ReadBoolean(backups, "confirm_source_deletion", true)
            };
        }

        var hiddenColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (root.TryGetProperty("database_editor", out var editor) &&
            editor.ValueKind == JsonValueKind.Object &&
            editor.TryGetProperty("hidden_columns", out var hidden) &&
            hidden.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in hidden.EnumerateArray())
                if (item.ValueKind == JsonValueKind.String &&
                    !string.IsNullOrWhiteSpace(item.GetString()))
                    hiddenColumns.Add(item.GetString()!.Trim());
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
            DatabaseBackups = databaseBackups,
            DatabaseEditor = new DatabaseEditorSettings { HiddenColumns = hiddenColumns },
            ParallelProcessing = parallel,
            MaxWorkers = workers
        };
    }

    private static bool ReadBoolean(JsonElement parent, string name, bool fallback) =>
        parent.TryGetProperty(name, out var value) &&
        value.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? value.GetBoolean()
            : fallback;

    private static int ReadInteger(
        JsonElement parent,
        string name,
        int fallback,
        int minimum,
        int maximum) =>
        parent.TryGetProperty(name, out var value) && value.TryGetInt32(out var parsed)
            ? Math.Clamp(parsed, minimum, maximum)
            : fallback;

    private static void ValidateDatabaseBackupSettings(DatabaseBackupSettings settings)
    {
        if (settings.IntervalDays is < 1 or > 3650)
            throw new ArgumentException("Интервал резервного копирования должен быть от 1 до 3650 дней.");
        if (settings.MaximumBackupFiles is < 1 or > 100)
            throw new ArgumentException("Количество резервных копий должно быть от 1 до 100.");
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

    private static Dictionary<string, Dictionary<string, string>> LoadEditableFioKeys(string path)
    {
        using var document = ReadDocument(path);
        if (document.RootElement.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("fio_keys.json должен содержать JSON-объект.");

        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var section in document.RootElement.EnumerateObject())
        {
            var sectionName = section.Name.Trim();
            if (sectionName.Length == 0 || section.Value.ValueKind != JsonValueKind.Object)
                continue;
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in section.Value.EnumerateObject())
                if (item.Value.ValueKind == JsonValueKind.String &&
                    !string.IsNullOrWhiteSpace(item.Name) &&
                    !string.IsNullOrWhiteSpace(item.Value.GetString()))
                    values[item.Name.Trim()] = item.Value.GetString()!.Trim();
            result[sectionName] = values;
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

    private static List<string> LoadEditableExcludedPeople(string path)
    {
        using var document = ReadDocument(path);
        var result = new List<string>();
        if (document.RootElement.ValueKind != JsonValueKind.Object ||
            !document.RootElement.TryGetProperty("people", out var people) ||
            people.ValueKind != JsonValueKind.Array)
            return result;
        foreach (var item in people.EnumerateArray())
            if (item.ValueKind == JsonValueKind.String)
                AddUnique(result, item.GetString());
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
        database_backups = new
        {
            automatic_enabled = true,
            interval_days = 15,
            maximum_files = 2,
            before_changes = true,
            confirm_row_deletion = true,
            confirm_source_deletion = true
        },
        database_editor = new { hidden_columns = Array.Empty<string>() },
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

public sealed class EditableConfigurationData
{
    public required Dictionary<string, List<string>> Installations { get; init; }
    public required Dictionary<string, Dictionary<string, string>> FioKeys { get; init; }
    public required List<string> ExcludedPeople { get; init; }
    public required AnalyzerSettings Settings { get; init; }
}

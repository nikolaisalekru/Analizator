using Microsoft.Data.Sqlite;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace Analizator.Core;

public sealed class TrainingDatabaseService
{
    private const int SchemaVersion = 6;
    private const string ManualSourceHash = "__ANALIZATOR_MANUAL_RECORDS__";
    private readonly string _connectionString;
    private IReadOnlyDictionary<string, Dictionary<string, string>> _fioKeys;
    private IReadOnlyDictionary<string, List<string>> _installations;

    public TrainingDatabaseService(
        string databasePath,
        IReadOnlyDictionary<string, Dictionary<string, string>>? fioKeys = null,
        IReadOnlyDictionary<string, List<string>>? installations = null)
    {
        if (string.IsNullOrWhiteSpace(databasePath))
            throw new ArgumentException("Не указан путь к базе данных.", nameof(databasePath));
        DatabasePath = Path.GetFullPath(databasePath);
        Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath)!);
        _connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = DatabasePath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Cache = SqliteCacheMode.Shared,
            Pooling = true
        }.ToString();
        _fioKeys = fioKeys ?? new Dictionary<string, Dictionary<string, string>>();
        _installations = installations ?? new Dictionary<string, List<string>>();
        Initialize();
        ReconcileConfiguredInstallations();
    }

    public string DatabasePath { get; }

    public void SetPersonNameMappings(
        IReadOnlyDictionary<string, Dictionary<string, string>> fioKeys,
        IReadOnlyDictionary<string, List<string>> installations)
    {
        ArgumentNullException.ThrowIfNull(fioKeys);
        ArgumentNullException.ThrowIfNull(installations);
        _fioKeys = fioKeys;
        _installations = installations;
        ReconcileConfiguredInstallations();
    }

    public DatabaseImportResult ImportFiles(
        IReadOnlyList<string> paths,
        IReadOnlyDictionary<string, List<string>> installations,
        IProgress<DatabaseImportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(installations);
        _installations = installations;
        ReconcileConfiguredInstallations();

        var result = new DatabaseImportResult();
        var files = paths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        for (var fileIndex = 0; fileIndex < files.Length; fileIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = files[fileIndex];
            var name = Path.GetFileName(path);
            var startPercent = files.Length == 0 ? 0 : fileIndex * 100 / files.Length;
            var endPercent = files.Length == 0 ? 100 : (fileIndex + 1) * 100 / files.Length;
            progress?.Report(new DatabaseImportProgress(startPercent, "Проверка файла", name));

            if (!File.Exists(path))
            {
                result.FailedFiles++;
                result.Files.Add(new DatabaseImportFileResult(name, "Ошибка", 0, 0, "Файл не найден."));
                continue;
            }

            if (!IsSupportedImportFile(path))
            {
                result.FailedFiles++;
                result.Files.Add(new DatabaseImportFileResult(name, "Ошибка", 0, 0,
                    "Поддерживаются книги Excel .xlsx/.xlsm и выгрузки JSON .json."));
                continue;
            }

            try
            {
                var hash = CalculateFileHash(path, cancellationToken);
                using var connection = OpenConnection();
                var existingSource = FindSourceFile(connection, hash);
                var parserVersion = ParserVersionFor(path);
                if (existingSource is not null &&
                    existingSource.ParserVersion >= parserVersion)
                {
                    result.SkippedFiles++;
                    result.Files.Add(new DatabaseImportFileResult(name, "Пропущен", 0, 0,
                        "Этот файл уже загружен в базу."));
                    progress?.Report(new DatabaseImportProgress(endPercent, "Файл уже был загружен", name));
                    continue;
                }

                void ReportFileProgress(int percent, string stage)
                {
                    var scaled = startPercent +
                                 (int)Math.Round((endPercent - startPercent) * percent / 100d);
                    progress?.Report(new DatabaseImportProgress(scaled, stage, name));
                }

                var parsed = Path.GetExtension(path).Equals(".json", StringComparison.OrdinalIgnoreCase)
                    ? JsonTrainingExportParser.Parse(path, installations, cancellationToken, ReportFileProgress)
                    : TrainingExportParser.Parse(path, installations, cancellationToken, ReportFileProgress);
                if (parsed.Attempts.Count == 0)
                {
                    result.FailedFiles++;
                    result.RowErrors += parsed.Errors.Count;
                    result.Files.Add(new DatabaseImportFileResult(name, "Ошибка", 0, parsed.Errors.Count,
                        ImportErrorSummary(parsed.Errors, "В файле нет подходящих записей для импорта.")));
                    continue;
                }

                progress?.Report(new DatabaseImportProgress(
                    Math.Min(endPercent, startPercent + (endPercent - startPercent) * 70 / 100),
                    "Сохранение записей в базу", name));
                var possibleDuplicates = SaveParsedFile(
                    connection, path, hash, parsed, existingSource?.Id, cancellationToken);
                result.ImportedFiles++;
                if (existingSource is not null)
                    result.UpdatedFiles++;
                result.ImportedAttempts += parsed.Attempts.Count;
                result.PossibleDuplicates += possibleDuplicates;
                result.RowErrors += parsed.Errors.Count;
                result.Files.Add(new DatabaseImportFileResult(
                    name, existingSource is null ? "Загружен" : "Обновлён", parsed.Attempts.Count,
                    parsed.Errors.Count,
                    parsed.Errors.Count == 0
                        ? "Все строки прочитаны."
                        : $"Замечаний при чтении: {parsed.Errors.Count}. " +
                          ImportErrorSummary(parsed.Errors, "Проверьте структуру файла.")));
                progress?.Report(new DatabaseImportProgress(endPercent,
                    existingSource is null ? "Файл загружен" : "Файл обновлён", name));
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                result.FailedFiles++;
                result.Files.Add(new DatabaseImportFileResult(name, "Ошибка", 0, 0,
                    FriendlyImportError(exception)));
            }
        }

        progress?.Report(new DatabaseImportProgress(100, "Импорт завершён"));
        return result;
    }

    public DatabaseSummary GetSummary()
    {
        using var connection = OpenConnection();
        var configuredNames = GetConfiguredInstallationNames();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                COUNT(*),
                COUNT(DISTINCT a.person_id),
                COUNT(DISTINCT a.source_file_id),
                COUNT(DISTINCT a.installation_id),
                COALESCE(SUM(a.possible_duplicate), 0),
                MIN(a.attempted_at),
                MAX(a.attempted_at)
            FROM attempts a
            JOIN installations i ON i.id = a.installation_id
            WHERE CONFIGURED_INSTALLATIONS;
            """.Replace("CONFIGURED_INSTALLATIONS",
                ConfiguredInstallationPredicate("i", configuredNames.Count), StringComparison.Ordinal);
        AddConfiguredInstallationParameters(command, configuredNames);
        using var reader = command.ExecuteReader();
        reader.Read();
        var summary = new DatabaseSummary(
            reader.GetInt32(0),
            reader.GetInt32(1),
            reader.GetInt32(2),
            reader.GetInt32(3),
            reader.GetInt32(4),
            ReadNullableDateTime(reader, 5),
            ReadNullableDateTime(reader, 6));
        reader.Close();
        var catalog = ReadPersonNameCatalog(connection);
        return summary with { People = catalog.CanonicalPeopleCount };
    }

    public IReadOnlyList<DatabaseSourceItem> GetImportHistory()
    {
        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT id, file_name, original_path, format, imported_at,
                   period_start, period_end, rows_count, error_count
            FROM source_files
            ORDER BY imported_at DESC, id DESC;
            """;
        using var reader = command.ExecuteReader();
        var result = new List<DatabaseSourceItem>();
        while (reader.Read())
        {
            result.Add(new DatabaseSourceItem
            {
                Id = reader.GetInt64(0),
                FileName = reader.GetString(1),
                OriginalPath = reader.GetString(2),
                Format = reader.GetString(3),
                ImportedAt = ParseStoredDateTime(reader.GetString(4)),
                PeriodStart = ReadNullableDateTime(reader, 5),
                PeriodEnd = ReadNullableDateTime(reader, 6),
                RowsCount = reader.GetInt32(7),
                ErrorCount = reader.GetInt32(8)
            });
        }
        return result;
    }

    public DatabaseDuplicateSummary GetDuplicateSummary()
    {
        using var connection = OpenConnection();
        var configuredNames = GetConfiguredInstallationNames();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                COALESCE(SUM(CASE WHEN duplicate_kind = 'Точный' THEN 1 ELSE 0 END), 0),
                COALESCE(SUM(CASE WHEN duplicate_kind = 'Возможный' THEN 1 ELSE 0 END), 0)
            FROM attempts a
            JOIN installations i ON i.id = a.installation_id
            WHERE CONFIGURED_INSTALLATIONS;
            """.Replace("CONFIGURED_INSTALLATIONS",
                ConfiguredInstallationPredicate("i", configuredNames.Count), StringComparison.Ordinal);
        AddConfiguredInstallationParameters(command, configuredNames);
        using var reader = command.ExecuteReader();
        reader.Read();
        return new DatabaseDuplicateSummary(reader.GetInt32(0), reader.GetInt32(1));
    }

    public IReadOnlyList<DatabaseAttempt> GetAttempts(DateTime periodStart, DateTime periodEnd)
    {
        if (periodEnd.Date < periodStart.Date)
            throw new ArgumentException("Дата окончания периода не может быть раньше даты начала.");

        using var connection = OpenConnection();
        var catalog = ReadPersonNameCatalog(connection);
        var configuredNames = GetConfiguredInstallationNames();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT a.id, i.name, p.display_name, s.display_name, a.attempted_at,
                   a.percent, a.mode, a.project, a.role, a.credited, a.duration,
                   a.actions, f.file_name, a.source_sheet, a.source_row,
                   a.possible_duplicate
            FROM attempts a
            JOIN installations i ON i.id = a.installation_id
            JOIN people p ON p.id = a.person_id
            JOIN scenarios s ON s.id = a.scenario_id
            JOIN source_files f ON f.id = a.source_file_id
            WHERE a.attempted_at >= $start AND a.attempted_at < $end
              AND a.duplicate_of_attempt_id IS NULL
              AND CONFIGURED_INSTALLATIONS
            ORDER BY a.attempted_at, i.name, p.display_name, a.id;
            """.Replace("CONFIGURED_INSTALLATIONS",
                ConfiguredInstallationPredicate("i", configuredNames.Count), StringComparison.Ordinal);
        command.Parameters.AddWithValue("$start", StoreDateTime(periodStart.Date));
        command.Parameters.AddWithValue("$end", StoreDateTime(periodEnd.Date.AddDays(1)));
        AddConfiguredInstallationParameters(command, configuredNames);
        using var reader = command.ExecuteReader();
        var result = new List<DatabaseAttempt>();
        while (reader.Read())
        {
            var attempt = new DatabaseAttempt(
                reader.GetInt64(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetString(3),
                ParseStoredDateTime(reader.GetString(4)),
                reader.GetDouble(5),
                ReadNullableString(reader, 6),
                ReadNullableString(reader, 7),
                ReadNullableString(reader, 8),
                ReadNullableString(reader, 9),
                ReadNullableString(reader, 10),
                reader.IsDBNull(11) ? null : reader.GetInt32(11),
                reader.GetString(12),
                reader.GetString(13),
                reader.GetInt32(14),
                reader.GetInt32(15) != 0);
            result.Add(attempt with
            {
                CanonicalPersonName = catalog.Resolve(attempt.Installation, attempt.PersonName)
            });
        }
        return result;
    }

    public IReadOnlyList<DatabasePersonSummary> GetPeopleForExport(
        DateTime periodStart,
        DateTime periodEnd)
    {
        if (periodEnd.Date < periodStart.Date)
            throw new ArgumentException("Дата окончания периода не может быть раньше даты начала.");

        using var connection = OpenConnection();
        var catalog = ReadPersonNameCatalog(connection);
        var configuredNames = GetConfiguredInstallationNames();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT i.name, p.display_name, COUNT(*)
            FROM attempts a
            JOIN installations i ON i.id = a.installation_id
            JOIN people p ON p.id = a.person_id
            WHERE a.attempted_at >= $start AND a.attempted_at < $end
              AND a.duplicate_of_attempt_id IS NULL
              AND CONFIGURED_INSTALLATIONS
            GROUP BY i.id, i.name, p.id, p.display_name
            ORDER BY i.name COLLATE NOCASE, p.display_name COLLATE NOCASE;
            """.Replace("CONFIGURED_INSTALLATIONS",
                ConfiguredInstallationPredicate("i", configuredNames.Count), StringComparison.Ordinal);
        command.Parameters.AddWithValue("$start", StoreDateTime(periodStart.Date));
        command.Parameters.AddWithValue("$end", StoreDateTime(periodEnd.Date.AddDays(1)));
        AddConfiguredInstallationParameters(command, configuredNames);
        using var reader = command.ExecuteReader();
        var rawResult = new List<DatabasePersonSummary>();
        while (reader.Read())
            rawResult.Add(new DatabasePersonSummary(
                reader.GetString(0),
                reader.GetString(1),
                reader.GetInt32(2)));
        return rawResult
            .Select(item => item with
            {
                PersonName = catalog.Resolve(item.Installation, item.PersonName)
            })
            .GroupBy(item => new DatabasePersonKey(item.Installation, item.PersonName).NormalizedValue,
                StringComparer.Ordinal)
            .Select(group => new DatabasePersonSummary(
                group.First().Installation,
                group.First().PersonName,
                group.Sum(item => item.Attempts)))
            .OrderBy(item => item.Installation, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(item => item.PersonName, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
    }

    public IReadOnlyList<DatabaseMonthSummary> GetAvailableMonths()
    {
        using var connection = OpenConnection();
        var configuredNames = GetConfiguredInstallationNames();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT CAST(substr(a.attempt_date, 1, 4) AS INTEGER),
                   CAST(substr(a.attempt_date, 6, 2) AS INTEGER),
                   COALESCE(SUM(CASE WHEN a.duplicate_of_attempt_id IS NULL THEN 1 ELSE 0 END), 0),
                   COALESCE(SUM(CASE WHEN a.duplicate_kind = 'Точный' THEN 1 ELSE 0 END), 0),
                   COALESCE(SUM(CASE WHEN a.duplicate_kind = 'Возможный' THEN 1 ELSE 0 END), 0),
                   COALESCE(SUM(CASE WHEN a.duplicate_of_attempt_id IS NULL AND
                                          (a.mode IS NULL OR trim(a.mode) = '') THEN 1 ELSE 0 END), 0)
            FROM attempts a
            JOIN installations i ON i.id = a.installation_id
            WHERE a.attempt_date IS NOT NULL
              AND CONFIGURED_INSTALLATIONS
            GROUP BY substr(a.attempt_date, 1, 7)
            ORDER BY substr(a.attempt_date, 1, 7) DESC;
            """.Replace("CONFIGURED_INSTALLATIONS",
                ConfiguredInstallationPredicate("i", configuredNames.Count), StringComparison.Ordinal);
        AddConfiguredInstallationParameters(command, configuredNames);
        using var reader = command.ExecuteReader();
        var result = new List<DatabaseMonthSummary>();
        while (reader.Read())
            result.Add(new DatabaseMonthSummary(
                reader.GetInt32(0), reader.GetInt32(1), reader.GetInt32(2),
                reader.GetInt32(3), reader.GetInt32(4), reader.GetInt32(5)));
        return result;
    }

    public IReadOnlyList<DatabaseUnmatchedInstallation> GetUnmatchedInstallations()
    {
        if (_installations.Count == 0)
            return [];

        using var connection = OpenConnection();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT i.name, COUNT(a.id)
            FROM installations i
            LEFT JOIN attempts a ON a.installation_id = i.id
            GROUP BY i.id, i.name
            ORDER BY i.name COLLATE NOCASE;
            """;
        using var reader = command.ExecuteReader();
        var result = new List<DatabaseUnmatchedInstallation>();
        while (reader.Read())
        {
            var name = reader.GetString(0);
            if (MatchingService.MatchInstallation(name, _installations).Name is null)
                result.Add(new DatabaseUnmatchedInstallation(name, reader.GetInt32(1)));
        }
        return result;
    }

    public DatabaseAttemptPage SearchAttempts(
        DateTime? periodStart,
        DateTime? periodEnd,
        string? search,
        int limit = 250,
        int offset = 0,
        TimeSpan? timeStart = null,
        TimeSpan? timeEnd = null,
        bool? duplicatesOnly = null)
    {
        if (periodStart.HasValue && periodEnd.HasValue && periodEnd.Value.Date < periodStart.Value.Date)
            throw new ArgumentException("Дата окончания периода не может быть раньше даты начала.");
        if (timeStart.HasValue && timeEnd.HasValue && timeEnd.Value < timeStart.Value)
            throw new ArgumentException("Время окончания не может быть раньше времени начала.");
        limit = Math.Max(1, limit);
        offset = Math.Max(0, offset);
        var normalizedSearch = TextNormalization.NormalizeText(search);
        var normalizedPerson = TextNormalization.NormalizePersonName(search);
        var installationSearch = MatchingService.MatchInstallation(search, _installations).Name;
        var normalizedInstallationSearch = TextNormalization.NormalizeText(installationSearch ?? search);

        using var connection = OpenConnection();
        var catalog = ReadPersonNameCatalog(connection);
        var configuredNames = GetConfiguredInstallationNames();
        var matchingPersonIds = catalog.FindPersonIds(normalizedPerson);
        var canonicalPersonFilter = matchingPersonIds.Count == 0
            ? "0"
            : "p.id IN (" + string.Join(", ", matchingPersonIds.Select((_, index) => $"$personId{index}")) + ")";
        const string joins = """
            FROM attempts a
            JOIN installations i ON i.id = a.installation_id
            JOIN people p ON p.id = a.person_id
            JOIN scenarios s ON s.id = a.scenario_id
            JOIN source_files f ON f.id = a.source_file_id
            """;
        var filters = """
            WHERE ($start IS NULL OR a.attempt_date >= $start)
              AND ($end IS NULL OR a.attempt_date <= $end)
              AND ($timeStart IS NULL OR a.attempt_time >= $timeStart)
              AND ($timeEnd IS NULL OR a.attempt_time <= $timeEnd)
              AND ($duplicateFilter = 0
                   OR ($duplicateFilter = 1 AND a.possible_duplicate = 1)
                   OR ($duplicateFilter = 2 AND a.possible_duplicate = 0))
              AND CONFIGURED_INSTALLATIONS
              AND ($search = '' OR i.normalized_name LIKE '%' || $installationSearch || '%'
                   OR p.normalized_name LIKE '%' || $personSearch || '%'
                   OR CANONICAL_PERSON_FILTER
                   OR s.normalized_name LIKE '%' || $search || '%'
                   OR f.file_name LIKE '%' || $rawSearch || '%'
                   OR a.session_guid LIKE '%' || $rawSearch || '%'
                   OR a.machine_name LIKE '%' || $rawSearch || '%'
                   OR a.plant_name LIKE '%' || $rawSearch || '%'
                   OR a.student_serial LIKE '%' || $rawSearch || '%'
                   OR a.file_path LIKE '%' || $rawSearch || '%')
            """.Replace("CANONICAL_PERSON_FILTER", canonicalPersonFilter, StringComparison.Ordinal)
                .Replace("CONFIGURED_INSTALLATIONS",
                    ConfiguredInstallationPredicate("i", configuredNames.Count), StringComparison.Ordinal);

        using var countCommand = connection.CreateCommand();
        countCommand.CommandText = "SELECT COUNT(*)\n" + joins + "\n" + filters + ";";
        AddSearchParameters(countCommand, periodStart, periodEnd, search, normalizedSearch,
            normalizedPerson, normalizedInstallationSearch, timeStart, timeEnd, duplicatesOnly);
        AddPersonIdParameters(countCommand, matchingPersonIds);
        AddConfiguredInstallationParameters(countCommand, configuredNames);
        var totalCount = Convert.ToInt32(countCommand.ExecuteScalar(), CultureInfo.InvariantCulture);

        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT a.id, i.name, p.display_name, s.display_name, a.attempted_at,
                   a.percent, a.mode, a.project, a.role, a.credited, a.duration,
                   a.actions, f.file_name, a.source_sheet, a.source_row,
                   a.possible_duplicate, a.session_guid, a.machine_name, a.plant_name,
                   a.student_serial, a.entry_date, a.session_start_date, a.session_end_date,
                   a.is_closed, a.file_path, a.resume, a.result, a.is_valid, a.report_type,
                   a.has_report_file, a.report_file, a.training_times_json, a.exercise_logs_json,
                   a.training_times_count, a.exercise_logs_count, a.duplicate_of_attempt_id,
                   a.duplicate_kind, a.duplicate_reason
            """ + "\n" + joins + "\n" + filters + "\n" + """
            ORDER BY a.attempted_at DESC, a.id DESC
            LIMIT $limit OFFSET $offset;
            """;
        AddSearchParameters(command, periodStart, periodEnd, search, normalizedSearch,
            normalizedPerson, normalizedInstallationSearch, timeStart, timeEnd, duplicatesOnly);
        AddPersonIdParameters(command, matchingPersonIds);
        AddConfiguredInstallationParameters(command, configuredNames);
        command.Parameters.AddWithValue("$limit", limit);
        command.Parameters.AddWithValue("$offset", offset);
        using var reader = command.ExecuteReader();
        var items = new List<DatabaseAttempt>();
        while (reader.Read())
        {
            var attempt = ReadAttempt(reader);
            items.Add(attempt with
            {
                CanonicalPersonName = catalog.Resolve(attempt.Installation, attempt.PersonName)
            });
        }
        return new DatabaseAttemptPage(items, totalCount);
    }

    public long AddAttempt(DatabaseAttemptEditModel model)
    {
        ValidateEditModel(model);
        var installation = ResolveConfiguredInstallation(model.Installation);
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();
        var sourceId = GetOrCreateManualSource(connection, transaction);
        var record = ToTrainingAttempt(model, "Ручной ввод", NextManualSourceRow(connection, transaction, sourceId),
            installation);
        var installationId = GetOrCreateInstallation(connection, transaction, installation,
            TextNormalization.NormalizeText(installation));
        var personId = GetOrCreatePerson(connection, transaction, installationId, model.PersonName);
        var scenarioId = GetOrCreateScenario(connection, transaction, model.ScenarioName,
            TextNormalization.NormalizeText(model.ScenarioName));
        var fingerprint = CalculateRecordFingerprint(record);
        InsertAttempt(connection, transaction, sourceId, installationId, personId, scenarioId,
            record, fingerprint, false);
        var id = SelectId(connection, transaction, "SELECT last_insert_rowid();");
        RefreshSourceStatistics(connection, transaction, sourceId);
        RecalculateDuplicateFlags(connection, transaction);
        transaction.Commit();
        return id;
    }

    public void UpdateAttempt(DatabaseAttemptEditModel model)
    {
        ValidateEditModel(model);
        var installation = ResolveConfiguredInstallation(model.Installation);
        if (!model.Id.HasValue)
            throw new ArgumentException("Не выбрана запись для изменения.", nameof(model));

        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();
        using var sourceCommand = connection.CreateCommand();
        sourceCommand.Transaction = transaction;
        sourceCommand.CommandText = "SELECT source_file_id, source_sheet, source_row FROM attempts WHERE id = $id;";
        sourceCommand.Parameters.AddWithValue("$id", model.Id.Value);
        using var reader = sourceCommand.ExecuteReader();
        if (!reader.Read())
            throw new InvalidOperationException("Запись уже была удалена из базы.");
        var sourceId = reader.GetInt64(0);
        var sourceSheet = reader.GetString(1);
        var sourceRow = reader.GetInt32(2);
        reader.Close();

        var installationId = GetOrCreateInstallation(connection, transaction, installation,
            TextNormalization.NormalizeText(installation));
        var personId = GetOrCreatePerson(connection, transaction, installationId, model.PersonName);
        var scenarioId = GetOrCreateScenario(connection, transaction, model.ScenarioName,
            TextNormalization.NormalizeText(model.ScenarioName));
        var record = ToTrainingAttempt(model, sourceSheet, sourceRow, installation);
        Execute(connection, transaction, """
            UPDATE attempts
            SET installation_id = $installation, person_id = $person, scenario_id = $scenario,
                attempted_at = $date, attempt_date = $attemptDate, attempt_time = $attemptTime,
                percent = $percent, mode = $mode, project = $project,
                role = $role, credited = $credited, duration = $duration, actions = $actions,
                session_guid = $sessionGuid, machine_name = $machineName, plant_name = $plantName,
                student_serial = $studentSerial, entry_date = $entryDate,
                session_start_date = $sessionStart, session_end_date = $sessionEnd,
                is_closed = $isClosed, file_path = $filePath, resume = $resume,
                result = $result, is_valid = $isValid, report_type = $reportType,
                has_report_file = $hasReportFile, report_file = $reportFile,
                training_times_json = $trainingTimesJson, exercise_logs_json = $exerciseLogsJson,
                training_times_count = $trainingTimesCount, exercise_logs_count = $exerciseLogsCount,
                record_fingerprint = $fingerprint
            WHERE id = $id;
            """,
            ("$installation", installationId), ("$person", personId), ("$scenario", scenarioId),
            ("$date", StoreDateTime(model.AttemptedAt)), ("$percent", model.Percent),
            ("$attemptDate", StoreDate(model.AttemptedAt)),
            ("$attemptTime", StoreTime(model.AttemptedAt.TimeOfDay)),
            ("$mode", NullIfWhiteSpace(model.Mode)), ("$project", NullIfWhiteSpace(model.Project)),
            ("$role", NullIfWhiteSpace(model.Role)), ("$credited", NullIfWhiteSpace(model.Credited)),
            ("$duration", NullIfWhiteSpace(model.Duration)), ("$actions", model.Actions),
            ("$sessionGuid", NullIfWhiteSpace(model.SessionGuid)),
            ("$machineName", NullIfWhiteSpace(model.MachineName)),
            ("$plantName", NullIfWhiteSpace(model.PlantName)),
            ("$studentSerial", NullIfWhiteSpace(model.StudentSerial)),
            ("$entryDate", DbValue(model.EntryDate)),
            ("$sessionStart", DbValue(model.SessionStartDate)),
            ("$sessionEnd", DbValue(model.SessionEndDate)),
            ("$isClosed", DbBoolean(model.IsClosed)),
            ("$filePath", NullIfWhiteSpace(model.FilePath)), ("$resume", model.Resume),
            ("$result", DbBoolean(model.Result)), ("$isValid", DbBoolean(model.IsValid)),
            ("$reportType", model.ReportType), ("$hasReportFile", DbBoolean(model.HasReportFile)),
            ("$reportFile", NullIfWhiteSpace(model.ReportFile)),
            ("$trainingTimesJson", NullIfWhiteSpace(model.TrainingTimesJson)),
            ("$exerciseLogsJson", NullIfWhiteSpace(model.ExerciseLogsJson)),
            ("$trainingTimesCount", model.TrainingTimesCount),
            ("$exerciseLogsCount", model.ExerciseLogsCount),
            ("$fingerprint", CalculateRecordFingerprint(record)), ("$id", model.Id.Value));
        RefreshSourceStatistics(connection, transaction, sourceId);
        CleanOrphans(connection, transaction);
        RecalculateDuplicateFlags(connection, transaction);
        transaction.Commit();
    }

    public void DeleteAttempt(long attemptId)
    {
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = "SELECT source_file_id FROM attempts WHERE id = $id;";
        command.Parameters.AddWithValue("$id", attemptId);
        var value = command.ExecuteScalar();
        if (value is null)
            return;
        var sourceId = Convert.ToInt64(value, CultureInfo.InvariantCulture);
        Execute(connection, transaction, "DELETE FROM attempts WHERE id = $id;", ("$id", attemptId));
        RefreshSourceStatistics(connection, transaction, sourceId);
        Execute(connection, transaction, """
            DELETE FROM source_files
            WHERE id = $source AND format = 'Manual'
              AND NOT EXISTS(SELECT 1 FROM attempts WHERE source_file_id = $source);
            """, ("$source", sourceId));
        CleanOrphans(connection, transaction);
        RecalculateDuplicateFlags(connection, transaction);
        transaction.Commit();
    }

    public ExportReadResult ReadForAnalysis(
        DateTime periodStart,
        DateTime periodEnd,
        ModeFilterSettings modeFilter)
    {
        var result = new ExportReadResult();
        var installationCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var attempt in GetAttempts(periodStart, periodEnd))
        {
            var normalizedMode = TextNormalization.NormalizeMode(attempt.Mode);
            if (modeFilter.Enabled && !TextNormalization.ModeIsAllowed(attempt.Mode, modeFilter))
            {
                result.FilteredReports++;
                if (normalizedMode.Length == 0)
                    result.ReportsWithoutMode++;
                continue;
            }

            if (normalizedMode.Length > 0 && TextNormalization.KnownMode(attempt.Mode) is { } knownMode)
            {
                result.ModeCounts.TryGetValue(knownMode, out var current);
                result.ModeCounts[knownMode] = current + 1;
            }

            result.TotalReports++;
            installationCounts.TryGetValue(attempt.Installation, out var count);
            installationCounts[attempt.Installation] = count + 1;
            if (!result.Data.TryGetValue(attempt.Installation, out var people))
                result.Data[attempt.Installation] = people = new(StringComparer.OrdinalIgnoreCase);
            if (!people.TryGetValue(attempt.DisplayPersonName, out var scenarios))
                people[attempt.DisplayPersonName] = scenarios = new(StringComparer.OrdinalIgnoreCase);
            var scenarioKey = TextNormalization.NormalizeText(attempt.ScenarioName);
            if (!scenarios.TryGetValue(scenarioKey, out var aggregate))
            {
                scenarios[scenarioKey] = new ScenarioAggregate
                {
                    DisplayName = attempt.ScenarioName,
                    BestPercent = attempt.Percent,
                    Attempts = 1
                };
            }
            else
            {
                aggregate.Attempts++;
                aggregate.BestPercent = Math.Max(aggregate.BestPercent, attempt.Percent);
            }
        }

        foreach (var item in installationCounts.OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase))
            result.InstallationStatistics.Add(new InstallationStatistic("База данных", item.Key, item.Value, "OK"));
        result.Month = periodStart.Month == periodEnd.Month && periodStart.Year == periodEnd.Year
            ? periodStart.Month
            : null;
        return result;
    }

    public void DeleteSourceFile(long sourceFileId)
    {
        using var connection = OpenConnection();
        using var transaction = connection.BeginTransaction();
        Execute(connection, transaction, "DELETE FROM source_files WHERE id = $id;", ("$id", sourceFileId));
        CleanOrphans(connection, transaction);
        RecalculateDuplicateFlags(connection, transaction);
        transaction.Commit();
    }

    public string CreateBackup(string backupDirectory, int maximumBackupFiles = int.MaxValue)
    {
        if (maximumBackupFiles < 1)
            throw new ArgumentOutOfRangeException(nameof(maximumBackupFiles),
                "Необходимо хранить хотя бы одну резервную копию.");
        Directory.CreateDirectory(backupDirectory);
        var backupPath = Path.Combine(backupDirectory,
            $"analizator_{DateTime.Now:yyyy-MM-dd_HH-mm-ss-fff}.db");
        using var source = OpenConnection();
        using var destination = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = backupPath,
            Mode = SqliteOpenMode.ReadWriteCreate
        }.ToString());
        destination.Open();
        source.BackupDatabase(destination);
        DeleteOldBackups(backupDirectory, maximumBackupFiles);
        return backupPath;
    }

    private static void DeleteOldBackups(string backupDirectory, int maximumBackupFiles)
    {
        if (maximumBackupFiles == int.MaxValue)
            return;
        var backups = new DirectoryInfo(backupDirectory)
            .EnumerateFiles("analizator_*.db", SearchOption.TopDirectoryOnly)
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .ThenByDescending(file => file.Name, StringComparer.OrdinalIgnoreCase)
            .Skip(maximumBackupFiles)
            .ToArray();
        foreach (var backup in backups)
            backup.Delete();
    }

    private void Initialize()
    {
        using var connection = OpenConnection();
        var currentVersion = ReadUserVersion(connection);
        Execute(connection, null, "PRAGMA journal_mode=WAL;");
        Execute(connection, null, "PRAGMA synchronous=NORMAL;");
        Execute(connection, null, """
            CREATE TABLE IF NOT EXISTS source_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                original_path TEXT NOT NULL,
                file_hash TEXT NOT NULL UNIQUE,
                format TEXT NOT NULL,
                imported_at TEXT NOT NULL,
                period_start TEXT,
                period_end TEXT,
                rows_count INTEGER NOT NULL DEFAULT 0,
                error_count INTEGER NOT NULL DEFAULT 0,
                parser_version INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS installations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                normalized_name TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS people (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                installation_id INTEGER NOT NULL REFERENCES installations(id),
                display_name TEXT NOT NULL,
                normalized_name TEXT NOT NULL,
                UNIQUE(installation_id, normalized_name)
            );

            CREATE TABLE IF NOT EXISTS scenarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                display_name TEXT NOT NULL,
                normalized_name TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS attempts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file_id INTEGER NOT NULL REFERENCES source_files(id) ON DELETE CASCADE,
                installation_id INTEGER NOT NULL REFERENCES installations(id),
                person_id INTEGER NOT NULL REFERENCES people(id),
                scenario_id INTEGER NOT NULL REFERENCES scenarios(id),
                attempted_at TEXT NOT NULL,
                attempt_date TEXT,
                attempt_time TEXT,
                percent REAL NOT NULL CHECK(percent >= 0 AND percent <= 100),
                mode TEXT,
                project TEXT,
                role TEXT,
                credited TEXT,
                duration TEXT,
                actions INTEGER,
                session_guid TEXT,
                machine_name TEXT,
                plant_name TEXT,
                student_serial TEXT,
                entry_date TEXT,
                session_start_date TEXT,
                session_end_date TEXT,
                is_closed INTEGER,
                file_path TEXT,
                resume INTEGER,
                result INTEGER,
                is_valid INTEGER,
                report_type INTEGER,
                has_report_file INTEGER,
                report_file TEXT,
                training_times_json TEXT,
                exercise_logs_json TEXT,
                training_times_count INTEGER,
                exercise_logs_count INTEGER,
                source_sheet TEXT NOT NULL,
                source_row INTEGER NOT NULL,
                record_fingerprint TEXT NOT NULL,
                possible_duplicate INTEGER NOT NULL DEFAULT 0 CHECK(possible_duplicate IN (0, 1)),
                duplicate_of_attempt_id INTEGER,
                duplicate_kind TEXT,
                duplicate_reason TEXT,
                UNIQUE(source_file_id, source_sheet, source_row)
            );

            CREATE TABLE IF NOT EXISTS import_errors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file_id INTEGER NOT NULL REFERENCES source_files(id) ON DELETE CASCADE,
                source_sheet TEXT NOT NULL,
                source_row INTEGER,
                message TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS ix_attempts_date ON attempts(attempted_at);
            CREATE INDEX IF NOT EXISTS ix_attempts_person_date ON attempts(person_id, attempted_at);
            CREATE INDEX IF NOT EXISTS ix_attempts_installation_date ON attempts(installation_id, attempted_at);
            CREATE INDEX IF NOT EXISTS ix_attempts_fingerprint ON attempts(record_fingerprint);
            CREATE INDEX IF NOT EXISTS ix_source_files_imported_at ON source_files(imported_at);
            """);

        if (!ColumnExists(connection, "source_files", "parser_version"))
            Execute(connection, null,
                "ALTER TABLE source_files ADD COLUMN parser_version INTEGER NOT NULL DEFAULT 1;");
        EnsureAttemptMetadataColumns(connection);
        Execute(connection, null, """
            UPDATE attempts
            SET attempt_date = substr(attempted_at, 1, 10),
                attempt_time = substr(attempted_at, 12, 8)
            WHERE attempt_date IS NULL OR attempt_time IS NULL;
            """);
        Execute(connection, null, """
            CREATE INDEX IF NOT EXISTS ix_attempts_date_time ON attempts(attempt_date, attempt_time);
            CREATE INDEX IF NOT EXISTS ix_attempts_duplicate_match
            ON attempts(installation_id, person_id, scenario_id, attempt_date, attempt_time, percent);
            """);
        if (currentVersion < 6)
            BackfillJsonModes(connection);
        if (currentVersion < SchemaVersion)
            RecalculateDuplicateFlags(connection, null);
        Execute(connection, null, $"PRAGMA user_version={SchemaVersion};");
    }

    private void ReconcileConfiguredInstallations()
    {
        if (_installations.Count == 0)
            return;

        using var connection = OpenConnection();
        var storedInstallations = new List<(long Id, string Name, string NormalizedName)>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT id, name, normalized_name FROM installations ORDER BY id;";
            using var reader = command.ExecuteReader();
            while (reader.Read())
                storedInstallations.Add((reader.GetInt64(0), reader.GetString(1), reader.GetString(2)));
        }

        if (storedInstallations.Count == 0)
            return;

        using var transaction = connection.BeginTransaction();
        var changed = false;
        foreach (var stored in storedInstallations)
        {
            var canonicalName = MatchingService.MatchInstallation(stored.Name, _installations).Name;
            if (canonicalName is null)
                continue;

            var canonicalNormalizedName = TextNormalization.NormalizeText(canonicalName);
            if (string.Equals(stored.NormalizedName, canonicalNormalizedName, StringComparison.Ordinal))
            {
                if (!string.Equals(stored.Name, canonicalName, StringComparison.Ordinal))
                {
                    Execute(connection, transaction,
                        "UPDATE installations SET name = $name WHERE id = $id;",
                        ("$name", canonicalName), ("$id", stored.Id));
                    changed = true;
                }
                continue;
            }

            var canonicalId = GetOrCreateInstallation(
                connection, transaction, canonicalName, canonicalNormalizedName);
            var people = new List<(long Id, string DisplayName)>();
            using (var command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = "SELECT id, display_name FROM people WHERE installation_id = $installation;";
                command.Parameters.AddWithValue("$installation", stored.Id);
                using var reader = command.ExecuteReader();
                while (reader.Read())
                    people.Add((reader.GetInt64(0), reader.GetString(1)));
            }

            foreach (var person in people)
            {
                var canonicalPersonId = GetOrCreatePerson(
                    connection, transaction, canonicalId, person.DisplayName);
                Execute(connection, transaction, """
                    UPDATE attempts
                    SET installation_id = $installation, person_id = $person
                    WHERE person_id = $sourcePerson;
                    """, ("$installation", canonicalId), ("$person", canonicalPersonId),
                    ("$sourcePerson", person.Id));
            }

            Execute(connection, transaction,
                "DELETE FROM people WHERE installation_id = $installation;",
                ("$installation", stored.Id));
            Execute(connection, transaction,
                "DELETE FROM installations WHERE id = $installation;",
                ("$installation", stored.Id));
            changed = true;
        }

        if (changed)
            RecalculateDuplicateFlags(connection, transaction);
        transaction.Commit();
    }

    private static void EnsureAttemptMetadataColumns(SqliteConnection connection)
    {
        var columns = new (string Name, string Type)[]
        {
            ("session_guid", "TEXT"),
            ("attempt_date", "TEXT"),
            ("attempt_time", "TEXT"),
            ("machine_name", "TEXT"),
            ("plant_name", "TEXT"),
            ("student_serial", "TEXT"),
            ("entry_date", "TEXT"),
            ("session_start_date", "TEXT"),
            ("session_end_date", "TEXT"),
            ("is_closed", "INTEGER"),
            ("file_path", "TEXT"),
            ("resume", "INTEGER"),
            ("result", "INTEGER"),
            ("is_valid", "INTEGER"),
            ("report_type", "INTEGER"),
            ("has_report_file", "INTEGER"),
            ("report_file", "TEXT"),
            ("training_times_json", "TEXT"),
            ("exercise_logs_json", "TEXT"),
            ("training_times_count", "INTEGER"),
            ("exercise_logs_count", "INTEGER")
            ,("duplicate_of_attempt_id", "INTEGER")
            ,("duplicate_kind", "TEXT")
            ,("duplicate_reason", "TEXT")
        };
        foreach (var column in columns)
        {
            if (!ColumnExists(connection, "attempts", column.Name))
                Execute(connection, null, $"ALTER TABLE attempts ADD COLUMN {column.Name} {column.Type};");
        }
    }

    private static void AddSearchParameters(
        SqliteCommand command,
        DateTime? periodStart,
        DateTime? periodEnd,
        string? rawSearch,
        string normalizedSearch,
        string normalizedPerson,
        string normalizedInstallationSearch,
        TimeSpan? timeStart,
        TimeSpan? timeEnd,
        bool? duplicatesOnly)
    {
        command.Parameters.AddWithValue("$start",
            periodStart.HasValue ? StoreDate(periodStart.Value) : DBNull.Value);
        command.Parameters.AddWithValue("$end",
            periodEnd.HasValue ? StoreDate(periodEnd.Value) : DBNull.Value);
        command.Parameters.AddWithValue("$timeStart",
            timeStart.HasValue ? StoreTime(timeStart.Value) : DBNull.Value);
        command.Parameters.AddWithValue("$timeEnd",
            timeEnd.HasValue ? StoreTime(timeEnd.Value) : DBNull.Value);
        command.Parameters.AddWithValue("$duplicateFilter", duplicatesOnly switch
        {
            true => 1,
            false => 2,
            null => 0
        });
        command.Parameters.AddWithValue("$search", normalizedSearch);
        command.Parameters.AddWithValue("$personSearch", normalizedPerson);
        command.Parameters.AddWithValue("$installationSearch", normalizedInstallationSearch);
        command.Parameters.AddWithValue("$rawSearch", rawSearch?.Trim() ?? "");
    }

    private static void AddPersonIdParameters(
        SqliteCommand command,
        IReadOnlyList<long> personIds)
    {
        for (var index = 0; index < personIds.Count; index++)
            command.Parameters.AddWithValue($"$personId{index}", personIds[index]);
    }

    private DatabasePersonNameCatalog ReadPersonNameCatalog(SqliteConnection connection)
    {
        var configuredNames = GetConfiguredInstallationNames();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT p.id, i.name, p.display_name
            FROM people p
            JOIN installations i ON i.id = p.installation_id
            WHERE CONFIGURED_INSTALLATIONS;
            """.Replace("CONFIGURED_INSTALLATIONS",
                ConfiguredInstallationPredicate("i", configuredNames.Count), StringComparison.Ordinal);
        AddConfiguredInstallationParameters(command, configuredNames);
        using var reader = command.ExecuteReader();
        var people = new List<PersonIdentity>();
        while (reader.Read())
            people.Add(new PersonIdentity(reader.GetInt64(0), reader.GetString(1), reader.GetString(2)));
        return new DatabasePersonNameCatalog(people, _fioKeys, _installations);
    }

    private static DatabaseAttempt ReadAttempt(SqliteDataReader reader) => new(
        reader.GetInt64(0),
        reader.GetString(1),
        reader.GetString(2),
        reader.GetString(3),
        ParseStoredDateTime(reader.GetString(4)),
        reader.GetDouble(5),
        ReadNullableString(reader, 6),
        ReadNullableString(reader, 7),
        ReadNullableString(reader, 8),
        ReadNullableString(reader, 9),
        ReadNullableString(reader, 10),
        reader.IsDBNull(11) ? null : reader.GetInt32(11),
        reader.GetString(12),
        reader.GetString(13),
        reader.GetInt32(14),
        reader.GetInt32(15) != 0,
        ReadNullableString(reader, 16),
        ReadNullableString(reader, 17),
        ReadNullableString(reader, 18),
        ReadNullableString(reader, 19),
        ReadNullableDateTime(reader, 20),
        ReadNullableDateTime(reader, 21),
        ReadNullableDateTime(reader, 22),
        ReadNullableBoolean(reader, 23),
        ReadNullableString(reader, 24),
        reader.IsDBNull(25) ? null : reader.GetInt32(25),
        ReadNullableBoolean(reader, 26),
        ReadNullableBoolean(reader, 27),
        reader.IsDBNull(28) ? null : reader.GetInt32(28),
        ReadNullableBoolean(reader, 29),
        ReadNullableString(reader, 30),
        ReadNullableString(reader, 31),
        ReadNullableString(reader, 32),
        reader.IsDBNull(33) ? null : reader.GetInt32(33),
        reader.IsDBNull(34) ? null : reader.GetInt32(34),
        reader.IsDBNull(35) ? null : reader.GetInt64(35),
        ReadNullableString(reader, 36),
        ReadNullableString(reader, 37));

    private static void ValidateEditModel(DatabaseAttemptEditModel model)
    {
        if (string.IsNullOrWhiteSpace(model.Installation))
            throw new ArgumentException("Укажите установку.");
        if (string.IsNullOrWhiteSpace(model.PersonName))
            throw new ArgumentException("Укажите ФИО.");
        if (string.IsNullOrWhiteSpace(model.ScenarioName))
            throw new ArgumentException("Укажите сценарий.");
        if (!double.IsFinite(model.Percent) || model.Percent is < 0 or > 100)
            throw new ArgumentException("Процент должен быть от 0 до 100.");
        if (model.Actions < 0)
            throw new ArgumentException("Количество действий не может быть отрицательным.");
    }

    private static TrainingAttemptRecord ToTrainingAttempt(
        DatabaseAttemptEditModel model,
        string sourceSheet,
        int sourceRow,
        string? installation = null) => new(
            installation ?? model.Installation.Trim(), model.PersonName.Trim(), model.ScenarioName.Trim(),
            model.AttemptedAt, model.Percent, NullIfWhiteSpace(model.Mode),
            NullIfWhiteSpace(model.Project), NullIfWhiteSpace(model.Role),
            NullIfWhiteSpace(model.Credited), NullIfWhiteSpace(model.Duration), model.Actions,
            sourceSheet, sourceRow,
            new JsonAttemptMetadata(
                NullIfWhiteSpace(model.SessionGuid), NullIfWhiteSpace(model.MachineName),
                NullIfWhiteSpace(model.PlantName), NullIfWhiteSpace(model.StudentSerial),
                model.EntryDate, model.SessionStartDate, model.SessionEndDate, model.IsClosed,
                NullIfWhiteSpace(model.FilePath), model.Resume, model.Result, model.IsValid,
                model.ReportType, model.HasReportFile, NullIfWhiteSpace(model.ReportFile),
                NullIfWhiteSpace(model.TrainingTimesJson), NullIfWhiteSpace(model.ExerciseLogsJson),
                model.TrainingTimesCount, model.ExerciseLogsCount));

    private string ResolveConfiguredInstallation(string value)
    {
        var trimmed = value.Trim();
        if (_installations.Count == 0)
            return trimmed;
        return MatchingService.MatchInstallation(trimmed, _installations).Name
               ?? throw new ArgumentException(
                   $"Установка «{trimmed}» не найдена в настройках. Добавьте это название как синоним нужной установки.");
    }

    private IReadOnlyList<string> GetConfiguredInstallationNames() => _installations.Keys
        .Select(TextNormalization.NormalizeText)
        .Where(name => name.Length > 0)
        .Distinct(StringComparer.Ordinal)
        .ToArray();

    private static string ConfiguredInstallationPredicate(string tableAlias, int count) => count == 0
        ? "1 = 1"
        : tableAlias + ".normalized_name IN (" +
          string.Join(", ", Enumerable.Range(0, count).Select(index => $"$configuredInstallation{index}")) + ")";

    private static void AddConfiguredInstallationParameters(
        SqliteCommand command,
        IReadOnlyList<string> configuredNames)
    {
        for (var index = 0; index < configuredNames.Count; index++)
            command.Parameters.AddWithValue($"$configuredInstallation{index}", configuredNames[index]);
    }

    private static string? NullIfWhiteSpace(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static long GetOrCreateManualSource(
        SqliteConnection connection,
        SqliteTransaction transaction)
    {
        Execute(connection, transaction, """
            INSERT INTO source_files(
                file_name, original_path, file_hash, format, imported_at,
                rows_count, error_count, parser_version)
            VALUES('Ручное редактирование', '', $hash, 'Manual', $imported, 0, 0, $parserVersion)
            ON CONFLICT(file_hash) DO NOTHING;
            """, ("$hash", ManualSourceHash), ("$imported", StoreDateTime(DateTime.Now)),
            ("$parserVersion", TrainingExportParser.Version));
        return SelectId(connection, transaction,
            "SELECT id FROM source_files WHERE file_hash = $hash;", ("$hash", ManualSourceHash));
    }

    private static int NextManualSourceRow(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long sourceId)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = "SELECT COALESCE(MAX(source_row), 0) + 1 FROM attempts WHERE source_file_id = $source;";
        command.Parameters.AddWithValue("$source", sourceId);
        return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    private static void RefreshSourceStatistics(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long sourceId)
    {
        Execute(connection, transaction, """
            UPDATE source_files
            SET rows_count = (SELECT COUNT(*) FROM attempts WHERE source_file_id = $source),
                period_start = (SELECT MIN(attempted_at) FROM attempts WHERE source_file_id = $source),
                period_end = (SELECT MAX(attempted_at) FROM attempts WHERE source_file_id = $source)
            WHERE id = $source;
            """, ("$source", sourceId));
    }

    private int SaveParsedFile(
        SqliteConnection connection,
        string path,
        string hash,
        ParsedTrainingExport parsed,
        long? replacedSourceFileId,
        CancellationToken cancellationToken)
    {
        using var transaction = connection.BeginTransaction();
        if (replacedSourceFileId.HasValue)
            Execute(connection, transaction, "DELETE FROM source_files WHERE id = $id;",
                ("$id", replacedSourceFileId.Value));
        var sourceFileId = InsertSourceFile(connection, transaction, path, hash, parsed);
        var installationIds = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        var personIds = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        var scenarioIds = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var attempt in parsed.Attempts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var installationKey = TextNormalization.NormalizeText(attempt.Installation);
            if (!installationIds.TryGetValue(installationKey, out var installationId))
            {
                installationId = GetOrCreateInstallation(connection, transaction, attempt.Installation, installationKey);
                installationIds[installationKey] = installationId;
            }

            var personKey = installationId.ToString(CultureInfo.InvariantCulture) + "|" +
                            TextNormalization.NormalizePersonName(attempt.PersonName);
            if (!personIds.TryGetValue(personKey, out var personId))
            {
                personId = GetOrCreatePerson(connection, transaction, installationId, attempt.PersonName);
                personIds[personKey] = personId;
            }

            var scenarioKey = TextNormalization.NormalizeText(attempt.ScenarioName);
            if (!scenarioIds.TryGetValue(scenarioKey, out var scenarioId))
            {
                scenarioId = GetOrCreateScenario(connection, transaction, attempt.ScenarioName, scenarioKey);
                scenarioIds[scenarioKey] = scenarioId;
            }

            var fingerprint = CalculateRecordFingerprint(attempt);
            InsertAttempt(connection, transaction, sourceFileId, installationId, personId,
                scenarioId, attempt, fingerprint, false);
        }

        foreach (var error in parsed.Errors)
        {
            Execute(connection, transaction, """
                INSERT INTO import_errors(source_file_id, source_sheet, source_row, message)
                VALUES($source, $sheet, $row, $message);
                """,
                ("$source", sourceFileId), ("$sheet", error.Sheet),
                ("$row", error.Row), ("$message", error.Message));
        }

        if (replacedSourceFileId.HasValue)
            CleanOrphans(connection, transaction);
        RecalculateDuplicateFlags(connection, transaction);
        var duplicateCount = CountPossibleDuplicates(connection, transaction, sourceFileId);

        transaction.Commit();
        return duplicateCount;
    }

    private static long InsertSourceFile(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string path,
        string hash,
        ParsedTrainingExport parsed)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO source_files(
                file_name, original_path, file_hash, format, imported_at,
                period_start, period_end, rows_count, error_count, parser_version)
            VALUES($name, $path, $hash, $format, $imported,
                   $start, $end, $rows, $errors, $parserVersion);
            SELECT last_insert_rowid();
            """;
        command.Parameters.AddWithValue("$name", Path.GetFileName(path));
        command.Parameters.AddWithValue("$path", path);
        command.Parameters.AddWithValue("$hash", hash);
        command.Parameters.AddWithValue("$format", parsed.Format.ToString());
        command.Parameters.AddWithValue("$imported", StoreDateTime(DateTime.Now));
        command.Parameters.AddWithValue("$start", DbValue(parsed.PeriodStart));
        command.Parameters.AddWithValue("$end", DbValue(parsed.PeriodEnd));
        command.Parameters.AddWithValue("$rows", parsed.Attempts.Count);
        command.Parameters.AddWithValue("$errors", parsed.Errors.Count);
        command.Parameters.AddWithValue("$parserVersion", ParserVersionFor(path));
        return (long)(command.ExecuteScalar()
            ?? throw new InvalidOperationException("Не удалось создать запись исходного файла."));
    }

    private static long GetOrCreateInstallation(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string displayName,
        string normalizedName)
    {
        Execute(connection, transaction, """
            INSERT INTO installations(name, normalized_name)
            VALUES($name, $normalized)
            ON CONFLICT(normalized_name) DO NOTHING;
            """, ("$name", displayName), ("$normalized", normalizedName));
        return SelectId(connection, transaction,
            "SELECT id FROM installations WHERE normalized_name = $normalized;",
            ("$normalized", normalizedName));
    }

    private static long GetOrCreatePerson(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long installationId,
        string displayName)
    {
        var normalizedName = TextNormalization.NormalizePersonName(displayName);
        Execute(connection, transaction, """
            INSERT INTO people(installation_id, display_name, normalized_name)
            VALUES($installation, $name, $normalized)
            ON CONFLICT(installation_id, normalized_name) DO NOTHING;
            """, ("$installation", installationId), ("$name", displayName),
            ("$normalized", normalizedName));
        return SelectId(connection, transaction, """
            SELECT id FROM people
            WHERE installation_id = $installation AND normalized_name = $normalized;
            """, ("$installation", installationId), ("$normalized", normalizedName));
    }

    private static long GetOrCreateScenario(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string displayName,
        string normalizedName)
    {
        Execute(connection, transaction, """
            INSERT INTO scenarios(display_name, normalized_name)
            VALUES($name, $normalized)
            ON CONFLICT(normalized_name) DO NOTHING;
            """, ("$name", displayName), ("$normalized", normalizedName));
        return SelectId(connection, transaction,
            "SELECT id FROM scenarios WHERE normalized_name = $normalized;",
            ("$normalized", normalizedName));
    }

    private static void InsertAttempt(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long sourceFileId,
        long installationId,
        long personId,
        long scenarioId,
        TrainingAttemptRecord attempt,
        string fingerprint,
        bool duplicate)
    {
        Execute(connection, transaction, """
            INSERT INTO attempts(
                source_file_id, installation_id, person_id, scenario_id,
                attempted_at, attempt_date, attempt_time, percent, mode, project, role, credited, duration,
                actions, session_guid, machine_name, plant_name, student_serial,
                entry_date, session_start_date, session_end_date, is_closed,
                file_path, resume, result, is_valid, report_type, has_report_file,
                report_file, training_times_json, exercise_logs_json,
                training_times_count, exercise_logs_count,
                source_sheet, source_row, record_fingerprint, possible_duplicate)
            VALUES($source, $installation, $person, $scenario,
                   $date, $attemptDate, $attemptTime, $percent, $mode, $project, $role, $credited, $duration,
                   $actions, $sessionGuid, $machineName, $plantName, $studentSerial,
                   $entryDate, $sessionStart, $sessionEnd, $isClosed,
                   $filePath, $resume, $result, $isValid, $reportType, $hasReportFile,
                   $reportFile, $trainingTimesJson, $exerciseLogsJson,
                   $trainingTimesCount, $exerciseLogsCount,
                   $sheet, $row, $fingerprint, $duplicate);
            """,
            ("$source", sourceFileId), ("$installation", installationId),
            ("$person", personId), ("$scenario", scenarioId),
            ("$date", StoreDateTime(attempt.AttemptedAt)), ("$percent", attempt.Percent),
            ("$attemptDate", StoreDate(attempt.AttemptedAt)),
            ("$attemptTime", StoreTime(attempt.AttemptedAt.TimeOfDay)),
            ("$mode", attempt.Mode), ("$project", attempt.Project),
            ("$role", attempt.Role), ("$credited", attempt.Credited),
            ("$duration", attempt.Duration), ("$actions", attempt.Actions),
            ("$sessionGuid", attempt.JsonMetadata?.Guid),
            ("$machineName", attempt.JsonMetadata?.MachineName),
            ("$plantName", attempt.JsonMetadata?.PlantName),
            ("$studentSerial", attempt.JsonMetadata?.StudentSerial),
            ("$entryDate", DbValue(attempt.JsonMetadata?.EntryDate)),
            ("$sessionStart", DbValue(attempt.JsonMetadata?.SessionStartDate)),
            ("$sessionEnd", DbValue(attempt.JsonMetadata?.SessionEndDate)),
            ("$isClosed", DbBoolean(attempt.JsonMetadata?.IsClosed)),
            ("$filePath", attempt.JsonMetadata?.FilePath),
            ("$resume", attempt.JsonMetadata?.Resume),
            ("$result", DbBoolean(attempt.JsonMetadata?.Result)),
            ("$isValid", DbBoolean(attempt.JsonMetadata?.IsValid)),
            ("$reportType", attempt.JsonMetadata?.ReportType),
            ("$hasReportFile", DbBoolean(attempt.JsonMetadata?.HasReportFile)),
            ("$reportFile", attempt.JsonMetadata?.ReportFile),
            ("$trainingTimesJson", attempt.JsonMetadata?.TrainingTimesJson),
            ("$exerciseLogsJson", attempt.JsonMetadata?.ExerciseLogsJson),
            ("$trainingTimesCount", attempt.JsonMetadata?.TrainingTimesCount),
            ("$exerciseLogsCount", attempt.JsonMetadata?.ExerciseLogsCount),
            ("$sheet", attempt.SourceSheet), ("$row", attempt.SourceRow),
            ("$fingerprint", fingerprint), ("$duplicate", duplicate ? 1 : 0));
    }

    private SqliteConnection OpenConnection()
    {
        var connection = new SqliteConnection(_connectionString);
        connection.Open();
        Execute(connection, null, "PRAGMA foreign_keys=ON;");
        Execute(connection, null, "PRAGMA busy_timeout=5000;");
        return connection;
    }

    private static ExistingSourceFile? FindSourceFile(SqliteConnection connection, string hash)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT id, parser_version FROM source_files WHERE file_hash = $hash;";
        command.Parameters.AddWithValue("$hash", hash);
        using var reader = command.ExecuteReader();
        return reader.Read()
            ? new ExistingSourceFile(reader.GetInt64(0), reader.GetInt32(1))
            : null;
    }

    private static int CountPossibleDuplicates(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long sourceFileId)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            SELECT COUNT(*) FROM attempts
            WHERE source_file_id = $source AND possible_duplicate = 1;
            """;
        command.Parameters.AddWithValue("$source", sourceFileId);
        return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    private static void BackfillJsonModes(SqliteConnection connection)
    {
        var rows = new List<(long Id, DateTime AttemptedAt, string Scenario, string TrainingTimesJson)>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = """
                SELECT attempts.id, attempts.attempted_at, scenarios.display_name,
                       attempts.training_times_json
                FROM attempts
                JOIN scenarios ON scenarios.id = attempts.scenario_id
                JOIN source_files ON source_files.id = attempts.source_file_id
                WHERE source_files.format = 'Json'
                  AND (attempts.mode IS NULL OR trim(attempts.mode) = '')
                  AND attempts.training_times_json IS NOT NULL
                  AND trim(attempts.training_times_json) <> '';
                """;
            using var reader = command.ExecuteReader();
            while (reader.Read())
                rows.Add((reader.GetInt64(0), ParseStoredDateTime(reader.GetString(1)),
                    reader.GetString(2), reader.GetString(3)));
        }

        if (rows.Count == 0)
            return;

        using var transaction = connection.BeginTransaction();
        foreach (var row in rows)
        {
            var mode = JsonTrainingExportParser.ResolveMode(
                row.TrainingTimesJson, row.AttemptedAt, row.Scenario);
            if (string.IsNullOrWhiteSpace(mode))
                continue;
            Execute(connection, transaction,
                "UPDATE attempts SET mode = $mode WHERE id = $id;",
                ("$mode", mode), ("$id", row.Id));
        }
        transaction.Commit();
    }

    private static void RecalculateDuplicateFlags(
        SqliteConnection connection,
        SqliteTransaction? transaction)
    {
        Execute(connection, transaction, """
            UPDATE attempts
            SET possible_duplicate = 0,
                duplicate_of_attempt_id = NULL,
                duplicate_kind = NULL,
                duplicate_reason = NULL;

            UPDATE attempts AS current
            SET duplicate_of_attempt_id = (
                SELECT prior.id
                FROM attempts AS prior
                JOIN source_files AS prior_source ON prior_source.id = prior.source_file_id
                JOIN source_files AS current_source ON current_source.id = current.source_file_id
                WHERE prior.id <> current.id
                  AND prior.installation_id = current.installation_id
                  AND prior.person_id = current.person_id
                  AND prior.scenario_id = current.scenario_id
                  AND prior.attempt_date = current.attempt_date
                  AND prior.attempt_time = current.attempt_time
                  AND ABS(prior.percent - current.percent) < 0.000001
                  AND (
                      CASE WHEN prior_source.format = 'Json' THEN 0 ELSE 1 END
                          < CASE WHEN current_source.format = 'Json' THEN 0 ELSE 1 END
                      OR (
                          CASE WHEN prior_source.format = 'Json' THEN 0 ELSE 1 END
                              = CASE WHEN current_source.format = 'Json' THEN 0 ELSE 1 END
                          AND prior.id < current.id
                      )
                  )
                ORDER BY CASE WHEN prior_source.format = 'Json' THEN 0 ELSE 1 END, prior.id
                LIMIT 1
            );

            UPDATE attempts
            SET possible_duplicate = 1,
                duplicate_kind = 'Точный',
                duplicate_reason = 'Совпадают установка, ФИО, сценарий, дата, время и процент.'
            WHERE duplicate_of_attempt_id IS NOT NULL;

            UPDATE attempts AS current
            SET possible_duplicate = 1,
                duplicate_kind = 'Возможный',
                duplicate_reason = 'Совпадают установка, ФИО, сценарий и процент; время отличается не более чем на 60 секунд.'
            WHERE current.duplicate_of_attempt_id IS NULL
              AND EXISTS(
                SELECT 1
                FROM attempts AS prior
                JOIN source_files AS prior_source ON prior_source.id = prior.source_file_id
                JOIN source_files AS current_source ON current_source.id = current.source_file_id
                WHERE prior.id <> current.id
                  AND prior.installation_id = current.installation_id
                  AND prior.person_id = current.person_id
                  AND prior.scenario_id = current.scenario_id
                  AND prior.attempt_date = current.attempt_date
                  AND ABS(prior.percent - current.percent) < 0.000001
                  AND ABS((julianday(prior.attempted_at) - julianday(current.attempted_at)) * 86400.0)
                      BETWEEN 0.000001 AND 60.0
                  AND (
                      CASE WHEN prior_source.format = 'Json' THEN 0 ELSE 1 END
                          < CASE WHEN current_source.format = 'Json' THEN 0 ELSE 1 END
                      OR (
                          CASE WHEN prior_source.format = 'Json' THEN 0 ELSE 1 END
                              = CASE WHEN current_source.format = 'Json' THEN 0 ELSE 1 END
                          AND prior.id < current.id
                      )
                  )
              );
            """);
    }

    private static void CleanOrphans(
        SqliteConnection connection,
        SqliteTransaction transaction)
    {
        Execute(connection, transaction,
            "DELETE FROM people WHERE NOT EXISTS (SELECT 1 FROM attempts WHERE attempts.person_id = people.id);");
        Execute(connection, transaction,
            "DELETE FROM scenarios WHERE NOT EXISTS (SELECT 1 FROM attempts WHERE attempts.scenario_id = scenarios.id);");
        Execute(connection, transaction,
            "DELETE FROM installations WHERE NOT EXISTS (SELECT 1 FROM attempts WHERE attempts.installation_id = installations.id);");
    }

    private static int ReadUserVersion(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA user_version;";
        return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    private static bool ColumnExists(SqliteConnection connection, string table, string column)
    {
        using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({table});";
        using var reader = command.ExecuteReader();
        while (reader.Read())
            if (string.Equals(reader.GetString(1), column, StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }

    private static long SelectId(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string sql,
        params (string Name, object? Value)[] parameters)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        AddParameters(command, parameters);
        return (long)(command.ExecuteScalar()
            ?? throw new InvalidOperationException("Не удалось получить идентификатор записи базы данных."));
    }

    private static void Execute(
        SqliteConnection connection,
        SqliteTransaction? transaction,
        string sql,
        params (string Name, object? Value)[] parameters)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        AddParameters(command, parameters);
        command.ExecuteNonQuery();
    }

    private static void AddParameters(
        SqliteCommand command,
        IEnumerable<(string Name, object? Value)> parameters)
    {
        foreach (var (name, value) in parameters)
            command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private static string CalculateFileHash(string path, CancellationToken cancellationToken)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var hash = SHA256.Create();
        var buffer = new byte[1024 * 128];
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            hash.TransformBlock(buffer, 0, read, null, 0);
        }
        hash.TransformFinalBlock([], 0, 0);
        return Convert.ToHexString(hash.Hash!);
    }

    private static string CalculateRecordFingerprint(TrainingAttemptRecord attempt)
    {
        var value = string.Join('|',
            TextNormalization.NormalizeText(attempt.Installation),
            TextNormalization.NormalizePersonName(attempt.PersonName),
            TextNormalization.NormalizeText(attempt.ScenarioName),
            StoreDate(attempt.AttemptedAt),
            StoreTime(attempt.AttemptedAt.TimeOfDay),
            attempt.Percent.ToString("R", CultureInfo.InvariantCulture));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
    }

    private static bool IsSupportedImportFile(string path) =>
        Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
        Path.GetExtension(path).Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
        Path.GetExtension(path).Equals(".json", StringComparison.OrdinalIgnoreCase);

    private static int ParserVersionFor(string path) =>
        Path.GetExtension(path).Equals(".json", StringComparison.OrdinalIgnoreCase)
            ? JsonTrainingExportParser.Version
            : TrainingExportParser.Version;

    private static string StoreDateTime(DateTime value) =>
        value.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff", CultureInfo.InvariantCulture);

    private static string StoreDate(DateTime value) =>
        value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static string StoreTime(TimeSpan value) =>
        value.ToString(@"hh\:mm\:ss", CultureInfo.InvariantCulture);

    private static object DbValue(DateTime? value) =>
        value.HasValue ? StoreDateTime(value.Value) : DBNull.Value;

    private static object DbBoolean(bool? value) =>
        value.HasValue ? value.Value ? 1 : 0 : DBNull.Value;

    private static DateTime ParseStoredDateTime(string value) =>
        DateTime.ParseExact(value, "yyyy-MM-dd'T'HH:mm:ss.fffffff",
            CultureInfo.InvariantCulture, DateTimeStyles.None);

    private static DateTime? ReadNullableDateTime(SqliteDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? null : ParseStoredDateTime(reader.GetString(ordinal));

    private static string? ReadNullableString(SqliteDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);

    private static bool? ReadNullableBoolean(SqliteDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? null : reader.GetInt32(ordinal) != 0;

    private static string ImportErrorSummary(IReadOnlyList<ExportImportError> errors, string fallback)
    {
        if (errors.Count == 0)
            return fallback;
        var summary = string.Join(" ", errors.Take(5)
            .Select(error => $"{error.Sheet}: {error.Message}"));
        return errors.Count > 5 ? summary + $" Ещё замечаний: {errors.Count - 5}." : summary;
    }

    private static string FriendlyImportError(Exception exception) => exception switch
    {
        IOException => "Файл занят другой программой или недоступен для чтения.",
        UnauthorizedAccessException => "Нет доступа к файлу или папке базы данных.",
        System.Text.Json.JsonException => "JSON повреждён или имеет некорректную структуру: " + exception.Message,
        InvalidDataException => "Неподдерживаемая структура файла: " + exception.Message,
        _ => exception.Message
    };

    private sealed record ExistingSourceFile(long Id, int ParserVersion);
}

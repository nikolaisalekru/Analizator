using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Analizator.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Data.Sqlite;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Writer;

var failures = new List<string>();
Run("Нормализация", () =>
{
    Equal("ф и о по отчету", TextNormalization.NormalizeHeader("  Ф.И.О.\nпо-отчёту  "));
    Equal(8, TextNormalization.MonthFromFileName("Выгрузка_КТК_август_2026.xlsx"));
    Equal("Выгрузка_КТК_МНПЗ_август_2026.xlsx",
        DatabaseSelectionExportService.SuggestedFileName(
            new DateTime(2026, 8, 1), new DateTime(2026, 8, 31)));
    Equal("Выгрузка_КТК_МНПЗ_2026-08-03_2026-08-10.xlsx",
        DatabaseSelectionExportService.SuggestedFileName(
            new DateTime(2026, 8, 3), new DateTime(2026, 8, 10)));
    Equal(92.5, TextNormalization.SafeDouble("92,5%"));
});
Run("Сопоставление", () =>
{
    var installations = new Dictionary<string, List<string>> { ["ГФУ-2"] = ["ГФУ", "ГФУ 2"] };
    Equal("ГФУ-2", MatchingService.MatchInstallation("ГФУ", installations).Name);
    var exact = MatchingService.MatchPerson("Иванов Иван", ["Иванов  Иван"], "ГФУ-2",
        new Dictionary<string, Dictionary<string, string>>(), installations);
    Equal(MatchKind.Exact, exact.Kind);
});

var root = Path.Combine(Path.GetTempPath(), "Analizator-CSharp-tests-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    Run("Сохранение настройки порогов", () =>
    {
        var config = Path.Combine(root, "threshold-save-config");
        Directory.CreateDirectory(config);
        var settingsPath = Path.Combine(config, "settings.json");
        File.WriteAllText(settingsPath, "{\"custom_setting\":\"preserve-me\"}");
        ConfigurationService.SavePassingThresholds(config, new PassingThresholdSettings
        {
            UsePerInstallation = true,
            GlobalPercent = 70d,
            Installations = new Dictionary<string, double>
            {
                ["ГФУ-2"] = 55d,
                ["ВБ"] = 65d
            }
        });

        var loaded = ConfigurationService.Load(config).Settings.PassingThreshold;
        True(loaded.UsePerInstallation, "Индивидуальный режим не сохранился");
        Equal(70d, loaded.GlobalPercent);
        Equal(55d, loaded.GetForInstallation("ГФУ-2"));
        Equal(65d, loaded.GetForInstallation("ВБ"));
        using var document = JsonDocument.Parse(File.ReadAllText(settingsPath));
        Equal("preserve-me", document.RootElement.GetProperty("custom_setting").GetString());

        ConfigurationService.SaveDatabaseBackupSettings(config, new DatabaseBackupSettings
        {
            AutomaticBackupsEnabled = false,
            IntervalDays = 7,
            MaximumBackupFiles = 4,
            BackupBeforeChanges = false,
            ConfirmRowDeletion = false,
            ConfirmSourceDeletion = true
        });
        ConfigurationService.SaveDatabaseEditorSettings(config, new DatabaseEditorSettings
        {
            HiddenColumns = ["Guid", "ReportFile"]
        });
        var databaseSettings = ConfigurationService.Load(config).Settings;
        True(!databaseSettings.DatabaseBackups.AutomaticBackupsEnabled,
            "Настройка автоматических копий не сохранилась");
        Equal(7, databaseSettings.DatabaseBackups.IntervalDays);
        Equal(4, databaseSettings.DatabaseBackups.MaximumBackupFiles);
        True(!databaseSettings.DatabaseBackups.BackupBeforeChanges,
            "Настройка копии перед изменением не сохранилась");
        True(!databaseSettings.DatabaseBackups.ConfirmRowDeletion,
            "Настройка подтверждения удаления строки не сохранилась");
        True(databaseSettings.DatabaseEditor.HiddenColumns.SetEquals(["Guid", "ReportFile"]),
            "Список скрытых столбцов не сохранился");
        using var updatedDocument = JsonDocument.Parse(File.ReadAllText(settingsPath));
        Equal("preserve-me", updatedDocument.RootElement.GetProperty("custom_setting").GetString());
    });

    Run("Нативное сохранение конфигураций", () =>
    {
        var config = Path.Combine(root, "native-config-editor");
        ConfigurationService.Load(config);

        ConfigurationService.SaveInstallations(config,
            new Dictionary<string, IReadOnlyCollection<string>>
            {
                ["Тестовая установка"] = ["ТУ", "Тест"]
            });
        ConfigurationService.SaveFioKeys(config,
            new Dictionary<string, IReadOnlyDictionary<string, string>>
            {
                ["Тестовая установка"] = new Dictionary<string, string>
                {
                    ["$Иванов Ив"] = "Иванов Иван Иванович"
                }
            });
        ConfigurationService.SaveExcludedPeople(config, ["Авто-оператор", "Служебная запись"]);
        ConfigurationService.SaveAnalysisSettings(config, false, 3, false);

        var editable = ConfigurationService.LoadEditable(config);
        True(editable.Installations["Тестовая установка"].SequenceEqual(["ТУ", "Тест"]),
            "Синонимы установки не сохранились");
        Equal("Иванов Иван Иванович", editable.FioKeys["Тестовая установка"]["$Иванов Ив"]);
        True(editable.ExcludedPeople.Contains("Служебная запись"),
            "Исключение не сохранилось");
        True(!editable.Settings.ParallelProcessing, "Настройка параллельной обработки не сохранилась");
        Equal(3, editable.Settings.MaxWorkers);
        True(!editable.Settings.AutoAddNewPeople.Default,
            "Настройка автодобавления не сохранилась");

        var normalized = ConfigurationService.Load(config);
        Equal("Тестовая установка", MatchingService.MatchInstallation("ТУ", normalized.Installations).Name);
        True(normalized.ExcludedPeople.Contains("служебная запись"),
            "Сохранённое исключение не загружается ядром");
    });

    Run("Установки базы ограничены настройками и синонимами", () =>
    {
        var directory = Path.Combine(root, "configured-installations");
        var databasePath = Path.Combine(directory, "analizator.db");
        var installations = new Dictionary<string, List<string>>
        {
            ["ГФУ-2"] = ["ГФУ", "ГФУ 2"]
        };
        var database = new TrainingDatabaseService(databasePath);
        database.AddAttempt(new DatabaseAttemptEditModel
        {
            AttemptedAt = new DateTime(2026, 8, 3, 10, 0, 0),
            Installation = "ГФУ",
            PersonName = "Иванов Иван Иванович",
            ScenarioName = "Сценарий А",
            Percent = 80
        });
        database.AddAttempt(new DatabaseAttemptEditModel
        {
            AttemptedAt = new DateTime(2026, 8, 4, 10, 0, 0),
            Installation = "Неизвестная установка",
            PersonName = "Петров Пётр Петрович",
            ScenarioName = "Сценарий Б",
            Percent = 90
        });

        database.SetPersonNameMappings(
            new Dictionary<string, Dictionary<string, string>>(), installations);
        Equal(1, database.GetSummary().Attempts);
        Equal("ГФУ-2", database.GetAttempts(
            new DateTime(2026, 8, 1), new DateTime(2026, 8, 31)).Single().Installation);
        Equal(1, database.SearchAttempts(null, null, "ГФУ").TotalCount);
        var unmatched = database.GetUnmatchedInstallations().Single();
        Equal("Неизвестная установка", unmatched.Name);
        Equal(1, unmatched.Attempts);

        database.AddAttempt(new DatabaseAttemptEditModel
        {
            AttemptedAt = new DateTime(2026, 8, 5, 10, 0, 0),
            Installation = "ГФУ 2",
            PersonName = "Сидоров Сидор Сидорович",
            ScenarioName = "Сценарий В",
            Percent = 95
        });
        True(database.GetAttempts(new DateTime(2026, 8, 1), new DateTime(2026, 8, 31))
                .All(item => item.Installation == "ГФУ-2"),
            "Синоним ручной записи не заменён каноническим названием");
        var rejected = false;
        try
        {
            database.AddAttempt(new DatabaseAttemptEditModel
            {
                AttemptedAt = new DateTime(2026, 8, 6, 10, 0, 0),
                Installation = "Совсем новая установка",
                PersonName = "Сотрудник",
                ScenarioName = "Сценарий",
                Percent = 75
            });
        }
        catch (ArgumentException exception)
        {
            rejected = exception.Message.Contains("не найдена в настройках", StringComparison.OrdinalIgnoreCase);
        }
        True(rejected, "Неизвестная установка ручной записи не отклонена с предупреждением");

        var exportPath = Path.Combine(directory, "установки.xlsx");
        using (var workbook = new XLWorkbook())
        {
            foreach (var sheetName in new[] { "ГФУ", "Неизвестная установка" })
            {
                var sheet = workbook.AddWorksheet(sheetName);
                sheet.Cell(1, 1).Value = "Тестовая выгрузка";
                sheet.Cell(2, 1).Value = "Дата";
                sheet.Cell(2, 2).Value = "Тема";
                sheet.Cell(2, 3).Value = "Ф.И.О. по отчёту";
                sheet.Cell(2, 4).Value = "Процент";
                sheet.Cell(2, 5).Value = "Режим";
                sheet.Cell(3, 1).Value = "07.08.2026";
                sheet.Cell(3, 2).Value = "Сценарий импорта";
                sheet.Cell(3, 3).Value = "Тестовый сотрудник";
                sheet.Cell(3, 4).Value = 88;
                sheet.Cell(3, 5).Value = "Экзамен";
            }
            workbook.SaveAs(exportPath);
        }
        var importedDatabase = new TrainingDatabaseService(
            Path.Combine(directory, "imported.db"), null, installations);
        var import = importedDatabase.ImportFiles([exportPath], installations);
        Equal(1, import.ImportedAttempts);
        Equal(1, import.RowErrors);
        True(import.Files.Single().Message.Contains("Неизвестная установка", StringComparison.Ordinal),
            "Предупреждение не содержит название несопоставленной установки");
        Equal("ГФУ-2", importedDatabase.SearchAttempts(null, null, null).Items.Single().Installation);

        var jsonPath = Path.Combine(directory, "установки.json");
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(new
        {
            ExerciseReports = new[]
            {
                new
                {
                    PlantName = "ГФУ",
                    StudentName = "Тестовый сотрудник",
                    ExerciseName = "Сценарий JSON",
                    Date = "2026-08-08T10:00:00",
                    Percent = 91
                },
                new
                {
                    PlantName = "Неизвестная установка",
                    StudentName = "Тестовый сотрудник",
                    ExerciseName = "Сценарий JSON",
                    Date = "2026-08-08T11:00:00",
                    Percent = 92
                }
            }
        }));
        var parsedJson = JsonTrainingExportParser.Parse(jsonPath, installations);
        Equal(1, parsedJson.Attempts.Count);
        Equal("ГФУ-2", parsedJson.Attempts.Single().Installation);
        Equal(1, parsedJson.Errors.Count);
        Equal("Неизвестная установка", parsedJson.Errors.Single().Sheet);
    });

    Run("Локальная SQLite-база и оба формата выгрузки", () =>
    {
        var current = Path.Combine(root, "Выгрузка_КТК_август_database.xlsx");
        var legacy = Path.Combine(root, "Общие_январь_database.xlsx");
        WriteExport(current,
        [
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 80d),
            ("04.08.2026", "Сценарий А", "Иванов Иван Иванович", 95d)
        ]);
        WriteLegacyExport(legacy,
        [
            ("02.01.2026 02:08", "Сценарий Б", "Иванов Иван Иванович", 100d),
            ("03.01.2026 07:42", "Сценарий Б", "Иванов Иван Иванович", 85d)
        ]);

        var installations = new Dictionary<string, List<string>>
        {
            ["ГФУ-2"] = ["ГФУ", "ГФУ 2"]
        };
        var database = new TrainingDatabaseService(Path.Combine(root, "database", "analizator.db"));
        var imported = database.ImportFiles([current, legacy], installations);
        Equal(2, imported.ImportedFiles);
        Equal(4, imported.ImportedAttempts);
        Equal(0, imported.FailedFiles);

        var summary = database.GetSummary();
        Equal(4, summary.Attempts);
        Equal(1, summary.People);
        Equal(2, summary.SourceFiles);
        Equal(new DateTime(2026, 1, 2, 2, 8, 0), summary.PeriodStart);
        Equal(new DateTime(2026, 8, 4), summary.PeriodEnd);
        var availableMonths = database.GetAvailableMonths();
        Equal(2, availableMonths.Count);
        Equal(new DateTime(2026, 8, 1), availableMonths[0].FirstDay);
        Equal(2, availableMonths[0].Attempts);
        Equal(0, availableMonths[0].ExactDuplicates);

        using (var connection = new SqliteConnection($"Data Source={database.DatabasePath}"))
        {
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT COUNT(*) FROM attempts
                WHERE attempt_date IS NOT NULL AND attempt_time IS NOT NULL;
                """;
            Equal(4L, (long)(command.ExecuteScalar() ?? 0L));
            command.CommandText = """
                UPDATE attempts SET possible_duplicate = 1;
                UPDATE source_files SET parser_version = 1;
                PRAGMA user_version = 1;
                """;
            command.ExecuteNonQuery();
        }
        database = new TrainingDatabaseService(database.DatabasePath);
        Equal(0, database.GetSummary().PossibleDuplicates);
        var refreshed = database.ImportFiles([current], installations);
        Equal(1, refreshed.UpdatedFiles);
        Equal(4, database.GetSummary().Attempts);

        var august = database.GetAttempts(new DateTime(2026, 8, 1), new DateTime(2026, 8, 31));
        Equal(2, august.Count);
        var aggregate = database.ReadForAnalysis(
            new DateTime(2026, 8, 1), new DateTime(2026, 8, 31), new ModeFilterSettings());
        Equal(2, aggregate.TotalReports);
        Equal(95d, aggregate.Data["ГФУ-2"]["Иванов Иван Иванович"]
            [TextNormalization.NormalizeText("Сценарий А")].BestPercent);

        var repeated = database.ImportFiles([current], installations);
        Equal(1, repeated.SkippedFiles);
        Equal(4, database.GetSummary().Attempts);

        var allRows = database.SearchAttempts(null, null, null, limit: 2);
        Equal(4, allRows.TotalCount);
        Equal(2, allRows.Items.Count);
        Equal(4, database.SearchAttempts(null, null, "Иванов").TotalCount);
        Equal(2, database.SearchAttempts(null, null, null, 100, 0,
            TimeSpan.Zero, TimeSpan.Zero).TotalCount);
        var manualId = database.AddAttempt(new DatabaseAttemptEditModel
        {
            AttemptedAt = new DateTime(2026, 9, 8, 12, 30, 0),
            Installation = "ГФУ-2",
            PersonName = "Петров Пётр Петрович",
            ScenarioName = "Ручной сценарий",
            Percent = 65d,
            Mode = "Экзамен",
            Actions = 3
        });
        Equal(5, database.GetSummary().Attempts);
        var manual = database.SearchAttempts(null, null, "Петров").Items.Single();
        Equal(manualId, manual.Id);
        Equal("Ручной", database.GetImportHistory().Single(item => item.Format == "Manual").FormatDisplay);
        database.UpdateAttempt(new DatabaseAttemptEditModel
        {
            Id = manualId,
            AttemptedAt = manual.AttemptedAt,
            Installation = manual.Installation,
            PersonName = manual.PersonName,
            ScenarioName = "Исправленный сценарий",
            Percent = 77d,
            Mode = manual.Mode,
            Actions = manual.Actions
        });
        var updatedManual = database.SearchAttempts(null, null, "Исправленный").Items.Single();
        Equal(77d, updatedManual.Percent);
        database.DeleteAttempt(manualId);
        Equal(4, database.GetSummary().Attempts);
        Equal(0, database.SearchAttempts(null, null, "Петров").TotalCount);
        True(database.GetImportHistory().All(item => item.Format != "Manual"),
            "Пустой источник ручных записей не был удалён");

        var overlap = Path.Combine(root, "Выгрузка_КТК_август_overlap.xlsx");
        WriteExport(overlap,
        [
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 80d),
            ("04.08.2026", "Сценарий А", "Иванов Иван Иванович", 95d)
        ]);
        using (var workbook = new XLWorkbook(overlap))
        {
            workbook.Worksheet(1).Cell("J1").Value = "Другой исходный файл";
            workbook.Save();
        }
        var overlapImport = database.ImportFiles([overlap], installations);
        Equal(1, overlapImport.ImportedFiles);
        Equal(2, overlapImport.PossibleDuplicates);
        Equal(2, database.GetSummary().PossibleDuplicates);
        var exactDuplicates = database.SearchAttempts(null, null, null, 100, 0,
            duplicatesOnly: true).Items;
        Equal(2, exactDuplicates.Count);
        True(exactDuplicates.All(item => item.DuplicateKind == "Точный" &&
                                         item.DuplicateOfAttemptId.HasValue),
            "Точные повторы не связаны с исходными записями");
        Equal(4, database.GetAttempts(new DateTime(2026, 1, 1),
            new DateTime(2026, 12, 31)).Count);
        var overlapHistory = database.GetImportHistory().Single(item => item.FileName == Path.GetFileName(overlap));
        database.DeleteSourceFile(overlapHistory.Id);
        Equal(0, database.GetSummary().PossibleDuplicates);
        Equal(4, database.GetSummary().Attempts);

        var nearOverlap = Path.Combine(root, "Выгрузка_КТК_август_near_overlap.xlsx");
        WriteExport(nearOverlap,
        [
            ("03.08.2026 00:00:30", "Сценарий А", "Иванов Иван Иванович", 80d)
        ]);
        var nearImport = database.ImportFiles([nearOverlap], installations);
        Equal(1, nearImport.PossibleDuplicates);
        var possibleDuplicate = database.SearchAttempts(null, null, null, 100, 0,
            duplicatesOnly: true).Items.Single();
        Equal("Возможный", possibleDuplicate.DuplicateKind);
        True(!possibleDuplicate.DuplicateOfAttemptId.HasValue,
            "Неточное совпадение ошибочно исключено из анализа");
        Equal(5, database.GetAttempts(new DateTime(2026, 1, 1),
            new DateTime(2026, 12, 31)).Count);
        var nearHistory = database.GetImportHistory().Single(item => item.FileName == Path.GetFileName(nearOverlap));
        database.DeleteSourceFile(nearHistory.Id);
        Equal(0, database.GetSummary().PossibleDuplicates);
        Equal(4, database.GetSummary().Attempts);

        var backup = database.CreateBackup(Path.Combine(root, "database", "backups"));
        True(File.Exists(backup), "Резервная копия базы не создана");
        Equal(4, new TrainingDatabaseService(backup).GetSummary().Attempts);
        var backupDirectory = Path.GetDirectoryName(backup)!;
        var obsoleteBackup = Path.Combine(backupDirectory, "analizator_2000-01-01_00-00-00-000.db");
        File.WriteAllText(obsoleteBackup, "old");
        File.SetLastWriteTimeUtc(obsoleteBackup, new DateTime(2000, 1, 1));
        database.CreateBackup(backupDirectory, 2);
        Equal(2, Directory.GetFiles(backupDirectory, "analizator_*.db").Length);
        True(!File.Exists(obsoleteBackup), "Старая резервная копия не удалена по лимиту");

        var legacyHistory = database.GetImportHistory().Single(item => item.FileName == Path.GetFileName(legacy));
        Equal("Старый", legacyHistory.FormatDisplay);
        database.DeleteSourceFile(legacyHistory.Id);
        Equal(2, database.GetSummary().Attempts);
    });

    Run("Выборочная выгрузка по сотрудникам из базы", () =>
    {
        var source = Path.Combine(root, "Выгрузка_КТК_выборка_source.xlsx");
        var databasePath = Path.Combine(root, "selection-export", "analizator.db");
        var output = Path.Combine(root, "selection-export", "result.xlsx");
        WriteExport(source,
        [
            ("03.08.2026 08:15:00", "Сценарий А", "Иванов Иван Иванович", 100d),
            ("04.08.2026 09:20:00", "Сценарий Б", "Иванов Иван Иванович", 85d),
            ("05.08.2026 10:30:00", "Сценарий В", "Петров Пётр Петрович", 90d)
        ]);
        var installations = new Dictionary<string, List<string>>
        {
            ["ГФУ-2"] = ["ГФУ", "ГФУ 2"],
            ["ЭЛОУ-АВТ-6"] = ["ЭЛОУ", "АВТ 6"]
        };
        var database = new TrainingDatabaseService(databasePath);
        database.ImportFiles([source], installations);
        database.AddAttempt(new DatabaseAttemptEditModel
        {
            AttemptedAt = new DateTime(2026, 8, 6, 11, 40, 0),
            Installation = "ЭЛОУ-АВТ-6",
            PersonName = "Сидоров Сидор Сидорович",
            ScenarioName = "Сценарий Г",
            Percent = 97d,
            Mode = "Экзамен",
            Duration = "00:04:10",
            Actions = 7
        });

        var people = database.GetPeopleForExport(
            new DateTime(2026, 8, 1), new DateTime(2026, 8, 31));
        Equal(3, people.Count);
        Equal(2, people.Single(item => item.PersonName.StartsWith("Иванов")).Attempts);

        var result = new DatabaseSelectionExportService().Export(
            database,
            new DatabaseSelectionExportRequest(
                output,
                new DateTime(2026, 8, 1),
                new DateTime(2026, 8, 31),
                ["ГФУ-2", "ЭЛОУ-АВТ-6"],
                [
                    new DatabasePersonKey("ГФУ-2", "Иванов Иван Иванович"),
                    new DatabasePersonKey("ЭЛОУ-АВТ-6", "Сидоров Сидор Сидорович")
                ],
                [
                    DatabaseExportColumn.Date,
                    DatabaseExportColumn.Scenario,
                    DatabaseExportColumn.Person,
                    DatabaseExportColumn.Duration,
                    DatabaseExportColumn.Actions,
                    DatabaseExportColumn.Percent,
                    DatabaseExportColumn.Mode
                ]));
        Equal(3, result.Attempts);
        Equal(2, result.People);
        Equal(2, result.Installations);
        True(File.Exists(output), "Выборочная выгрузка не создана");

        using (var workbook = new XLWorkbook(output))
        {
            Equal(2, workbook.Worksheets.Count);
            var gfu = workbook.Worksheet("ГФУ-2");
            Equal("Генератор Отчетов", gfu.Cell("A1").GetString());
            Equal("Дата", gfu.Cell("B14").GetString());
            Equal("Тема", gfu.Cell("C14").GetString());
            Equal("ФИО по отчёту", gfu.Cell("D14").GetString());
            Equal("Время", gfu.Cell("E14").GetString());
            Equal("Действий", gfu.Cell("F14").GetString());
            Equal("Процент", gfu.Cell("G14").GetString());
            Equal("Режим", gfu.Cell("H14").GetString());
            Equal(2, gfu.Cell("B6").GetValue<int>());
            Equal("Иванов Иван Иванович", gfu.Cell("D15").GetString());
            Equal("Иванов Иван Иванович", gfu.Cell("D16").GetString());
            True(gfu.Cell("B15").DataType == XLDataType.DateTime,
                "Дата записана текстом и не будет нормально фильтроваться в Excel");
            Equal(100d, gfu.Cell("G15").GetDouble());
            True(gfu.AutoFilter.IsEnabled, "В итоговой таблице не включён фильтр");
            True(gfu.SheetView.SplitRow == 14, "Заголовок таблицы не закреплён");
            var elou = workbook.Worksheet("ЭЛОУ-АВТ-6");
            True(elou.Cell("E15").DataType == XLDataType.TimeSpan,
                "Продолжительность записана текстом, а не временем Excel");
        }

        var parsed = TrainingExportParser.Parse(output, installations);
        Equal(3, parsed.Attempts.Count);
        True(parsed.Attempts.All(item => !item.PersonName.StartsWith("Петров")),
            "В выгрузку попал невыбранный сотрудник");

        var fullInstallationOutput = Path.Combine(root, "selection-export", "all-people.xlsx");
        var fullInstallationResult = new DatabaseSelectionExportService().Export(
            database,
            new DatabaseSelectionExportRequest(
                fullInstallationOutput,
                new DateTime(2026, 8, 1),
                new DateTime(2026, 8, 31),
                ["ГФУ-2"],
                [],
                [DatabaseExportColumn.Person, DatabaseExportColumn.Percent]));
        Equal(3, fullInstallationResult.Attempts);
        Equal(2, fullInstallationResult.People);
        Equal(1, fullInstallationResult.Installations);
        using (var fullInstallationWorkbook = new XLWorkbook(fullInstallationOutput))
        {
            Equal(1, fullInstallationWorkbook.Worksheets.Count);
            var sheet = fullInstallationWorkbook.Worksheet("ГФУ-2");
            Equal("Все", sheet.Cell("B10").GetString());
            Equal("ФИО по отчёту", sheet.Cell("B14").GetString());
            Equal("Процент", sheet.Cell("C14").GetString());
            Equal("Петров Пётр Петрович", sheet.Cell("B17").GetString());
        }
    });

    Run("Полные ФИО при хранении исходных сокращений", () =>
    {
        var directory = Path.Combine(root, "person-name-aliases");
        var databasePath = Path.Combine(directory, "analizator.db");
        var output = Path.Combine(directory, "result.xlsx");
        var installations = new Dictionary<string, List<string>>
        {
            ["Г-43-107"] = ["Г-43-107", "Г 43 107"]
        };
        var fioKeys = new Dictionary<string, Dictionary<string, string>>
        {
            ["Г-43-107"] = new()
            {
                ["$Офицеров Илья Алек"] = "Офицеров Илья Александрович"
            }
        };
        var database = new TrainingDatabaseService(databasePath, fioKeys, installations);
        foreach (var (person, minute) in new[]
                 {
                     ("Назаров Павел Валерьевич", 1),
                     ("$Назаров Павел Вале", 2),
                     ("$Офицеров Илья Алек", 3)
                 })
        {
            database.AddAttempt(new DatabaseAttemptEditModel
            {
                AttemptedAt = new DateTime(2026, 8, 3, 8, minute, 0),
                Installation = "Г-43-107",
                PersonName = person,
                ScenarioName = "Сценарий А",
                Percent = 100d
            });
        }

        using (var connection = new SqliteConnection($"Data Source={databasePath}"))
        {
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "SELECT COUNT(*) FROM people WHERE display_name LIKE '$%';";
            Equal(2L, (long)(command.ExecuteScalar() ?? 0L));
        }

        Equal(2, database.GetSummary().People);
        var people = database.GetPeopleForExport(
            new DateTime(2026, 8, 1), new DateTime(2026, 8, 31));
        Equal(2, people.Count);
        Equal(2, people.Single(item => item.PersonName == "Назаров Павел Валерьевич").Attempts);
        Equal(1, people.Single(item => item.PersonName == "Офицеров Илья Александрович").Attempts);

        var nazarov = database.SearchAttempts(null, null, "Назаров Павел Валерьевич", 100);
        Equal(2, nazarov.TotalCount);
        True(nazarov.Items.All(item => item.DisplayPersonName == "Назаров Павел Валерьевич"),
            "Сокращение не отображается как полное ФИО");
        Equal(2, database.SearchAttempts(null, null, "$Назаров Павел Вале", 100).TotalCount);
        Equal(1, database.SearchAttempts(null, null, "Офицеров Илья Александрович", 100).TotalCount);

        var exportResult = new DatabaseSelectionExportService().Export(
            database,
            new DatabaseSelectionExportRequest(
                output,
                new DateTime(2026, 8, 1),
                new DateTime(2026, 8, 31),
                ["Г-43-107"],
                [new DatabasePersonKey("Г-43-107", "Назаров Павел Валерьевич")],
                [
                    DatabaseExportColumn.Date,
                    DatabaseExportColumn.Scenario,
                    DatabaseExportColumn.Person,
                    DatabaseExportColumn.Percent
                ]));
        Equal(2, exportResult.Attempts);
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("Г-43-107");
        Equal("Назаров Павел Валерьевич", sheet.Cell("D15").GetString());
        Equal("Назаров Павел Валерьевич", sheet.Cell("D16").GetString());
    });

    Run("Полные ФИО на рабочей базе", () =>
    {
        const string liveRoot = @"C:\Users\MSI\Desktop\АП";
        var liveDatabase = Path.Combine(liveRoot, "data", "analizator.db");
        var liveConfig = Path.Combine(liveRoot, "config");
        if (!File.Exists(liveDatabase) || !Directory.Exists(liveConfig))
            return;

        var databaseCopy = Path.Combine(root, "live-person-names", "analizator.db");
        Directory.CreateDirectory(Path.GetDirectoryName(databaseCopy)!);
        using (var source = new SqliteConnection(new SqliteConnectionStringBuilder
               {
                   DataSource = liveDatabase,
                   Mode = SqliteOpenMode.ReadOnly
               }.ToString()))
        using (var destination = new SqliteConnection($"Data Source={databaseCopy}"))
        {
            source.Open();
            destination.Open();
            source.BackupDatabase(destination);
        }

        var configuration = ConfigurationService.Load(liveConfig);
        var database = new TrainingDatabaseService(
            databaseCopy, configuration.FioKeys, configuration.Installations);
        foreach (var fullName in new[]
                 {
                     "Офицеров Илья Александрович",
                     "Назаров Вадим Александрович"
                 })
        {
            var page = database.SearchAttempts(null, null, fullName, 10_000);
            True(page.TotalCount > 0, $"Рабочая база не нашла «{fullName}»");
            True(page.Items.Any(item => item.PersonName.StartsWith('$')),
                $"Для «{fullName}» в рабочей базе не найдено исходное сокращение");
            True(page.Items.All(item => item.DisplayPersonName == fullName),
                $"Для «{fullName}» рабочая база показывает: " + string.Join(" | ", page.Items
                    .Select(item => $"{item.PersonName} => {item.DisplayPersonName}")
                    .Distinct(StringComparer.Ordinal)));
        }
    });

    Run("Подробный Excel-отчёт фактической статистики", () =>
    {
        var source = Path.Combine(root, "Выгрузка_КТК_статистика_source.xlsx");
        var databasePath = Path.Combine(root, "statistics-report", "analizator.db");
        var output = Path.Combine(root, "statistics-report", "statistics.xlsx");
        WriteExport(source,
        [
            ("03.07.2026 08:15:00", "Сценарий А", "Иванов Иван Иванович", 70d),
            ("04.08.2026 09:20:00", "Сценарий Б", "Иванов Иван Иванович", 90d),
            ("05.08.2026 10:30:00", "Сценарий В", "Петров Пётр Петрович", 100d)
        ]);
        var installations = new Dictionary<string, List<string>>
        {
            ["ГФУ-2"] = ["ГФУ", "ГФУ 2"],
            ["ЭЛОУ-АВТ-6"] = ["ЭЛОУ", "АВТ 6"]
        };
        var database = new TrainingDatabaseService(databasePath);
        database.ImportFiles([source], installations);
        database.AddAttempt(new DatabaseAttemptEditModel
        {
            AttemptedAt = new DateTime(2026, 8, 6, 11, 40, 0),
            Installation = "ЭЛОУ-АВТ-6",
            PersonName = "Сидоров Сидор Сидорович",
            ScenarioName = "Сценарий Г",
            Percent = 85d,
            Mode = "Экзамен",
            Duration = "00:04:10",
            Actions = 7
        });
        var result = new StatisticsReportService().Run(new StatisticsReportRequest(
            databasePath,
            output,
            new DateTime(2026, 7, 1),
            new DateTime(2026, 8, 31),
            ["ГФУ-2", "ЭЛОУ-АВТ-6"],
            true,
            new PassingThresholdSettings { GlobalPercent = 80d }));

        Equal(4, result.Attempts);
        Equal(3, result.People);
        Equal(2, result.Installations);
        Equal(4, result.Scenarios);
        Equal(3, result.SuccessfulAttempts);
        True(File.Exists(output), "Файл статистики не создан");

        using (var workbook = new XLWorkbook(output))
        {
            string[] expectedSheets =
            [
                "Сводка", "Сотрудники", "Установки", "Сценарии", "Динамика",
                "Режимы", "Качество данных", "Исходные данные", "Методика"
            ];
            Equal(expectedSheets.Length, workbook.Worksheets.Count);
            True(expectedSheets.All(name => workbook.TryGetWorksheet(name, out _)),
                "В статистическом отчёте отсутствует обязательный лист");
            var summary = workbook.Worksheet("Сводка");
            Equal("Подробная статистика прохождения КТК", summary.Cell("A2").GetString());
            Equal(4, summary.Cell("A7").GetValue<int>());
            Equal(3, summary.Cell("B7").GetValue<int>());
            True(summary.Cell("F12").Style.Fill.BackgroundColor.Equals(
                    XLColor.FromHtml("#FCE4D6")),
                "Низкая успешность на сводке не выделена цветом");
            True(summary.Cell("F13").Style.Fill.BackgroundColor.Equals(
                    XLColor.FromHtml("#E2F0D9")),
                "Высокая успешность на сводке не выделена цветом");
            Equal(0, summary.ConditionalFormats.Count());
            var peopleSheet = workbook.Worksheet("Сотрудники");
            Equal("Иванов Иван Иванович", peopleSheet.Cell("B7").GetString());
            Equal(0.5d, peopleSheet.Cell("I7").GetDouble());
            var dynamics = workbook.Worksheet("Динамика");
            Equal("июль 2026", dynamics.Cell("A7").GetString());
            Equal(1, dynamics.Cell("B7").GetValue<int>());
            Equal(3, dynamics.Cell("B8").GetValue<int>());
            var sourceSheet = workbook.Worksheet("Исходные данные");
            Equal(XLDataType.DateTime, sourceSheet.Cell("D7").DataType);
            Equal(XLDataType.TimeSpan, sourceSheet.Cell("E7").DataType);
            True(sourceSheet.Tables.Single().ShowAutoFilter,
                "На исходных данных не включён фильтр");
        }

        using var document = SpreadsheetDocument.Open(output, false);
        var chartCount = document.WorkbookPart!.WorksheetParts
            .Where(part => part.DrawingsPart is not null)
            .Sum(part => part.DrawingsPart!.ChartParts.Count());
        Equal(2, chartCount);
    });

    Run("Мониторинг КТК и заполнение квартальной формы", () =>
    {
        var directory = Path.Combine(root, "monitoring-report");
        Directory.CreateDirectory(directory);
        var databasePath = Path.Combine(directory, "analizator.db");
        var reportPath = Path.Combine(directory, "monitoring.xlsx");
        var templatePath = Path.Combine(directory, "form-template.xlsx");
        var formPath = Path.Combine(directory, "form-filled.xlsx");
        var database = new TrainingDatabaseService(databasePath);

        void Add(DateTime date, string installation, string person, string scenario) =>
            database.AddAttempt(new DatabaseAttemptEditModel
            {
                AttemptedAt = date,
                Installation = installation,
                PersonName = person,
                ScenarioName = scenario,
                Percent = 100d,
                Mode = "Экзамен"
            });

        Add(new DateTime(2026, 1, 5), "Объект-1", "Иванов И.И.", "Сценарий А");
        Add(new DateTime(2026, 1, 6), "Объект-1", "Иванов И.И.", "Сценарий А");
        Add(new DateTime(2026, 1, 7), "Объект-1", "Петров П.П.", "Сценарий А");
        Add(new DateTime(2026, 2, 8), "Объект-2", "Иванов И.И.", "Сценарий А");
        Add(new DateTime(2026, 3, 9), "Объект-1", "Иванов И.И.", "Сценарий Б");
        Add(new DateTime(2026, 4, 10), "Объект-1", "Иванов И.И.", "Сценарий А");
        Add(new DateTime(2026, 5, 11), "Объект-1", "Иванов И.И.", "Сценарий А");
        Add(new DateTime(2026, 6, 12), "Объект-1", "Петров П.П.", "Сценарий А");

        using (var template = new XLWorkbook())
        {
            template.AddWorksheet("Инструкция").Cell("A1").Value = "Тест";
            var form = template.AddWorksheet("Форма сбора");
            form.Cell("C6").Value = "Наименование показателя";
            form.Cell("I6").Value = "Q1";
            form.Cell("J6").Value = "Q2";
            form.Cell("K6").Value = "Q3";
            form.Cell("L6").Value = "Q4";
            form.Cell("C19").Value = "Кол-во прохождения сценариев";
            form.Cell("I19").Value = 999;
            form.Cell("J19").Value = 999;
            form.Cell("I20").FormulaA1 = "I19/2";
            template.AddWorksheet("Список ДО").Cell("A1").Value = "Объекты";
            template.SaveAs(templatePath);
        }

        var result = new MonitoringReportService().Run(new MonitoringReportRequest(
            databasePath,
            reportPath,
            templatePath,
            formPath,
            new DateTime(2026, 1, 1),
            new DateTime(2026, 6, 30),
            ["Объект-1", "Объект-2"],
            true));

        Equal(8, result.Attempts);
        Equal(4, result.IndividualScenarios);
        Equal(4, result.Quarters.Single(x => x.Quarter == 1).IndividualScenarios);
        Equal(2, result.Quarters.Single(x => x.Quarter == 2).IndividualScenarios);
        True(result.Quarters.Where(x => x.Quarter <= 2).All(x => x.WrittenToForm),
            "Полные кварталы не отмечены для записи в форму");

        using (var report = new XLWorkbook(reportPath))
        {
            string[] expected = ["Статистика", "Сотрудники", "Детализация", "Диагностика"];
            True(expected.All(name => report.TryGetWorksheet(name, out _)),
                "В мониторинге отсутствует обязательный лист");
            Equal(8, report.Worksheet("Детализация").Tables.Single().DataRange!.RowCount());
        }
        using (var form = new XLWorkbook(formPath))
        {
            var sheet = form.Worksheet("Форма сбора");
            Equal(4, sheet.Cell("I19").GetValue<int>());
            Equal(2, sheet.Cell("J19").GetValue<int>());
            True(sheet.Cell("K19").IsEmpty(), "Неполный Q3 не должен заполняться");
            Equal("I19/2", sheet.Cell("I20").FormulaA1);
        }
    });

    Run("Контроль фотографий в PDF-отчётах", () =>
    {
        var source = Path.Combine(root, "pdf-photo", "Отчёты");
        var withPhoto = Path.Combine(source, "Установка А", "2026_08_03 08-15-00",
            "Протокол (Иванов Иван Иванович).pdf");
        var withoutPhoto = Path.Combine(source, "Установка Б", "2026_08_04 09-30-00",
            "Протокол (Петров Пётр Петрович).pdf");
        WritePdfReport(withPhoto, "Иванов Иван Иванович", "Останов оборудования", true);
        WritePdfReport(withoutPhoto, "Петров Пётр Петрович", "Пуск оборудования", false);

        var output = Path.Combine(root, "pdf-photo", "Контроль фотографий.xlsx");
        var result = new PdfPhotoReportService().Run(new PdfPhotoReportRequest(source, output));

        Equal(2, result.PdfFiles);
        Equal(1, result.WithPhoto);
        Equal(1, result.WithoutPhoto);
        Equal(0, result.ReadErrors);
        Equal(2, result.Installations);
        using var workbook = new XLWorkbook(output);
        True(workbook.TryGetWorksheet("Сводка", out _), "Нет листа сводки по PDF");
        var details = workbook.Worksheet("Отчёты PDF");
        Equal("Установка А", details.Cell("B6").GetString());
        Equal("Иванов Иван Иванович", details.Cell("C6").GetString());
        Equal("Останов оборудования", details.Cell("E6").GetString());
        Equal("Есть", details.Cell("G6").GetString());
        Equal("Открыть PDF", details.Cell("L6").GetString());
        True(details.Cell("L6").HasHyperlink, "В итоговой строке нет ссылки на PDF");
        Equal("Установка Б", details.Cell("B7").GetString());
        Equal("Нет", details.Cell("G7").GetString());
    });

    Run("Изолированный контроль фотографий в PDF", () =>
    {
        var source = Path.Combine(root, "pdf-photo-worker", "Установка В", "2026_08_05 11-00-00");
        WritePdfReport(Path.Combine(source, "Протокол (Сидоров Сидор Сидорович).pdf"),
            "Сидоров Сидор Сидорович", "Штатный пуск", true);
        var output = Path.Combine(root, "pdf-photo-worker", "Результат.xlsx");
        var request = new PdfPhotoReportRequest(
            Path.Combine(root, "pdf-photo-worker"), output, null, null);
        var payload = Convert.ToBase64String(
            Encoding.UTF8.GetBytes(JsonSerializer.Serialize(request)));
        var executable = Path.Combine(AppContext.BaseDirectory, "Analizator.exe");
        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        startInfo.ArgumentList.Add("--pdf-photo-worker");
        startInfo.ArgumentList.Add(payload);
        startInfo.Environment["DOTNET_EnableDiagnostics_Debugger"] = "0";
        startInfo.Environment["DOTNET_MODIFIABLE_ASSEMBLIES"] = "none";

        using var process = new Process { StartInfo = startInfo };
        True(process.Start(), "Не удалось запустить изолированную проверку PDF");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("Изолированная проверка PDF не завершилась за 30 секунд");
        }
        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        Equal(0, process.ExitCode);
        True(string.IsNullOrWhiteSpace(stderr), "Ошибка изолированной проверки PDF: " + stderr);
        True(stdout.Contains("[PDF_PHOTO_RESULT]", StringComparison.Ordinal),
            "Изолированная проверка PDF не вернула результат");
        True(stdout.Contains("\"WithPhoto\":1", StringComparison.Ordinal),
            "Изолированная проверка PDF не обнаружила фотографию");
        True(File.Exists(output), "Изолированный процесс не создал таблицу контроля PDF");
    });

    Run("Реальные архивная и текущая выгрузки", () =>
    {
        var dataRoot = @"C:\Users\MSI\Desktop\АП";
        var legacy = Path.Combine(dataRoot, "input", "Old", "Общие_январь.xlsx");
        var current = Path.Combine(dataRoot, "input", "Выгрузка_КТК_МНПЗ_август_2026.xlsx");
        if (!File.Exists(legacy) || !File.Exists(current))
            return;

        var installations = ConfigurationService.Load(Path.Combine(dataRoot, "config")).Installations;
        var legacyParsed = TrainingExportParser.Parse(legacy, installations);
        var currentParsed = TrainingExportParser.Parse(current, installations);
        Equal(TrainingExportFormat.Legacy, legacyParsed.Format);
        True(currentParsed.Format is TrainingExportFormat.Current or TrainingExportFormat.Mixed,
            "Текущая выгрузка ошибочно распознана только как старый формат");
        True(legacyParsed.Attempts.Count > 100, "Архивная выгрузка не распознана");
        True(currentParsed.Attempts.Count > 100, "Текущая выгрузка не распознана");
        True(legacyParsed.PeriodStart?.Month == 1, "Неверный период архивной выгрузки");
        True(currentParsed.PeriodStart?.Month == 8, "Неверный период текущей выгрузки");

        var realDatabase = new TrainingDatabaseService(
            Path.Combine(root, "real-export-database", "analizator.db"));
        var imported = realDatabase.ImportFiles([legacy, current], installations);
        Equal(2, imported.ImportedFiles);
        Equal(legacyParsed.Attempts.Count + currentParsed.Attempts.Count,
            realDatabase.GetSummary().Attempts);

        foreach (var month in new[] { "март", "апрель" })
        {
            var path = Path.Combine(dataRoot, "input", "Old", $"Выгрузка_КТК_МНПЗ_{month}_2026.xlsx");
            if (!File.Exists(path))
                continue;
            var sourceHash = Hash(path);
            var parsed = TrainingExportParser.Parse(path, installations);
            True(parsed.Attempts.Count > 100, $"Не распознана реальная выгрузка за {month}");
            Equal(sourceHash, Hash(path));
        }
    });

    Run("Приоритет и режим основной JSON-записи", () =>
    {
        var directory = Path.Combine(root, "json-priority");
        Directory.CreateDirectory(directory);
        var excel = Path.Combine(directory, "Выгрузка_КТК_август_2026.xlsx");
        var json = Path.Combine(directory, "КТК Лаунчер.json");
        WriteExport(excel,
        [
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 80d)
        ]);
        File.WriteAllText(json, JsonSerializer.Serialize(new[]
        {
            new
            {
                Guid = "json-session",
                MachineName = "АРМ-1",
                PlantName = "ГФУ-2",
                StudentSerial = "student-1",
                StudentName = "Иванов Иван Иванович",
                EntryDate = "2026-08-03T00:00:00",
                StartDate = "2026-08-03T00:00:00",
                EndDate = "2026-08-03T00:10:00",
                IsClosed = true,
                TrainingTimes = new[]
                {
                    new
                    {
                        ProgramTitle = "АРМ Инструктора",
                        MachineName = "АРМ-1",
                        Mode = 1,
                        Exercises = "Сценарий А",
                        StartDate = "2026-08-03T00:00:00",
                        EndDate = "2026-08-03T00:10:00"
                    }
                },
                ExerciseLogs = Array.Empty<object>(),
                ExerciseReports = new[]
                {
                    new
                    {
                        PlantName = "ГФУ-2",
                        StudentName = "Иванов Иван Иванович",
                        ExerciseName = "Сценарий А",
                        Date = "2026-08-03T00:00:00",
                        Percent = 80d
                    }
                }
            }
        }));

        var installations = new Dictionary<string, List<string>>
        {
            ["ГФУ-2"] = ["ГФУ-2"]
        };
        var databasePath = Path.Combine(directory, "analizator.db");
        var database = new TrainingDatabaseService(databasePath);
        var imported = database.ImportFiles([excel, json], installations);
        Equal(2, imported.ImportedFiles);
        Equal(1, database.GetDuplicateSummary().ExactDuplicates);

        var rows = database.SearchAttempts(null, null, null, 100).Items;
        var primary = rows.Single(item => !item.DuplicateOfAttemptId.HasValue);
        var duplicate = rows.Single(item => item.DuplicateOfAttemptId.HasValue);
        Equal("КТК Лаунчер.json", primary.SourceFile);
        Equal("Ежемесячный тренинг", primary.Mode);
        Equal("Выгрузка_КТК_август_2026.xlsx", duplicate.SourceFile);
        Equal(primary.Id, duplicate.DuplicateOfAttemptId);

        using (var connection = new SqliteConnection($"Data Source={databasePath}"))
        {
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE attempts SET mode = NULL WHERE id = $id; PRAGMA user_version = 5;";
            command.Parameters.AddWithValue("$id", primary.Id);
            command.ExecuteNonQuery();
        }
        database = new TrainingDatabaseService(databasePath);
        primary = database.SearchAttempts(null, null, null, 100).Items
            .Single(item => item.Id == primary.Id);
        Equal("Ежемесячный тренинг", primary.Mode);
    });

    Run("JSON-выгрузки с одной и несколькими установками", () =>
    {
        var dataRoot = @"C:\Users\MSI\Desktop\АП";
        var singlePlant = Path.Combine(dataRoot, "input", "Old",
            "АСБиКТ-oper 2026_08 КТК Лаунчер.json");
        var multiplePlants = Path.Combine(dataRoot, "input", "Old",
            "БРиДТ-РХиПНО-СНМиВГ-oper 2026_08 КТК Лаунчер.json");
        if (!File.Exists(singlePlant) || !File.Exists(multiplePlants))
            return;

        var installations = ConfigurationService.Load(Path.Combine(dataRoot, "config")).Installations;
        var database = new TrainingDatabaseService(Path.Combine(root, "json-database", "analizator.db"));
        var imported = database.ImportFiles([singlePlant, multiplePlants], installations);
        Equal(2, imported.ImportedFiles);
        Equal(1123, imported.ImportedAttempts);
        Equal(0, imported.FailedFiles);
        Equal(1123, database.GetSummary().Attempts);
        True(database.GetSummary().Installations >= 4,
            "JSON с несколькими установками загружен как одна установка");
        True(database.GetImportHistory().All(item => item.FormatDisplay == "JSON"),
            "JSON-источник неверно указан в истории");

        var asb = database.SearchAttempts(null, null, "АСБиКТ", 10_000);
        Equal(183, asb.TotalCount);
        var metadata = asb.Items.First(item => !string.IsNullOrWhiteSpace(item.SessionGuid));
        True(!string.IsNullOrWhiteSpace(metadata.MachineName), "Не сохранён MachineName");
        True(!string.IsNullOrWhiteSpace(metadata.StudentSerial), "Не сохранён StudentSerial");
        True(metadata.EntryDate.HasValue, "Не сохранён EntryDate");
        True(metadata.SessionStartDate.HasValue, "Не сохранён StartDate сеанса");
        True(metadata.Result.HasValue, "Не сохранён Result");
        True(metadata.IsValid.HasValue, "Не сохранён IsValid");
        True(metadata.ReportType.HasValue, "Не сохранён ReportType");
        True(metadata.TrainingTimesCount.HasValue, "Не сохранён TrainingTimes");
        True(asb.Items.Any(item => TextNormalization.KnownMode(item.Mode) is not null),
            "Режим из TrainingTimes.Mode не сохранён");

        var repeated = database.ImportFiles([singlePlant], installations);
        Equal(1, repeated.SkippedFiles);
        Equal(1123, database.GetSummary().Attempts);

        var excel = Path.Combine(dataRoot, "input", "Выгрузка_КТК_МНПЗ_август_2026.xlsx");
        if (File.Exists(excel))
        {
            var excelImport = database.ImportFiles([excel], installations);
            Equal(1, excelImport.ImportedFiles);
            var duplicates = database.GetDuplicateSummary();
            True(duplicates.Total > 0,
                "Не обнаружены повторы между исходными JSON и сформированной Excel-выгрузкой");
            Console.WriteLine($"INFO: JSON/Excel повторы — точных {duplicates.ExactDuplicates}, " +
                              $"возможных {duplicates.PossibleDuplicates}");
        }
    });

    Run("Сквозная обработка месяца из базы данных", () =>
    {
        var analysis = Path.Combine(root, "Анализ_database_source.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_database_source.xlsx");
        var output = Path.Combine(root, "database-source-output");
        var config = Path.Combine(root, "database-source-config");
        var databasePath = Path.Combine(root, "database-source", "analizator.db");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export,
        [
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 80d),
            ("04.08.2026", "Сценарий Б", "Иванов Иван Иванович", 95d)
        ]);
        var installations = new Dictionary<string, List<string>>
        {
            ["ГФУ-2"] = ["ГФУ", "ГФУ 2"]
        };
        new TrainingDatabaseService(databasePath).ImportFiles([export], installations);

        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis,
            "",
            output,
            config,
            false,
            false,
            AnalysisDataSourceKind.Database,
            databasePath,
            new DateTime(2026, 8, 1)));
        Equal(2, result.Reports);
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        Equal(95d, sheet.Cell("C3").GetDouble());
        Equal(80d, sheet.Cell("D3").GetDouble());
    });

    Run("Сквозная обработка XLSX", () =>
    {
        var analysis = Path.Combine(root, "Анализ_прохождения 2026_08.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_2026.xlsx");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export,
        [
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 80d),
            ("04.08.2026", "Сценарий А", "Иванов Иван Иванович", 95d),
            ("05.08.2026", "Сценарий Б", "Иванов Иван Иванович", 75d)
        ]);
        var sourceHash = Hash(analysis);
        var service = new AnalyzerService();
        var progressUpdates = new List<AnalyzerProgress>();
        var result = service.Run(new AnalyzerRequest(
                analysis, export, Path.Combine(root, "output"), Path.Combine(root, "config"), false, false),
            new InlineProgress<AnalyzerProgress>(progressUpdates.Add));
        Equal(sourceHash, Hash(analysis));
        Equal(3, result.Reports);
        Equal(1, result.MatchedPeople);
        True(progressUpdates.Any(item => item.Percent is >= 83 and <= 93),
            "Нет промежуточного прогресса при формировании служебных листов");
        int firstConditionalFormats;
        using (var workbook = new XLWorkbook(result.OutputFile))
        {
            var sheet = workbook.Worksheet("ГФУ-2");
            Equal(95d, sheet.Cell("C3").GetDouble());
            Equal(75d, sheet.Cell("D3").GetDouble());
            Equal("ДА", sheet.Cell("H3").GetString());
            True(workbook.TryGetWorksheet("Новые сотрудники", out _), "Нет листа новых сотрудников");
            True(workbook.TryGetWorksheet("Отчёт обработки", out _), "Нет отчёта обработки");
            firstConditionalFormats = sheet.ConditionalFormats.Count();
        }
        var second = service.Run(new AnalyzerRequest(
            result.OutputFile, export, Path.Combine(root, "second-output"), Path.Combine(root, "config"), false, false));
        using var secondWorkbook = new XLWorkbook(second.OutputFile);
        var secondSheet = secondWorkbook.Worksheet("ГФУ-2");
        Equal(firstConditionalFormats, secondSheet.ConditionalFormats.Count());
        Equal(95d, secondSheet.Cell("C3").GetDouble());
    });

    Run("Создание отсутствующего месяца", () =>
    {
        var analysis = Path.Combine(root, "Анализ_июль.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август.xlsx");
        WriteAnalysis(analysis, "Июль", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 90d)]);
        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "month-output"), Path.Combine(root, "config"), false, false));
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        var blocks = AnalysisWorkbookProcessor.FindMonthBlocks(sheet);
        True(blocks.Any(block => block.Month == 7), "Пропал исходный месяц");
        True(blocks.Any(block => block.Month == 8), "Новый месяц не создан");
        var august = blocks.First(block => block.Month == 8);
        Equal(90d, sheet.Cell(3, august.ScenarioColumns[1]).GetDouble());
    });

    Run("Автодобавление сотрудника", () =>
    {
        var analysis = Path.Combine(root, "Анализ_add.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_add.xlsx");
        WriteAnalysis(analysis, "Август", ["Алексеев Алексей Алексеевич", "Яковлев Яков Яковлевич"]);
        WriteExport(export,
        [
            ("03.08.2026", "Сценарий А", "Алексеев Алексей Алексеевич", 90d),
            ("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 91d),
            ("03.08.2026", "Сценарий Б", "Иванов Иван Иванович", 82d),
            ("03.08.2026", "Сценарий А", "Яковлев Яков Яковлевич", 80d)
        ]);
        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "add-output"), Path.Combine(root, "config"), false, true));
        Equal(1, result.NewPeople);
        Equal(1, result.AddedPeople);
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        Equal("Иванов Иван Иванович", sheet.Cell("A4").GetString());
        Equal(91d, sheet.Cell("C4").GetDouble());
        Equal(82d, sheet.Cell("D4").GetDouble());
    });

    Run("Единый порог прохождения", () =>
    {
        var analysis = Path.Combine(root, "Анализ_threshold_global.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_threshold_global.xlsx");
        var config = Path.Combine(root, "threshold-global-config");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 60d)]);
        WriteThresholdSettings(config, false, 75d, new Dictionary<string, double>());

        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "threshold-global-output"), config, false, false));
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        Equal(XLColor.FromHtml("#FFF2CC").Color.ToArgb() & 0xFFFFFF,
            sheet.Cell("C3").Style.Fill.BackgroundColor.Color.ToArgb() & 0xFFFFFF);
        True(sheet.Cell("C4").FormulaA1.Contains(">=75", StringComparison.Ordinal),
            "Итоговая формула не использует единый порог 75%");
    });

    Run("Индивидуальный порог установки", () =>
    {
        var analysis = Path.Combine(root, "Анализ_threshold_installation.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_threshold_installation.xlsx");
        var config = Path.Combine(root, "threshold-installation-config");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 60d)]);
        WriteThresholdSettings(config, true, 90d,
            new Dictionary<string, double> { ["ГФУ-2"] = 50d });

        var result = new AnalyzerService().Run(new AnalyzerRequest(
            analysis, export, Path.Combine(root, "threshold-installation-output"), config, false, false));
        using var workbook = new XLWorkbook(result.OutputFile);
        var sheet = workbook.Worksheet("ГФУ-2");
        True(sheet.Cell("C4").FormulaA1.Contains(">=50", StringComparison.Ordinal),
            "Итоговая формула не использует индивидуальный порог 50%");
        True(!sheet.Cell("C4").FormulaA1.Contains(">=90", StringComparison.Ordinal),
            "Для установки ошибочно применён общий порог");
    });

    Run("Понятная ошибка занятого файла результата", () =>
    {
        var analysis = Path.Combine(root, "Анализ_locked.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_locked.xlsx");
        var output = Path.Combine(root, "locked-output");
        Directory.CreateDirectory(output);
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 90d)]);
        var resultFile = Path.Combine(output, "Анализ_прохождения_результат.xlsx");
        File.WriteAllBytes(resultFile, [0]);

        using var lockStream = new FileStream(
            resultFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        try
        {
            new AnalyzerService().Run(new AnalyzerRequest(
                analysis, export, output, Path.Combine(root, "config"), false, false));
            throw new InvalidOperationException("Ожидалась ошибка занятого файла");
        }
        catch (AnalysisFileAccessException exception)
        {
            True(exception.Message.Contains("Закройте файл", StringComparison.Ordinal),
                "Ошибка не содержит понятной инструкции пользователю");
            Equal(resultFile, exception.FilePath);
        }
    });

    Run("Изолированный процесс анализа", () =>
    {
        var analysis = Path.Combine(root, "Анализ_worker.xlsx");
        var export = Path.Combine(root, "Выгрузка_КТК_август_worker.xlsx");
        var output = Path.Combine(root, "worker-output");
        WriteAnalysis(analysis, "Август", ["Иванов Иван Иванович"]);
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 90d)]);

        var request = new AnalyzerRequest(
            analysis, export, output, Path.Combine(root, "config"), false, false);
        var payload = Convert.ToBase64String(
            Encoding.UTF8.GetBytes(JsonSerializer.Serialize(request)));
        var executable = Path.Combine(AppContext.BaseDirectory, "Analizator.exe");
        True(File.Exists(executable), "Не найден исполняемый файл изолированного процесса");

        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        startInfo.ArgumentList.Add("--analysis-worker");
        startInfo.ArgumentList.Add(payload);
        startInfo.Environment["DOTNET_EnableDiagnostics_Debugger"] = "0";
        startInfo.Environment["DOTNET_MODIFIABLE_ASSEMBLIES"] = "none";

        using var process = new Process { StartInfo = startInfo };
        True(process.Start(), "Не удалось запустить изолированный процесс");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("Изолированный процесс не завершился за 30 секунд");
        }

        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        Equal(0, process.ExitCode);
        True(string.IsNullOrWhiteSpace(stderr), "Ошибка изолированного процесса: " + stderr);
        True(stdout.Contains("\"Percent\":100", StringComparison.Ordinal),
            "Изолированный процесс не сообщил о 100% прогресса");
        var resultFile = Path.Combine(output, "Анализ_прохождения_результат.xlsx");
        True(File.Exists(resultFile),
            "Изолированный процесс не создал результат");

        using var lockStream = new FileStream(
            resultFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var failingProcess = new Process { StartInfo = startInfo };
        True(failingProcess.Start(), "Не удалось повторно запустить изолированный процесс");
        var failingStdoutTask = failingProcess.StandardOutput.ReadToEndAsync();
        var failingStderrTask = failingProcess.StandardError.ReadToEndAsync();
        if (!failingProcess.WaitForExit(30_000))
        {
            failingProcess.Kill(entireProcessTree: true);
            throw new TimeoutException("Проверка занятого файла не завершилась за 30 секунд");
        }

        var failingStdout = failingStdoutTask.GetAwaiter().GetResult();
        var failingStderr = failingStderrTask.GetAwaiter().GetResult();
        Equal(1, failingProcess.ExitCode);
        True(string.IsNullOrWhiteSpace(failingStderr),
            "Техническая ошибка попала в поток диагностики: " + failingStderr);
        var errorLine = failingStdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
            .SingleOrDefault(line => line.StartsWith("[ANALIZATOR_ERROR]", StringComparison.Ordinal));
        True(errorLine is not null, "Изолированный процесс не передал понятную ошибку интерфейсу");
        var message = JsonSerializer.Deserialize<string>(errorLine!["[ANALIZATOR_ERROR]".Length..]);
        True(message?.Contains("Закройте файл", StringComparison.Ordinal) == true,
            "Ошибка изолированного процесса не содержит инструкции закрыть файл");
        True(!failingStdout.Contains("System.IO.IOException", StringComparison.Ordinal),
            "Пользователю передаётся технический стек вызовов");
    });

    Run("Изолированный импорт базы данных", () =>
    {
        var export = Path.Combine(root, "Выгрузка_КТК_август_database_worker.xlsx");
        var databasePath = Path.Combine(root, "database-worker", "analizator.db");
        WriteExport(export, [("03.08.2026", "Сценарий А", "Иванов Иван Иванович", 90d)]);
        var request = new
        {
            DatabasePath = databasePath,
            Files = new[] { export },
            Installations = new Dictionary<string, List<string>>
            {
                ["ГФУ-2"] = ["ГФУ", "ГФУ 2"]
            }
        };
        var payload = Convert.ToBase64String(
            Encoding.UTF8.GetBytes(JsonSerializer.Serialize(request)));
        var executable = Path.Combine(AppContext.BaseDirectory, "Analizator.exe");
        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        startInfo.ArgumentList.Add("--database-import-worker");
        startInfo.ArgumentList.Add(payload);
        startInfo.Environment["DOTNET_EnableDiagnostics_Debugger"] = "0";
        startInfo.Environment["DOTNET_MODIFIABLE_ASSEMBLIES"] = "none";

        using var process = new Process { StartInfo = startInfo };
        True(process.Start(), "Не удалось запустить изолированный импорт");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("Изолированный импорт не завершился за 30 секунд");
        }

        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        Equal(0, process.ExitCode);
        True(string.IsNullOrWhiteSpace(stderr), "Ошибка изолированного импорта: " + stderr);
        True(stdout.Contains("[DATABASE_RESULT]", StringComparison.Ordinal),
            "Изолированный импорт не вернул результат");
        True(stdout.Contains("\"Percent\":100", StringComparison.Ordinal),
            "Изолированный импорт не сообщил о завершении");
        Equal(1, new TrainingDatabaseService(databasePath).GetSummary().Attempts);
    });

    Run("Изолированное формирование статистики", () =>
    {
        var databasePath = Path.Combine(root, "statistics-worker", "analizator.db");
        var output = Path.Combine(root, "statistics-worker", "statistics.xlsx");
        var database = new TrainingDatabaseService(databasePath);
        database.AddAttempt(new DatabaseAttemptEditModel
        {
            AttemptedAt = new DateTime(2026, 8, 3, 8, 15, 0),
            Installation = "ГФУ-2",
            PersonName = "Иванов Иван Иванович",
            ScenarioName = "Сценарий А",
            Percent = 90d,
            Mode = "Экзамен"
        });
        var request = new StatisticsReportRequest(
            databasePath,
            output,
            new DateTime(2026, 8, 1),
            new DateTime(2026, 8, 31),
            ["ГФУ-2"],
            true,
            new PassingThresholdSettings { GlobalPercent = 80d });
        var payload = Convert.ToBase64String(
            Encoding.UTF8.GetBytes(JsonSerializer.Serialize(request)));
        var executable = Path.Combine(AppContext.BaseDirectory, "Analizator.exe");
        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        startInfo.ArgumentList.Add("--statistics-worker");
        startInfo.ArgumentList.Add(payload);
        startInfo.Environment["DOTNET_EnableDiagnostics_Debugger"] = "0";
        startInfo.Environment["DOTNET_MODIFIABLE_ASSEMBLIES"] = "none";

        using var process = new Process { StartInfo = startInfo };
        True(process.Start(), "Не удалось запустить изолированное формирование статистики");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(30_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException(
                "Изолированное формирование статистики не завершилось за 30 секунд");
        }

        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        Equal(0, process.ExitCode);
        True(string.IsNullOrWhiteSpace(stderr),
            "Ошибка изолированного формирования статистики: " + stderr);
        True(stdout.Contains("[STATISTICS_RESULT]", StringComparison.Ordinal),
            "Изолированное формирование статистики не вернуло результат");
        True(stdout.Contains("\"Percent\":100", StringComparison.Ordinal),
            "Изолированное формирование статистики не сообщило о завершении");
        True(File.Exists(output), "Изолированный процесс не создал статистику");
    });
}
finally
{
    try { Directory.Delete(root, true); } catch { }
}

if (failures.Count > 0)
{
    Console.Error.WriteLine($"FAILED ({failures.Count})");
    foreach (var failure in failures)
        Console.Error.WriteLine(failure);
    return 1;
}
Console.WriteLine("OK: все C#-тесты пройдены");
return 0;

void Run(string name, Action test)
{
    try
    {
        test();
        Console.WriteLine("PASS: " + name);
    }
    catch (Exception exception)
    {
        failures.Add($"{name}: {exception.Message}\n{exception}");
        Console.WriteLine("FAIL: " + name);
    }
}

static void WriteAnalysis(string path, string month, IReadOnlyList<string> people)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.AddWorksheet("ГФУ-2");
    sheet.Cell("C1").Value = month;
    sheet.Range("C1:H1").Merge();
    string[] headers = ["ФИО", "Должность", "1", "2", "3", "4", "5", "Посещение КТК"];
    for (var column = 1; column <= headers.Length; column++)
    {
        sheet.Cell(2, column).Value = headers[column - 1];
        sheet.Cell(2, column).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAF7");
    }
    var row = 3;
    foreach (var person in people)
    {
        sheet.Cell(row, 1).Value = person;
        sheet.Cell(row, 2).Value = "Оператор";
        for (var column = 3; column <= 7; column++)
            sheet.Cell(row, column).Value = 0;
        sheet.Cell(row, 8).Value = "Нет";
        row++;
    }
    sheet.Cell(row, 1).Value = "Итого";
    sheet.Cell(row, 3).FormulaA1 = $"COUNT(C3:C{row - 1})";
    sheet.SheetView.FreezeRows(2);
    sheet.SheetView.FreezeColumns(2);
    sheet.Range(2, 1, row - 1, 8).SetAutoFilter();
    workbook.SaveAs(path);
}

static void WriteExport(string path, IReadOnlyList<(string Date, string Scenario, string Person, double Percent)> rows)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.AddWorksheet("ГФУ-2");
    sheet.Cell(1, 1).Value = "Тестовая выгрузка";
    string[] headers = ["Дата", "Тема", "Ф.И.О. по отчёту", "Процент", "Режим"];
    for (var column = 1; column <= headers.Length; column++)
        sheet.Cell(2, column).Value = headers[column - 1];
    var row = 3;
    foreach (var item in rows)
    {
        sheet.Cell(row, 1).Value = item.Date;
        sheet.Cell(row, 2).Value = item.Scenario;
        sheet.Cell(row, 3).Value = item.Person;
        sheet.Cell(row, 4).Value = item.Percent;
        sheet.Cell(row, 5).Value = "Случайный выбор";
        row++;
    }
    workbook.SaveAs(path);
}

static void WriteLegacyExport(
    string path,
    IReadOnlyList<(string Date, string Scenario, string Person, double Percent)> rows)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.AddWorksheet("ГФУ-2");
    sheet.Cell(1, 1).Value = "Результаты тренинга";
    string[] headers =
        ["Дата", "Время", "Проект", "Тема", "Студент", "Роль", "Зачтено", "Действий", "Процент"];
    for (var column = 1; column <= headers.Length; column++)
        sheet.Cell(13, column).Value = headers[column - 1];
    var row = 14;
    foreach (var item in rows)
    {
        sheet.Cell(row, 1).Value = item.Date;
        sheet.Cell(row, 2).Value = "00:05:00";
        sheet.Cell(row, 3).Value = "GFU-2";
        sheet.Cell(row, 4).Value = item.Scenario;
        sheet.Cell(row, 5).Value = item.Person;
        sheet.Cell(row, 6).Value = "назначенный оператор";
        sheet.Cell(row, 7).Value = item.Percent >= 100 ? "Да" : "Нет";
        sheet.Cell(row, 8).Value = 10;
        sheet.Cell(row, 9).Value = item.Percent;
        row++;
    }
    workbook.SaveAs(path);
}

static void WritePdfReport(
    string path,
    string person,
    string scenario,
    bool includePhoto)
{
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    using var builder = new PdfDocumentBuilder();
    var fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
        "Fonts", "arial.ttf");
    var font = builder.AddTrueTypeFont(File.ReadAllBytes(fontPath));
    var page = builder.AddPage(PageSize.A4);
    page.AddText("ФИО: " + person, 12, new PdfPoint(48, 790), font);
    page.AddText("Сценарий: " + scenario, 12, new PdfPoint(48, 765), font);
    page.AddText("Режим: Экзамен", 12, new PdfPoint(48, 740), font);
    if (includePhoto)
    {
        const string photo = "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAIAAAD/gAIDAAAA6klEQVR4nO3SMREAIRAEQXhlyEIYAt/CTd4dbzS1+9y3mPmGO8RqPCsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArECsQKxArEGvN/ZsDAjSJr20rAAAAAElFTkSuQmCC";
        page.AddPng(Convert.FromBase64String(photo), new PdfRectangle(48, 500, 248, 700));
    }
    File.WriteAllBytes(path, builder.Build());
}

static void WriteThresholdSettings(
    string directory,
    bool usePerInstallation,
    double global,
    IReadOnlyDictionary<string, double> installations)
{
    Directory.CreateDirectory(directory);
    File.WriteAllText(Path.Combine(directory, "settings.json"), JsonSerializer.Serialize(new
    {
        passing_threshold = new
        {
            use_per_installation = usePerInstallation,
            global,
            installations
        }
    }));
}

static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

static void Equal<T>(T expected, T actual)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"Ожидалось: {expected}; получено: {actual}");
}

static void True(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

file sealed class InlineProgress<T>(Action<T> handler) : IProgress<T>
{
    public void Report(T value) => handler(value);
}

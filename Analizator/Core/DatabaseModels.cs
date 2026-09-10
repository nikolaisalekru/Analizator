namespace Analizator.Core;

public sealed record DatabaseImportProgress(int Percent, string Stage, string? FileName = null);

public sealed class DatabaseImportResult
{
    public int ImportedFiles { get; set; }
    public int UpdatedFiles { get; set; }
    public int SkippedFiles { get; set; }
    public int FailedFiles { get; set; }
    public int ImportedAttempts { get; set; }
    public int PossibleDuplicates { get; set; }
    public int RowErrors { get; set; }
    public List<DatabaseImportFileResult> Files { get; init; } = [];
}

public sealed record DatabaseImportFileResult(
    string FileName,
    string Status,
    int Attempts,
    int Errors,
    string Message);

public sealed record DatabaseSummary(
    int Attempts,
    int People,
    int SourceFiles,
    int Installations,
    int PossibleDuplicates,
    DateTime? PeriodStart,
    DateTime? PeriodEnd);

public sealed record DatabaseMonthSummary(
    int Year,
    int Month,
    int Attempts,
    int ExactDuplicates,
    int PossibleDuplicates,
    int ReportsWithoutMode)
{
    public DateTime FirstDay => new(Year, Month, 1);
    public DateTime LastDay => FirstDay.AddMonths(1).AddDays(-1);
    public string Display =>
        $"{TextNormalization.MonthNames.GetValueOrDefault(Month, Month.ToString())} {Year} · {Attempts:N0} записей";
}

public sealed record DatabaseUnmatchedInstallation(
    string Name,
    int Attempts);

public sealed record DatabasePersonSummary(
    string Installation,
    string PersonName,
    int Attempts)
{
    public DatabasePersonKey Key => new(Installation, PersonName);
}

public sealed record DatabasePersonKey(string Installation, string PersonName)
{
    public string NormalizedValue =>
        TextNormalization.NormalizeText(Installation) + "\u001F" +
        TextNormalization.NormalizePersonName(PersonName);
}

public enum DatabaseExportColumn
{
    Date,
    Scenario,
    Person,
    Duration,
    Actions,
    Percent,
    Mode,
    Project,
    Role,
    Credited,
    Source
}

public sealed record DatabaseSelectionExportRequest(
    string OutputFile,
    DateTime PeriodStart,
    DateTime PeriodEnd,
    IReadOnlyList<string> Installations,
    IReadOnlyList<DatabasePersonKey> People,
    IReadOnlyList<DatabaseExportColumn> Columns);

public sealed record DatabaseSelectionExportResult(
    string OutputFile,
    int Attempts,
    int People,
    int Installations);

public sealed class DatabaseSourceItem
{
    public long Id { get; init; }
    public required string FileName { get; init; }
    public required string OriginalPath { get; init; }
    public required string Format { get; init; }
    public DateTime ImportedAt { get; init; }
    public DateTime? PeriodStart { get; init; }
    public DateTime? PeriodEnd { get; init; }
    public int RowsCount { get; init; }
    public int ErrorCount { get; init; }

    public string FormatDisplay => Format switch
    {
        nameof(TrainingExportFormat.Legacy) => "Старый",
        nameof(TrainingExportFormat.Mixed) => "Смешанный",
        nameof(TrainingExportFormat.Json) => "JSON",
        "Manual" => "Ручной",
        _ => "Текущий"
    };

    public string PeriodDisplay => PeriodStart.HasValue && PeriodEnd.HasValue
        ? PeriodStart.Value.Date == PeriodEnd.Value.Date
            ? PeriodStart.Value.ToString("dd.MM.yyyy")
            : $"{PeriodStart.Value:dd.MM.yyyy} — {PeriodEnd.Value:dd.MM.yyyy}"
        : "—";

    public string ImportedAtDisplay => ImportedAt.ToString("dd.MM.yyyy HH:mm");
}

public sealed record DatabaseAttempt(
    long Id,
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
    string SourceFile,
    string SourceSheet,
    int SourceRow,
    bool PossibleDuplicate,
    string? SessionGuid = null,
    string? MachineName = null,
    string? PlantName = null,
    string? StudentSerial = null,
    DateTime? EntryDate = null,
    DateTime? SessionStartDate = null,
    DateTime? SessionEndDate = null,
    bool? IsClosed = null,
    string? FilePath = null,
    int? Resume = null,
    bool? Result = null,
    bool? IsValid = null,
    int? ReportType = null,
    bool? HasReportFile = null,
    string? ReportFile = null,
    string? TrainingTimesJson = null,
    string? ExerciseLogsJson = null,
    int? TrainingTimesCount = null,
    int? ExerciseLogsCount = null,
    long? DuplicateOfAttemptId = null,
    string? DuplicateKind = null,
    string? DuplicateReason = null,
    string? CanonicalPersonName = null)
{
    public string PossibleDuplicateDisplay => PossibleDuplicate
        ? string.IsNullOrWhiteSpace(DuplicateKind) ? "Возможный" : DuplicateKind
        : "—";
    public string AttemptDateDisplay => AttemptedAt.ToString("dd.MM.yyyy");
    public string AttemptTimeDisplay => AttemptedAt.ToString("HH:mm:ss");
    public string PlantNameDisplay => string.IsNullOrWhiteSpace(PlantName) ? Installation : PlantName;
    public string DisplayPersonName => string.IsNullOrWhiteSpace(CanonicalPersonName)
        ? PersonName
        : CanonicalPersonName;
    public string IsClosedDisplay => BooleanDisplay(IsClosed);
    public string ResultDisplay => BooleanDisplay(Result);
    public string IsValidDisplay => BooleanDisplay(IsValid);
    public string HasReportFileDisplay => BooleanDisplay(HasReportFile);
    public string TrainingTimesDisplay => CountDisplay(TrainingTimesCount);
    public string ExerciseLogsDisplay => CountDisplay(ExerciseLogsCount);

    private static string BooleanDisplay(bool? value) => value.HasValue ? value.Value ? "Да" : "Нет" : "—";
    private static string CountDisplay(int? value) => value.HasValue ? value.Value.ToString("N0") : "—";
}

public sealed record DatabaseAttemptPage(
    IReadOnlyList<DatabaseAttempt> Items,
    int TotalCount);

public sealed record DatabaseDuplicateSummary(
    int ExactDuplicates,
    int PossibleDuplicates)
{
    public int Total => ExactDuplicates + PossibleDuplicates;
}

public sealed class DatabaseAttemptEditModel
{
    public long? Id { get; init; }
    public DateTime AttemptedAt { get; set; } = DateTime.Now;
    public string Installation { get; set; } = "";
    public string PersonName { get; set; } = "";
    public string ScenarioName { get; set; } = "";
    public double Percent { get; set; }
    public string? Mode { get; set; }
    public string? Project { get; set; }
    public string? Role { get; set; }
    public string? Credited { get; set; }
    public string? Duration { get; set; }
    public int? Actions { get; set; }
    public string? SessionGuid { get; set; }
    public string? MachineName { get; set; }
    public string? PlantName { get; set; }
    public string? StudentSerial { get; set; }
    public DateTime? EntryDate { get; set; }
    public DateTime? SessionStartDate { get; set; }
    public DateTime? SessionEndDate { get; set; }
    public bool? IsClosed { get; set; }
    public string? FilePath { get; set; }
    public int? Resume { get; set; }
    public bool? Result { get; set; }
    public bool? IsValid { get; set; }
    public int? ReportType { get; set; }
    public bool? HasReportFile { get; set; }
    public string? ReportFile { get; set; }
    public string? TrainingTimesJson { get; set; }
    public string? ExerciseLogsJson { get; set; }
    public int? TrainingTimesCount { get; set; }
    public int? ExerciseLogsCount { get; set; }

    public static DatabaseAttemptEditModel FromAttempt(DatabaseAttempt attempt) => new()
    {
        Id = attempt.Id,
        AttemptedAt = attempt.AttemptedAt,
        Installation = attempt.Installation,
        PersonName = attempt.PersonName,
        ScenarioName = attempt.ScenarioName,
        Percent = attempt.Percent,
        Mode = attempt.Mode,
        Project = attempt.Project,
        Role = attempt.Role,
        Credited = attempt.Credited,
        Duration = attempt.Duration,
        Actions = attempt.Actions,
        SessionGuid = attempt.SessionGuid,
        MachineName = attempt.MachineName,
        PlantName = attempt.PlantName,
        StudentSerial = attempt.StudentSerial,
        EntryDate = attempt.EntryDate,
        SessionStartDate = attempt.SessionStartDate,
        SessionEndDate = attempt.SessionEndDate,
        IsClosed = attempt.IsClosed,
        FilePath = attempt.FilePath,
        Resume = attempt.Resume,
        Result = attempt.Result,
        IsValid = attempt.IsValid,
        ReportType = attempt.ReportType,
        HasReportFile = attempt.HasReportFile,
        ReportFile = attempt.ReportFile,
        TrainingTimesJson = attempt.TrainingTimesJson,
        ExerciseLogsJson = attempt.ExerciseLogsJson,
        TrainingTimesCount = attempt.TrainingTimesCount,
        ExerciseLogsCount = attempt.ExerciseLogsCount
    };
}

namespace Analizator.Core;

public sealed record MonitoringReportProgress(int Percent, string Stage);

public sealed record MonitoringReportRequest(
    string DatabasePath,
    string ReportOutputFile,
    string FormTemplateFile,
    string FormOutputFile,
    DateTime PeriodStart,
    DateTime PeriodEnd,
    string[] Installations,
    bool IncludePossibleDuplicates);

public sealed record MonitoringQuarterResult(
    int Year,
    int Quarter,
    int IndividualScenarios,
    bool WrittenToForm)
{
    public string DisplayName => $"Q{Quarter} {Year}";
}

public sealed class MonitoringReportResult
{
    public required string ReportOutputFile { get; init; }
    public required string FormOutputFile { get; init; }
    public int Attempts { get; init; }
    public int People { get; init; }
    public int Installations { get; init; }
    public int IndividualScenarios { get; init; }
    public int ExactDuplicatesExcluded { get; init; }
    public int PossibleDuplicatesIncluded { get; init; }
    public IReadOnlyList<MonitoringQuarterResult> Quarters { get; init; } = [];
}

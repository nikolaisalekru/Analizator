namespace Analizator.Core;

public sealed record StatisticsReportProgress(int Percent, string Stage);

public sealed record StatisticsReportRequest(
    string DatabasePath,
    string OutputFile,
    DateTime PeriodStart,
    DateTime PeriodEnd,
    string[] Installations,
    bool IncludePossibleDuplicates,
    PassingThresholdSettings PassingThreshold);

public sealed class StatisticsReportResult
{
    public required string OutputFile { get; init; }
    public int Attempts { get; init; }
    public int People { get; init; }
    public int Installations { get; init; }
    public int Scenarios { get; init; }
    public int SuccessfulAttempts { get; init; }
    public int ExactDuplicatesExcluded { get; init; }
    public int PossibleDuplicatesIncluded { get; init; }
}

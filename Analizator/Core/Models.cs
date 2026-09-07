using System.Collections.ObjectModel;

namespace Analizator.Core;

public sealed record AnalyzerRequest(
    string AnalysisFile,
    string ExportFile,
    string OutputDirectory,
    string ConfigurationDirectory,
    bool CreateBackup = true,
    bool AutoAddNewPeople = true);

public sealed record AnalyzerProgress(int Percent, string Stage, string? Detail = null);

public sealed class AnalyzerResult
{
    public required string OutputFile { get; init; }
    public int Reports { get; init; }
    public int MatchedPeople { get; init; }
    public int NewPeople { get; init; }
    public int AddedPeople { get; init; }
    public int AbsentPeople { get; init; }
    public int ExactMatches { get; init; }
    public int KeyMatches { get; init; }
    public int FuzzyMatches { get; init; }
    public IReadOnlyList<string> UnmatchedInstallations { get; init; } = [];
    public required string LogFile { get; init; }
}

public sealed class AnalyzerSettings
{
    public ModeFilterSettings ModeFilter { get; init; } = new();
    public AutoAddSettings AutoAddNewPeople { get; init; } = new();
    public PassingThresholdSettings PassingThreshold { get; init; } = new();
    public bool ParallelProcessing { get; init; } = true;
    public int MaxWorkers { get; init; } = 5;
}

public sealed class PassingThresholdSettings
{
    public const double DefaultPercent = 10d;

    public bool UsePerInstallation { get; init; }
    public double GlobalPercent { get; init; } = DefaultPercent;
    public Dictionary<string, double> Installations { get; init; } =
        new(StringComparer.OrdinalIgnoreCase);

    public double GetForInstallation(string installation) =>
        UsePerInstallation && Installations.TryGetValue(installation, out var threshold)
            ? threshold
            : GlobalPercent;
}

public sealed class ModeFilterSettings
{
    public bool Enabled { get; init; }
    public HashSet<string> Modes { get; init; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class AutoAddSettings
{
    public bool Default { get; init; } = true;
    public Dictionary<string, bool> Installations { get; init; } = new(StringComparer.OrdinalIgnoreCase);

    public bool IsEnabled(string installation) =>
        Installations.TryGetValue(installation, out var enabled) ? enabled : Default;
}

public sealed class ExportReadResult
{
    public Dictionary<string, Dictionary<string, Dictionary<string, ScenarioAggregate>>> Data { get; } =
        new(StringComparer.OrdinalIgnoreCase);
    public List<InstallationStatistic> InstallationStatistics { get; } = [];
    public Dictionary<string, int> ModeCounts { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, int> UnknownModes { get; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> SheetsWithMode { get; } = [];
    public List<string> SheetsWithoutMode { get; } = [];
    public int FilteredReports { get; set; }
    public int ReportsWithoutMode { get; set; }
    public int TotalReports { get; set; }
    public int? Month { get; set; }
}

public sealed class ScenarioAggregate
{
    public required string DisplayName { get; init; }
    public double BestPercent { get; set; }
    public int Attempts { get; set; }
}

public sealed record InstallationStatistic(
    string Sheet,
    string? Installation,
    int Reports,
    string Status,
    double MatchScore = 0,
    string MatchType = "");

public sealed record PersonRow(string Name, string Position, int RowNumber);

public sealed record MonthBlock(
    int Month,
    string MonthName,
    int HeaderRow,
    int StartColumn,
    int EndColumn,
    IReadOnlyDictionary<int, int> ScenarioColumns,
    int? AttendanceColumn);

public sealed class AnalysisSheet
{
    public required string Title { get; init; }
    public required string Installation { get; init; }
    public required ClosedXML.Excel.IXLWorksheet Worksheet { get; init; }
    public int HeaderRow { get; init; }
    public int FioColumn { get; init; }
    public int? PositionColumn { get; init; }
    public List<PersonRow> People { get; } = [];
    public List<MonthBlock> MonthBlocks { get; } = [];
}

public sealed record MatchResult(string? Name, double Score, MatchKind Kind)
{
    public bool IsMatch => Kind is MatchKind.Exact or MatchKind.Key or MatchKind.Fuzzy;
}

public enum MatchKind
{
    New,
    Exact,
    Key,
    Fuzzy
}

public sealed class NewPerson
{
    public required string Installation { get; init; }
    public required string Name { get; init; }
    public string Position { get; set; } = "";
    public List<ScenarioAggregate> Scenarios { get; init; } = [];
    public bool Visited => Scenarios.Count > 0;
    public bool AutoAdd { get; set; }
    public bool AddedToAnalysis { get; set; }
    public bool AddedToAttendanceChart { get; set; }
    public int? AttendanceChartRow { get; set; }
}

public sealed class ProcessingStatistics
{
    public int Reports { get; set; }
    public int MatchedPeople { get; set; }
    public int NewPeople { get; set; }
    public int AddedPeople { get; set; }
    public int AbsentPeople { get; set; }
    public int ExactMatches { get; set; }
    public int KeyMatches { get; set; }
    public int FuzzyMatches { get; set; }
    public List<(string Installation, string ExportName, string AnalysisName, double Score)> FuzzyDetails { get; } = [];
    public List<(string Installation, string ExportName, string AnalysisName)> KeyDetails { get; } = [];
}

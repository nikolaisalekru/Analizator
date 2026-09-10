namespace Analizator.Core;

public sealed record PdfPhotoProgress(int Percent, string Stage, string? FileName = null);

public sealed record PdfPhotoReportRequest(
    string SourceFolder,
    string OutputFile,
    DateTime? PeriodStart = null,
    DateTime? PeriodEnd = null,
    int MinimumPhotoWidth = 100,
    int MinimumPhotoHeight = 100);

public sealed record PdfPhotoEntry(
    string Installation,
    string PersonName,
    DateTime? AttemptedAt,
    string Scenario,
    string Mode,
    bool HasPhoto,
    int MatchingImages,
    int Pages,
    string ReportFolder,
    string FileName,
    string PdfPath,
    string RecognitionNote);

public sealed class PdfPhotoReportResult
{
    public required string OutputFile { get; init; }
    public int PdfFiles { get; init; }
    public int WithPhoto { get; init; }
    public int WithoutPhoto { get; init; }
    public int ReadErrors { get; init; }
    public int MissingPerson { get; init; }
    public int MissingScenario { get; init; }
    public int Installations { get; init; }
}

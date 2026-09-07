namespace Analizator.Core;

public sealed class AnalysisFileAccessException : IOException
{
    public AnalysisFileAccessException(string filePath, Exception innerException)
        : base(
            $"Не удалось записать файл результата, потому что он открыт в Excel или недоступен для записи:\n\n" +
            $"{filePath}\n\nЗакройте файл и повторите запуск.",
            innerException)
    {
        FilePath = filePath;
    }

    public string FilePath { get; }
}

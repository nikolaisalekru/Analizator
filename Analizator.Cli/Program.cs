using Analizator.Core;

try
{
    var options = ParseOptions(args);
    if (!options.TryGetValue("analysis", out var analysis) ||
        !options.TryGetValue("export", out var export))
    {
        PrintHelp();
        return 2;
    }

    var output = options.GetValueOrDefault("output-dir") ?? Path.Combine(Environment.CurrentDirectory, "output");
    var config = options.GetValueOrDefault("config-dir") ?? Path.Combine(Environment.CurrentDirectory, "config");
    var createBackup = options.ContainsKey("create-backup");
    var autoAdd = !options.ContainsKey("do-not-auto-add-new-people");
    if (options.ContainsKey("auto-add-new-people"))
        autoAdd = true;

    var progress = new Progress<AnalyzerProgress>(item =>
        Console.WriteLine($"[UI_PROGRESS] {item.Percent}|{item.Stage}"));
    var result = new AnalyzerService().Run(
        new AnalyzerRequest(analysis!, export!, output, config, createBackup, autoAdd),
        progress);

    Console.WriteLine();
    Console.WriteLine("ОБРАБОТКА ЗАВЕРШЕНА");
    Console.WriteLine($"Результат: {result.OutputFile}");
    Console.WriteLine($"Всего отчётов: {result.Reports}");
    Console.WriteLine($"Сопоставлено сотрудников: {result.MatchedPeople}");
    Console.WriteLine($"Новых сотрудников: {result.NewPeople}; добавлено: {result.AddedPeople}");
    Console.WriteLine($"Отсутствовали в выгрузке: {result.AbsentPeople}");
    Console.WriteLine($"Журнал: {result.LogFile}");
    return 0;
}
catch (Exception exception)
{
    Console.Error.WriteLine("КРИТИЧЕСКАЯ ОШИБКА");
    Console.Error.WriteLine(exception);
    return 1;
}

static Dictionary<string, string?> ParseOptions(string[] arguments)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    for (var index = 0; index < arguments.Length; index++)
    {
        var argument = arguments[index];
        if (!argument.StartsWith("--", StringComparison.Ordinal))
            throw new ArgumentException($"Неизвестный аргумент: {argument}");
        var name = argument[2..];
        if (name is "create-backup" or "auto-add-new-people" or "do-not-auto-add-new-people")
        {
            result[name] = null;
            continue;
        }
        if (index + 1 >= arguments.Length || arguments[index + 1].StartsWith("--", StringComparison.Ordinal))
            throw new ArgumentException($"Для --{name} требуется значение.");
        result[name] = arguments[++index];
    }
    return result;
}

static void PrintHelp()
{
    Console.WriteLine("Анализатор прохождения КТК — C# CLI");
    Console.WriteLine();
    Console.WriteLine("Обязательные параметры:");
    Console.WriteLine("  --analysis <файл.xlsx>  книга анализа");
    Console.WriteLine("  --export <файл.xlsx>    выгрузка КТК");
    Console.WriteLine();
    Console.WriteLine("Дополнительно:");
    Console.WriteLine("  --output-dir <папка>    каталог результата");
    Console.WriteLine("  --config-dir <папка>    каталог JSON-конфигурации");
    Console.WriteLine("  --create-backup         создать резервную копию");
    Console.WriteLine("  --auto-add-new-people | --do-not-auto-add-new-people");
}

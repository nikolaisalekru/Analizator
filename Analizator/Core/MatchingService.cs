namespace Analizator.Core;

public static class MatchingService
{
    public const double PersonFuzzyThreshold = 88;
    public const double InstallationFuzzyThreshold = 70;

    public static (string? Name, double Score, string Type) MatchInstallation(
        string? name,
        IReadOnlyDictionary<string, List<string>> installations)
    {
        var normalized = TextNormalization.NormalizeText(name);
        if (normalized.Length == 0)
            return (null, 0, "не найдено");

        foreach (var canonical in installations.Keys)
            if (TextNormalization.NormalizeText(canonical) == normalized)
                return (canonical, 100, "точное совпадение");

        foreach (var pair in installations)
            if (pair.Value.Any(item => TextNormalization.NormalizeText(item) == normalized))
                return (pair.Key, 100, "точное совпадение");

        var inputTokens = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (var pair in installations)
        foreach (var candidate in pair.Value.Prepend(pair.Key))
        {
            var candidateTokens = TextNormalization.NormalizeText(candidate)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (inputTokens.Length == candidateTokens.Length &&
                inputTokens.Zip(candidateTokens).All(tokens => tokens.Second.StartsWith(tokens.First, StringComparison.Ordinal)))
                return (pair.Key, 99, "сокращённое совпадение");
        }

        string? bestCanonical = null;
        var bestScore = 0d;
        foreach (var pair in installations)
        foreach (var candidate in pair.Value.Prepend(pair.Key))
        {
            var score = Similarity(normalized, TextNormalization.NormalizeText(candidate));
            if (score > bestScore)
            {
                bestScore = score;
                bestCanonical = pair.Key;
            }
        }
        return bestScore >= InstallationFuzzyThreshold
            ? (bestCanonical, bestScore, "нечёткое совпадение")
            : (null, bestScore, "не найдено");
    }

    public static MatchResult MatchPerson(
        string exportName,
        IEnumerable<string> analysisPeople,
        string installation,
        IReadOnlyDictionary<string, Dictionary<string, string>> fioKeys,
        IReadOnlyDictionary<string, List<string>> installations)
    {
        var normalizedExport = TextNormalization.NormalizePersonName(exportName);
        if (normalizedExport.Length == 0)
            return new MatchResult(null, 0, MatchKind.New);

        var people = analysisPeople
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .GroupBy(TextNormalization.NormalizePersonName)
            .Where(group => group.Key.Length > 0)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

        if (people.TryGetValue(normalizedExport, out var exact))
            return new MatchResult(exact, 100, MatchKind.Exact);

        foreach (var section in InstallationCandidates(installation, installations))
        {
            if (!fioKeys.TryGetValue(section, out var keys))
                continue;
            foreach (var key in keys)
            {
                // The legacy loader normalizes '$' away. Preserve that observable behavior.
                if (!string.Equals(normalizedExport, TextNormalization.NormalizePersonName(key.Key), StringComparison.OrdinalIgnoreCase))
                    continue;
                var normalizedTarget = TextNormalization.NormalizePersonName(key.Value);
                if (people.TryGetValue(normalizedTarget, out var target))
                    return new MatchResult(target, 100, MatchKind.Key);
                var keyFuzzy = FindBest(normalizedTarget, people);
                if (keyFuzzy.Score >= PersonFuzzyThreshold)
                    return new MatchResult(keyFuzzy.Name, keyFuzzy.Score, MatchKind.Key);
            }
        }

        var fuzzy = FindBest(normalizedExport, people);
        return fuzzy.Score >= PersonFuzzyThreshold
            ? new MatchResult(fuzzy.Name, fuzzy.Score, MatchKind.Fuzzy)
            : new MatchResult(null, 0, MatchKind.New);
    }

    private static IEnumerable<string> InstallationCandidates(
        string installation,
        IReadOnlyDictionary<string, List<string>> installations)
    {
        var values = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            TextNormalization.NormalizeText(installation)
        };
        var matched = MatchInstallation(installation, installations).Name;
        if (matched is not null)
        {
            values.Add(TextNormalization.NormalizeText(matched));
            foreach (var alias in installations[matched])
                values.Add(TextNormalization.NormalizeText(alias));
        }
        return values;
    }

    private static (string? Name, double Score) FindBest(
        string query,
        IReadOnlyDictionary<string, string> choices)
    {
        string? best = null;
        var score = 0d;
        foreach (var choice in choices)
        {
            var candidateScore = Similarity(query, choice.Key);
            if (candidateScore > score)
            {
                score = candidateScore;
                best = choice.Value;
            }
        }
        return (best, score);
    }

    // RapidFuzz ratio is a normalized indel similarity. LCS yields the same scale.
    public static double Similarity(string left, string right)
    {
        if (left == right)
            return 100;
        if (left.Length == 0 || right.Length == 0)
            return 0;
        var previous = new int[right.Length + 1];
        var current = new int[right.Length + 1];
        for (var i = 1; i <= left.Length; i++)
        {
            for (var j = 1; j <= right.Length; j++)
                current[j] = left[i - 1] == right[j - 1]
                    ? previous[j - 1] + 1
                    : Math.Max(previous[j], current[j - 1]);
            (previous, current) = (current, previous);
            Array.Clear(current);
        }
        var totalLength = left.Length + right.Length;
        var indelDistance = totalLength - 2d * previous[right.Length];
        return 100d * (1d - indelDistance / totalLength);
    }
}

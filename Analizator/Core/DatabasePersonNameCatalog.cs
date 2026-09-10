namespace Analizator.Core;

internal sealed class DatabasePersonNameCatalog
{
    private const char KeySeparator = '\u001F';
    private readonly IReadOnlyList<PersonIdentity> _people;
    private readonly Dictionary<string, string> _canonicalNames = new(StringComparer.Ordinal);

    public DatabasePersonNameCatalog(
        IReadOnlyList<PersonIdentity> people,
        IReadOnlyDictionary<string, Dictionary<string, string>> fioKeys,
        IReadOnlyDictionary<string, List<string>> installations)
    {
        _people = people;
        var globalKeys = BuildUnambiguousGlobalKeys(fioKeys);

        foreach (var installationGroup in people.GroupBy(
                     item => TextNormalization.NormalizeText(item.Installation)))
        {
            var explicitKeys = GetExplicitKeys(
                installationGroup.First().Installation, fioKeys, installations, globalKeys);
            var observedNames = installationGroup
                .Select(item => item.RawName)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToArray();

            foreach (var person in installationGroup)
            {
                var rawNormalized = TextNormalization.NormalizePersonName(person.RawName);
                var canonical = explicitKeys.TryGetValue(rawNormalized, out var configured)
                    ? CleanDisplayName(configured)
                    : InferUnambiguousFullName(person.RawName, observedNames);
                _canonicalNames[Key(person.Installation, person.RawName)] =
                    canonical ?? CleanDisplayName(person.RawName);
            }
        }
    }

    public int CanonicalPeopleCount => _people
        .Select(item => Key(item.Installation, Resolve(item.Installation, item.RawName)))
        .Distinct(StringComparer.Ordinal)
        .Count();

    public string Resolve(string installation, string rawName) =>
        _canonicalNames.GetValueOrDefault(Key(installation, rawName), CleanDisplayName(rawName));

    public IReadOnlyList<long> FindPersonIds(string normalizedSearch)
    {
        if (normalizedSearch.Length == 0)
            return [];
        return _people
            .Where(item =>
                !TextNormalization.NormalizePersonName(item.RawName)
                    .Contains(normalizedSearch, StringComparison.Ordinal) &&
                TextNormalization.NormalizePersonName(Resolve(item.Installation, item.RawName))
                    .Contains(normalizedSearch, StringComparison.Ordinal))
            .Select(item => item.Id)
            .Distinct()
            .ToArray();
    }

    private static Dictionary<string, string> GetExplicitKeys(
        string installation,
        IReadOnlyDictionary<string, Dictionary<string, string>> fioKeys,
        IReadOnlyDictionary<string, List<string>> installations,
        IReadOnlyDictionary<string, string> globalKeys)
    {
        var result = new Dictionary<string, string>(globalKeys, StringComparer.Ordinal);
        var sections = new HashSet<string>(StringComparer.Ordinal)
        {
            TextNormalization.NormalizeText(installation)
        };
        var matchedInstallation = MatchingService.MatchInstallation(installation, installations).Name;
        if (matchedInstallation is not null)
        {
            sections.Add(TextNormalization.NormalizeText(matchedInstallation));
            foreach (var alias in installations[matchedInstallation])
                sections.Add(TextNormalization.NormalizeText(alias));
        }

        foreach (var fioSection in fioKeys.Where(item =>
                     sections.Contains(TextNormalization.NormalizeText(item.Key))))
        {
            foreach (var pair in fioSection.Value)
                result[TextNormalization.NormalizePersonName(pair.Key)] = pair.Value.Trim();
        }
        return result;
    }

    private static Dictionary<string, string> BuildUnambiguousGlobalKeys(
        IReadOnlyDictionary<string, Dictionary<string, string>> fioKeys) =>
        fioKeys.Values
            .SelectMany(section => section)
            .GroupBy(pair => TextNormalization.NormalizePersonName(pair.Key), StringComparer.Ordinal)
            .Where(group => group.Key.Length > 0 && group
                .Select(pair => TextNormalization.NormalizePersonName(pair.Value))
                .Distinct(StringComparer.Ordinal).Count() == 1)
            .ToDictionary(group => group.Key, group => group.First().Value.Trim(), StringComparer.Ordinal);

    private static string? InferUnambiguousFullName(
        string rawName,
        IReadOnlyList<string> observedNames)
    {
        var raw = TextNormalization.NormalizePersonName(rawName);
        var rawTokens = Tokens(raw);
        if (rawTokens.Length < 2)
            return null;

        var candidates = observedNames
            .Select(name => new { Display = name.Trim(), Normalized = TextNormalization.NormalizePersonName(name) })
            .Where(item => item.Normalized.Length > raw.Length &&
                           IsSamePersonPrefix(rawTokens, Tokens(item.Normalized)))
            .GroupBy(item => item.Normalized, StringComparer.Ordinal)
            .Select(group => group.First())
            .ToArray();
        if (candidates.Length == 0)
            return null;

        var terminalCandidates = candidates
            .Where(candidate => !candidates.Any(other =>
                other.Normalized.Length > candidate.Normalized.Length &&
                IsSamePersonPrefix(Tokens(candidate.Normalized), Tokens(other.Normalized))))
            .ToArray();
        return terminalCandidates.Length == 1
            ? CleanDisplayName(terminalCandidates[0].Display)
            : null;
    }

    private static bool IsSamePersonPrefix(string[] abbreviated, string[] full)
    {
        if (abbreviated.Length != full.Length)
            return false;
        return abbreviated.Zip(full).All(pair =>
            pair.Second.StartsWith(pair.First, StringComparison.Ordinal));
    }

    private static string[] Tokens(string value) =>
        value.Split(' ', StringSplitOptions.RemoveEmptyEntries);

    private static string CleanDisplayName(string rawName) =>
        rawName.Trim().TrimStart('$').TrimStart();

    private static string Key(string installation, string personName) =>
        TextNormalization.NormalizeText(installation) + KeySeparator +
        TextNormalization.NormalizePersonName(personName);
}

internal sealed record PersonIdentity(long Id, string Installation, string RawName);

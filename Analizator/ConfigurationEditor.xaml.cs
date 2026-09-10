using Analizator.Core;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Analizator;

public partial class ConfigurationEditor : UserControl
{
    private string? _configurationDirectory;
    private string? _fileName;
    private Dictionary<string, List<string>> _installations = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, Dictionary<string, string>> _fioKeys = new(StringComparer.OrdinalIgnoreCase);
    private List<string> _excludedPeople = [];

    public event EventHandler? Saved;

    public ConfigurationEditor() => InitializeComponent();

    public void LoadConfiguration(string configurationDirectory, string fileName, string displayName)
    {
        _configurationDirectory = configurationDirectory;
        _fileName = fileName;
        TitleText.Text = displayName;
        PathText.Text = "Папка конфигурации: " + configurationDirectory;

        try
        {
            var data = ConfigurationService.LoadEditable(configurationDirectory);
            _installations = data.Installations.ToDictionary(
                item => item.Key,
                item => item.Value.ToList(),
                StringComparer.OrdinalIgnoreCase);
            _fioKeys = data.FioKeys.ToDictionary(
                item => item.Key,
                item => new Dictionary<string, string>(item.Value, StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);
            _excludedPeople = data.ExcludedPeople.ToList();

            ShowPanel(fileName);
            if (fileName.Equals("settings.json", StringComparison.OrdinalIgnoreCase))
            {
                ParallelProcessingCheckBox.IsChecked = data.Settings.ParallelProcessing;
                AutoAddPeopleDefaultCheckBox.IsChecked = data.Settings.AutoAddNewPeople.Default;
                MaximumWorkersTextBox.Text = data.Settings.MaxWorkers.ToString();
            }

            RefreshInstallationList();
            RefreshFioSectionList();
            RefreshExcludedPeople();
            SaveButton.IsEnabled = true;
            ReloadButton.IsEnabled = true;
            StatusText.Text = "Данные загружены";
        }
        catch (Exception exception)
        {
            ShowPanel(null);
            SaveButton.IsEnabled = false;
            ReloadButton.IsEnabled = false;
            StatusText.Text = "Ошибка: " + exception.Message;
        }
    }

    public void Reload()
    {
        if (_configurationDirectory is not null && _fileName is not null)
            LoadConfiguration(_configurationDirectory, _fileName, TitleText.Text);
    }

    private void ShowPanel(string? fileName)
    {
        EmptyText.Visibility = fileName is null ? Visibility.Visible : Visibility.Collapsed;
        InstallationsPanel.Visibility = fileName == "installations.json" ? Visibility.Visible : Visibility.Collapsed;
        SettingsPanel.Visibility = fileName == "settings.json" ? Visibility.Visible : Visibility.Collapsed;
        FioKeysPanel.Visibility = fileName == "fio_keys.json" ? Visibility.Visible : Visibility.Collapsed;
        ExcludedPeoplePanel.Visibility = fileName == "excluded_people.json" ? Visibility.Visible : Visibility.Collapsed;
    }

    private void RefreshInstallationList(string? selected = null)
    {
        selected ??= InstallationsComboBox.SelectedItem as string;
        var names = _installations.Keys.OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase).ToList();
        InstallationsComboBox.ItemsSource = names;
        InstallationsComboBox.SelectedItem = names.FirstOrDefault(value =>
            string.Equals(value, selected, StringComparison.OrdinalIgnoreCase)) ?? names.FirstOrDefault();
        RefreshAliases();
    }

    private void RefreshAliases()
    {
        var installation = InstallationsComboBox.SelectedItem as string;
        InstallationAliasesListBox.ItemsSource = installation is not null && _installations.TryGetValue(installation, out var aliases)
            ? aliases.OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase).ToList()
            : null;
    }

    private void InstallationsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => RefreshAliases();

    private void AddInstallationButton_Click(object sender, RoutedEventArgs e)
    {
        var name = NewInstallationTextBox.Text.Trim();
        if (name.Length == 0)
            return;
        if (_installations.ContainsKey(name))
        {
            ShowInformation($"Установка «{name}» уже существует.");
            return;
        }
        _installations[name] = [];
        NewInstallationTextBox.Clear();
        RefreshInstallationList(name);
        MarkChanged();
    }

    private void DeleteInstallationButton_Click(object sender, RoutedEventArgs e)
    {
        if (InstallationsComboBox.SelectedItem is not string installation)
            return;
        if (_installations.Count == 1)
        {
            ShowInformation("Должна остаться хотя бы одна установка.");
            return;
        }
        if (MessageBox.Show(Window.GetWindow(this), $"Удалить установку «{installation}» и все её синонимы?",
                "Удаление установки", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;
        _installations.Remove(installation);
        RefreshInstallationList();
        MarkChanged();
    }

    private void AddAliasButton_Click(object sender, RoutedEventArgs e) => AddAlias();

    private void NewAliasTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            AddAlias();
            e.Handled = true;
        }
    }

    private void AddAlias()
    {
        if (InstallationsComboBox.SelectedItem is not string installation)
            return;
        var alias = NewAliasTextBox.Text.Trim();
        if (alias.Length == 0)
            return;
        var aliases = _installations[installation];
        if (string.Equals(alias, installation, StringComparison.OrdinalIgnoreCase) ||
            aliases.Contains(alias, StringComparer.OrdinalIgnoreCase))
        {
            ShowInformation($"Синоним «{alias}» уже указан.");
            return;
        }
        aliases.Add(alias);
        NewAliasTextBox.Clear();
        RefreshAliases();
        MarkChanged();
    }

    private void DeleteAliasButton_Click(object sender, RoutedEventArgs e)
    {
        if (InstallationsComboBox.SelectedItem is not string installation ||
            InstallationAliasesListBox.SelectedItem is not string alias)
            return;
        _installations[installation].RemoveAll(value =>
            string.Equals(value, alias, StringComparison.OrdinalIgnoreCase));
        RefreshAliases();
        MarkChanged();
    }

    private void RefreshFioSectionList(string? selected = null)
    {
        selected ??= FioSectionsComboBox.SelectedItem as string;
        var sections = _fioKeys.Keys.OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase).ToList();
        FioSectionsComboBox.ItemsSource = sections;
        FioSectionsComboBox.SelectedItem = sections.FirstOrDefault(value =>
            string.Equals(value, selected, StringComparison.OrdinalIgnoreCase)) ?? sections.FirstOrDefault();
        RefreshFioKeys();
    }

    private void RefreshFioKeys()
    {
        var section = FioSectionsComboBox.SelectedItem as string;
        FioKeysDataGrid.ItemsSource = section is not null && _fioKeys.TryGetValue(section, out var values)
            ? values.OrderBy(item => item.Key, StringComparer.CurrentCultureIgnoreCase)
                .Select(item => new FioKeyRow(item.Key, item.Value)).ToList()
            : null;
    }

    private void FioSectionsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => RefreshFioKeys();

    private void AddFioSectionButton_Click(object sender, RoutedEventArgs e)
    {
        var name = NewFioSectionTextBox.Text.Trim();
        if (name.Length == 0)
            return;
        if (_fioKeys.ContainsKey(name))
        {
            ShowInformation($"Раздел «{name}» уже существует.");
            return;
        }
        _fioKeys[name] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        NewFioSectionTextBox.Clear();
        RefreshFioSectionList(name);
        MarkChanged();
    }

    private void DeleteFioSectionButton_Click(object sender, RoutedEventArgs e)
    {
        if (FioSectionsComboBox.SelectedItem is not string section)
            return;
        if (MessageBox.Show(Window.GetWindow(this), $"Удалить раздел «{section}» и все его сокращения?",
                "Удаление раздела", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;
        _fioKeys.Remove(section);
        RefreshFioSectionList();
        MarkChanged();
    }

    private void AddFioKeyButton_Click(object sender, RoutedEventArgs e)
    {
        if (FioSectionsComboBox.SelectedItem is not string section)
        {
            ShowInformation("Сначала добавьте раздел для сокращений.");
            return;
        }
        var abbreviation = NewAbbreviationTextBox.Text.Trim();
        var fullName = NewFullNameTextBox.Text.Trim();
        if (abbreviation.Length == 0 || fullName.Length == 0)
        {
            ShowInformation("Заполните сокращение и полное ФИО.");
            return;
        }
        if (_fioKeys[section].ContainsKey(abbreviation))
        {
            ShowInformation($"Сокращение «{abbreviation}» уже существует в этом разделе.");
            return;
        }
        _fioKeys[section][abbreviation] = fullName;
        NewAbbreviationTextBox.Clear();
        NewFullNameTextBox.Clear();
        RefreshFioKeys();
        MarkChanged();
    }

    private void DeleteFioKeyButton_Click(object sender, RoutedEventArgs e)
    {
        if (FioSectionsComboBox.SelectedItem is not string section ||
            FioKeysDataGrid.SelectedItem is not FioKeyRow row)
            return;
        _fioKeys[section].Remove(row.Abbreviation);
        RefreshFioKeys();
        MarkChanged();
    }

    private void RefreshExcludedPeople() =>
        ExcludedPeopleListBox.ItemsSource = _excludedPeople
            .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase).ToList();

    private void AddExcludedPersonButton_Click(object sender, RoutedEventArgs e) => AddExcludedPerson();

    private void NewExcludedPersonTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            AddExcludedPerson();
            e.Handled = true;
        }
    }

    private void AddExcludedPerson()
    {
        var value = NewExcludedPersonTextBox.Text.Trim();
        if (value.Length == 0)
            return;
        if (_excludedPeople.Contains(value, StringComparer.OrdinalIgnoreCase))
        {
            ShowInformation($"Значение «{value}» уже добавлено.");
            return;
        }
        _excludedPeople.Add(value);
        NewExcludedPersonTextBox.Clear();
        RefreshExcludedPeople();
        MarkChanged();
    }

    private void DeleteExcludedPersonButton_Click(object sender, RoutedEventArgs e)
    {
        if (ExcludedPeopleListBox.SelectedItem is not string value)
            return;
        _excludedPeople.RemoveAll(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase));
        RefreshExcludedPeople();
        MarkChanged();
    }

    private void MarkChanged() => StatusText.Text = "Есть несохранённые изменения";

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        if (_configurationDirectory is null || _fileName is null)
            return;
        try
        {
            switch (_fileName)
            {
                case "installations.json":
                    ConfigurationService.SaveInstallations(_configurationDirectory,
                        _installations.ToDictionary(
                            item => item.Key,
                            item => (IReadOnlyCollection<string>)item.Value,
                            StringComparer.OrdinalIgnoreCase));
                    break;
                case "settings.json":
                    if (!int.TryParse(MaximumWorkersTextBox.Text.Trim(), out var maximumWorkers))
                        throw new ArgumentException("Количество параллельных обработчиков должно быть целым числом.");
                    ConfigurationService.SaveAnalysisSettings(
                        _configurationDirectory,
                        ParallelProcessingCheckBox.IsChecked == true,
                        maximumWorkers,
                        AutoAddPeopleDefaultCheckBox.IsChecked == true);
                    break;
                case "fio_keys.json":
                    ConfigurationService.SaveFioKeys(_configurationDirectory,
                        _fioKeys.ToDictionary(
                            item => item.Key,
                            item => (IReadOnlyDictionary<string, string>)item.Value,
                            StringComparer.OrdinalIgnoreCase));
                    break;
                case "excluded_people.json":
                    ConfigurationService.SaveExcludedPeople(_configurationDirectory, _excludedPeople);
                    break;
            }

            StatusText.Text = "Сохранено · " + DateTime.Now.ToString("HH:mm:ss");
            Saved?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception exception)
        {
            MessageBox.Show(Window.GetWindow(this), exception.Message, "Ошибка сохранения",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ReloadButton_Click(object sender, RoutedEventArgs e) => Reload();

    private void ShowInformation(string message) =>
        MessageBox.Show(Window.GetWindow(this), message, "Редактор конфигурации",
            MessageBoxButton.OK, MessageBoxImage.Information);

    private sealed record FioKeyRow(string Abbreviation, string FullName);
}

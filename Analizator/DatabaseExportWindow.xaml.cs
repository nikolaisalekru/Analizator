using Analizator.Core;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace Analizator;

public partial class DatabaseExportWindow : Window
{
    private readonly TrainingDatabaseService _database;
    private readonly string _defaultOutputDirectory;
    private readonly ObservableCollection<DatabaseInstallationOption> _installations = [];
    private readonly ObservableCollection<DatabasePersonOption> _people = [];
    private readonly ObservableCollection<ExportColumnOption> _columns;
    private readonly ICollectionView _peopleView;
    private bool _isBusy;

    public DatabaseExportWindow(
        TrainingDatabaseService database,
        string defaultOutputDirectory)
    {
        InitializeComponent();
        _database = database;
        _defaultOutputDirectory = defaultOutputDirectory;
        _columns = CreateColumns();
        InstallationsListBox.ItemsSource = _installations;
        PeopleDataGrid.ItemsSource = _people;
        ColumnsListBox.ItemsSource = _columns;
        _peopleView = CollectionViewSource.GetDefaultView(_people);
        _peopleView.Filter = IsPersonVisible;

        var availableMonths = _database.GetAvailableMonths();
        MonthComboBox.ItemsSource = availableMonths;
        var latestMonth = availableMonths.FirstOrDefault();
        if (latestMonth is not null)
        {
            MonthComboBox.SelectedItem = latestMonth;
            MonthPeriodRadioButton.IsChecked = true;
            ApplyMonthPeriod(latestMonth);
        }
        else
        {
            var summary = _database.GetSummary();
            RangePeriodRadioButton.IsChecked = true;
            StartDatePicker.SelectedDate = summary.PeriodStart?.Date ?? DateTime.Today;
            EndDatePicker.SelectedDate = summary.PeriodEnd?.Date ?? DateTime.Today;
        }

        UpdatePeriodModeControls();
        LoadPeopleForPeriod();
    }

    public DatabaseSelectionExportResult? ExportResult { get; private set; }

    private bool IsMonthMode => MonthPeriodRadioButton.IsChecked == true;

    private DateTime PeriodStart => IsMonthMode
        ? (MonthComboBox.SelectedItem as DatabaseMonthSummary)?.FirstDay
          ?? throw new InvalidOperationException("Выберите месяц.")
        : StartDatePicker.SelectedDate?.Date
          ?? throw new InvalidOperationException("Укажите дату начала периода.");

    private DateTime PeriodEnd => IsMonthMode
        ? (MonthComboBox.SelectedItem as DatabaseMonthSummary)?.LastDay
          ?? throw new InvalidOperationException("Выберите месяц.")
        : EndDatePicker.SelectedDate?.Date
          ?? throw new InvalidOperationException("Укажите дату окончания периода.");

    private void ApplyMonthPeriod(DatabaseMonthSummary month)
    {
        StartDatePicker.SelectedDate = month.FirstDay;
        EndDatePicker.SelectedDate = month.LastDay;
    }

    private void UpdatePeriodModeControls()
    {
        MonthComboBox.IsEnabled = !_isBusy && IsMonthMode;
        StartDatePicker.IsEnabled = !_isBusy && !IsMonthMode;
        EndDatePicker.IsEnabled = !_isBusy && !IsMonthMode;
        RefreshPeopleButton.IsEnabled = !_isBusy;
    }

    private void LoadPeopleForPeriod()
    {
        try
        {
            if (PeriodEnd < PeriodStart)
                throw new InvalidOperationException(
                    "Дата окончания периода не может быть раньше даты начала.");

            var rows = _database.GetPeopleForExport(PeriodStart, PeriodEnd);
            _installations.Clear();
            foreach (var installation in rows.Select(item => item.Installation)
                         .Distinct(StringComparer.CurrentCultureIgnoreCase)
                         .OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase))
                _installations.Add(new DatabaseInstallationOption(installation, true));

            _people.Clear();
            foreach (var row in rows)
                _people.Add(new DatabasePersonOption(
                    row.Installation, row.PersonName, row.Attempts));

            PeopleSearchTextBox.Clear();
            RefreshPeopleFilter();
            PeriodSummaryText.Text = rows.Count == 0
                ? "За этот период записей нет"
                : $"{PeriodStart:dd.MM.yyyy} — {PeriodEnd:dd.MM.yyyy} · " +
                  $"{rows.Sum(item => item.Attempts):N0} записей · " +
                  $"{rows.Count:N0} сотрудников · {_installations.Count:N0} установок";
            ExportStatusText.Text = rows.Count == 0
                ? "Измените даты и обновите список"
                : "Настройте листы, столбцы и, при необходимости, фильтр по ФИО";
            UpdateSelectionSummary();
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось подготовить выборку",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private bool IsPersonVisible(object item)
    {
        if (item is not DatabasePersonOption person)
            return false;
        var installationSelected = _installations.Any(option =>
            option.IsSelected && string.Equals(option.Name, person.Installation,
                StringComparison.CurrentCultureIgnoreCase));
        if (!installationSelected)
            return false;

        var search = TextNormalization.NormalizePersonName(PeopleSearchTextBox.Text);
        return search.Length == 0 ||
               TextNormalization.NormalizePersonName(person.PersonName).Contains(
                   search, StringComparison.Ordinal);
    }

    private void RefreshPeopleFilter()
    {
        _peopleView.Refresh();
        VisiblePeopleText.Text = $"Найдено: {_peopleView.Cast<object>().Count():N0}";
    }

    private void UpdateSelectionSummary()
    {
        var selectedInstallations = _installations.Where(item => item.IsSelected)
            .Select(item => item.Name)
            .ToHashSet(StringComparer.CurrentCultureIgnoreCase);
        var selected = _people.Where(item => item.IsSelected).ToArray();
        var selectedColumnCount = _columns.Count(item => item.IsSelected);
        var scopedPeople = _people.Where(item => selectedInstallations.Contains(item.Installation));
        var reports = selected.Length == 0
            ? scopedPeople.Sum(item => item.Attempts)
            : selected.Sum(item => item.Attempts);
        SelectionSummaryText.Text = selectedColumnCount == 0
            ? "Столбцы Excel не выбраны"
            : selectedInstallations.Count == 0
            ? "Установки не выбраны"
            : selected.Length == 0
                ? $"Все сотрудники · {reports:N0} записей · " +
                  $"{selectedInstallations.Count:N0} установок · {selectedColumnCount:N0} столбцов"
                : $"Выбрано: {selected.Length:N0} сотрудников · {reports:N0} записей · " +
                  $"{selectedInstallations.Count:N0} установок · {selectedColumnCount:N0} столбцов";
        ExportButton.IsEnabled = !_isBusy && selectedInstallations.Count > 0 &&
                                 reports > 0 && selectedColumnCount > 0;
    }

    private void RefreshPeopleButton_Click(object sender, RoutedEventArgs e) =>
        LoadPeopleForPeriod();

    private void PeriodModeButton_Click(object sender, RoutedEventArgs e)
    {
        UpdatePeriodModeControls();
        if (IsMonthMode && MonthComboBox.SelectedItem is DatabaseMonthSummary month)
            ApplyMonthPeriod(month);
        LoadPeopleForPeriod();
    }

    private void MonthComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!IsLoaded || !IsMonthMode ||
            MonthComboBox.SelectedItem is not DatabaseMonthSummary month)
            return;
        ApplyMonthPeriod(month);
        LoadPeopleForPeriod();
    }

    private void InstallationCheckBox_Click(object sender, RoutedEventArgs e)
    {
        if (sender is CheckBox { DataContext: DatabaseInstallationOption installation } &&
            !installation.IsSelected)
        {
            foreach (var person in _people.Where(item => string.Equals(
                         item.Installation, installation.Name,
                         StringComparison.CurrentCultureIgnoreCase)))
                person.IsSelected = false;
        }
        RefreshPeopleFilter();
        UpdateSelectionSummary();
    }

    private void SelectAllInstallationsButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var installation in _installations)
            installation.IsSelected = true;
        RefreshPeopleFilter();
        UpdateSelectionSummary();
    }

    private void ClearInstallationsButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var installation in _installations)
            installation.IsSelected = false;
        foreach (var person in _people)
            person.IsSelected = false;
        RefreshPeopleFilter();
        UpdateSelectionSummary();
    }

    private void PeopleSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_peopleView is not null)
            RefreshPeopleFilter();
    }

    private void PersonCheckBox_Click(object sender, RoutedEventArgs e)
    {
        if (sender is CheckBox
            {
                DataContext: DatabasePersonOption person,
                IsChecked: var isChecked
            })
        {
            person.IsSelected = isChecked == true;
        }
        UpdateSelectionSummary();
    }

    private void SelectVisiblePeopleButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var person in _peopleView.Cast<DatabasePersonOption>())
            person.IsSelected = true;
        UpdateSelectionSummary();
    }

    private void ClearPeopleButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var person in _people)
            person.IsSelected = false;
        UpdateSelectionSummary();
    }

    private void ColumnCheckBox_Click(object sender, RoutedEventArgs e) =>
        UpdateSelectionSummary();

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
            return;

        try
        {
            var selectedInstallations = _installations.Where(item => item.IsSelected)
                .Select(item => item.Name)
                .ToArray();
            if (selectedInstallations.Length == 0)
                throw new InvalidOperationException("Выберите хотя бы одну установку.");

            var selectedInstallationSet = selectedInstallations
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase);
            var selectedPeople = _people.Where(item =>
                    item.IsSelected && selectedInstallationSet.Contains(item.Installation))
                .Select(item => new DatabasePersonKey(item.Installation, item.PersonName))
                .ToArray();

            var selectedColumns = _columns.Where(item => item.IsSelected)
                .Select(item => item.Column).ToArray();
            if (selectedColumns.Length == 0)
                throw new InvalidOperationException("Выберите хотя бы один столбец Excel.");

            Directory.CreateDirectory(_defaultOutputDirectory);
            var dialog = new SaveFileDialog
            {
                Title = "Сохранить выгрузку КТК из базы",
                Filter = "Книга Excel (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                OverwritePrompt = true,
                InitialDirectory = _defaultOutputDirectory,
                FileName = DatabaseSelectionExportService.SuggestedFileName(PeriodStart, PeriodEnd)
            };
            if (dialog.ShowDialog(this) != true)
                return;

            var request = new DatabaseSelectionExportRequest(
                dialog.FileName,
                PeriodStart,
                PeriodEnd,
                selectedInstallations,
                selectedPeople,
                selectedColumns);

            SetBusy(true);
            ExportStatusText.Text = "Формирование книги Excel…";
            ExportResult = await Task.Run(() =>
                new DatabaseSelectionExportService().Export(_database, request));
            ExportStatusText.Text = "Файл успешно сформирован";

            var answer = MessageBox.Show(this,
                $"Выгрузка готова.\n\n" +
                $"Записей: {ExportResult.Attempts:N0}\n" +
                $"Сотрудников: {ExportResult.People:N0}\n" +
                $"Установок: {ExportResult.Installations:N0}\n\n" +
                "Открыть файл сейчас?",
                "Выгрузка сформирована",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);
            if (answer == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(ExportResult.OutputFile)
                    { UseShellExecute = true });
            SetBusy(false);
            DialogResult = true;
        }
        catch (Exception exception)
        {
            ExportStatusText.Text = "Не удалось сформировать файл";
            MessageBox.Show(this, exception.Message, "Ошибка выгрузки",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (IsVisible)
                SetBusy(false);
        }
    }

    private void SetBusy(bool busy)
    {
        _isBusy = busy;
        SelectionWorkspace.IsEnabled = !busy;
        InstallationsListBox.IsEnabled = !busy;
        PeopleDataGrid.IsEnabled = !busy;
        ColumnsListBox.IsEnabled = !busy;
        CancelButton.IsEnabled = !busy;
        ExportProgressBar.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        UpdatePeriodModeControls();
        UpdateSelectionSummary();
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_isBusy)
            e.Cancel = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => Close();

    private static ObservableCollection<ExportColumnOption> CreateColumns() =>
    [
        new(DatabaseExportColumn.Date, "Дата", false, true),
        new(DatabaseExportColumn.Scenario, "Тема", false, true),
        new(DatabaseExportColumn.Person, "ФИО по отчёту", false, true),
        new(DatabaseExportColumn.Duration, "Время прохождения", false, true),
        new(DatabaseExportColumn.Actions, "Количество действий", false, true),
        new(DatabaseExportColumn.Percent, "Процент", false, true),
        new(DatabaseExportColumn.Mode, "Режим", false, true),
        new(DatabaseExportColumn.Project, "Проект", false, false),
        new(DatabaseExportColumn.Role, "Роль", false, false),
        new(DatabaseExportColumn.Credited, "Зачтено", false, false),
        new(DatabaseExportColumn.Source, "Источник", false, false)
    ];
}

internal abstract class SelectableOption : INotifyPropertyChanged
{
    private bool _isSelected;

    protected SelectableOption(bool isSelected) => _isSelected = isSelected;

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value)
                return;
            _isSelected = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

internal sealed class DatabaseInstallationOption(string name, bool isSelected)
    : SelectableOption(isSelected)
{
    public string Name { get; } = name;
}

internal sealed class DatabasePersonOption(
    string installation,
    string personName,
    int attempts) : SelectableOption(false)
{
    public string Installation { get; } = installation;
    public string PersonName { get; } = personName;
    public int Attempts { get; } = attempts;
}

internal sealed class ExportColumnOption : SelectableOption
{
    public ExportColumnOption(
        DatabaseExportColumn column,
        string displayName,
        bool required,
        bool selected) : base(selected)
    {
        Column = column;
        DisplayName = displayName;
        CanChange = !required;
        Note = required ? "обязательно для формата КТК" : "";
    }

    public DatabaseExportColumn Column { get; }
    public string DisplayName { get; }
    public bool CanChange { get; }
    public string Note { get; }
}

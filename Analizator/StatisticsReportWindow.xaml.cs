using Analizator.Core;
using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;

namespace Analizator;

public partial class StatisticsReportWindow : Window
{
    private readonly TrainingDatabaseService _database;
    private readonly string _configurationDirectory;
    private readonly string _outputDirectory;
    private readonly List<StatisticsInstallationOption> _installations = [];
    private CancellationTokenSource? _cancellation;
    private bool _isBusy;

    public StatisticsReportWindow(
        TrainingDatabaseService database,
        string configurationDirectory,
        string outputDirectory)
    {
        InitializeComponent();
        _database = database;
        _configurationDirectory = configurationDirectory;
        _outputDirectory = outputDirectory;
        var summary = _database.GetSummary();
        StartDatePicker.SelectedDate = summary.PeriodStart?.Date ?? DateTime.Today;
        EndDatePicker.SelectedDate = summary.PeriodEnd?.Date ?? DateTime.Today;
        LoadScope();
    }

    public StatisticsReportResult? ReportResult { get; private set; }

    private DateTime PeriodStart => StartDatePicker.SelectedDate?.Date
        ?? throw new InvalidOperationException("Укажите дату начала периода.");

    private DateTime PeriodEnd => EndDatePicker.SelectedDate?.Date
        ?? throw new InvalidOperationException("Укажите дату окончания периода.");

    private void LoadScope()
    {
        if (_isBusy)
            return;
        try
        {
            if (PeriodEnd < PeriodStart)
                throw new InvalidOperationException(
                    "Дата окончания периода не может быть раньше даты начала.");

            var selected = _installations.Count == 0
                ? null
                : _installations.Where(item => item.IsSelected)
                    .Select(item => TextNormalization.NormalizeText(item.Name))
                    .ToHashSet(StringComparer.Ordinal);
            var people = _database.GetPeopleForExport(PeriodStart, PeriodEnd);
            var rows = people.GroupBy(item => item.Installation,
                    StringComparer.CurrentCultureIgnoreCase)
                .Select(group => new StatisticsInstallationOption(
                    group.Key,
                    group.Sum(item => item.Attempts),
                    group.Count(),
                    selected is null || selected.Contains(TextNormalization.NormalizeText(group.Key))))
                .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            _installations.Clear();
            _installations.AddRange(rows);
            InstallationsListBox.ItemsSource = null;
            InstallationsListBox.ItemsSource = _installations;
            ScopeSummaryText.Text = rows.Count == 0
                ? "За выбранный период записей нет"
                : $"{rows.Sum(item => item.Attempts):N0} записей · " +
                  $"{people.Count:N0} сотрудников · {rows.Count:N0} установок";
            StatusText.Text = rows.Count == 0
                ? "Измените период и обновите данные"
                : "Выберите установки и сформируйте отчёт";
            UpdateSelection();
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось прочитать базу",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void UpdateSelection()
    {
        var selected = _installations.Where(item => item.IsSelected).ToArray();
        GenerateButton.IsEnabled = !_isBusy && selected.Length > 0;
        if (_isBusy)
            return;
        StatusText.Text = selected.Length == 0
            ? "Выберите хотя бы одну установку"
            : $"Выбрано установок: {selected.Length:N0} · " +
              $"записей: {selected.Sum(item => item.Attempts):N0} · " +
              $"сотрудников: {selected.Sum(item => item.People):N0}";
    }

    private void RefreshScopeButton_Click(object sender, RoutedEventArgs e) => LoadScope();

    private void InstallationCheckBox_Click(object sender, RoutedEventArgs e)
    {
        if (sender is CheckBox
            {
                DataContext: StatisticsInstallationOption option,
                IsChecked: var isChecked
            })
        {
            option.IsSelected = isChecked == true;
        }
        UpdateSelection();
    }

    private void SelectAllInstallationsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
            return;
        foreach (var item in _installations)
            item.IsSelected = true;
        UpdateSelection();
    }

    private void ClearInstallationsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
            return;
        foreach (var item in _installations)
            item.IsSelected = false;
        UpdateSelection();
    }

    private async void GenerateButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
            return;
        try
        {
            var selected = _installations.Where(item => item.IsSelected)
                .Select(item => item.Name).ToArray();
            if (selected.Length == 0)
                throw new InvalidOperationException("Выберите хотя бы одну установку.");

            Directory.CreateDirectory(_outputDirectory);
            var dialog = new SaveFileDialog
            {
                Title = "Сохранить подробную статистику КТК",
                Filter = "Книга Excel (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                OverwritePrompt = true,
                InitialDirectory = _outputDirectory,
                FileName = DefaultFileName(PeriodStart, PeriodEnd)
            };
            if (dialog.ShowDialog(this) != true)
                return;

            var configuration = ConfigurationService.Load(_configurationDirectory);
            var request = new StatisticsReportRequest(
                _database.DatabasePath,
                dialog.FileName,
                PeriodStart,
                PeriodEnd,
                selected,
                IncludePossibleDuplicatesCheckBox.IsChecked == true,
                configuration.Settings.PassingThreshold);
            _cancellation = new CancellationTokenSource();
            SetBusy(true);
            ProgressBar.Value = 1;
            StatusText.Text = "Подготовка данных…";
            var progress = new Progress<StatisticsReportProgress>(item =>
            {
                ProgressBar.Value = Math.Clamp(item.Percent, 0, 100);
                StatusText.Text = item.Stage;
            });
            ReportResult = await RunReportAsync(request, progress, _cancellation.Token);
            ProgressBar.Value = 100;
            StatusText.Text = "Отчёт успешно сформирован";

            if (OpenResultCheckBox.IsChecked == true)
                Process.Start(new ProcessStartInfo(ReportResult.OutputFile)
                    { UseShellExecute = true });
            SetBusy(false);
            MessageBox.Show(this,
                $"Подробная статистика сформирована.\n\n" +
                $"Попыток: {ReportResult.Attempts:N0}\n" +
                $"Сотрудников: {ReportResult.People:N0}\n" +
                $"Установок: {ReportResult.Installations:N0}\n" +
                $"Сценариев: {ReportResult.Scenarios:N0}",
                "Отчёт готов",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            DialogResult = true;
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = "Формирование отчёта отменено";
        }
        catch (Exception exception)
        {
            StatusText.Text = "Не удалось сформировать отчёт";
            MessageBox.Show(this, exception.Message, "Ошибка статистики",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _cancellation?.Dispose();
            _cancellation = null;
            if (IsVisible)
                SetBusy(false);
        }
    }

    private static Task<StatisticsReportResult> RunReportAsync(
        StatisticsReportRequest request,
        IProgress<StatisticsReportProgress> progress,
        CancellationToken cancellationToken)
    {
        var forceWorker = string.Equals(
            Environment.GetEnvironmentVariable("ANALIZATOR_FORCE_WORKER"),
            "1",
            StringComparison.Ordinal);
        if (Debugger.IsAttached || forceWorker)
            return StatisticsWorker.RunIsolatedAsync(request, progress, cancellationToken);
        return Task.Run(
            () => new StatisticsReportService().Run(request, progress, cancellationToken),
            cancellationToken);
    }

    private void SetBusy(bool busy)
    {
        _isBusy = busy;
        StartDatePicker.IsEnabled = !busy;
        EndDatePicker.IsEnabled = !busy;
        RefreshScopeButton.IsEnabled = !busy;
        InstallationsListBox.IsEnabled = !busy;
        SelectAllButton.IsEnabled = !busy;
        ClearButton.IsEnabled = !busy;
        IncludePossibleDuplicatesCheckBox.IsEnabled = !busy;
        OpenResultCheckBox.IsEnabled = !busy;
        CancelButton.IsEnabled = busy;
        UpdateSelection();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => _cancellation?.Cancel();

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (!_isBusy)
            return;
        _cancellation?.Cancel();
        e.Cancel = true;
    }

    private static string DefaultFileName(DateTime start, DateTime end) =>
        $"Статистика_КТК_{start:yyyy-MM-dd}_{end:yyyy-MM-dd}.xlsx";
}

internal sealed class StatisticsInstallationOption : INotifyPropertyChanged
{
    private bool _isSelected;

    public StatisticsInstallationOption(string name, int attempts, int people, bool isSelected)
    {
        Name = name;
        Attempts = attempts;
        People = people;
        _isSelected = isSelected;
    }

    public string Name { get; }
    public int Attempts { get; }
    public int People { get; }
    public string Summary => $"{Attempts:N0} записей · {People:N0} сотрудников";

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

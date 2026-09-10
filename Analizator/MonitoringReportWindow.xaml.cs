using Analizator.Core;
using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;

namespace Analizator;

public partial class MonitoringReportWindow : Window
{
    private readonly TrainingDatabaseService _database;
    private readonly string _outputDirectory;
    private readonly string _dataFolder;
    private readonly List<StatisticsInstallationOption> _installations = [];
    private CancellationTokenSource? _cancellation;
    private bool _isBusy;

    public MonitoringReportWindow(TrainingDatabaseService database, string outputDirectory, string dataFolder)
    {
        InitializeComponent();
        _database = database;
        _outputDirectory = outputDirectory;
        _dataFolder = dataFolder;
        var summary = _database.GetSummary();
        var end = summary.PeriodEnd?.Date ?? DateTime.Today;
        StartDatePicker.SelectedDate = new[] { summary.PeriodStart?.Date ?? end, new DateTime(end.Year, 1, 1) }.Max();
        EndDatePicker.SelectedDate = end;
        TemplatePathTextBox.Text = LocateTemplate();
        LoadScope();
    }

    public MonitoringReportResult? ReportResult { get; private set; }
    private DateTime PeriodStart => StartDatePicker.SelectedDate?.Date ?? throw new InvalidOperationException("Укажите дату начала периода.");
    private DateTime PeriodEnd => EndDatePicker.SelectedDate?.Date ?? throw new InvalidOperationException("Укажите дату окончания периода.");

    private string LocateTemplate()
    {
        var name = "1. Мониторинг КТК 2026_форма сбора.xlsx";
        string[] candidates =
        [
            Path.Combine(_dataFolder, "input", name),
            Path.Combine(_dataFolder, "0. Дополнительные модули", "14. Мониторинг КТК", "input", name)
        ];
        return candidates.FirstOrDefault(File.Exists) ?? "";
    }

    private void LoadScope()
    {
        if (_isBusy) return;
        try
        {
            if (PeriodEnd < PeriodStart) throw new InvalidOperationException("Дата окончания периода не может быть раньше даты начала.");
            if (PeriodStart.Year != PeriodEnd.Year) throw new InvalidOperationException("Для формы выберите даты одного календарного года.");
            var selected = _installations.Count == 0 ? null : _installations.Where(x => x.IsSelected).Select(x => TextNormalization.NormalizeText(x.Name)).ToHashSet(StringComparer.Ordinal);
            var people = _database.GetPeopleForExport(PeriodStart, PeriodEnd);
            var rows = people.GroupBy(x => x.Installation, StringComparer.CurrentCultureIgnoreCase)
                .Select(g => new StatisticsInstallationOption(g.Key, g.Sum(x => x.Attempts), g.Count(), selected is null || selected.Contains(TextNormalization.NormalizeText(g.Key))))
                .OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
            _installations.Clear(); _installations.AddRange(rows);
            InstallationsListBox.ItemsSource = null; InstallationsListBox.ItemsSource = _installations;
            ScopeSummaryText.Text = rows.Count == 0 ? "За период данных нет" : $"{rows.Sum(x => x.Attempts):N0} записей · {people.Count:N0} сотрудников · {rows.Count:N0} установок";
            UpdateSelection();
        }
        catch (Exception exception) { MessageBox.Show(this, exception.Message, "Мониторинг КТК", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    private void UpdateSelection()
    {
        var selected = _installations.Where(x => x.IsSelected).ToArray();
        GenerateButton.IsEnabled = !_isBusy && selected.Length > 0 && File.Exists(TemplatePathTextBox.Text);
        if (!_isBusy) StatusText.Text = selected.Length == 0 ? "Выберите хотя бы одну установку" : $"Выбрано установок: {selected.Length:N0}";
    }

    private void RefreshScopeButton_Click(object sender, RoutedEventArgs e) => LoadScope();
    private void InstallationCheckBox_Click(object sender, RoutedEventArgs e) { if (sender is CheckBox { DataContext: StatisticsInstallationOption option, IsChecked: var value }) option.IsSelected = value == true; UpdateSelection(); }
    private void SelectAllButton_Click(object sender, RoutedEventArgs e) { foreach (var item in _installations) item.IsSelected = true; InstallationsListBox.Items.Refresh(); UpdateSelection(); }
    private void ClearButton_Click(object sender, RoutedEventArgs e) { foreach (var item in _installations) item.IsSelected = false; InstallationsListBox.Items.Refresh(); UpdateSelection(); }

    private void BrowseTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Title = "Выберите форму сбора КТК", Filter = "Книга Excel (*.xlsx)|*.xlsx", CheckFileExists = true };
        if (dialog.ShowDialog(this) == true) { TemplatePathTextBox.Text = dialog.FileName; UpdateSelection(); }
    }

    private async void GenerateButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy) return;
        try
        {
            var selected = _installations.Where(x => x.IsSelected).Select(x => x.Name).ToArray();
            if (selected.Length == 0) throw new InvalidOperationException("Выберите хотя бы одну установку.");
            if (!File.Exists(TemplatePathTextBox.Text)) throw new FileNotFoundException("Выберите файл формы сбора.");
            Directory.CreateDirectory(_outputDirectory);
            var dialog = new SaveFileDialog { Title = "Сохранить мониторинг КТК", Filter = "Книга Excel (*.xlsx)|*.xlsx", DefaultExt = ".xlsx", AddExtension = true, OverwritePrompt = true, InitialDirectory = _outputDirectory, FileName = $"Расчёт_мониторинга_КТК_{PeriodStart:yyyy-MM-dd}_{PeriodEnd:yyyy-MM-dd}.xlsx" };
            if (dialog.ShowDialog(this) != true) return;
            var formOutput = Path.Combine(Path.GetDirectoryName(dialog.FileName)!, $"Мониторинг_КТК_форма_{PeriodStart:yyyy-MM-dd}_{PeriodEnd:yyyy-MM-dd}.xlsx");
            if (File.Exists(formOutput) && MessageBox.Show(this, $"Файл {Path.GetFileName(formOutput)} уже существует. Заменить его?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            var request = new MonitoringReportRequest(_database.DatabasePath, dialog.FileName, TemplatePathTextBox.Text, formOutput, PeriodStart, PeriodEnd, selected, IncludePossibleDuplicatesCheckBox.IsChecked == true);
            _cancellation = new CancellationTokenSource(); SetBusy(true);
            var progress = new Progress<MonitoringReportProgress>(x => { ProgressBar.Value = x.Percent; StatusText.Text = x.Stage; });
            ReportResult = await RunReportAsync(request, progress, _cancellation.Token);
            ProgressBar.Value = 100; StatusText.Text = "Мониторинг сформирован";
            if (OpenResultsCheckBox.IsChecked == true)
            {
                Process.Start(new ProcessStartInfo(ReportResult.ReportOutputFile) { UseShellExecute = true });
                Process.Start(new ProcessStartInfo(ReportResult.FormOutputFile) { UseShellExecute = true });
            }
            var written = string.Join(", ", ReportResult.Quarters.Where(x => x.WrittenToForm).Select(x => $"Q{x.Quarter}: {x.IndividualScenarios:N0}"));
            MessageBox.Show(this, $"Созданы два Excel-файла.\n\nИндивидуальных сценариев: {ReportResult.IndividualScenarios:N0}\nВ форму записано: {written}", "Мониторинг готов", MessageBoxButton.OK, MessageBoxImage.Information);
            DialogResult = true;
        }
        catch (OperationCanceledException) { StatusText.Text = "Формирование отменено"; }
        catch (Exception exception) { StatusText.Text = "Не удалось сформировать мониторинг"; MessageBox.Show(this, exception.Message, "Ошибка мониторинга", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { _cancellation?.Dispose(); _cancellation = null; if (IsVisible) SetBusy(false); }
    }

    private static Task<MonitoringReportResult> RunReportAsync(MonitoringReportRequest request, IProgress<MonitoringReportProgress> progress, CancellationToken cancellationToken)
    {
        var forceWorker = string.Equals(Environment.GetEnvironmentVariable("ANALIZATOR_FORCE_WORKER"), "1", StringComparison.Ordinal);
        return Debugger.IsAttached || forceWorker
            ? MonitoringWorker.RunIsolatedAsync(request, progress, cancellationToken)
            : Task.Run(() => new MonitoringReportService().Run(request, progress, cancellationToken), cancellationToken);
    }

    private void SetBusy(bool busy)
    {
        _isBusy = busy; StartDatePicker.IsEnabled = !busy; EndDatePicker.IsEnabled = !busy; RefreshScopeButton.IsEnabled = !busy;
        InstallationsListBox.IsEnabled = !busy; SelectAllButton.IsEnabled = !busy; ClearButton.IsEnabled = !busy; BrowseTemplateButton.IsEnabled = !busy;
        IncludePossibleDuplicatesCheckBox.IsEnabled = !busy; OpenResultsCheckBox.IsEnabled = !busy; CancelButton.IsEnabled = busy; UpdateSelection();
    }
    private void CancelButton_Click(object sender, RoutedEventArgs e) => _cancellation?.Cancel();
    private void Window_Closing(object? sender, CancelEventArgs e) { if (_isBusy) { _cancellation?.Cancel(); e.Cancel = true; } }
}

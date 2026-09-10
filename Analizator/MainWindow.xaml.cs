using Analizator.Core;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Analizator;

public partial class MainWindow : Window
{
    private static readonly Brush ReadyBackground = new SolidColorBrush(Color.FromRgb(0xE8, 0xF6, 0xEF));
    private static readonly Brush ReadyForeground = new SolidColorBrush(Color.FromRgb(0x16, 0x76, 0x4D));
    private static readonly Brush WorkingBackground = new SolidColorBrush(Color.FromRgb(0xEA, 0xF0, 0xFF));
    private static readonly Brush WorkingForeground = new SolidColorBrush(Color.FromRgb(0x2E, 0x57, 0xC8));
    private static readonly Brush AttentionBackground = new SolidColorBrush(Color.FromRgb(0xFF, 0xF4, 0xE5));
    private static readonly Brush AttentionForeground = new SolidColorBrush(Color.FromRgb(0xA1, 0x5C, 0x00));
    private static readonly Brush ErrorBackground = new SolidColorBrush(Color.FromRgb(0xFD, 0xEC, 0xEC));
    private static readonly Brush ErrorForeground = new SolidColorBrush(Color.FromRgb(0xB4, 0x23, 0x18));

    private readonly ConfigFileItem[] _configFiles =
    [
        new("installations.json", "Установки"),
        new("settings.json", "Настройки анализа"),
        new("fio_keys.json", "Известные сокращения ФИО"),
        new("excluded_people.json", "Исключения")
    ];

    private readonly List<FileInfo> _logs = [];
    private readonly List<FileInfo> _databaseFiles = [];
    private CancellationTokenSource? _cancellation;
    private CancellationTokenSource? _databaseCancellation;
    private ConfigurationBundle? _configuration;
    private TrainingDatabaseService? _database;
    private string? _configPath;
    private string? _logPath;
    private bool _loaded;
    private bool _isProcessing;
    private bool _isDatabaseImporting;

    public MainWindow() => InitializeComponent();

    private string DataFolder => EngineFolderTextBox.Text.Trim();
    private string InputFolder => Path.Combine(DataFolder, "input");
    private string ConfigFolder => Path.Combine(DataFolder, "config");
    private string LogsFolder => Path.Combine(DataFolder, "logs");
    private string DatabaseFolder => Path.Combine(DataFolder, "data");
    private string DatabasePath => Path.Combine(DatabaseFolder, "analizator.db");
    private bool UseDatabaseForAnalysis => AnalysisSourceComboBox.SelectedIndex == 1;
    private DatabaseMonthSummary? SelectedDatabaseMonth =>
        DatabaseMonthComboBox.SelectedItem as DatabaseMonthSummary;

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        RuntimeTextBox.Text = ".NET 8 · встроенное C#-ядро";
        RuntimeTextBox.IsReadOnly = true;
        EngineFolderTextBox.Text = LocateDefaultDataFolder();
        OutputFolderTextBox.Text = Path.Combine(DataFolder, "output");
        EnsureDataDirectories();
        _loaded = true;
        RefreshFiles();
        LoadConfigList();
        LoadLogs();
        await CreateScheduledDatabaseBackupIfDueAsync();
        UpdateEngineStatus();
    }

    private static string LocateDefaultDataFolder()
    {
        var legacy = @"C:\Users\MSI\Desktop\АП";
        if (Directory.Exists(Path.Combine(legacy, "input")))
            return legacy;
        return AppContext.BaseDirectory;
    }

    private void EnsureDataDirectories()
    {
        if (string.IsNullOrWhiteSpace(DataFolder))
            return;
        Directory.CreateDirectory(InputFolder);
        Directory.CreateDirectory(ConfigFolder);
        Directory.CreateDirectory(LogsFolder);
        Directory.CreateDirectory(DatabaseFolder);
        Directory.CreateDirectory(Path.Combine(DataFolder, "output"));
        _configuration = ConfigurationService.Load(ConfigFolder);
        _database = new TrainingDatabaseService(
            DatabasePath, _configuration.FioKeys, _configuration.Installations);
        UpdatePassingThresholdSummary();
        UpdateDatabaseBackupSummary();
    }

    private void ShowView(UIElement view, Button button, string status)
    {
        DashboardView.Visibility = Visibility.Collapsed;
        DatabaseView.Visibility = Visibility.Collapsed;
        SettingsView.Visibility = Visibility.Collapsed;
        LogsView.Visibility = Visibility.Collapsed;
        view.Visibility = Visibility.Visible;
        DashboardNavButton.Background = Brushes.Transparent;
        DatabaseNavButton.Background = Brushes.Transparent;
        StatisticsNavButton.Background = Brushes.Transparent;
        MonitoringNavButton.Background = Brushes.Transparent;
        PdfPhotoNavButton.Background = Brushes.Transparent;
        SettingsNavButton.Background = Brushes.Transparent;
        LogsNavButton.Background = Brushes.Transparent;
        button.Background = (Brush)FindResource("SidebarHoverBrush");
        FooterStatusText.Text = status;
    }

    private void DashboardNavButton_Click(object sender, RoutedEventArgs e) =>
        ShowView(DashboardView, DashboardNavButton, "Подготовка обработки");

    private void DatabaseNavButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshDatabaseView();
        ShowView(DatabaseView, DatabaseNavButton, "Локальная база данных");
    }

    private void StatisticsNavButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting || _isProcessing)
        {
            MessageBox.Show(this, "Дождитесь завершения текущей операции.",
                "Операция выполняется", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database ??= new TrainingDatabaseService(DatabasePath);
            _database.SetPersonNameMappings(
                _configuration.FioKeys, _configuration.Installations);
            if (_database.GetSummary().Attempts == 0)
            {
                MessageBox.Show(this, "Сначала загрузите данные в локальную базу.",
                    "База данных пуста", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            StatisticsNavButton.Background = (Brush)FindResource("SidebarHoverBrush");
            var dialog = new StatisticsReportWindow(
                _database,
                ConfigFolder,
                Path.Combine(DataFolder, "output"))
            {
                Owner = this
            };
            if (dialog.ShowDialog() == true && dialog.ReportResult is not null)
            {
                FooterStatusText.Text =
                    $"Статистика сформирована: {Path.GetFileName(dialog.ReportResult.OutputFile)}";
            }
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось открыть статистику",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            StatisticsNavButton.Background = Brushes.Transparent;
            var active = DashboardView.Visibility == Visibility.Visible
                ? DashboardNavButton
                : DatabaseView.Visibility == Visibility.Visible
                    ? DatabaseNavButton
                    : SettingsView.Visibility == Visibility.Visible
                        ? SettingsNavButton
                        : LogsNavButton;
            active.Background = (Brush)FindResource("SidebarHoverBrush");
        }
    }

    private void MonitoringNavButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting || _isProcessing)
        {
            MessageBox.Show(this, "Дождитесь завершения текущей операции.",
                "Операция выполняется", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database ??= new TrainingDatabaseService(DatabasePath);
            _database.SetPersonNameMappings(
                _configuration.FioKeys, _configuration.Installations);
            if (_database.GetSummary().Attempts == 0)
            {
                MessageBox.Show(this, "Сначала загрузите данные в локальную базу.",
                    "База данных пуста", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            MonitoringNavButton.Background = (Brush)FindResource("SidebarHoverBrush");
            var dialog = new MonitoringReportWindow(
                _database,
                Path.Combine(DataFolder, "output"),
                DataFolder)
            {
                Owner = this
            };
            if (dialog.ShowDialog() == true && dialog.ReportResult is not null)
                FooterStatusText.Text = $"Мониторинг сформирован: {Path.GetFileName(dialog.ReportResult.ReportOutputFile)}";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось открыть мониторинг",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            MonitoringNavButton.Background = Brushes.Transparent;
            var active = DashboardView.Visibility == Visibility.Visible
                ? DashboardNavButton
                : DatabaseView.Visibility == Visibility.Visible
                    ? DatabaseNavButton
                    : SettingsView.Visibility == Visibility.Visible
                        ? SettingsNavButton
                        : LogsNavButton;
            active.Background = (Brush)FindResource("SidebarHoverBrush");
        }
    }

    private void PdfPhotoNavButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting || _isProcessing)
        {
            MessageBox.Show(this, "Дождитесь завершения текущей операции.",
                "Операция выполняется", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            PdfPhotoNavButton.Background = (Brush)FindResource("SidebarHoverBrush");
            var dialog = new PdfPhotoReportWindow(Path.Combine(DataFolder, "output"), DataFolder)
            {
                Owner = this
            };
            if (dialog.ShowDialog() == true && dialog.ReportResult is not null)
                FooterStatusText.Text = $"PDF проверены: {dialog.ReportResult.PdfFiles:N0}";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось открыть проверку PDF",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            PdfPhotoNavButton.Background = Brushes.Transparent;
            var active = DashboardView.Visibility == Visibility.Visible
                ? DashboardNavButton
                : DatabaseView.Visibility == Visibility.Visible
                    ? DatabaseNavButton
                    : SettingsView.Visibility == Visibility.Visible
                        ? SettingsNavButton
                        : LogsNavButton;
            active.Background = (Brush)FindResource("SidebarHoverBrush");
        }
    }

    private void SettingsNavButton_Click(object sender, RoutedEventArgs e)
    {
        LoadConfigList();
        UpdatePassingThresholdSummary();
        UpdateDatabaseBackupSummary();
        ShowView(SettingsView, SettingsNavButton, "Настройки программы");
    }

    private void LogsNavButton_Click(object sender, RoutedEventArgs e)
    {
        LoadLogs();
        ShowView(LogsView, LogsNavButton, "Просмотр журналов");
    }

    private void RefreshDatabaseButton_Click(object sender, RoutedEventArgs e) => RefreshDatabaseView();

    private void RefreshDatabaseView()
    {
        try
        {
            _database ??= new TrainingDatabaseService(DatabasePath);
            var summary = _database.GetSummary();
            var duplicateSummary = _database.GetDuplicateSummary();
            DatabaseAttemptsText.Text = summary.Attempts.ToString("N0");
            DatabasePeopleText.Text = summary.People.ToString("N0");
            DatabaseFilesCountText.Text = summary.SourceFiles.ToString("N0");
            DatabasePeriodText.Text = summary.PeriodStart.HasValue && summary.PeriodEnd.HasValue
                ? $"{summary.PeriodStart.Value:dd.MM.yyyy}\n{summary.PeriodEnd.Value:dd.MM.yyyy}"
                : "—";
            DatabaseHistoryDataGrid.ItemsSource = _database.GetImportHistory();
            DatabasePathText.Text = _database.DatabasePath;
            DatabaseDuplicateText.Text = summary.PossibleDuplicates == 0
                ? "Повторов не найдено. Проверяются также совпадения внутри одного JSON-файла."
                : $"Повторов: {duplicateSummary.Total:N0}. Точных: {duplicateSummary.ExactDuplicates:N0}; " +
                  $"требуют проверки: {duplicateSummary.PossibleDuplicates:N0}. " +
                  "Точные исключаются из анализа, но сохраняются в базе.";
            DatabaseDuplicateBorder.Background = summary.PossibleDuplicates == 0
                ? ReadyBackground
                : AttentionBackground;
            DatabaseDuplicateBorder.BorderBrush = summary.PossibleDuplicates == 0
                ? new SolidColorBrush(Color.FromRgb(0xBC, 0xE4, 0xD0))
                : new SolidColorBrush(Color.FromRgb(0xF2, 0xCF, 0x94));
            DatabaseDuplicateText.Foreground = summary.PossibleDuplicates == 0
                ? ReadyForeground
                : AttentionForeground;
            var unmatchedInstallations = _database.GetUnmatchedInstallations();
            DatabaseInstallationWarningBorder.Visibility = unmatchedInstallations.Count == 0
                ? Visibility.Collapsed
                : Visibility.Visible;
            if (unmatchedInstallations.Count > 0)
            {
                var names = string.Join(", ", unmatchedInstallations.Take(8)
                    .Select(item => $"{item.Name} ({item.Attempts:N0})"));
                if (unmatchedInstallations.Count > 8)
                    names += $" и ещё {unmatchedInstallations.Count - 8}";
                DatabaseInstallationWarningText.Text =
                    "Не удалось сопоставить установки с перечнем в настройках: " + names + ". " +
                    "Эти записи исключены из поиска, статистики и выгрузок. " +
                    "Добавьте названия как синонимы нужных установок в настройках.";
            }
            if (_loaded && UseDatabaseForAnalysis)
                RefreshDatabaseMonths();
            FooterStatusText.Text = "Данные базы обновлены";
        }
        catch (Exception exception)
        {
            DatabaseImportStatusText.Text = "База недоступна: " + exception.Message;
            FooterStatusText.Text = "Не удалось открыть базу данных";
        }
    }

    private void SelectDatabaseFilesButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Выберите выгрузки КТК",
            Filter = "Выгрузки КТК (*.xlsx;*.xlsm;*.json)|*.xlsx;*.xlsm;*.json|" +
                     "Книги Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|JSON (*.json)|*.json|Все файлы (*.*)|*.*",
            Multiselect = true,
            InitialDirectory = Directory.Exists(InputFolder)
                ? InputFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        };
        if (dialog.ShowDialog(this) != true)
            return;

        foreach (var path in dialog.FileNames)
            if (!_databaseFiles.Any(file =>
                    string.Equals(file.FullName, path, StringComparison.OrdinalIgnoreCase)))
                _databaseFiles.Add(new FileInfo(path));
        RefreshDatabaseFileList();
    }

    private void ClearDatabaseFilesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting)
            return;
        _databaseFiles.Clear();
        RefreshDatabaseFileList();
    }

    private void RefreshDatabaseFileList()
    {
        DatabaseFilesListBox.ItemsSource = null;
        DatabaseFilesListBox.ItemsSource = _databaseFiles.ToArray();
        DatabaseImportStatusText.Text = _databaseFiles.Count == 0
            ? "Выберите одну или несколько выгрузок"
            : $"Выбрано файлов: {_databaseFiles.Count}";
        ImportDatabaseFilesButton.IsEnabled = _databaseFiles.Count > 0 && !_isDatabaseImporting;
    }

    private async void ImportDatabaseFilesButton_Click(object sender, RoutedEventArgs e)
    {
        if (_databaseFiles.Count == 0 || _isDatabaseImporting)
            return;

        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database ??= new TrainingDatabaseService(DatabasePath);
            _database.SetPersonNameMappings(
                _configuration.FioKeys, _configuration.Installations);
            var database = _database;
            var files = _databaseFiles.Select(file => file.FullName).ToArray();
            var installations = _configuration.Installations;
            _databaseCancellation = new CancellationTokenSource();
            _isDatabaseImporting = true;
            SetDatabaseImportControls(false);
            DatabaseImportProgressBar.Value = 0;
            DatabaseImportStatusText.Text = "Подготовка импорта…";
            FooterStatusText.Text = "Загрузка выгрузок в базу данных";

            var uiProgress = new Progress<DatabaseImportProgress>(item =>
            {
                DatabaseImportProgressBar.Value = Math.Clamp(item.Percent, 0, 100);
                DatabaseImportStatusText.Text = item.FileName is null
                    ? item.Stage
                    : $"{item.Stage}: {item.FileName}";
            });
            var result = await RunDatabaseImportAsync(
                database, files, installations, uiProgress, _databaseCancellation.Token);

            _databaseFiles.Clear();
            RefreshDatabaseFileList();
            RefreshDatabaseView();
            DatabaseImportProgressBar.Value = 100;
            DatabaseImportStatusText.Text =
                $"Обработано файлов: {result.ImportedFiles}; обновлено: {result.UpdatedFiles}; " +
                $"попыток: {result.ImportedAttempts:N0}; " +
                $"повторно выбранных файлов: {result.SkippedFiles}; ошибок файлов: {result.FailedFiles}.";
            FooterStatusText.Text = result.FailedFiles == 0
                ? "Импорт в базу завершён"
                : "Импорт завершён с замечаниями";

            if (result.FailedFiles > 0 || result.RowErrors > 0)
            {
                var details = string.Join(Environment.NewLine,
                    result.Files.Where(file => file.Status == "Ошибка" || file.Errors > 0)
                        .Take(8)
                        .Select(file => $"{file.FileName}: {file.Message}"));
                MessageBox.Show(this,
                    DatabaseImportStatusText.Text +
                    (details.Length == 0 ? "" : Environment.NewLine + Environment.NewLine + details),
                    "Импорт завершён с замечаниями", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (OperationCanceledException)
        {
            DatabaseImportStatusText.Text = "Импорт отменён. Уже завершённые файлы сохранены.";
            FooterStatusText.Text = "Импорт в базу отменён";
            RefreshDatabaseView();
        }
        catch (Exception exception)
        {
            DatabaseImportStatusText.Text = "Ошибка импорта: " + exception.Message;
            FooterStatusText.Text = "Ошибка импорта в базу";
            MessageBox.Show(this, exception.Message, "Ошибка импорта",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _isDatabaseImporting = false;
            SetDatabaseImportControls(true);
            _databaseCancellation?.Dispose();
            _databaseCancellation = null;
        }
    }

    private void SetDatabaseImportControls(bool enabled)
    {
        SelectDatabaseFilesButton.IsEnabled = enabled;
        ClearDatabaseFilesButton.IsEnabled = enabled;
        ImportDatabaseFilesButton.IsEnabled = enabled && _databaseFiles.Count > 0;
        CancelDatabaseImportButton.IsEnabled = !enabled;
        DeleteDatabaseSourceButton.IsEnabled = enabled;
        EditDatabaseRecordsButton.IsEnabled = enabled;
        ExportDatabaseSelectionButton.IsEnabled = enabled;
    }

    private static Task<DatabaseImportResult> RunDatabaseImportAsync(
        TrainingDatabaseService database,
        string[] files,
        Dictionary<string, List<string>> installations,
        IProgress<DatabaseImportProgress> progress,
        CancellationToken cancellationToken)
    {
        var forceWorker = string.Equals(
            Environment.GetEnvironmentVariable("ANALIZATOR_FORCE_WORKER"),
            "1",
            StringComparison.Ordinal);
        if (Debugger.IsAttached || forceWorker)
        {
            return DatabaseImportWorker.RunIsolatedAsync(
                new DatabaseImportWorkerRequest(database.DatabasePath, files, installations),
                progress,
                cancellationToken);
        }

        return Task.Run(
            () => database.ImportFiles(files, installations, progress, cancellationToken),
            cancellationToken);
    }

    private void CancelDatabaseImportButton_Click(object sender, RoutedEventArgs e) =>
        _databaseCancellation?.Cancel();

    private void OpenDatabaseFolderButton_Click(object sender, RoutedEventArgs e) =>
        OpenFolder(DatabaseFolder);

    private void ExportDatabaseSelectionButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting)
            return;
        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database ??= new TrainingDatabaseService(DatabasePath);
            _database.SetPersonNameMappings(
                _configuration.FioKeys, _configuration.Installations);
            if (_database.GetSummary().Attempts == 0)
            {
                MessageBox.Show(this,
                    "Сначала загрузите в базу хотя бы одну выгрузку КТК.",
                    "База данных пуста",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var dialog = new DatabaseExportWindow(_database, Path.Combine(DataFolder, "output"))
            {
                Owner = this
            };
            if (dialog.ShowDialog() == true && dialog.ExportResult is not null)
            {
                FooterStatusText.Text =
                    $"Выгрузка сформирована: {Path.GetFileName(dialog.ExportResult.OutputFile)}";
            }
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось открыть выгрузку",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void EditDatabaseRecordsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting)
            return;
        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database ??= new TrainingDatabaseService(DatabasePath);
            _database.SetPersonNameMappings(
                _configuration.FioKeys, _configuration.Installations);
            var editor = new DatabaseEditorWindow(
                _database,
                Path.Combine(DatabaseFolder, "backups"),
                ConfigFolder,
                _configuration.Installations.Keys,
                _configuration.Settings.DatabaseBackups,
                _configuration.Settings.DatabaseEditor)
            {
                Owner = this
            };
            editor.ShowDialog();
            RefreshDatabaseView();
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось открыть редактор базы",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void DeleteDatabaseSourceButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting || DatabaseHistoryDataGrid.SelectedItem is not DatabaseSourceItem source)
        {
            MessageBox.Show(this, "Выберите загруженный файл в таблице.", "База данных",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _configuration = ConfigurationService.Load(ConfigFolder);
        var backupSettings = _configuration.Settings.DatabaseBackups;
        if (backupSettings.ConfirmSourceDeletion)
        {
            var backupMessage = backupSettings.BackupBeforeChanges
                ? "\n\nПеред удалением программа создаст резервную копию базы."
                : "";
            var answer = MessageBox.Show(this,
                $"Удалить из базы все {source.RowsCount:N0} записей файла «{source.FileName}»?" +
                backupMessage,
                "Удаление загруженного файла", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (answer != MessageBoxResult.Yes)
                return;
        }

        try
        {
            _database ??= new TrainingDatabaseService(DatabasePath);
            SetDatabaseImportControls(false);
            DatabaseImportStatusText.Text = backupSettings.BackupBeforeChanges
                ? "Создание резервной копии и удаление…"
                : "Удаление данных…";
            var backupDirectory = Path.Combine(DatabaseFolder, "backups");
            var backup = await Task.Run(() =>
            {
                var backupPath = backupSettings.BackupBeforeChanges
                    ? _database.CreateBackup(backupDirectory, backupSettings.MaximumBackupFiles)
                    : null;
                _database.DeleteSourceFile(source.Id);
                return backupPath;
            });
            RefreshDatabaseView();
            DatabaseImportStatusText.Text = backup is null
                ? "Файл удалён из базы."
                : $"Файл удалён из базы. Резервная копия: {Path.GetFileName(backup)}";
            FooterStatusText.Text = "Данные выбранного файла удалены";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось удалить данные",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetDatabaseImportControls(true);
        }
    }

    private void AutoSelectCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (_loaded && AutoSelectCheckBox.IsChecked == true)
            RefreshFiles();
    }

    private void AnalysisSourceComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ExcelSourcePanel is null || DatabaseSourcePanel is null)
            return;
        var fromDatabase = UseDatabaseForAnalysis;
        ExcelSourcePanel.Visibility = fromDatabase ? Visibility.Collapsed : Visibility.Visible;
        DatabaseSourcePanel.Visibility = fromDatabase ? Visibility.Visible : Visibility.Collapsed;
        AnalysisSourceDescriptionText.Text = fromDatabase
            ? "Шаблон анализа и выбранный месяц локальной базы"
            : "Шаблон анализа и Excel-выгрузка";
        AutoSelectCheckBox.Content = fromDatabase
            ? "Автовыбор шаблона из input"
            : "Автовыбор из input";
        if (_loaded && fromDatabase)
            RefreshDatabaseMonths();
        if (_loaded)
            UpdateEngineStatus();
    }

    private void RefreshDatabaseMonthsButton_Click(object sender, RoutedEventArgs e) =>
        RefreshDatabaseMonths();

    private void RefreshDatabaseMonths()
    {
        try
        {
            var previous = SelectedDatabaseMonth?.FirstDay;
            _database ??= new TrainingDatabaseService(DatabasePath);
            var months = _database.GetAvailableMonths();
            DatabaseMonthComboBox.ItemsSource = months;
            DatabaseMonthComboBox.SelectedItem = previous.HasValue
                ? months.FirstOrDefault(item => item.FirstDay == previous.Value) ?? months.FirstOrDefault()
                : months.FirstOrDefault();
            UpdateDatabaseMonthDetails();
        }
        catch (Exception exception)
        {
            DatabaseMonthComboBox.ItemsSource = null;
            DatabaseMonthDetailsText.Text = "Не удалось прочитать базу: " + exception.Message;
        }
    }

    private void DatabaseMonthComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateDatabaseMonthDetails();
        if (_loaded)
            UpdateEngineStatus();
    }

    private void UpdateDatabaseMonthDetails()
    {
        if (SelectedDatabaseMonth is not { } month)
        {
            DatabaseMonthDetailsText.Text = "В базе пока нет доступных месяцев";
            return;
        }

        DatabaseMonthDetailsText.Text =
            $"Будет учтено: {month.Attempts:N0} · точных повторов исключено: " +
            $"{month.ExactDuplicates:N0} · возможных повторов: {month.PossibleDuplicates:N0} · " +
            $"без режима: {month.ReportsWithoutMode:N0}";
    }

    private void RefreshFilesButton_Click(object sender, RoutedEventArgs e) => RefreshFiles();

    private void RefreshFiles()
    {
        if (!Directory.Exists(InputFolder))
        {
            AnalysisFileTextBox.Text = "";
            ExportFileTextBox.Text = "";
            FooterStatusText.Text = "Папка input не найдена";
            UpdateEngineStatus();
            return;
        }
        var files = new DirectoryInfo(InputFolder)
            .EnumerateFiles()
            .Where(file => file.Extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                           file.Extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            .Where(file => !file.Name.StartsWith("~$", StringComparison.OrdinalIgnoreCase))
            .OrderBy(file => file.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
        var analysis = files.FirstOrDefault(file =>
            TextNormalization.NormalizeText(file.Name).Contains("анализ"));
        var export = files.FirstOrDefault(file =>
            TextNormalization.NormalizeText(file.Name).Contains("выгрузка"));
        AnalysisFileTextBox.Text = analysis?.FullName ?? "";
        ExportFileTextBox.Text = export?.FullName ?? "";
        FooterStatusText.Text = "Поиск файлов завершён";
        UpdateEngineStatus();
    }

    private string? SelectExcel(string title)
    {
        var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = "Книги Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Все файлы (*.*)|*.*",
            InitialDirectory = Directory.Exists(InputFolder)
                ? InputFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        };
        return dialog.ShowDialog(this) == true ? dialog.FileName : null;
    }

    private void BrowseAnalysisButton_Click(object sender, RoutedEventArgs e)
    {
        var path = SelectExcel("Выберите таблицу анализа");
        if (path is null)
            return;
        AutoSelectCheckBox.IsChecked = false;
        AnalysisFileTextBox.Text = path;
        UpdateEngineStatus();
    }

    private void BrowseExportButton_Click(object sender, RoutedEventArgs e)
    {
        var path = SelectExcel("Выберите выгрузку КТК");
        if (path is null)
            return;
        AutoSelectCheckBox.IsChecked = false;
        ExportFileTextBox.Text = path;
        UpdateEngineStatus();
    }

    private async void StartButton_Click(object sender, RoutedEventArgs e)
    {
        if (!File.Exists(AnalysisFileTextBox.Text))
        {
            UpdateEngineStatus();
            MessageBox.Show(this, "Выберите существующую таблицу анализа.",
                "Не хватает данных", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (!UseDatabaseForAnalysis && !File.Exists(ExportFileTextBox.Text))
        {
            UpdateEngineStatus();
            MessageBox.Show(this, "Выберите существующую Excel-выгрузку КТК.",
                "Не хватает данных", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (UseDatabaseForAnalysis && (!File.Exists(DatabasePath) || SelectedDatabaseMonth is null))
        {
            UpdateEngineStatus();
            MessageBox.Show(this, "Выберите месяц, содержащий записи в базе данных.",
                "Не хватает данных", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (string.IsNullOrWhiteSpace(OutputFolderTextBox.Text))
        {
            UpdateEngineStatus();
            MessageBox.Show(this, "Укажите папку результата.", "Не хватает данных",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _cancellation = new CancellationTokenSource();
        _isProcessing = true;
        StartButton.IsEnabled = false;
        CancelButton.IsEnabled = true;
        SetSystemStatus("●  Обработка: 1%", "Запуск встроенного C#-ядра",
            WorkingBackground, WorkingForeground);
        SetProgress(1, "Запуск встроенного ядра");

        try
        {
            var request = new AnalyzerRequest(
                AnalysisFileTextBox.Text,
                ExportFileTextBox.Text,
                OutputFolderTextBox.Text.Trim(),
                ConfigFolder,
                CreateBackupCheckBox.IsChecked == true,
                AutoAddPeopleCheckBox.IsChecked == true,
                UseDatabaseForAnalysis
                    ? AnalysisDataSourceKind.Database
                    : AnalysisDataSourceKind.Excel,
                UseDatabaseForAnalysis ? DatabasePath : null,
                UseDatabaseForAnalysis ? SelectedDatabaseMonth?.FirstDay : null);
            var uiProgress = new Progress<AnalyzerProgress>(item => SetProgress(item.Percent, item.Stage));
            var result = await RunAnalysisAsync(request, uiProgress, _cancellation.Token);

            LastRunText.Text = $"{DateTime.Now:dd.MM.yyyy · HH:mm}\n" +
                               $"{Path.GetFileName(result.OutputFile)}\n" +
                               $"Отчётов: {result.Reports}; сотрудников: {result.MatchedPeople}";
            FooterStatusText.Text = "Результат успешно сформирован";
            SetSystemStatus("●  Анализ завершён", "Результат успешно сформирован",
                ReadyBackground, ReadyForeground);
            LoadLogs();
            if (OpenResultCheckBox.IsChecked == true)
                OpenWithShell(result.OutputFile);
        }
        catch (OperationCanceledException)
        {
            SetProgress((int)ProcessingProgressBar.Value, "Операция отменена");
            FooterStatusText.Text = "Отменено пользователем";
            SetSystemStatus("●  Операция отменена", "Можно запустить анализ повторно",
                AttentionBackground, AttentionForeground);
        }
        catch (Exception exception)
        {
            ProgressStageText.Text = "Обработка завершилась с ошибкой";
            FooterStatusText.Text = exception.Message;
            SetSystemStatus("●  Ошибка обработки", "Требуется внимание пользователя",
                ErrorBackground, ErrorForeground);
            MessageBox.Show(this, exception.Message, "Ошибка обработки",
                MessageBoxButton.OK, MessageBoxImage.Error);
            LoadLogs();
        }
        finally
        {
            _isProcessing = false;
            StartButton.IsEnabled = true;
            CancelButton.IsEnabled = false;
            _cancellation?.Dispose();
            _cancellation = null;
        }
    }

    private static Task<AnalyzerResult> RunAnalysisAsync(
        AnalyzerRequest request,
        IProgress<AnalyzerProgress> progress,
        CancellationToken cancellationToken)
    {
        // ClosedXML stays in-process in a normal launch. Under the Visual Studio debugger it runs
        // in an isolated C# worker so a debugger wait timeout cannot terminate the WPF process.
        var forceWorker = string.Equals(
            Environment.GetEnvironmentVariable("ANALIZATOR_FORCE_WORKER"),
            "1",
            StringComparison.Ordinal);

        if (Debugger.IsAttached || forceWorker)
            return AnalysisWorker.RunIsolatedAsync(request, progress, cancellationToken);

        return Task.Run(
            () => new AnalyzerService().Run(request, progress, cancellationToken),
            cancellationToken);
    }

    private void SetProgress(int value, string stage)
    {
        value = Math.Clamp(value, 0, 100);
        ProcessingProgressBar.Value = value;
        ProgressPercentText.Text = value + "%";
        ProgressStageText.Text = stage;
        if (_isProcessing)
        {
            SystemStatusText.Text = $"●  Обработка: {value}%";
            SidebarStatusText.Text = stage;
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => _cancellation?.Cancel();

    private void OutputFolderTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_loaded)
            UpdateEngineStatus();
    }

    private async void ApplyEngineFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isDatabaseImporting)
        {
            MessageBox.Show(this, "Дождитесь завершения импорта в базу данных или отмените его.",
                "Импорт выполняется", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            EnsureDataDirectories();
            OutputFolderTextBox.Text = Path.Combine(DataFolder, "output");
            RefreshFiles();
            LoadConfigList();
            LoadLogs();
            await CreateScheduledDatabaseBackupIfDueAsync();
            UpdateEngineStatus();
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Ошибка рабочей папки",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void EditPassingThresholdButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            var dialog = new PassingThresholdWindow(
                _configuration.Settings.PassingThreshold,
                _configuration.Installations.Keys)
            {
                Owner = this
            };
            if (dialog.ShowDialog() != true || dialog.ResultSettings is null)
                return;

            ConfigurationService.SavePassingThresholds(ConfigFolder, dialog.ResultSettings);
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database?.SetPersonNameMappings(
                _configuration.FioKeys, _configuration.Installations);
            UpdatePassingThresholdSummary();
            if (string.Equals(Path.GetFileName(_configPath), "settings.json", StringComparison.OrdinalIgnoreCase))
                LoadSelectedConfig();
            FooterStatusText.Text = "Порог прохождения сохранён";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Ошибка настройки порога",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void UpdatePassingThresholdSummary()
    {
        if (_configuration is null)
        {
            PassingThresholdSummaryText.Text = "Настройка недоступна";
            return;
        }

        var threshold = _configuration.Settings.PassingThreshold;
        PassingThresholdSummaryText.Text = threshold.UsePerInstallation
            ? $"Индивидуальный режим · {threshold.Installations.Count} установок"
            : $"Единый порог: {threshold.GlobalPercent:0.##}%";
    }

    private void EditDatabaseBackupSettingsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            var dialog = new DatabaseBackupSettingsWindow(
                _configuration.Settings.DatabaseBackups)
            {
                Owner = this
            };
            if (dialog.ShowDialog() != true || dialog.ResultSettings is null)
                return;

            ConfigurationService.SaveDatabaseBackupSettings(ConfigFolder, dialog.ResultSettings);
            _configuration = ConfigurationService.Load(ConfigFolder);
            UpdateDatabaseBackupSummary();
            if (string.Equals(Path.GetFileName(_configPath), "settings.json",
                    StringComparison.OrdinalIgnoreCase))
                LoadSelectedConfig();
            FooterStatusText.Text = "Настройки резервных копий сохранены";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Ошибка настройки резервных копий",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void CreateDatabaseBackupNowButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database ??= new TrainingDatabaseService(DatabasePath);
            var settings = _configuration.Settings.DatabaseBackups;
            FooterStatusText.Text = "Создание резервной копии базы…";
            var backup = await Task.Run(() => _database.CreateBackup(
                Path.Combine(DatabaseFolder, "backups"), settings.MaximumBackupFiles));
            FooterStatusText.Text = $"Создана резервная копия: {Path.GetFileName(backup)}";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось создать резервную копию",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void UpdateDatabaseBackupSummary()
    {
        if (_configuration is null)
        {
            DatabaseBackupSummaryText.Text = "Настройка недоступна";
            return;
        }

        var settings = _configuration.Settings.DatabaseBackups;
        var automatic = settings.AutomaticBackupsEnabled
            ? $"автоматически раз в {settings.IntervalDays} дн."
            : "автоматическое копирование выключено";
        DatabaseBackupSummaryText.Text =
            $"{automatic} · хранить {settings.MaximumBackupFiles} · " +
            (settings.BackupBeforeChanges ? "копия перед изменением" : "без копии перед изменением");
    }

    private async Task CreateScheduledDatabaseBackupIfDueAsync()
    {
        try
        {
            if (_configuration is null || _database is null ||
                _database.GetSummary().Attempts == 0)
                return;
            var settings = _configuration.Settings.DatabaseBackups;
            if (!settings.AutomaticBackupsEnabled)
                return;

            var backupDirectory = Path.Combine(DatabaseFolder, "backups");
            var latestBackup = Directory.Exists(backupDirectory)
                ? new DirectoryInfo(backupDirectory)
                    .EnumerateFiles("analizator_*.db", SearchOption.TopDirectoryOnly)
                    .OrderByDescending(file => file.LastWriteTimeUtc)
                    .FirstOrDefault()
                : null;
            if (latestBackup is not null &&
                DateTime.UtcNow - latestBackup.LastWriteTimeUtc < TimeSpan.FromDays(settings.IntervalDays))
                return;

            var backup = await Task.Run(() => _database.CreateBackup(
                backupDirectory, settings.MaximumBackupFiles));
            FooterStatusText.Text = $"Автоматическая копия базы: {Path.GetFileName(backup)}";
        }
        catch (Exception exception)
        {
            FooterStatusText.Text = "Не удалось создать автоматическую копию базы";
            Debug.WriteLine(exception);
        }
    }

    private void UpdateEngineStatus()
    {
        if (_isProcessing || !_loaded)
            return;

        if (!Directory.Exists(DataFolder) || !Directory.Exists(ConfigFolder))
        {
            SetSystemStatus("●  Система не готова", "Проверьте рабочую папку",
                ErrorBackground, ErrorForeground);
            return;
        }

        if (!File.Exists(AnalysisFileTextBox.Text) ||
            (!UseDatabaseForAnalysis && !File.Exists(ExportFileTextBox.Text)) ||
            (UseDatabaseForAnalysis && SelectedDatabaseMonth is null))
        {
            SetSystemStatus("●  Нужны исходные данные",
                UseDatabaseForAnalysis
                    ? "Выберите таблицу анализа и месяц базы данных"
                    : "Выберите таблицу анализа и выгрузку КТК",
                AttentionBackground, AttentionForeground);
            return;
        }

        if (string.IsNullOrWhiteSpace(OutputFolderTextBox.Text))
        {
            SetSystemStatus("●  Укажите папку результата", "Папка результата не выбрана",
                AttentionBackground, AttentionForeground);
            return;
        }

        SetSystemStatus("●  Система готова", "Все данные выбраны · C#-ядро готово",
            ReadyBackground, ReadyForeground);
    }

    private void SetSystemStatus(
        string badgeText,
        string sidebarText,
        Brush background,
        Brush foreground)
    {
        SystemStatusText.Text = badgeText;
        SystemStatusText.Foreground = foreground;
        SystemStatusBadge.Background = background;
        SidebarStatusText.Text = sidebarText;
    }

    private void LoadConfigList()
    {
        if (!_loaded)
            return;
        ConfigFilesListBox.Items.Clear();
        foreach (var item in _configFiles)
            ConfigFilesListBox.Items.Add(item);
        if (ConfigFilesListBox.Items.Count > 0)
            ConfigFilesListBox.SelectedIndex = 0;
    }

    private void ConfigFilesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ConfigFilesListBox.SelectedItem is not ConfigFileItem item)
            return;
        _configPath = Path.Combine(ConfigFolder, item.FileName);
        LoadSelectedConfig();
    }

    private void LoadSelectedConfig()
    {
        if (_configPath is null)
            return;
        var fileName = Path.GetFileName(_configPath);
        var displayName = _configFiles.FirstOrDefault(item =>
            item.FileName.Equals(fileName, StringComparison.OrdinalIgnoreCase))?.DisplayName ?? fileName;
        ConfigEditor.LoadConfiguration(ConfigFolder, fileName, displayName);
    }

    private void ConfigEditor_Saved(object? sender, EventArgs e)
    {
        try
        {
            _configuration = ConfigurationService.Load(ConfigFolder);
            _database?.SetPersonNameMappings(
                _configuration.FioKeys, _configuration.Installations);
            UpdatePassingThresholdSummary();
            UpdateDatabaseBackupSummary();
            if (DatabaseView.Visibility == Visibility.Visible)
                RefreshDatabaseView();
            FooterStatusText.Text = "Конфигурация сохранена";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Ошибка загрузки конфигурации",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadLogs()
    {
        if (!_loaded)
            return;
        _logs.Clear();
        if (Directory.Exists(LogsFolder))
            _logs.AddRange(new DirectoryInfo(LogsFolder).GetFiles("*.txt").OrderByDescending(file => file.LastWriteTime));
        ApplyLogFilter();
    }

    private void ApplyLogFilter()
    {
        var filter = LogSearchTextBox.Text.Trim();
        LogsListBox.ItemsSource = _logs.Where(file => string.IsNullOrEmpty(filter) ||
            file.Name.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    private void LogSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_loaded)
            ApplyLogFilter();
    }

    private void LogsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LogsListBox.SelectedItem is not FileInfo file)
            return;
        _logPath = file.FullName;
        LogTitleText.Text = file.Name;
        LogMetaText.Text = $"{file.LastWriteTime:dd.MM.yyyy HH:mm:ss} · {file.Length:N0} байт";
        try
        {
            LogContentTextBox.Text = File.ReadAllText(file.FullName, Encoding.UTF8);
            LogContentTextBox.ScrollToHome();
            OpenLogExternalButton.IsEnabled = true;
        }
        catch (Exception exception)
        {
            LogContentTextBox.Text = "Ошибка чтения:\r\n" + exception.Message;
        }
    }

    private void RefreshLogsButton_Click(object sender, RoutedEventArgs e)
    {
        LoadLogs();
        FooterStatusText.Text = "Список журналов обновлён";
    }

    private void OpenLatestLogButton_Click(object sender, RoutedEventArgs e)
    {
        LoadLogs();
        var latest = _logs.FirstOrDefault();
        if (latest is null)
        {
            MessageBox.Show(this, "Журналы пока отсутствуют.", "Журналы",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        OpenWithShell(latest.FullName);
    }

    private void OpenLogExternalButton_Click(object sender, RoutedEventArgs e)
    {
        if (_logPath is not null)
            OpenWithShell(_logPath);
    }

    private void OpenInputFolderButton_Click(object sender, RoutedEventArgs e) => OpenFolder(InputFolder);
    private void OpenOutputFolderButton_Click(object sender, RoutedEventArgs e) => OpenFolder(OutputFolderTextBox.Text.Trim());
    private void OpenConfigFolderButton_Click(object sender, RoutedEventArgs e) => OpenFolder(ConfigFolder);
    private void OpenLogsFolderButton_Click(object sender, RoutedEventArgs e) => OpenFolder(LogsFolder);

    private void OpenFolder(string path)
    {
        try
        {
            Directory.CreateDirectory(path);
            Process.Start(new ProcessStartInfo("explorer.exe", path) { UseShellExecute = true });
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось открыть папку",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static void OpenWithShell(string path)
    {
        if (File.Exists(path))
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }
}

internal sealed record ConfigFileItem(string FileName, string DisplayName)
{
    public override string ToString() => DisplayName;
}

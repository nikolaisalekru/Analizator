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

    private readonly string[] _configNames =
    [
        "installations.json", "settings.json", "fio_keys.json", "excluded_people.json"
    ];

    private readonly List<FileInfo> _logs = [];
    private CancellationTokenSource? _cancellation;
    private ConfigurationBundle? _configuration;
    private string? _configPath;
    private string? _logPath;
    private bool _loaded;
    private bool _isProcessing;

    public MainWindow() => InitializeComponent();

    private string DataFolder => EngineFolderTextBox.Text.Trim();
    private string InputFolder => Path.Combine(DataFolder, "input");
    private string ConfigFolder => Path.Combine(DataFolder, "config");
    private string LogsFolder => Path.Combine(DataFolder, "logs");

    private void Window_Loaded(object sender, RoutedEventArgs e)
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
        UpdateEngineStatus();
    }

    private static string LocateDefaultDataFolder()
    {
        var legacy = @"C:\Users\MSI\Desktop\АП";
        if (Directory.Exists(Path.Combine(legacy, "input")) && Directory.Exists(Path.Combine(legacy, "config")))
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
        Directory.CreateDirectory(Path.Combine(DataFolder, "output"));
        _configuration = ConfigurationService.Load(ConfigFolder);
        UpdatePassingThresholdSummary();
    }

    private void ShowView(UIElement view, Button button, string status)
    {
        DashboardView.Visibility = Visibility.Collapsed;
        SettingsView.Visibility = Visibility.Collapsed;
        LogsView.Visibility = Visibility.Collapsed;
        view.Visibility = Visibility.Visible;
        DashboardNavButton.Background = Brushes.Transparent;
        SettingsNavButton.Background = Brushes.Transparent;
        LogsNavButton.Background = Brushes.Transparent;
        button.Background = (Brush)FindResource("SidebarHoverBrush");
        FooterStatusText.Text = status;
    }

    private void DashboardNavButton_Click(object sender, RoutedEventArgs e) =>
        ShowView(DashboardView, DashboardNavButton, "Подготовка обработки");

    private void SettingsNavButton_Click(object sender, RoutedEventArgs e)
    {
        LoadConfigList();
        UpdatePassingThresholdSummary();
        ShowView(SettingsView, SettingsNavButton, "Настройки программы");
    }

    private void LogsNavButton_Click(object sender, RoutedEventArgs e)
    {
        LoadLogs();
        ShowView(LogsView, LogsNavButton, "Просмотр журналов");
    }

    private void AutoSelectCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (_loaded && AutoSelectCheckBox.IsChecked == true)
            RefreshFiles();
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
        if (!File.Exists(AnalysisFileTextBox.Text) || !File.Exists(ExportFileTextBox.Text))
        {
            UpdateEngineStatus();
            MessageBox.Show(this, "Выберите существующие файлы анализа и выгрузки.",
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
                AutoAddPeopleCheckBox.IsChecked == true);
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

    private void ApplyEngineFolderButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            EnsureDataDirectories();
            OutputFolderTextBox.Text = Path.Combine(DataFolder, "output");
            RefreshFiles();
            LoadConfigList();
            LoadLogs();
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

        if (!File.Exists(AnalysisFileTextBox.Text) || !File.Exists(ExportFileTextBox.Text))
        {
            SetSystemStatus("●  Нужны исходные файлы", "Выберите таблицу анализа и выгрузку КТК",
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
        foreach (var name in _configNames)
            ConfigFilesListBox.Items.Add(name);
        if (ConfigFilesListBox.Items.Count > 0)
            ConfigFilesListBox.SelectedIndex = 0;
    }

    private void ConfigFilesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ConfigFilesListBox.SelectedItem is null)
            return;
        _configPath = Path.Combine(ConfigFolder, ConfigFilesListBox.SelectedItem.ToString()!);
        LoadSelectedConfig();
    }

    private void LoadSelectedConfig()
    {
        if (_configPath is null)
            return;
        ConfigEditorTitleText.Text = Path.GetFileName(_configPath);
        ConfigEditorPathText.Text = _configPath;
        try
        {
            ConfigEditorTextBox.Text = File.Exists(_configPath)
                ? File.ReadAllText(_configPath, Encoding.UTF8)
                : "{\r\n  \r\n}";
            ConfigEditorTextBox.IsEnabled = true;
            SaveConfigButton.IsEnabled = true;
            ReloadConfigButton.IsEnabled = File.Exists(_configPath);
            ConfigStatusText.Text = File.Exists(_configPath) ? "Файл загружен" : "Будет создан при сохранении";
        }
        catch (Exception exception)
        {
            ConfigStatusText.Text = "Ошибка: " + exception.Message;
            ConfigEditorTextBox.IsEnabled = false;
        }
    }

    private void SaveConfigButton_Click(object sender, RoutedEventArgs e)
    {
        if (_configPath is null)
            return;
        try
        {
            // Validate all edited JSON before replacing a working configuration.
            using var _ = System.Text.Json.JsonDocument.Parse(ConfigEditorTextBox.Text);
            Directory.CreateDirectory(ConfigFolder);
            File.WriteAllText(_configPath, ConfigEditorTextBox.Text, new UTF8Encoding(false));
            _configuration = ConfigurationService.Load(ConfigFolder);
            UpdatePassingThresholdSummary();
            ConfigStatusText.Text = "Сохранено · " + DateTime.Now.ToString("HH:mm:ss");
            ReloadConfigButton.IsEnabled = true;
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Ошибка сохранения",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ReloadConfigButton_Click(object sender, RoutedEventArgs e) => LoadSelectedConfig();

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

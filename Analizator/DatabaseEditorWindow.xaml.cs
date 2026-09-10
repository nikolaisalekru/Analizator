using Analizator.Core;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Analizator;

public partial class DatabaseEditorWindow : Window
{
    private const int MaximumRows = int.MaxValue;
    private readonly TrainingDatabaseService _database;
    private readonly string _backupDirectory;
    private readonly string _configurationDirectory;
    private readonly string[] _installations;
    private readonly DatabaseBackupSettings _backupSettings;
    private HashSet<string> _hiddenColumns;

    public DatabaseEditorWindow(
        TrainingDatabaseService database,
        string backupDirectory,
        string configurationDirectory,
        IEnumerable<string> installations,
        DatabaseBackupSettings backupSettings,
        DatabaseEditorSettings editorSettings)
    {
        InitializeComponent();
        _database = database;
        _backupDirectory = backupDirectory;
        _configurationDirectory = configurationDirectory;
        _backupSettings = backupSettings;
        _hiddenColumns = new HashSet<string>(editorSettings.HiddenColumns,
            StringComparer.OrdinalIgnoreCase);
        _installations = installations.Distinct(StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase).ToArray();
        ApplyColumnVisibility();
        var summary = _database.GetSummary();
        StartDatePicker.SelectedDate = summary.PeriodStart?.Date;
        EndDatePicker.SelectedDate = summary.PeriodEnd?.Date;
        RefreshPage();
    }

    private DatabaseAttempt? SelectedAttempt => AttemptsDataGrid.SelectedItem as DatabaseAttempt;

    private void RefreshPage()
    {
        try
        {
            var page = _database.SearchAttempts(
                StartDatePicker.SelectedDate,
                EndDatePicker.SelectedDate,
                SearchTextBox.Text,
                MaximumRows,
                0,
                ParseOptionalTime(StartTimeTextBox.Text, "Время от", false),
                ParseOptionalTime(EndTimeTextBox.Text, "Время до", true),
                DuplicateFilterComboBox.SelectedIndex switch
                {
                    1 => true,
                    2 => false,
                    _ => null
                });
            AttemptsDataGrid.ItemsSource = page.Items;
            ResultText.Text = $"Показано записей: {page.TotalCount:N0}";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось прочитать базу",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SearchButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshPage();
    }

    private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;
        RefreshPage();
    }

    private void ClearFiltersButton_Click(object sender, RoutedEventArgs e)
    {
        SearchTextBox.Clear();
        StartDatePicker.SelectedDate = null;
        EndDatePicker.SelectedDate = null;
        StartTimeTextBox.Clear();
        EndTimeTextBox.Clear();
        DuplicateFilterComboBox.SelectedIndex = 0;
        RefreshPage();
    }

    private static TimeSpan? ParseOptionalTime(string? value, string fieldName, bool includeWholeMinute)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (!TimeSpan.TryParse(value.Trim(), CultureInfo.CurrentCulture, out var result) &&
            !TimeSpan.TryParse(value.Trim(), CultureInfo.InvariantCulture, out result))
            throw new ArgumentException($"{fieldName}: укажите время в формате ЧЧ:ММ.");
        if (result < TimeSpan.Zero || result >= TimeSpan.FromDays(1))
            throw new ArgumentException($"{fieldName}: время должно быть от 00:00 до 23:59.");
        if (includeWholeMinute && value.Count(character => character == ':') == 1)
            result = result.Add(TimeSpan.FromSeconds(59));
        return result;
    }

    private async void AddButton_Click(object sender, RoutedEventArgs e)
    {
        var editor = new DatabaseAttemptWindow(new DatabaseAttemptEditModel(), _installations)
        {
            Owner = this
        };
        if (editor.ShowDialog() != true || editor.Result is null)
            return;
        try
        {
            IsEnabled = false;
            await Task.Run(() => _database.AddAttempt(editor.Result));
            RefreshPage();
            ResultText.Text = "Запись добавлена";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось добавить запись",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsEnabled = true;
        }
    }

    private async void EditButton_Click(object sender, RoutedEventArgs e) => await EditSelectedAsync();

    private async void AttemptsDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e) =>
        await EditSelectedAsync();

    private async Task EditSelectedAsync()
    {
        if (SelectedAttempt is not { } selected)
        {
            MessageBox.Show(this, "Выберите запись в таблице.", "Редактор базы данных",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        var editor = new DatabaseAttemptWindow(DatabaseAttemptEditModel.FromAttempt(selected), _installations)
        {
            Owner = this
        };
        if (editor.ShowDialog() != true || editor.Result is null)
            return;
        await RunProtectedChangeAsync(
            () => _database.UpdateAttempt(editor.Result),
            "Запись изменена");
    }

    private async void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedAttempt is not { } selected)
        {
            MessageBox.Show(this, "Выберите запись в таблице.", "Редактор базы данных",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (_backupSettings.ConfirmRowDeletion)
        {
            var backupMessage = _backupSettings.BackupBeforeChanges
                ? "\n\nПеред удалением будет создана резервная копия базы."
                : "";
            var answer = MessageBox.Show(this,
                $"Удалить попытку «{selected.PersonName} — {selected.ScenarioName}» от " +
                $"{selected.AttemptedAt:dd.MM.yyyy HH:mm}?{backupMessage}",
                "Удаление записи", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (answer != MessageBoxResult.Yes)
                return;
        }
        await RunProtectedChangeAsync(
            () => _database.DeleteAttempt(selected.Id),
            "Запись удалена");
    }

    private void ColumnsButton_Click(object sender, RoutedEventArgs e)
    {
        var options = AttemptsDataGrid.Columns
            .Select(column => new DatabaseColumnOption
            {
                Key = ColumnKey(column),
                DisplayName = ColumnKey(column),
                IsVisible = column.Visibility == Visibility.Visible
            })
            .ToArray();
        var dialog = new DatabaseColumnsWindow(options) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.HiddenColumns is null)
            return;

        try
        {
            _hiddenColumns = dialog.HiddenColumns;
            ConfigurationService.SaveDatabaseEditorSettings(
                _configurationDirectory,
                new DatabaseEditorSettings { HiddenColumns = _hiddenColumns });
            ApplyColumnVisibility();
            ResultText.Text = _hiddenColumns.Count == 0
                ? "Показаны все столбцы"
                : $"Скрыто столбцов: {_hiddenColumns.Count}";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось сохранить столбцы",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ApplyColumnVisibility()
    {
        foreach (var column in AttemptsDataGrid.Columns)
            column.Visibility = _hiddenColumns.Contains(ColumnKey(column))
                ? Visibility.Collapsed
                : Visibility.Visible;
    }

    private static string ColumnKey(DataGridColumn column) =>
        column.Header?.ToString()?.Trim() is { Length: > 0 } header ? header : "Столбец";

    private async Task RunProtectedChangeAsync(Action change, string successText)
    {
        try
        {
            IsEnabled = false;
            var backup = await Task.Run(() =>
            {
                var backupPath = _backupSettings.BackupBeforeChanges
                    ? _database.CreateBackup(
                        _backupDirectory,
                        _backupSettings.MaximumBackupFiles)
                    : null;
                change();
                return backupPath;
            });
            RefreshPage();
            ResultText.Text = backup is null
                ? successText
                : $"{successText} · копия: {Path.GetFileName(backup)}";
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Не удалось изменить запись",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsEnabled = true;
        }
    }
}

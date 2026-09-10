using Analizator.Core;
using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;

namespace Analizator;

public partial class PdfPhotoReportWindow : Window
{
    private readonly string _outputDirectory;
    private CancellationTokenSource? _cancellation;
    private bool _isBusy;

    public PdfPhotoReportWindow(string outputDirectory, string dataFolder)
    {
        InitializeComponent();
        _outputDirectory = outputDirectory;
        var reports = Path.Combine(dataFolder, "reports");
        if (Directory.Exists(reports)) SourceFolderTextBox.Text = reports;
        var previousMonth = DateTime.Today.AddMonths(-1);
        StartDatePicker.SelectedDate = new DateTime(previousMonth.Year, previousMonth.Month, 1);
        EndDatePicker.SelectedDate = new DateTime(previousMonth.Year, previousMonth.Month,
            DateTime.DaysInMonth(previousMonth.Year, previousMonth.Month));
        UpdateState();
    }

    public PdfPhotoReportResult? ReportResult { get; private set; }

    private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "Выберите папку с PDF-отчётами",
            Multiselect = false
        };
        if (Directory.Exists(SourceFolderTextBox.Text)) dialog.InitialDirectory = SourceFolderTextBox.Text;
        if (dialog.ShowDialog(this) == true) SourceFolderTextBox.Text = dialog.FolderName;
    }

    private void SourceFolderTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) => UpdateState();
    private void UsePeriodCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        var enabled = UsePeriodCheckBox.IsChecked == true && !_isBusy;
        StartDatePicker.IsEnabled = enabled;
        EndDatePicker.IsEnabled = enabled;
    }

    private async void GenerateButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy) return;
        try
        {
            var source = SourceFolderTextBox.Text.Trim();
            if (!Directory.Exists(source)) throw new DirectoryNotFoundException("Выберите существующую папку с PDF-отчётами.");
            DateTime? start = UsePeriodCheckBox.IsChecked == true ? StartDatePicker.SelectedDate?.Date : null;
            DateTime? end = UsePeriodCheckBox.IsChecked == true ? EndDatePicker.SelectedDate?.Date : null;
            if (UsePeriodCheckBox.IsChecked == true && (!start.HasValue || !end.HasValue))
                throw new InvalidOperationException("Укажите обе даты периода.");
            Directory.CreateDirectory(_outputDirectory);
            var dialog = new SaveFileDialog
            {
                Title = "Сохранить таблицу контроля PDF",
                Filter = "Книга Excel (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                OverwritePrompt = true,
                InitialDirectory = _outputDirectory,
                FileName = $"Контроль_фотографий_PDF_{DateTime.Now:yyyy-MM-dd}.xlsx"
            };
            if (dialog.ShowDialog(this) != true) return;

            var request = new PdfPhotoReportRequest(source, dialog.FileName, start, end);
            _cancellation = new CancellationTokenSource();
            SetBusy(true);
            var progress = new Progress<PdfPhotoProgress>(item =>
            {
                ProgressBar.Value = Math.Clamp(item.Percent, 0, 100);
                StatusText.Text = item.FileName is null ? item.Stage : $"{item.Stage}: {item.FileName}";
            });
            ReportResult = await RunReportAsync(request, progress, _cancellation.Token);
            ProgressBar.Value = 100;
            StatusText.Text = "Таблица сформирована";
            if (OpenResultCheckBox.IsChecked == true)
                Process.Start(new ProcessStartInfo(ReportResult.OutputFile) { UseShellExecute = true });
            MessageBox.Show(this,
                $"Проверено PDF: {ReportResult.PdfFiles:N0}\n" +
                $"С фотографией: {ReportResult.WithPhoto:N0}\n" +
                $"Без фотографии: {ReportResult.WithoutPhoto:N0}\n" +
                $"Ошибки чтения: {ReportResult.ReadErrors:N0}",
                "Таблица готова", MessageBoxButton.OK, MessageBoxImage.Information);
            DialogResult = true;
        }
        catch (OperationCanceledException) { StatusText.Text = "Проверка отменена"; }
        catch (Exception exception)
        {
            StatusText.Text = "Не удалось проверить PDF";
            MessageBox.Show(this, exception.Message, "Ошибка проверки PDF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _cancellation?.Dispose(); _cancellation = null;
            if (IsVisible) SetBusy(false);
        }
    }

    private static Task<PdfPhotoReportResult> RunReportAsync(
        PdfPhotoReportRequest request,
        IProgress<PdfPhotoProgress> progress,
        CancellationToken cancellationToken)
    {
        var forceWorker = string.Equals(Environment.GetEnvironmentVariable("ANALIZATOR_FORCE_WORKER"), "1", StringComparison.Ordinal);
        return Debugger.IsAttached || forceWorker
            ? PdfPhotoWorker.RunIsolatedAsync(request, progress, cancellationToken)
            : Task.Run(() => new PdfPhotoReportService().Run(request, progress, cancellationToken), cancellationToken);
    }

    private void SetBusy(bool busy)
    {
        _isBusy = busy;
        SourceFolderTextBox.IsEnabled = !busy;
        BrowseFolderButton.IsEnabled = !busy;
        UsePeriodCheckBox.IsEnabled = !busy;
        StartDatePicker.IsEnabled = !busy && UsePeriodCheckBox.IsChecked == true;
        EndDatePicker.IsEnabled = !busy && UsePeriodCheckBox.IsChecked == true;
        OpenResultCheckBox.IsEnabled = !busy;
        CancelButton.IsEnabled = busy;
        UpdateState();
    }

    private void UpdateState()
    {
        GenerateButton.IsEnabled = !_isBusy && Directory.Exists(SourceFolderTextBox.Text.Trim());
        if (!_isBusy) StatusText.Text = GenerateButton.IsEnabled ? "Папка готова к проверке" : "Выберите папку с PDF";
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => _cancellation?.Cancel();
    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_isBusy) { _cancellation?.Cancel(); e.Cancel = true; }
    }
}

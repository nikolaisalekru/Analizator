using Analizator.Core;
using System.Globalization;
using System.Windows;

namespace Analizator;

public partial class DatabaseBackupSettingsWindow : Window
{
    public DatabaseBackupSettingsWindow(DatabaseBackupSettings settings)
    {
        InitializeComponent();
        AutomaticBackupsCheckBox.IsChecked = settings.AutomaticBackupsEnabled;
        IntervalDaysTextBox.Text = settings.IntervalDays.ToString(CultureInfo.CurrentCulture);
        MaximumFilesTextBox.Text = settings.MaximumBackupFiles.ToString(CultureInfo.CurrentCulture);
        BackupBeforeChangesCheckBox.IsChecked = settings.BackupBeforeChanges;
        ConfirmRowDeletionCheckBox.IsChecked = settings.ConfirmRowDeletion;
        ConfirmSourceDeletionCheckBox.IsChecked = settings.ConfirmSourceDeletion;
        UpdateAutomaticOptions();
    }

    public DatabaseBackupSettings? ResultSettings { get; private set; }

    private void AutomaticBackups_Changed(object sender, RoutedEventArgs e) =>
        UpdateAutomaticOptions();

    private void UpdateAutomaticOptions()
    {
        if (AutomaticOptionsGrid is not null)
            AutomaticOptionsGrid.IsEnabled = AutomaticBackupsCheckBox.IsChecked == true;
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!int.TryParse(IntervalDaysTextBox.Text.Trim(), NumberStyles.Integer,
                    CultureInfo.CurrentCulture, out var intervalDays) ||
                intervalDays is < 1 or > 3650)
                throw new ArgumentException("Интервал должен быть целым числом от 1 до 3650 дней.");
            if (!int.TryParse(MaximumFilesTextBox.Text.Trim(), NumberStyles.Integer,
                    CultureInfo.CurrentCulture, out var maximumFiles) ||
                maximumFiles is < 1 or > 100)
                throw new ArgumentException("Количество копий должно быть целым числом от 1 до 100.");

            ResultSettings = new DatabaseBackupSettings
            {
                AutomaticBackupsEnabled = AutomaticBackupsCheckBox.IsChecked == true,
                IntervalDays = intervalDays,
                MaximumBackupFiles = maximumFiles,
                BackupBeforeChanges = BackupBeforeChangesCheckBox.IsChecked == true,
                ConfirmRowDeletion = ConfirmRowDeletionCheckBox.IsChecked == true,
                ConfirmSourceDeletion = ConfirmSourceDeletionCheckBox.IsChecked == true
            };
            DialogResult = true;
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Проверьте настройки",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}

using Analizator.Core;
using System.Globalization;
using System.Windows;

namespace Analizator;

public partial class DatabaseAttemptWindow : Window
{
    private readonly DatabaseAttemptEditModel _source;

    public DatabaseAttemptWindow(DatabaseAttemptEditModel source, IEnumerable<string> installations)
    {
        InitializeComponent();
        _source = source;
        HeadingText.Text = source.Id.HasValue ? "Изменение попытки" : "Новая попытка";
        InstallationComboBox.ItemsSource = installations
            .Append(source.Installation)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
        AttemptDatePicker.SelectedDate = source.AttemptedAt.Date;
        AttemptTimeTextBox.Text = source.AttemptedAt.ToString("HH:mm", CultureInfo.CurrentCulture);
        InstallationComboBox.Text = source.Installation;
        PersonNameTextBox.Text = source.PersonName;
        ScenarioTextBox.Text = source.ScenarioName;
        PercentTextBox.Text = source.Percent.ToString("0.##", CultureInfo.CurrentCulture);
        ActionsTextBox.Text = source.Actions?.ToString(CultureInfo.CurrentCulture) ?? "";
        ModeTextBox.Text = source.Mode ?? "";
        ProjectTextBox.Text = source.Project ?? "";
        RoleTextBox.Text = source.Role ?? "";
        CreditedTextBox.Text = source.Credited ?? "";
        DurationTextBox.Text = source.Duration ?? "";
    }

    public DatabaseAttemptEditModel? Result { get; private set; }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!AttemptDatePicker.SelectedDate.HasValue)
                throw new ArgumentException("Укажите дату попытки.");
            if (!TimeSpan.TryParse(AttemptTimeTextBox.Text.Trim(), CultureInfo.CurrentCulture, out var time) &&
                !TimeSpan.TryParse(AttemptTimeTextBox.Text.Trim(), CultureInfo.InvariantCulture, out time))
                throw new ArgumentException("Укажите время в формате ЧЧ:ММ.");
            if (time < TimeSpan.Zero || time >= TimeSpan.FromDays(1))
                throw new ArgumentException("Время должно быть в пределах от 00:00 до 23:59.");
            if (!TryParseDouble(PercentTextBox.Text, out var percent))
                throw new ArgumentException("Укажите процент числом от 0 до 100.");
            int? actions = null;
            if (!string.IsNullOrWhiteSpace(ActionsTextBox.Text))
            {
                if (!int.TryParse(ActionsTextBox.Text.Trim(), out var parsedActions))
                    throw new ArgumentException("Количество действий должно быть целым числом.");
                actions = parsedActions;
            }

            Result = new DatabaseAttemptEditModel
            {
                Id = _source.Id,
                AttemptedAt = AttemptDatePicker.SelectedDate.Value.Date.Add(time),
                Installation = InstallationComboBox.Text.Trim(),
                PersonName = PersonNameTextBox.Text.Trim(),
                ScenarioName = ScenarioTextBox.Text.Trim(),
                Percent = percent,
                Actions = actions,
                Mode = ModeTextBox.Text,
                Project = ProjectTextBox.Text,
                Role = RoleTextBox.Text,
                Credited = CreditedTextBox.Text,
                Duration = DurationTextBox.Text,
                SessionGuid = _source.SessionGuid,
                MachineName = _source.MachineName,
                PlantName = _source.PlantName,
                StudentSerial = _source.StudentSerial,
                EntryDate = _source.EntryDate,
                SessionStartDate = _source.SessionStartDate,
                SessionEndDate = _source.SessionEndDate,
                IsClosed = _source.IsClosed,
                FilePath = _source.FilePath,
                Resume = _source.Resume,
                Result = _source.Result,
                IsValid = _source.IsValid,
                ReportType = _source.ReportType,
                HasReportFile = _source.HasReportFile,
                ReportFile = _source.ReportFile,
                TrainingTimesJson = _source.TrainingTimesJson,
                ExerciseLogsJson = _source.ExerciseLogsJson,
                TrainingTimesCount = _source.TrainingTimesCount,
                ExerciseLogsCount = _source.ExerciseLogsCount
            };
            ValidateRequired(Result);
            DialogResult = true;
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Проверьте запись",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private static void ValidateRequired(DatabaseAttemptEditModel model)
    {
        if (string.IsNullOrWhiteSpace(model.Installation))
            throw new ArgumentException("Укажите установку.");
        if (string.IsNullOrWhiteSpace(model.PersonName))
            throw new ArgumentException("Укажите ФИО.");
        if (string.IsNullOrWhiteSpace(model.ScenarioName))
            throw new ArgumentException("Укажите сценарий.");
        if (!double.IsFinite(model.Percent) || model.Percent is < 0 or > 100)
            throw new ArgumentException("Процент должен быть от 0 до 100.");
        if (model.Actions < 0)
            throw new ArgumentException("Количество действий не может быть отрицательным.");
    }

    private static bool TryParseDouble(string value, out double result) =>
        double.TryParse(value.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out result) ||
        double.TryParse(value.Trim().Replace(',', '.'), NumberStyles.Float,
            CultureInfo.InvariantCulture, out result);
}

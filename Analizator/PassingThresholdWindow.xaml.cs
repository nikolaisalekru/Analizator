using Analizator.Core;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;

namespace Analizator;

public partial class PassingThresholdWindow : Window
{
    private readonly ObservableCollection<ThresholdRow> _rows = [];

    public PassingThresholdWindow(
        PassingThresholdSettings settings,
        IEnumerable<string> installations)
    {
        InitializeComponent();
        GlobalThresholdTextBox.Text = FormatPercent(settings.GlobalPercent);
        UsePerInstallationCheckBox.IsChecked = settings.UsePerInstallation;

        foreach (var installation in installations
                     .Concat(settings.Installations.Keys)
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase))
        {
            _rows.Add(new ThresholdRow
            {
                Installation = installation,
                ThresholdText = FormatPercent(settings.GetForInstallation(installation))
            });
        }

        ThresholdDataGrid.ItemsSource = _rows;
        UpdateModeControls();
    }

    public PassingThresholdSettings? ResultSettings { get; private set; }

    private void ThresholdMode_Changed(object sender, RoutedEventArgs e) => UpdateModeControls();

    private void UpdateModeControls()
    {
        if (ThresholdDataGrid is null || GlobalThresholdTextBox is null)
            return;
        var perInstallation = UsePerInstallationCheckBox.IsChecked == true;
        ThresholdDataGrid.IsEnabled = perInstallation;
        ThresholdDataGrid.Opacity = perInstallation ? 1 : 0.55;
        GlobalThresholdTextBox.IsEnabled = !perInstallation;
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        ThresholdDataGrid.CommitEdit();
        ThresholdDataGrid.CommitEdit();

        if (!TryParsePercent(GlobalThresholdTextBox.Text, out var global))
        {
            ShowValidation("Укажите единый порог числом от 0 до 100.", GlobalThresholdTextBox);
            return;
        }

        var values = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in _rows)
        {
            if (!TryParsePercent(row.ThresholdText, out var value))
            {
                ThresholdDataGrid.ScrollIntoView(row);
                ThresholdDataGrid.SelectedItem = row;
                ShowValidation($"Проверьте порог для установки «{row.Installation}».", ThresholdDataGrid);
                return;
            }
            values[row.Installation] = value;
        }

        ResultSettings = new PassingThresholdSettings
        {
            UsePerInstallation = UsePerInstallationCheckBox.IsChecked == true,
            GlobalPercent = global,
            Installations = values
        };
        DialogResult = true;
    }

    private void ShowValidation(string message, UIElement target)
    {
        ValidationStatusText.Text = message;
        target.Focus();
    }

    private static bool TryParsePercent(string? text, out double value)
    {
        var normalized = (text ?? "").Trim().Replace(',', '.');
        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
               !double.IsNaN(value) && !double.IsInfinity(value) && value is >= 0 and <= 100;
    }

    private static string FormatPercent(double value) =>
        value.ToString("0.##", CultureInfo.CurrentCulture);

    private sealed class ThresholdRow
    {
        public required string Installation { get; init; }
        public required string ThresholdText { get; set; }
    }
}

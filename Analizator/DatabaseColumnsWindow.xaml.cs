using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace Analizator;

public partial class DatabaseColumnsWindow : Window
{
    private readonly DatabaseColumnOption[] _columns;

    public DatabaseColumnsWindow(IEnumerable<DatabaseColumnOption> columns)
    {
        InitializeComponent();
        _columns = columns.ToArray();
        ColumnsItemsControl.ItemsSource = _columns;
    }

    public HashSet<string>? HiddenColumns { get; private set; }

    private void ShowAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var column in _columns)
            column.IsVisible = true;
    }

    private void HideAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var column in _columns)
            column.IsVisible = false;
    }

    private void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        if (_columns.All(column => !column.IsVisible))
        {
            MessageBox.Show(this, "Оставьте видимым хотя бы один столбец.", "Столбцы",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        HiddenColumns = _columns
            .Where(column => !column.IsVisible)
            .Select(column => column.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        DialogResult = true;
    }
}

public sealed class DatabaseColumnOption : INotifyPropertyChanged
{
    private bool _isVisible;

    public required string Key { get; init; }
    public required string DisplayName { get; init; }

    public bool IsVisible
    {
        get => _isVisible;
        set
        {
            if (_isVisible == value)
                return;
            _isVisible = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsVisible)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

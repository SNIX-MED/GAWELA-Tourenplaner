using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.Infrastructure.Backups;

namespace Tourenplaner.CSharp.App.Views;

public sealed class RestoreSelectionWindow : Window
{
    private readonly Dictionary<string, CheckBox> _groupBoxes = new(StringComparer.OrdinalIgnoreCase);

    public RestoreSelectionWindow()
    {
        Title = "Wiederherstellung auswählen";
        Width = 520;
        Height = 520;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(15, 23, 42));
        Build();
    }

    public IReadOnlyList<string>? SelectedGroups { get; private set; }

    private void Build()
    {
        var shell = new Grid { Margin = new Thickness(20) };
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        shell.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        shell.Children.Add(ViewStyling.Text("Welche Daten sollen wiederhergestellt werden?", 18, FontWeights.SemiBold, Brushes.White));
        shell.Children.Add(ViewStyling.Text("Du kannst alle Daten oder einzelne Datengruppen aus dem ausgewählten Backup zurückspielen.", 12, FontWeights.Normal, Brushes.Gainsboro, new Thickness(0, 32, 0, 0)));

        var list = new StackPanel { Margin = new Thickness(0, 74, 0, 0) };
        var selectAll = new CheckBox
        {
            Content = "Alle Daten wiederherstellen",
            Foreground = Brushes.White,
            IsChecked = true,
            Margin = new Thickness(0, 0, 0, 12),
        };
        selectAll.Checked += (_, _) => SetAllBoxes(true);
        selectAll.Unchecked += (_, _) => SetAllBoxes(false);
        list.Children.Add(selectAll);

        foreach (var pair in BackupManager.RestoreLabels)
        {
            var box = new CheckBox
            {
                Content = pair.Value,
                Foreground = Brushes.White,
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 8),
            };
            box.Checked += (_, _) => SyncSelectAll(selectAll);
            box.Unchecked += (_, _) => SyncSelectAll(selectAll);
            _groupBoxes[pair.Key] = box;
            list.Children.Add(box);
        }

        Grid.SetRow(list, 2);
        shell.Children.Add(list);

        var buttons = new UniformGrid { Columns = 2, Margin = new Thickness(0, 20, 0, 0) };
        var cancel = new Button { Content = "Abbrechen", Margin = new Thickness(0, 0, 8, 0), Padding = new Thickness(12, 8, 12, 8) };
        cancel.Click += (_, _) => DialogResult = false;
        buttons.Children.Add(cancel);

        var restore = new Button { Content = "Wiederherstellen", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(12, 8, 12, 8) };
        restore.Click += (_, _) => ConfirmSelection();
        buttons.Children.Add(restore);
        Grid.SetRow(buttons, 3);
        shell.Children.Add(buttons);

        Content = shell;
    }

    private void SetAllBoxes(bool value)
    {
        foreach (var box in _groupBoxes.Values)
        {
            box.IsChecked = value;
        }
    }

    private void SyncSelectAll(CheckBox selectAll)
    {
        selectAll.IsChecked = _groupBoxes.Values.All(box => box.IsChecked == true);
    }

    private void ConfirmSelection()
    {
        var selected = _groupBoxes.Where(pair => pair.Value.IsChecked == true).Select(pair => pair.Key).ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show(this, "Bitte mindestens einen Datenbereich auswählen.", "Wiederherstellung", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SelectedGroups = selected;
        DialogResult = true;
    }
}

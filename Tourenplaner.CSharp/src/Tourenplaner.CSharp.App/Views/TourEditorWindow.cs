using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class TourEditorWindow : Window
{
    private sealed record SelectionOption(string Id, string Label);

    private readonly TextBox _dateTextBox = new();
    private readonly TextBox _nameTextBox = new();
    private readonly TextBox _startTimeTextBox = new();
    private readonly ComboBox _vehicleComboBox = new();
    private readonly ComboBox _trailerComboBox = new();
    private readonly ListBox _employeesListBox = new() { SelectionMode = SelectionMode.Multiple, Height = 120 };
    private readonly Tour? _existingTour;
    public TourEditorWindow(Tour? tour, IReadOnlyList<Employee> employees, VehicleCatalog catalog)
    {
        _existingTour = tour;
        Title = tour is null ? "Tour hinzufügen" : "Tour bearbeiten";
        Width = 520;
        Height = 560;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush(Color.FromRgb(15, 23, 42));

        Build(tour, employees, catalog);
    }

    public Tour? EditedTour { get; private set; }

    private void Build(Tour? tour, IReadOnlyList<Employee> employees, VehicleCatalog catalog)
    {
        var shell = new Grid { Margin = new Thickness(20) };
        for (var i = 0; i < 12; i++)
        {
            shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }
        shell.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        AddLabeledTextBox(shell, "Datum (DD-MM-YYYY)", _dateTextBox, tour?.Date ?? DateTime.Today.ToString("dd-MM-yyyy"), 0, 1);
        AddLabeledTextBox(shell, "Name", _nameTextBox, tour?.Name ?? string.Empty, 2, 3);
        AddLabeledTextBox(shell, "Startzeit", _startTimeTextBox, tour?.StartTime ?? "08:00", 4, 5);

        AddLabel(shell, "Fahrzeug", 6);
        _vehicleComboBox.DisplayMemberPath = nameof(SelectionOption.Label);
        _vehicleComboBox.SelectedValuePath = nameof(SelectionOption.Id);
        _vehicleComboBox.ItemsSource = catalog.Vehicles.Select(vehicle => new SelectionOption(vehicle.Id, vehicle.Name)).ToList();
        _vehicleComboBox.SelectedValue = tour?.VehicleId;
        Grid.SetRow(_vehicleComboBox, 7);
        shell.Children.Add(_vehicleComboBox);

        AddLabel(shell, "Anhänger", 8);
        _trailerComboBox.DisplayMemberPath = nameof(SelectionOption.Label);
        _trailerComboBox.SelectedValuePath = nameof(SelectionOption.Id);
        _trailerComboBox.ItemsSource = new[] { new SelectionOption(string.Empty, "(kein Anhänger)") }
            .Concat(catalog.Trailers.Select(trailer => new SelectionOption(trailer.Id, trailer.Name)))
            .ToList();
        _trailerComboBox.SelectedValue = tour?.TrailerId ?? string.Empty;
        Grid.SetRow(_trailerComboBox, 9);
        shell.Children.Add(_trailerComboBox);

        AddLabel(shell, "Mitarbeiter (1-2 auswählen)", 10);
        foreach (var employee in employees)
        {
            var item = new ListBoxItem { Content = employee.Name, Tag = employee.Id };
            if ((tour?.EmployeeIds ?? Array.Empty<string>()).Contains(employee.Id))
            {
                item.IsSelected = true;
            }
            _employeesListBox.Items.Add(item);
        }
        Grid.SetRow(_employeesListBox, 11);
        shell.Children.Add(_employeesListBox);

        var buttons = new UniformGrid { Columns = 2, Margin = new Thickness(0, 20, 0, 0) };
        var cancelButton = new Button { Content = "Abbrechen", Margin = new Thickness(0, 0, 8, 0), Padding = new Thickness(12, 8, 12, 8) };
        cancelButton.Click += (_, _) => DialogResult = false;
        buttons.Children.Add(cancelButton);

        var saveButton = new Button { Content = "Speichern", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(12, 8, 12, 8) };
        saveButton.Click += (_, _) => Save();
        buttons.Children.Add(saveButton);
        Grid.SetRow(buttons, 13);
        shell.Children.Add(buttons);

        Content = new ScrollViewer { Content = shell, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
    }

    private void AddLabeledTextBox(Grid shell, string text, TextBox textBox, string value, int labelRow, int textBoxRow)
    {
        AddLabel(shell, text, labelRow);
        textBox.Text = value;
        textBox.Padding = new Thickness(8);
        Grid.SetRow(textBox, textBoxRow);
        shell.Children.Add(textBox);
    }

    private void AddLabel(Grid shell, string text, int row)
    {
        var label = new Label { Content = text, Foreground = Brushes.White, Margin = new Thickness(0, row == 0 ? 0 : 10, 0, 4) };
        Grid.SetRow(label, row);
        shell.Children.Add(label);
    }

    private void Save()
    {
        var selectedVehicleId = _vehicleComboBox.SelectedValue?.ToString() ?? string.Empty;
        var selectedTrailerId = _trailerComboBox.SelectedValue?.ToString() ?? string.Empty;
        var employeeIds = _employeesListBox.SelectedItems.OfType<ListBoxItem>().Select(item => item.Tag?.ToString() ?? string.Empty).Where(id => !string.IsNullOrWhiteSpace(id)).ToList();

        if (string.IsNullOrWhiteSpace(_dateTextBox.Text) || string.IsNullOrWhiteSpace(_nameTextBox.Text) || string.IsNullOrWhiteSpace(selectedVehicleId))
        {
            MessageBox.Show(this, "Bitte Datum, Name und Fahrzeug auswählen.", "Touren", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (employeeIds.Count is < 1 or > 2)
        {
            MessageBox.Show(this, "Bitte 1 oder 2 Mitarbeiter auswählen.", "Touren", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        EditedTour = (_existingTour ?? new Tour { Id = 0, Stops = Array.Empty<TourStop>(), TravelTimeCache = new Dictionary<string, int>(), RouteMode = "car" }) with
        {
            Date = _dateTextBox.Text.Trim(),
            Name = _nameTextBox.Text.Trim(),
            StartTime = _startTimeTextBox.Text.Trim(),
            VehicleId = selectedVehicleId,
            TrailerId = string.IsNullOrWhiteSpace(selectedTrailerId) ? null : selectedTrailerId,
            EmployeeIds = employeeIds,
        };

        DialogResult = true;
    }
}

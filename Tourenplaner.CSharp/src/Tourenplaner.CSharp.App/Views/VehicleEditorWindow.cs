using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class VehicleEditorWindow : Window
{
    private readonly bool _isTrailer;
    private readonly TextBox _nameTextBox = new();
    private readonly TextBox _licensePlateTextBox = new();
    private readonly TextBox _payloadTextBox = new();
    private readonly TextBox _trailerLoadTextBox = new();
    private readonly TextBox _volumeTextBox = new();
    private readonly CheckBox _activeCheckBox = new() { Content = "Eintrag ist aktiv" };

    public VehicleEditorWindow(Vehicle? vehicle = null)
    {
        _isTrailer = false;
        Title = vehicle is null ? "Fahrzeug hinzufügen" : "Fahrzeug bearbeiten";
        Width = 460;
        Height = 420;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(15, 23, 42));
        EditedVehicle = vehicle;
        Build(vehicle, null);
    }

    public VehicleEditorWindow(Trailer? trailer = null)
    {
        _isTrailer = true;
        Title = trailer is null ? "Anhänger hinzufügen" : "Anhänger bearbeiten";
        Width = 460;
        Height = 380;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(15, 23, 42));
        EditedTrailer = trailer;
        Build(null, trailer);
    }

    public Vehicle? EditedVehicle { get; private set; }
    public Trailer? EditedTrailer { get; private set; }

    private void Build(Vehicle? vehicle, Trailer? trailer)
    {
        var shell = new Grid { Margin = new Thickness(20) };
        for (var i = 0; i < 10; i++)
        {
            shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }

        AddLabeledTextBox(shell, "Name", _nameTextBox, vehicle?.Name ?? trailer?.Name ?? string.Empty, 0, 1);
        AddLabeledTextBox(shell, "Kennzeichen", _licensePlateTextBox, vehicle?.LicensePlate ?? trailer?.LicensePlate ?? string.Empty, 2, 3);
        AddLabeledTextBox(shell, "Nutzlast (kg)", _payloadTextBox, (vehicle?.MaxPayloadKg ?? trailer?.MaxPayloadKg ?? 0).ToString(CultureInfo.InvariantCulture), 4, 5);

        if (!_isTrailer)
        {
            AddLabeledTextBox(shell, "Anhängelast (kg)", _trailerLoadTextBox, (vehicle?.MaxTrailerLoadKg ?? 0).ToString(CultureInfo.InvariantCulture), 6, 7);
        }

        AddLabeledTextBox(shell, "Volumen (m³)", _volumeTextBox, (vehicle?.VolumeM3 ?? trailer?.VolumeM3 ?? 0).ToString(CultureInfo.InvariantCulture), 8, 9);

        _activeCheckBox.Margin = new Thickness(0, 12, 0, 0);
        _activeCheckBox.Foreground = Brushes.White;
        _activeCheckBox.IsChecked = vehicle?.Active ?? trailer?.Active ?? true;
        Grid.SetRow(_activeCheckBox, _isTrailer ? 10 : 10);
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        shell.Children.Add(_activeCheckBox);

        var buttons = new UniformGrid { Columns = 2, Margin = new Thickness(0, 20, 0, 0) };
        var cancelButton = new Button { Content = "Abbrechen", Margin = new Thickness(0, 0, 8, 0), Padding = new Thickness(12, 8, 12, 8) };
        cancelButton.Click += (_, _) => DialogResult = false;
        buttons.Children.Add(cancelButton);

        var saveButton = new Button { Content = "Speichern", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(12, 8, 12, 8) };
        saveButton.Click += (_, _) => Save();
        buttons.Children.Add(saveButton);

        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        Grid.SetRow(buttons, 11);
        shell.Children.Add(buttons);
        Content = shell;
    }

    private void AddLabeledTextBox(Grid shell, string labelText, TextBox textBox, string value, int labelRow, int textBoxRow)
    {
        var label = new Label { Content = labelText, Foreground = Brushes.White, Margin = new Thickness(0, labelRow == 0 ? 0 : 10, 0, 4) };
        Grid.SetRow(label, labelRow);
        shell.Children.Add(label);

        textBox.Text = value;
        textBox.Padding = new Thickness(8);
        Grid.SetRow(textBox, textBoxRow);
        shell.Children.Add(textBox);
    }

    private void Save()
    {
        var name = _nameTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(this, "Bitte einen Namen eingeben.", "Fahrzeuge", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!int.TryParse(_payloadTextBox.Text.Trim(), out var payload))
        {
            payload = 0;
        }
        if (!int.TryParse(_trailerLoadTextBox.Text.Trim(), out var trailerLoad))
        {
            trailerLoad = 0;
        }
        if (!int.TryParse(_volumeTextBox.Text.Trim(), out var volume))
        {
            volume = 0;
        }

        if (_isTrailer)
        {
            EditedTrailer = (EditedTrailer ?? new Trailer(Guid.NewGuid().ToString(), string.Empty, string.Empty, 0, true, string.Empty, 0, null, DateTime.Now, null)) with
            {
                Name = name,
                LicensePlate = _licensePlateTextBox.Text.Trim(),
                MaxPayloadKg = Math.Max(0, payload),
                VolumeM3 = Math.Max(0, volume),
                Active = _activeCheckBox.IsChecked ?? true,
                UpdatedAt = DateTime.Now,
            };
        }
        else
        {
            EditedVehicle = (EditedVehicle ?? new Vehicle(Guid.NewGuid().ToString(), "other", string.Empty, string.Empty, 0, 0, true, string.Empty, 0, null, DateTime.Now, null)) with
            {
                Name = name,
                LicensePlate = _licensePlateTextBox.Text.Trim(),
                MaxPayloadKg = Math.Max(0, payload),
                MaxTrailerLoadKg = Math.Max(0, trailerLoad),
                VolumeM3 = Math.Max(0, volume),
                Active = _activeCheckBox.IsChecked ?? true,
                UpdatedAt = DateTime.Now,
            };
        }

        DialogResult = true;
    }
}

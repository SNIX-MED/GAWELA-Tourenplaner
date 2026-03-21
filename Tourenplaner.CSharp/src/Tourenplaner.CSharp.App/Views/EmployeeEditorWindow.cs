using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class EmployeeEditorWindow : Window
{
    private readonly TextBox _nameTextBox = new();
    private readonly TextBox _shortTextBox = new();
    private readonly TextBox _phoneTextBox = new();
    private readonly CheckBox _activeCheckBox = new() { Content = "Mitarbeiter ist aktiv" };

    public EmployeeEditorWindow(Employee? employee = null)
    {
        Title = employee is null ? "Mitarbeiter hinzufügen" : "Mitarbeiter bearbeiten";
        Width = 420;
        Height = 320;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(15, 23, 42));

        EditedEmployee = employee;
        Build(employee);
    }

    public Employee? EditedEmployee { get; private set; }

    private void Build(Employee? employee)
    {
        var shell = new Grid { Margin = new Thickness(20) };
        for (var i = 0; i < 6; i++)
        {
            shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }
        shell.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        shell.Children.Add(CreateLabel("Name", 0));
        ConfigureTextBox(_nameTextBox, employee?.Name ?? string.Empty, 1);
        shell.Children.Add(_nameTextBox);

        shell.Children.Add(CreateLabel("Kürzel", 2));
        ConfigureTextBox(_shortTextBox, employee?.Short ?? string.Empty, 3);
        shell.Children.Add(_shortTextBox);

        shell.Children.Add(CreateLabel("Telefon", 4));
        ConfigureTextBox(_phoneTextBox, employee?.Phone ?? string.Empty, 5);
        shell.Children.Add(_phoneTextBox);

        _activeCheckBox.Margin = new Thickness(0, 12, 0, 0);
        _activeCheckBox.Foreground = Brushes.White;
        _activeCheckBox.IsChecked = employee?.Active ?? true;
        Grid.SetRow(_activeCheckBox, 6);
        shell.Children.Add(_activeCheckBox);

        var buttons = new UniformGrid { Columns = 2, Margin = new Thickness(0, 20, 0, 0) };
        buttons.Children.Add(new Button
        {
            Content = "Abbrechen",
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(12, 8, 12, 8),
        });
        ((Button)buttons.Children[0]).Click += (_, _) => DialogResult = false;

        var saveButton = new Button
        {
            Content = "Speichern",
            Margin = new Thickness(8, 0, 0, 0),
            Padding = new Thickness(12, 8, 12, 8),
        };
        saveButton.Click += (_, _) => Save();
        buttons.Children.Add(saveButton);
        Grid.SetRow(buttons, 7);
        shell.Children.Add(buttons);

        Content = shell;
    }

    private Label CreateLabel(string text, int row)
    {
        var label = new Label
        {
            Content = text,
            Foreground = Brushes.White,
            Margin = new Thickness(0, row == 0 ? 0 : 12, 0, 4),
        };
        Grid.SetRow(label, row);
        return label;
    }

    private static void ConfigureTextBox(TextBox textBox, string text, int row)
    {
        textBox.Text = text;
        textBox.Padding = new Thickness(8);
        Grid.SetRow(textBox, row);
    }

    private void Save()
    {
        var name = _nameTextBox.Text.Trim();
        var shortName = _shortTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(shortName))
        {
            MessageBox.Show(this, "Bitte Name und Kürzel ausfüllen.", "Mitarbeiter", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        EditedEmployee = (EditedEmployee ?? new Employee(Guid.NewGuid().ToString(), string.Empty, string.Empty, string.Empty, true, DateTime.Now)) with
        {
            Name = name,
            Short = shortName,
            Phone = _phoneTextBox.Text.Trim(),
            Active = _activeCheckBox.IsChecked ?? true,
        };

        DialogResult = true;
    }
}

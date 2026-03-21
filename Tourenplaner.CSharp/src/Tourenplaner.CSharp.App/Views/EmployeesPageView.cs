using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class EmployeesPageView : ScrollViewer
{
    public EmployeesPageView(
        IReadOnlyList<Employee> employees,
        Brush panelBrush,
        Brush textBrush,
        Brush subTextBrush,
        Action? onAdd,
        Action<Employee>? onEdit,
        Action<Employee>? onDelete)
    {
        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build(employees, panelBrush, textBrush, subTextBrush, onAdd, onEdit, onDelete);
    }

    private static UIElement Build(
        IReadOnlyList<Employee> employees,
        Brush panelBrush,
        Brush textBrush,
        Brush subTextBrush,
        Action? onAdd,
        Action<Employee>? onEdit,
        Action<Employee>? onDelete)
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Mitarbeiterverwaltung", 22, FontWeights.SemiBold, textBrush));
        shell.Children.Add(ViewStyling.Text("Erste echte Bereichsansicht mit geladenen Mitarbeiterdaten und grundlegenden CRUD-Aktionen.", 13, FontWeights.Normal, subTextBrush, new Thickness(0, 10, 0, 0)));
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Mitarbeiter gesamt", employees.Count.ToString(), "Alle geladenen Mitarbeiter"),
            new PageStatViewModel("Aktiv", employees.Count(employee => employee.Active).ToString(), "Aktiv markierte Mitarbeiter"),
        }, panelBrush, textBrush, subTextBrush));

        var actions = new UniformGrid { Columns = 3, Margin = new Thickness(0, 16, 0, 0) };
        var addButton = new Button { Content = "Hinzufügen", Margin = new Thickness(0, 0, 8, 0), Padding = new Thickness(12, 8, 12, 8) };
        var editButton = new Button { Content = "Bearbeiten", Margin = new Thickness(4, 0, 4, 0), Padding = new Thickness(12, 8, 12, 8) };
        var deleteButton = new Button { Content = "Löschen", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(12, 8, 12, 8) };
        actions.Children.Add(addButton);
        actions.Children.Add(editButton);
        actions.Children.Add(deleteButton);
        shell.Children.Add(actions);

        var grid = ViewStyling.CreateReadOnlyGrid();
        grid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(Employee.Name)) });
        grid.Columns.Add(new DataGridTextColumn { Header = "Kürzel", Binding = new System.Windows.Data.Binding(nameof(Employee.Short)) });
        grid.Columns.Add(new DataGridTextColumn { Header = "Telefon", Binding = new System.Windows.Data.Binding(nameof(Employee.Phone)) });
        grid.Columns.Add(new DataGridCheckBoxColumn { Header = "Aktiv", Binding = new System.Windows.Data.Binding(nameof(Employee.Active)) });
        grid.Columns.Add(new DataGridTextColumn { Header = "Erstellt", Binding = new System.Windows.Data.Binding(nameof(Employee.CreatedAt)) { StringFormat = "dd.MM.yyyy HH:mm" } });
        grid.ItemsSource = employees;
        shell.Children.Add(grid);

        addButton.Click += (_, _) => onAdd?.Invoke();
        editButton.Click += (_, _) =>
        {
            if (grid.SelectedItem is Employee employee)
            {
                onEdit?.Invoke(employee);
            }
        };
        deleteButton.Click += (_, _) =>
        {
            if (grid.SelectedItem is Employee employee)
            {
                onDelete?.Invoke(employee);
            }
        };

        return shell;
    }
}

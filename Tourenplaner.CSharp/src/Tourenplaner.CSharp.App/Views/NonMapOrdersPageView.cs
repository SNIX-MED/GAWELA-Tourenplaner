using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class NonMapOrdersPageView : ScrollViewer
{
    public NonMapOrdersPageView(SqlImportWorkspace sqlImportWorkspace, Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build(sqlImportWorkspace, panelBrush, textBrush, subTextBrush);
    }

    private static UIElement Build(SqlImportWorkspace sqlImportWorkspace, Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Nicht-Karten-Aufträge", 22, FontWeights.SemiBold, textBrush));
        shell.Children.Add(ViewStyling.Text("Die Seite arbeitet direkt mit data/non_map_sql_orders.json und gruppiert die vom Python-SQL-Import separierten Lieferarten in der gleichen Informationsarchitektur.", 13, FontWeights.Normal, subTextBrush, new Thickness(0, 10, 0, 0)));

        var grouped = sqlImportWorkspace.NonMapOrders
            .Select(order => new
            {
                DeliveryType = FirstNonEmpty(order.NonMapCategory, order.DeliveryType, "Nicht klassifiziert"),
                OrderNumber = FirstNonEmpty(order.OrderNumber, "-"),
                Name = FirstNonEmpty(order.Name, "-"),
                Address = FirstNonEmpty(order.Address, "-"),
                Status = string.IsNullOrWhiteSpace(order.Status) ? "nicht festgelegt" : order.Status,
            })
            .GroupBy(item => item.DeliveryType)
            .OrderBy(group => group.Key)
            .ToList();

        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Gruppen", grouped.Count.ToString(), "Ermittelte Lieferart-Kategorien"),
            new PageStatViewModel("Ohne Lieferart", grouped.FirstOrDefault(group => group.Key == "Nicht klassifiziert")?.Count().ToString() ?? "0", "Datensätze ohne Lieferart"),
            new PageStatViewModel("Nicht-Karte gesamt", sqlImportWorkspace.NonMapOrders.Count.ToString(), "SQL-Aufträge außerhalb der Kartenansicht"),
            new PageStatViewModel("Pending parallel", sqlImportWorkspace.PendingOrders.Count.ToString(), "Noch ungeocodete Aufträge in separater Datei"),
        }, panelBrush, textBrush, subTextBrush));

        if (grouped.Count == 0)
        {
            shell.Children.Add(new Border
            {
                Background = panelBrush,
                CornerRadius = new CornerRadius(14),
                Padding = new Thickness(16),
                Margin = new Thickness(0, 16, 0, 0),
                Child = ViewStyling.Text("Aktuell liegen keine Nicht-Karten-Aufträge in data/non_map_sql_orders.json vor.", 12, FontWeights.Normal, subTextBrush),
            });
        }

        foreach (var group in grouped)
        {
            var card = new Border
            {
                Background = panelBrush,
                CornerRadius = new CornerRadius(14),
                Padding = new Thickness(16),
                Margin = new Thickness(0, 16, 0, 0),
            };

            var stack = new StackPanel();
            stack.Children.Add(ViewStyling.Text(group.Key, 16, FontWeights.SemiBold, textBrush));
            stack.Children.Add(ViewStyling.Text($"{group.Count()} Datensätze in dieser Kategorie.", 12, FontWeights.Normal, subTextBrush, new Thickness(0, 6, 0, 0)));

            var grid = ViewStyling.CreateReadOnlyGrid();
            grid.MinHeight = 120;
            grid.Columns.Add(new DataGridTextColumn { Header = "Auftrag", Binding = new System.Windows.Data.Binding("OrderNumber") });
            grid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding("Name") });
            grid.Columns.Add(new DataGridTextColumn { Header = "Adresse", Binding = new System.Windows.Data.Binding("Address"), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header = "Status", Binding = new System.Windows.Data.Binding("Status") });
            grid.ItemsSource = group.ToList();
            stack.Children.Add(grid);

            card.Child = stack;
            shell.Children.Add(card);
        }

        return shell;
    }

    private static string FirstNonEmpty(params string?[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
}

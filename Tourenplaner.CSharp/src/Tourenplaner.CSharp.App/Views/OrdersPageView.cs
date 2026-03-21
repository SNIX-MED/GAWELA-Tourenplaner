using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class OrdersPageView : ScrollViewer
{
    private readonly IReadOnlyList<PinRecord> _pins;
    private readonly IReadOnlyList<Tour> _tours;
    private readonly SqlImportWorkspace _sqlImportWorkspace;
    private readonly Brush _panelBrush;
    private readonly Brush _textBrush;
    private readonly Brush _subTextBrush;
    private readonly TextBox _searchBox = new() { MinWidth = 220, Padding = new Thickness(8, 6, 8, 6) };
    private readonly DataGrid _grid = ViewStyling.CreateReadOnlyGrid();
    private readonly DataGrid _pendingGrid = ViewStyling.CreateReadOnlyGrid();
    private readonly TextBlock _detailText = new() { TextWrapping = TextWrapping.Wrap };

    public OrdersPageView(IReadOnlyList<PinRecord> pins, IReadOnlyList<Tour> tours, SqlImportWorkspace sqlImportWorkspace, Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        _pins = pins;
        _tours = tours;
        _sqlImportWorkspace = sqlImportWorkspace;
        _panelBrush = panelBrush;
        _textBrush = textBrush;
        _subTextBrush = subTextBrush;

        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build();
        RefreshGrid();
    }

    private UIElement Build()
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Auftragsliste", 22, FontWeights.SemiBold, _textBrush));
        shell.Children.Add(ViewStyling.Text("Die Auftragsseite kombiniert jetzt kartierte Pins mit den echten Pending-SQL-Aufträgen, dem Geocode-Cache und dem letzten Fehlerreport aus den Produktivdateien.", 13, FontWeights.Normal, _subTextBrush, new Thickness(0, 10, 0, 0)));
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Pins", _pins.Count.ToString(), "Kartierte/geocodierte Aufträge"),
            new PageStatViewModel("In Touren", CountAssignedPins().ToString(), "Pins mit gespeicherter Tourzuordnung"),
            new PageStatViewModel("Status gesetzt", _pins.Count(pin => !string.Equals(CleanStatus(pin.Status), "nicht festgelegt", StringComparison.OrdinalIgnoreCase)).ToString(), "Aufträge mit aktivem Status"),
            new PageStatViewModel("E-Mail hinterlegt", _pins.Count(pin => pin.Data.TryGetValue("Email", out var email) && !string.IsNullOrWhiteSpace(email)).ToString(), "Kontaktdatensätze mit E-Mail"),
            new PageStatViewModel("Pending SQL", _sqlImportWorkspace.PendingOrders.Count.ToString(), "Offene Import-Aufträge ohne Koordinaten"),
            new PageStatViewModel("Geocode-Cache", _sqlImportWorkspace.GeocodeCacheEntries.ToString(), "Gespeicherte Geocoding-Treffer"),
        }, _panelBrush, _textBrush, _subTextBrush));

        shell.Children.Add(BuildSqlWorkspaceCard());

        var searchBar = new WrapPanel { Margin = new Thickness(0, 16, 0, 0) };
        searchBar.Children.Add(ViewStyling.Text("Suche", 12, FontWeights.SemiBold, _textBrush, new Thickness(0, 8, 8, 0)));
        _searchBox.TextChanged += (_, _) => RefreshGrid();
        searchBar.Children.Add(_searchBox);
        shell.Children.Add(searchBar);

        var layout = new Grid { Margin = new Thickness(0, 16, 0, 0) };
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.2, GridUnitType.Star) });
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.8, GridUnitType.Star) });

        ConfigureGrid();
        ConfigurePendingGrid();
        var gridCard = new Border { Background = _panelBrush, CornerRadius = new CornerRadius(14), Padding = new Thickness(16), Margin = new Thickness(0, 0, 12, 0), Child = _grid };
        Grid.SetColumn(gridCard, 0);
        layout.Children.Add(gridCard);

        var detailCard = new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(16),
            Margin = new Thickness(12, 0, 0, 0),
            Child = new StackPanel
            {
                Children =
                {
                    ViewStyling.Text("Auftragsdetails", 16, FontWeights.SemiBold, _textBrush),
                    _detailText,
                },
            },
        };
        _detailText.Foreground = _subTextBrush;
        _detailText.Margin = new Thickness(0, 10, 0, 0);
        Grid.SetColumn(detailCard, 1);
        layout.Children.Add(detailCard);

        shell.Children.Add(layout);

        shell.Children.Add(ViewStyling.Text("Pending SQL-Aufträge", 16, FontWeights.SemiBold, _textBrush, new Thickness(0, 18, 0, 0)));
        shell.Children.Add(ViewStyling.Text("Diese Liste stammt direkt aus data/pending_sql_orders.json und zeigt Aufträge, die im Python-System noch kein erfolgreiches Geocoding erhalten haben.", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 8, 0, 0)));
        shell.Children.Add(new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(16),
            Margin = new Thickness(0, 12, 0, 0),
            Child = _pendingGrid,
        });
        return shell;
    }

    private void ConfigureGrid()
    {
        _grid.Columns.Add(new DataGridTextColumn { Header = "Auftrag", Binding = new System.Windows.Data.Binding(nameof(OrderRow.OrderNumber)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(OrderRow.Name)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Adresse", Binding = new System.Windows.Data.Binding(nameof(OrderRow.Address)), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Status", Binding = new System.Windows.Data.Binding(nameof(OrderRow.Status)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Gewicht", Binding = new System.Windows.Data.Binding(nameof(OrderRow.Weight)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Tour", Binding = new System.Windows.Data.Binding(nameof(OrderRow.Assignment)), Width = new DataGridLength(1.5, DataGridLengthUnitType.Star) });
        _grid.SelectionChanged += (_, _) => RefreshDetails(_grid.SelectedItem as OrderRow);
    }

    private void ConfigurePendingGrid()
    {
        _pendingGrid.MinHeight = 140;
        _pendingGrid.Columns.Add(new DataGridTextColumn { Header = "Auftrag", Binding = new System.Windows.Data.Binding(nameof(PendingOrderRow.OrderNumber)) });
        _pendingGrid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(PendingOrderRow.Name)) });
        _pendingGrid.Columns.Add(new DataGridTextColumn { Header = "Adresse", Binding = new System.Windows.Data.Binding(nameof(PendingOrderRow.Address)), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
        _pendingGrid.Columns.Add(new DataGridTextColumn { Header = "Lieferart", Binding = new System.Windows.Data.Binding(nameof(PendingOrderRow.DeliveryType)) });
        _pendingGrid.Columns.Add(new DataGridTextColumn { Header = "Status", Binding = new System.Windows.Data.Binding(nameof(PendingOrderRow.Status)) });
    }

    private void RefreshGrid()
    {
        var search = (_searchBox.Text ?? string.Empty).Trim();
        var rows = _pins
            .Select(BuildRow)
            .Where(row => string.IsNullOrWhiteSpace(search)
                || row.OrderNumber.Contains(search, StringComparison.OrdinalIgnoreCase)
                || row.Name.Contains(search, StringComparison.OrdinalIgnoreCase)
                || row.Address.Contains(search, StringComparison.OrdinalIgnoreCase)
                || row.Status.Contains(search, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var pendingRows = _sqlImportWorkspace.PendingOrders
            .Select(order => new PendingOrderRow(
                order.OrderNumber,
                FirstNonEmpty(order.Name, "-"),
                FirstNonEmpty(order.Address, "-"),
                FirstNonEmpty(order.DeliveryType, "-"),
                CleanStatus(order.Status)))
            .Where(row => string.IsNullOrWhiteSpace(search)
                || row.OrderNumber.Contains(search, StringComparison.OrdinalIgnoreCase)
                || row.Name.Contains(search, StringComparison.OrdinalIgnoreCase)
                || row.Address.Contains(search, StringComparison.OrdinalIgnoreCase)
                || row.Status.Contains(search, StringComparison.OrdinalIgnoreCase)
                || row.DeliveryType.Contains(search, StringComparison.OrdinalIgnoreCase))
            .ToList();
        _grid.ItemsSource = rows;
        _pendingGrid.ItemsSource = pendingRows;
        RefreshDetails(rows.FirstOrDefault());
        _grid.SelectedItem = rows.FirstOrDefault();
    }

    private void RefreshDetails(OrderRow? row)
    {
        if (row is null)
        {
            _detailText.Text = "Keine Aufträge gefunden.";
            return;
        }

        _detailText.Text = string.Join(Environment.NewLine, new[]
        {
            $"Auftrag: {row.OrderNumber}",
            $"Name: {row.Name}",
            $"Adresse: {row.Address}",
            $"Status: {row.Status}",
            $"Gewicht: {row.Weight}",
            $"Telefon: {row.Phone}",
            $"E-Mail: {row.Email}",
            $"Tourzuordnung: {row.Assignment}",
        });
    }

    private OrderRow BuildRow(PinRecord pin)
    {
        var data = pin.Data;
        var assignment = _tours.FirstOrDefault(tour => tour.Stops.Any(stop => string.Equals(stop.Id, GetOrderNumber(data), StringComparison.Ordinal) || string.Equals(stop.OrderNumber, GetOrderNumber(data), StringComparison.Ordinal)));
        return new OrderRow(
            GetOrderNumber(data),
            FirstNonEmpty(data.TryGetValue("Name", out var name) ? name : null, "-"),
            BuildAddress(data),
            CleanStatus(pin.Status),
            FirstNonEmpty(data.TryGetValue("Gewicht", out var weight) ? weight : null, "-"),
            FirstNonEmpty(data.TryGetValue("Telefon", out var phone) ? phone : null, "-"),
            FirstNonEmpty(data.TryGetValue("Email", out var email) ? email : null, "-"),
            assignment is null ? "Nicht zugeordnet" : $"{assignment.Date} · {assignment.Name}");
    }

    private int CountAssignedPins()
        => _pins.Count(pin => _tours.Any(tour => tour.Stops.Any(stop => string.Equals(stop.Id, GetOrderNumber(pin.Data), StringComparison.Ordinal) || string.Equals(stop.OrderNumber, GetOrderNumber(pin.Data), StringComparison.Ordinal))));

    private UIElement BuildSqlWorkspaceCard()
    {
        var latestReport = string.IsNullOrWhiteSpace(_sqlImportWorkspace.LatestGeocodeFailureReport)
            ? "Kein Geocoding-Fehlerreport gefunden."
            : _sqlImportWorkspace.LatestGeocodeFailureAt is DateTimeOffset timestamp
                ? $"{System.IO.Path.GetFileName(_sqlImportWorkspace.LatestGeocodeFailureReport)} · {timestamp.ToLocalTime():dd.MM.yyyy HH:mm}"
                : System.IO.Path.GetFileName(_sqlImportWorkspace.LatestGeocodeFailureReport);

        return new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(16),
            Margin = new Thickness(0, 16, 0, 0),
            Child = new StackPanel
            {
                Children =
                {
                    ViewStyling.Text("SQL-Import-Arbeitsstand", 16, FontWeights.SemiBold, _textBrush),
                    ViewStyling.Text($"Pending-Datei: {_sqlImportWorkspace.PendingOrders.Count} Einträge", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 8, 0, 0)),
                    ViewStyling.Text($"Nicht-Karten-Datei: {_sqlImportWorkspace.NonMapOrders.Count} Einträge", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 6, 0, 0)),
                    ViewStyling.Text($"Geocode-Cache: {_sqlImportWorkspace.GeocodeCacheEntries} Schlüssel", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 6, 0, 0)),
                    ViewStyling.Text($"Letzter Fehlerreport: {latestReport}", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 6, 0, 0)),
                },
            },
        };
    }

    private static string BuildAddress(IReadOnlyDictionary<string, string?> data)
    {
        var street = FirstNonEmpty(data.TryGetValue("Strasse", out var street) ? street : null, string.Empty);
        var zip = FirstNonEmpty(data.TryGetValue("PLZ", out var zip) ? zip : null, string.Empty);
        var city = FirstNonEmpty(data.TryGetValue("Ort", out var city) ? city : null, string.Empty);
        var zipCity = string.Join(" ", new[] { zip, city }.Where(value => !string.IsNullOrWhiteSpace(value)));
        return string.Join(", ", new[] { street, zipCity }.Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static string GetOrderNumber(IReadOnlyDictionary<string, string?> data)
        => FirstNonEmpty(data.TryGetValue("Auftragsnummer", out var orderNumber) ? orderNumber : null, "-");

    private static string CleanStatus(string? status)
        => string.IsNullOrWhiteSpace(status) ? "nicht festgelegt" : status.Trim();

    private static string FirstNonEmpty(params string?[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;

    private sealed record OrderRow(string OrderNumber, string Name, string Address, string Status, string Weight, string Phone, string Email, string Assignment);
    private sealed record PendingOrderRow(string OrderNumber, string Name, string Address, string DeliveryType, string Status);
}

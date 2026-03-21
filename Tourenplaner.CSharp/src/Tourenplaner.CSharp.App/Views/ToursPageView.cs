using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class ToursPageView : ScrollViewer
{
    private readonly IReadOnlyList<Tour> _tours;
    private readonly IReadOnlyList<Employee> _employees;
    private readonly VehicleCatalog _catalog;
    private readonly Brush _panelBrush;
    private readonly Brush _textBrush;
    private readonly Brush _subTextBrush;
    private readonly Action? _onAdd;
    private readonly Action? _onSaveFromCurrentRoute;
    private readonly Action<Tour>? _onShowOnMap;
    private readonly Action<Tour>? _onEdit;
    private readonly Action<Tour>? _onDelete;
    private readonly TextBox _dateFilter = new() { MinWidth = 120, Padding = new Thickness(8, 6, 8, 6) };
    private readonly TextBox _rangeStartFilter = new() { MinWidth = 120, Padding = new Thickness(8, 6, 8, 6) };
    private readonly TextBox _rangeEndFilter = new() { MinWidth = 120, Padding = new Thickness(8, 6, 8, 6) };
    private readonly TextBox _searchFilter = new() { MinWidth = 200, Padding = new Thickness(8, 6, 8, 6) };
    private readonly StackPanel _statsHost = new() { Margin = new Thickness(0, 16, 0, 0) };
    private readonly DataGrid _grid = ViewStyling.CreateReadOnlyGrid();
    private readonly TextBlock _detailText = new() { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 10, 0, 0) };
    private readonly TextBlock _conflictText = new() { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 10, 0, 0) };

    public ToursPageView(
        IReadOnlyList<Tour> tours,
        IReadOnlyList<Employee> employees,
        VehicleCatalog catalog,
        Brush panelBrush,
        Brush textBrush,
        Brush subTextBrush,
        Action? onAdd,
        Action? onSaveFromCurrentRoute,
        Action<Tour>? onShowOnMap,
        Action<Tour>? onEdit,
        Action<Tour>? onDelete)
    {
        _tours = tours;
        _employees = employees;
        _catalog = catalog;
        _panelBrush = panelBrush;
        _textBrush = textBrush;
        _subTextBrush = subTextBrush;
        _onAdd = onAdd;
        _onSaveFromCurrentRoute = onSaveFromCurrentRoute;
        _onShowOnMap = onShowOnMap;
        _onEdit = onEdit;
        _onDelete = onDelete;

        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build();
        RefreshView();
    }

    private UIElement Build()
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Liefertouren", 22, FontWeights.SemiBold, _textBrush));
        shell.Children.Add(ViewStyling.Text("Die Tourenansicht bietet jetzt Fachfilter, Totalgewicht, Ressourcenkonflikte, Karten-Sprung und das Speichern aus der aktuellen Kartenroute.", 13, FontWeights.Normal, _subTextBrush, new Thickness(0, 10, 0, 0)));
        shell.Children.Add(_statsHost);

        var filters = new WrapPanel { Margin = new Thickness(0, 16, 0, 0) };
        filters.Children.Add(ViewStyling.Text("Datum", 12, FontWeights.SemiBold, _textBrush, new Thickness(0, 8, 8, 0)));
        filters.Children.Add(_dateFilter);
        filters.Children.Add(ViewStyling.Text("Von", 12, FontWeights.SemiBold, _textBrush, new Thickness(12, 8, 8, 0)));
        filters.Children.Add(_rangeStartFilter);
        filters.Children.Add(ViewStyling.Text("Bis", 12, FontWeights.SemiBold, _textBrush, new Thickness(12, 8, 8, 0)));
        filters.Children.Add(_rangeEndFilter);
        filters.Children.Add(ViewStyling.Text("Suche", 12, FontWeights.SemiBold, _textBrush, new Thickness(12, 8, 8, 0)));
        filters.Children.Add(_searchFilter);
        var filterButton = CreateButton("Filter anwenden", (_, _) => RefreshView(), new Thickness(12, 0, 0, 0));
        filters.Children.Add(filterButton);
        foreach (var box in new[] { _dateFilter, _rangeStartFilter, _rangeEndFilter, _searchFilter })
        {
            box.TextChanged += (_, _) => RefreshView();
        }
        shell.Children.Add(filters);

        var actions = new WrapPanel { Margin = new Thickness(0, 16, 0, 0) };
        actions.Children.Add(CreateButton("Tour hinzufügen", (_, _) => _onAdd?.Invoke()));
        actions.Children.Add(CreateButton("Aus aktueller Route speichern", (_, _) => _onSaveFromCurrentRoute?.Invoke(), new Thickness(8, 0, 0, 0)));
        actions.Children.Add(CreateButton("Auf Karte zeigen", (_, _) => { if (_grid.SelectedItem is TourRow row) _onShowOnMap?.Invoke(row.Source); }, new Thickness(8, 0, 0, 0)));
        actions.Children.Add(CreateButton("Bearbeiten", (_, _) => { if (_grid.SelectedItem is TourRow row) _onEdit?.Invoke(row.Source); }, new Thickness(8, 0, 0, 0)));
        actions.Children.Add(CreateButton("Löschen", (_, _) => { if (_grid.SelectedItem is TourRow row) _onDelete?.Invoke(row.Source); }, new Thickness(8, 0, 0, 0)));
        shell.Children.Add(actions);

        var layout = new Grid { Margin = new Thickness(0, 16, 0, 0) };
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.25, GridUnitType.Star) });
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.75, GridUnitType.Star) });

        ConfigureGrid();
        var gridCard = new Border { Background = _panelBrush, CornerRadius = new CornerRadius(14), Padding = new Thickness(16), Margin = new Thickness(0, 0, 12, 0), Child = _grid };
        Grid.SetColumn(gridCard, 0);
        layout.Children.Add(gridCard);

        _detailText.Foreground = _subTextBrush;
        _conflictText.Foreground = Brushes.Orange;
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
                    ViewStyling.Text("Tourdetails", 16, FontWeights.SemiBold, _textBrush),
                    _detailText,
                    ViewStyling.Text("Ressourcenkonflikte", 16, FontWeights.SemiBold, _textBrush, new Thickness(0, 16, 0, 0)),
                    _conflictText,
                },
            },
        };
        Grid.SetColumn(detailCard, 1);
        layout.Children.Add(detailCard);

        shell.Children.Add(layout);
        return shell;
    }

    private void ConfigureGrid()
    {
        _grid.Columns.Add(new DataGridTextColumn { Header = "ID", Binding = new System.Windows.Data.Binding(nameof(TourRow.Id)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Datum", Binding = new System.Windows.Data.Binding(nameof(TourRow.Date)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(TourRow.Name)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Fahrzeug", Binding = new System.Windows.Data.Binding(nameof(TourRow.Vehicle)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Anhänger", Binding = new System.Windows.Data.Binding(nameof(TourRow.Trailer)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Mitarbeiter", Binding = new System.Windows.Data.Binding(nameof(TourRow.EmployeeSummary)), Width = new DataGridLength(1.5, DataGridLengthUnitType.Star) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Stopps", Binding = new System.Windows.Data.Binding(nameof(TourRow.StopCount)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Totalgewicht", Binding = new System.Windows.Data.Binding(nameof(TourRow.TotalWeight)) });
        _grid.Columns.Add(new DataGridTextColumn { Header = "Konflikte", Binding = new System.Windows.Data.Binding(nameof(TourRow.ConflictSummary)), Width = new DataGridLength(1.3, DataGridLengthUnitType.Star) });
        _grid.SelectionChanged += (_, _) => RefreshDetails(_grid.SelectedItem as TourRow);
    }

    private void RefreshView()
    {
        var rows = _tours.Select(BuildRow).Where(MatchesFilters).OrderByDescending(row => ParseDate(row.Date)).ThenBy(row => row.Name).ToList();
        _grid.ItemsSource = rows;
        _grid.SelectedItem = rows.FirstOrDefault();
        RefreshStats(rows);
        RefreshDetails(rows.FirstOrDefault());
    }

    private void RefreshStats(IReadOnlyList<TourRow> rows)
    {
        _statsHost.Children.Clear();
        _statsHost.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Touren", rows.Count.ToString(), "Gefilterte Tourdatensätze"),
            new PageStatViewModel("Stopps gesamt", rows.Sum(row => row.StopCount).ToString(), "Stopps über die aktuelle Auswahl"),
            new PageStatViewModel("Gewicht gesamt", $"{rows.Sum(row => row.TotalWeightValue):0} kg", "Summiertes Tourgewicht in der Auswahl"),
            new PageStatViewModel("Mit Konflikt", rows.Count(row => row.ConflictCount > 0).ToString(), "Touren mit Ressourcenkonflikt"),
        }, _panelBrush, _textBrush, _subTextBrush));
    }

    private void RefreshDetails(TourRow? row)
    {
        if (row is null)
        {
            _detailText.Text = "Keine Touren gefunden.";
            _conflictText.Text = "-";
            return;
        }

        _detailText.Text = string.Join(Environment.NewLine, new[]
        {
            $"Tour: {row.Date} · {row.Name}",
            $"Startzeit: {row.Source.StartTime}",
            $"Fahrzeug: {row.Vehicle}",
            $"Anhänger: {row.Trailer}",
            $"Mitarbeiter: {row.EmployeeSummary}",
            $"Stopps: {row.StopCount}",
            $"Totalgewicht: {row.TotalWeight}",
            $"Route: {BuildRoutePreview(row.Source)}",
        });
        _conflictText.Text = row.ConflictDetails.Count == 0 ? "Keine Ressourcenkonflikte erkannt." : string.Join(Environment.NewLine, row.ConflictDetails);
    }

    private TourRow BuildRow(Tour tour)
    {
        var employees = tour.EmployeeIds.Select(id => _employees.FirstOrDefault(employee => employee.Id == id)?.Name ?? id).ToList();
        var vehicle = _catalog.Vehicles.FirstOrDefault(item => item.Id == tour.VehicleId);
        var trailer = _catalog.Trailers.FirstOrDefault(item => item.Id == tour.TrailerId);
        var totalWeightValue = tour.Stops.Sum(stop => ParseWeight(stop.Weight));
        var conflicts = FindResourceConflicts(tour);
        return new TourRow(
            tour,
            tour.Id,
            tour.Date,
            string.IsNullOrWhiteSpace(tour.Name) ? $"Tour #{tour.Id}" : tour.Name,
            vehicle is null ? tour.VehicleId : $"{vehicle.Name} ({vehicle.LicensePlate})",
            trailer is null ? (string.IsNullOrWhiteSpace(tour.TrailerId) ? "-" : tour.TrailerId!) : $"{trailer.Name} ({trailer.LicensePlate})",
            employees.Count == 0 ? "-" : string.Join(", ", employees),
            tour.Stops.Count,
            totalWeightValue <= 0 ? "-" : $"{totalWeightValue:0} kg",
            totalWeightValue,
            conflicts.Count == 0 ? "-" : string.Join("; ", conflicts.Select(conflict => conflict.Kind)),
            conflicts.Count,
            conflicts.Select(conflict => conflict.Message).ToList());
    }

    private bool MatchesFilters(TourRow row)
    {
        if (!MatchesDateFilter(row.Date, _dateFilter.Text))
        {
            return false;
        }

        if (!MatchesRangeFilter(row.Date, _rangeStartFilter.Text, _rangeEndFilter.Text))
        {
            return false;
        }

        var search = (_searchFilter.Text ?? string.Empty).Trim();
        return string.IsNullOrWhiteSpace(search)
            || row.Name.Contains(search, StringComparison.OrdinalIgnoreCase)
            || row.Vehicle.Contains(search, StringComparison.OrdinalIgnoreCase)
            || row.EmployeeSummary.Contains(search, StringComparison.OrdinalIgnoreCase)
            || row.Date.Contains(search, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesDateFilter(string date, string filter)
        => string.IsNullOrWhiteSpace(filter) || string.Equals(date, filter.Trim(), StringComparison.Ordinal);

    private static bool MatchesRangeFilter(string date, string startFilter, string endFilter)
    {
        var current = ParseDate(date);
        if (current is null)
        {
            return false;
        }

        var start = ParseDate(startFilter);
        var end = ParseDate(endFilter);
        if (start is not null && current < start)
        {
            return false;
        }

        if (end is not null && current > end)
        {
            return false;
        }

        return true;
    }

    private List<ResourceConflict> FindResourceConflicts(Tour tour)
    {
        var conflicts = new List<ResourceConflict>();
        foreach (var other in _tours.Where(item => item.Id != tour.Id && string.Equals(item.Date, tour.Date, StringComparison.Ordinal)))
        {
            if (!string.IsNullOrWhiteSpace(tour.VehicleId) && string.Equals(other.VehicleId, tour.VehicleId, StringComparison.Ordinal))
            {
                conflicts.Add(new ResourceConflict("Fahrzeug", $"Fahrzeugkonflikt mit {other.Date} · {(string.IsNullOrWhiteSpace(other.Name) ? $"Tour #{other.Id}" : other.Name)}."));
            }

            if (!string.IsNullOrWhiteSpace(tour.TrailerId) && string.Equals(other.TrailerId, tour.TrailerId, StringComparison.Ordinal))
            {
                conflicts.Add(new ResourceConflict("Anhänger", $"Anhängerkonflikt mit {other.Date} · {(string.IsNullOrWhiteSpace(other.Name) ? $"Tour #{other.Id}" : other.Name)}."));
            }
        }

        return conflicts
            .GroupBy(conflict => conflict.Message)
            .Select(group => group.First())
            .ToList();
    }

    private static string BuildRoutePreview(Tour tour)
    {
        var stops = tour.Stops.OrderBy(stop => stop.Order).Take(4).Select(stop => stop.Name).Where(name => !string.IsNullOrWhiteSpace(name)).ToList();
        return stops.Count == 0 ? "Noch keine Stopps." : string.Join(" → ", stops) + (tour.Stops.Count > 4 ? " → …" : string.Empty);
    }

    private static DateTime? ParseDate(string? value)
        => DateTime.TryParseExact(value, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed
            : null;

    private static double ParseWeight(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }

        var cleaned = new string(value.Where(ch => char.IsDigit(ch) || ch is '.' or ',').ToArray()).Replace(',', '.');
        return double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) ? result : 0;
    }

    private Button CreateButton(string text, RoutedEventHandler handler, Thickness? margin = null)
    {
        var button = new Button
        {
            Content = text,
            Margin = margin ?? new Thickness(0),
            Padding = new Thickness(12, 8, 12, 8),
            Background = new SolidColorBrush(Color.FromRgb(49, 130, 206)),
            Foreground = Brushes.White,
            BorderBrush = Brushes.Transparent,
        };
        button.Click += handler;
        return button;
    }

    private sealed record ResourceConflict(string Kind, string Message);

    private sealed record TourRow(
        Tour Source,
        int Id,
        string Date,
        string Name,
        string Vehicle,
        string Trailer,
        string EmployeeSummary,
        int StopCount,
        string TotalWeight,
        double TotalWeightValue,
        string ConflictSummary,
        int ConflictCount,
        IReadOnlyList<string> ConflictDetails);
}

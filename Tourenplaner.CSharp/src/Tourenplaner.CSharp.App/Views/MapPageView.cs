using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Domain.Services;
using Tourenplaner.CSharp.Domain.ValueObjects;

namespace Tourenplaner.CSharp.App.Views;

public sealed class MapPageView : ScrollViewer
{
    private readonly AppSnapshotLike _snapshot;
    private readonly Brush _panelBrush;
    private readonly Brush _textBrush;
    private readonly Brush _subTextBrush;
    private readonly Dictionary<string, Employee> _employeeLookup;
    private readonly Dictionary<string, Vehicle> _vehicleLookup;
    private readonly Dictionary<string, Trailer> _trailerLookup;
    private readonly Dictionary<string, Tour> _tourByStopId;
    private readonly List<RouteStopDraft> _currentStops = [];
    private readonly Canvas _markerCanvas = new() { Width = 420, Height = 280, Background = Brushes.Transparent };
    private readonly StackPanel _routeStatsHost = new() { Margin = new Thickness(0, 16, 0, 0) };
    private readonly TextBlock _routeTitleText = new() { FontSize = 16, FontWeight = FontWeights.SemiBold };
    private readonly TextBlock _routeSummaryText = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _resourceSummaryText = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _pinDetailText = new() { TextWrapping = TextWrapping.Wrap };
    private readonly DataGrid _routeGrid = ViewStyling.CreateReadOnlyGrid();
    private readonly DataGrid _segmentGrid = ViewStyling.CreateReadOnlyGrid();
    private readonly DataGrid _pinGrid = ViewStyling.CreateReadOnlyGrid();
    private readonly ComboBox _tourSelector = new() { MinWidth = 320 };
    private readonly ComboBox _statusFilterSelector = new() { MinWidth = 220 };
    private readonly TextBox _startTimeTextBox = new() { MinWidth = 90, Text = "08:00" };
    private readonly Button _showTourButton = new() { Content = "Zugehörige Tour laden", Margin = new Thickness(6, 0, 0, 0), Padding = new Thickness(12, 6, 12, 6) };
    private readonly Button _addPinToRouteButton = new() { Content = "Zur Route hinzufügen", Padding = new Thickness(12, 6, 12, 6) };
    private readonly Action<RouteWorkspaceSnapshot>? _onRouteChanged;
    private readonly RouteWorkspaceSnapshot? _initialRoute;
    private readonly int? _initialTourId;

    private Tour? _currentTour;
    private PinRecord? _selectedPin;
    private string _currentRouteMode = "car";
    private Dictionary<string, int> _travelTimeCache = [];
    private IReadOnlyList<string> _currentEmployeeIds = Array.Empty<string>();
    private string _currentVehicleId = string.Empty;
    private string? _currentTrailerId;

    public MapPageView(AppSnapshotLike snapshot, Brush panelBrush, Brush textBrush, Brush subTextBrush, RouteWorkspaceSnapshot? initialRoute = null, Action<RouteWorkspaceSnapshot>? onRouteChanged = null, int? initialTourId = null)
    {
        _snapshot = snapshot;
        _panelBrush = panelBrush;
        _textBrush = textBrush;
        _subTextBrush = subTextBrush;
        _employeeLookup = snapshot.Employees.ToDictionary(employee => employee.Id, employee => employee);
        _vehicleLookup = snapshot.VehicleCatalog.Vehicles.ToDictionary(vehicle => vehicle.Id, vehicle => vehicle);
        _trailerLookup = snapshot.VehicleCatalog.Trailers.ToDictionary(trailer => trailer.Id, trailer => trailer);
        _tourByStopId = BuildTourStopLookup(snapshot.Tours);
        _initialRoute = initialRoute;
        _onRouteChanged = onRouteChanged;
        _initialTourId = initialTourId;

        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build();

        InitializeState();
    }

    public sealed record AppSnapshotLike(
        int PinCount,
        int TourCount,
        int EmployeeCount,
        int VehicleCount,
        IReadOnlyList<Tour> Tours,
        IReadOnlyList<Employee> Employees,
        VehicleCatalog VehicleCatalog,
        IReadOnlyList<PinRecord> Pins);

    public sealed record RouteWorkspaceSnapshot(
        int? TourId,
        string StartTime,
        string RouteMode,
        IReadOnlyList<TourStop> Stops,
        IReadOnlyList<string> EmployeeIds,
        string VehicleId,
        string? TrailerId,
        IReadOnlyDictionary<string, int> TravelTimeCache);

    private UIElement Build()
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Karte & Routenpanel", 22, FontWeights.SemiBold, _textBrush));
        shell.Children.Add(ViewStyling.Text(
            "Der Kartenbereich bildet jetzt die Kernlogik des Originalprogramms in WPF deutlich weiter ab: Tourkontext, Stoppliste, Zeitplanung, Optimierung, Markerübersicht und Auftragsdetails arbeiten zusammen auf Basis der echten JSON-Daten.",
            13,
            FontWeights.Normal,
            _subTextBrush,
            new Thickness(0, 10, 0, 0)));
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Pins", _snapshot.PinCount.ToString(), "Kartierte Aufträge"),
            new PageStatViewModel("Touren", _snapshot.TourCount.ToString(), "Gespeicherte Touren"),
            new PageStatViewModel("Mitarbeiter", _snapshot.EmployeeCount.ToString(), "Verfügbare Ressourcen"),
            new PageStatViewModel("Fahrzeuge", _snapshot.VehicleCount.ToString(), "Zugfahrzeuge im Bestand"),
        }, _panelBrush, _textBrush, _subTextBrush));

        var workspace = new Grid { Margin = new Thickness(0, 20, 0, 0) };
        workspace.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.25, GridUnitType.Star) });
        workspace.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.0, GridUnitType.Star) });

        var routePanel = BuildRoutePanel();
        Grid.SetColumn(routePanel, 0);
        workspace.Children.Add(routePanel);

        var markerPanel = BuildMarkerPanel();
        Grid.SetColumn(markerPanel, 1);
        workspace.Children.Add(markerPanel);

        shell.Children.Add(workspace);
        return shell;
    }

    private UIElement BuildRoutePanel()
    {
        var border = new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(18),
            Margin = new Thickness(0, 0, 12, 0),
        };

        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Routenpanel", 18, FontWeights.SemiBold, _textBrush));
        shell.Children.Add(ViewStyling.Text(
            "Tour laden, Stopps anpassen, Zeitfenster pflegen und die Reihenfolge optimieren – alles in einer gemeinsamen WPF-Arbeitsfläche.",
            12,
            FontWeights.Normal,
            _subTextBrush,
            new Thickness(0, 6, 0, 0)));

        var selectorGrid = new Grid { Margin = new Thickness(0, 14, 0, 0) };
        selectorGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        selectorGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        selectorGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        _tourSelector.Margin = new Thickness(0, 0, 8, 0);
        _tourSelector.DisplayMemberPath = nameof(TourSelectionItem.Label);
        _tourSelector.SelectionChanged += (_, _) =>
        {
            if (_tourSelector.SelectedItem is TourSelectionItem item)
            {
                LoadTour(item.Tour);
            }
        };
        Grid.SetColumn(_tourSelector, 0);
        selectorGrid.Children.Add(_tourSelector);

        var optimizeButton = CreateActionButton("Reihenfolge optimieren", (_, _) => OptimizeCurrentRoute(), new Thickness(0, 0, 8, 0));
        Grid.SetColumn(optimizeButton, 1);
        selectorGrid.Children.Add(optimizeButton);

        var unloadButton = CreateActionButton("Tour verlassen", (_, _) => ResetToManualRoute());
        Grid.SetColumn(unloadButton, 2);
        selectorGrid.Children.Add(unloadButton);
        shell.Children.Add(selectorGrid);

        var toolbar = new WrapPanel { Margin = new Thickness(0, 12, 0, 0), ItemHeight = 36 };
        toolbar.Children.Add(ViewStyling.Text("Startzeit", 12, FontWeights.SemiBold, _textBrush, new Thickness(0, 8, 8, 0)));
        _startTimeTextBox.Padding = new Thickness(8, 4, 8, 4);
        toolbar.Children.Add(_startTimeTextBox);
        toolbar.Children.Add(CreateActionButton("Übernehmen", (_, _) => ApplyRouteStartTime(), new Thickness(8, 0, 8, 0)));
        toolbar.Children.Add(CreateActionButton("Stopp bearbeiten", (_, _) => EditSelectedStop(), new Thickness(0, 0, 8, 0)));
        toolbar.Children.Add(CreateActionButton("Hoch", (_, _) => MoveSelectedStop(-1), new Thickness(0, 0, 8, 0)));
        toolbar.Children.Add(CreateActionButton("Runter", (_, _) => MoveSelectedStop(1), new Thickness(0, 0, 8, 0)));
        toolbar.Children.Add(CreateActionButton("Stopp entfernen", (_, _) => RemoveSelectedStop()));
        shell.Children.Add(toolbar);

        _routeStatsHost.Margin = new Thickness(0, 16, 0, 0);
        shell.Children.Add(_routeStatsHost);

        var detailGrid = new UniformGrid { Columns = 2, Margin = new Thickness(0, 16, 0, 0) };
        detailGrid.Children.Add(CreateInfoCard("Aktuelle Route", _routeTitleText, _routeSummaryText));
        detailGrid.Children.Add(CreateInfoCard("Ressourcen & Kontext", ViewStyling.Text(string.Empty, 1, FontWeights.Normal, Brushes.Transparent), _resourceSummaryText));
        shell.Children.Add(detailGrid);

        shell.Children.Add(ViewStyling.Text("Stoppliste", 16, FontWeights.SemiBold, _textBrush, new Thickness(0, 18, 0, 0)));
        ConfigureRouteGrid();
        shell.Children.Add(_routeGrid);

        var exportBar = new WrapPanel { Margin = new Thickness(0, 10, 0, 0) };
        exportBar.Children.Add(CreateActionButton("Routenexport kopieren", (_, _) => ExportRouteToClipboard()));
        shell.Children.Add(exportBar);

        shell.Children.Add(ViewStyling.Text("Segment- & Fahrzeiten", 16, FontWeights.SemiBold, _textBrush, new Thickness(0, 18, 0, 0)));
        ConfigureSegmentGrid();
        shell.Children.Add(_segmentGrid);

        border.Child = shell;
        return border;
    }

    private UIElement BuildMarkerPanel()
    {
        var border = new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(18),
            Margin = new Thickness(12, 0, 0, 0),
        };

        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Marker- & Auftragsbereich", 18, FontWeights.SemiBold, _textBrush));
        shell.Children.Add(ViewStyling.Text(
            "Die Geo-Vorschau zeigt alle kartierten Aufträge relativ zu ihren Koordinaten. Stopps der aktuellen Route werden hervorgehoben und können direkt aus dem Auftragskatalog übernommen werden.",
            12,
            FontWeights.Normal,
            _subTextBrush,
            new Thickness(0, 6, 0, 0)));

        shell.Children.Add(CreateLegendCard());

        var mapBorder = new Border
        {
            Margin = new Thickness(0, 14, 0, 0),
            Padding = new Thickness(12),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.FromRgb(74, 85, 104)),
            CornerRadius = new CornerRadius(12),
            Child = _markerCanvas,
        };
        shell.Children.Add(mapBorder);

        var filterBar = new WrapPanel { Margin = new Thickness(0, 14, 0, 0) };
        filterBar.Children.Add(ViewStyling.Text("Statusfilter", 12, FontWeights.SemiBold, _textBrush, new Thickness(0, 8, 8, 0)));
        _statusFilterSelector.DisplayMemberPath = nameof(FilterOption.Label);
        _statusFilterSelector.SelectionChanged += (_, _) => RefreshPinData();
        filterBar.Children.Add(_statusFilterSelector);
        shell.Children.Add(filterBar);

        shell.Children.Add(ViewStyling.Text("Auftragskatalog", 16, FontWeights.SemiBold, _textBrush, new Thickness(0, 18, 0, 0)));
        ConfigurePinGrid();
        shell.Children.Add(_pinGrid);

        var actions = new WrapPanel { Margin = new Thickness(0, 10, 0, 0) };
        _addPinToRouteButton.Click += (_, _) => AddSelectedPinToRoute();
        actions.Children.Add(_addPinToRouteButton);
        _showTourButton.Click += (_, _) => ShowSelectedPinTour();
        actions.Children.Add(_showTourButton);
        shell.Children.Add(actions);

        shell.Children.Add(ViewStyling.Text("Auftragsdetails", 16, FontWeights.SemiBold, _textBrush, new Thickness(0, 18, 0, 0)));
        shell.Children.Add(CreateInfoCard("Selektierter Auftrag", ViewStyling.Text(string.Empty, 1, FontWeights.Normal, Brushes.Transparent), _pinDetailText));

        border.Child = shell;
        return border;
    }

    private void InitializeState()
    {
        var tourOptions = _snapshot.Tours
            .OrderByDescending(tour => ParseTourDate(tour.Date))
            .ThenByDescending(tour => tour.Id)
            .Select(tour => new TourSelectionItem(tour, BuildTourLabel(tour)))
            .ToList();
        _tourSelector.ItemsSource = tourOptions;

        var statusOptions = new List<FilterOption>
        {
            new("Alle Status", null),
        };
        statusOptions.AddRange(_snapshot.Pins
            .Select(pin => CleanStatus(pin.Status))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value)
            .Select(value => new FilterOption(value, value)));
        _statusFilterSelector.ItemsSource = statusOptions;
        _statusFilterSelector.SelectedIndex = 0;

        if (_initialRoute is not null && _initialRoute.Stops.Count > 0)
        {
            ApplyRouteSnapshot(_initialRoute);
            if (_initialRoute.TourId is not null)
            {
                var selected = tourOptions.FirstOrDefault(item => item.Tour.Id == _initialRoute.TourId.Value);
                if (selected is not null)
                {
                    _tourSelector.SelectedItem = selected;
                }
            }
        }
        else if (_initialTourId is not null)
        {
            var selected = tourOptions.FirstOrDefault(item => item.Tour.Id == _initialTourId.Value);
            if (selected is not null)
            {
                _tourSelector.SelectedItem = selected;
            }
            else if (tourOptions.Count > 0)
            {
                _tourSelector.SelectedIndex = 0;
            }
            else
            {
                ResetToManualRoute();
            }
        }
        else if (tourOptions.Count > 0)
        {
            _tourSelector.SelectedIndex = 0;
        }
        else
        {
            ResetToManualRoute();
        }

        _selectedPin = _snapshot.Pins.FirstOrDefault();
        RefreshAll();
        NotifyRouteChanged();
    }

    private void LoadTour(Tour tour)
    {
        _currentTour = tour;
        _currentRouteMode = string.IsNullOrWhiteSpace(tour.RouteMode) ? "car" : tour.RouteMode;
        _startTimeTextBox.Text = string.IsNullOrWhiteSpace(tour.StartTime) ? "08:00" : tour.StartTime;
        _travelTimeCache = new Dictionary<string, int>(tour.TravelTimeCache);
        _currentEmployeeIds = tour.EmployeeIds;
        _currentVehicleId = tour.VehicleId;
        _currentTrailerId = tour.TrailerId;
        _currentStops.Clear();
        _currentStops.AddRange(tour.Stops
            .OrderBy(stop => stop.Order)
            .Select(stop => RouteStopDraft.FromTourStop(stop)));
        RecalculateRouteSchedule();
        RefreshAll();
        NotifyRouteChanged();
    }

    private void ResetToManualRoute()
    {
        _currentTour = null;
        _currentRouteMode = "car";
        _travelTimeCache = [];
        _currentEmployeeIds = Array.Empty<string>();
        _currentVehicleId = string.Empty;
        _currentTrailerId = null;
        if (_tourSelector.SelectedItem is not null)
        {
            _tourSelector.SelectedItem = null;
        }

        if (string.IsNullOrWhiteSpace(_startTimeTextBox.Text))
        {
            _startTimeTextBox.Text = "08:00";
        }

        _currentStops.Clear();
        RefreshAll();
        NotifyRouteChanged();
    }

    private void ApplyRouteStartTime()
    {
        if (TimeParser.ToMinutes(_startTimeTextBox.Text) is null)
        {
            MessageBox.Show("Bitte eine gültige Startzeit im Format HH:MM eingeben.", "Zeitplan", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        RecalculateRouteSchedule();
        RefreshAll();
        NotifyRouteChanged();
    }

    private void OptimizeCurrentRoute()
    {
        if (_currentStops.Count < 2)
        {
            MessageBox.Show("Für die Optimierung werden mindestens 2 Stopps benötigt.", "Route-Optimierung", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var startNode = new RouteNode("depot_start", "Start (Depot)", null, null, string.Empty, string.Empty, 0);
        var endNode = new RouteNode("depot_end", "Ende (Depot)", null, null, string.Empty, string.Empty, 0);
        var stopNodes = _currentStops.Select(stop => stop.ToRouteNode()).ToList();
        var result = RouteOptimizationService.Optimize(startNode, stopNodes, endNode, _startTimeTextBox.Text, _travelTimeCache);
        var reordered = new List<RouteStopDraft>();
        foreach (var node in result.Stops)
        {
            var stop = _currentStops.FirstOrDefault(item => item.Id == node.Id);
            if (stop is not null)
            {
                reordered.Add(stop);
            }
        }

        if (reordered.Count == _currentStops.Count)
        {
            _currentStops.Clear();
            _currentStops.AddRange(reordered);
            RecalculateRouteSchedule();
            RefreshAll();
            NotifyRouteChanged();
            MessageBox.Show(
                $"Optimierung abgeschlossen. Fahrt: {result.Metrics.DriveMinutes} Min., Warten: {result.Metrics.WaitMinutes} Min., Verspätung: {result.Metrics.LateMinutes} Min.",
                "Route-Optimierung",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }

    private void EditSelectedStop()
    {
        if (_routeGrid.SelectedItem is not RouteStopRow row)
        {
            MessageBox.Show("Bitte zuerst einen Stopp auswählen.", "Stopp bearbeiten", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var stop = _currentStops.FirstOrDefault(item => item.Id == row.StopId);
        if (stop is null)
        {
            return;
        }

        var dialog = new Window
        {
            Title = "Stopp bearbeiten",
            Width = 420,
            Height = 280,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize,
            Background = new SolidColorBrush(Color.FromRgb(27, 38, 59)),
        };

        var shell = new Grid { Margin = new Thickness(18) };
        for (var i = 0; i < 5; i++)
        {
            shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }
        shell.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var name = ViewStyling.Text(stop.Name, 16, FontWeights.SemiBold, Brushes.White);
        Grid.SetRow(name, 0);
        shell.Children.Add(name);

        var address = ViewStyling.Text(stop.Address, 12, FontWeights.Normal, new SolidColorBrush(Color.FromRgb(203, 213, 225)), new Thickness(0, 4, 0, 12));
        Grid.SetRow(address, 1);
        shell.Children.Add(address);

        var startBox = CreateDialogTextBox(stop.TimeWindowStart);
        var endBox = CreateDialogTextBox(stop.TimeWindowEnd);
        var serviceBox = CreateDialogTextBox(stop.ServiceMinutes.ToString(CultureInfo.InvariantCulture));

        shell.Children.Add(CreateDialogField("Zeitfenster Start", startBox, 2));
        shell.Children.Add(CreateDialogField("Zeitfenster Ende", endBox, 3));
        shell.Children.Add(CreateDialogField("Service-Minuten", serviceBox, 4));

        var statusText = ViewStyling.Text(string.Empty, 12, FontWeights.Normal, new SolidColorBrush(Color.FromRgb(191, 219, 254)), new Thickness(0, 10, 0, 0));
        Grid.SetRow(statusText, 5);
        shell.Children.Add(statusText);

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
        var cancel = CreateActionButton("Abbrechen", (_, _) => dialog.Close(), new Thickness(0, 0, 8, 0));
        var save = CreateActionButton("Speichern", (_, _) =>
        {
            if (!ValidateTimeWindow(startBox.Text, endBox.Text, out var error))
            {
                statusText.Text = error;
                statusText.Foreground = Brushes.OrangeRed;
                return;
            }

            if (!int.TryParse(serviceBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var serviceMinutes) || serviceMinutes < 0)
            {
                statusText.Text = "Bitte eine gültige nicht-negative Servicezeit eingeben.";
                statusText.Foreground = Brushes.OrangeRed;
                return;
            }

            stop.TimeWindowStart = startBox.Text.Trim();
            stop.TimeWindowEnd = endBox.Text.Trim();
            stop.ServiceMinutes = serviceMinutes;
            RecalculateRouteSchedule();
            RefreshAll();
            NotifyRouteChanged();
            dialog.Close();
        });
        buttons.Children.Add(cancel);
        buttons.Children.Add(save);
        Grid.SetRow(buttons, 6);
        shell.Children.Add(buttons);

        dialog.Content = shell;
        dialog.Owner = Window.GetWindow(this);
        dialog.ShowDialog();
    }

    private void MoveSelectedStop(int delta)
    {
        if (_routeGrid.SelectedItem is not RouteStopRow row)
        {
            return;
        }

        var index = _currentStops.FindIndex(stop => stop.Id == row.StopId);
        if (index < 0)
        {
            return;
        }

        var newIndex = index + delta;
        if (newIndex < 0 || newIndex >= _currentStops.Count)
        {
            return;
        }

        (_currentStops[index], _currentStops[newIndex]) = (_currentStops[newIndex], _currentStops[index]);
        RecalculateRouteSchedule();
        RefreshAll();
        NotifyRouteChanged();
        SelectRouteStop(row.StopId);
    }

    private void RemoveSelectedStop()
    {
        if (_routeGrid.SelectedItem is not RouteStopRow row)
        {
            return;
        }

        _currentStops.RemoveAll(stop => stop.Id == row.StopId);
        RecalculateRouteSchedule();
        RefreshAll();
        NotifyRouteChanged();
    }

    private void AddSelectedPinToRoute()
    {
        if (_selectedPin is null)
        {
            return;
        }

        var stopId = BuildStopIdFromPin(_selectedPin);
        if (_currentStops.Any(stop => stop.Id == stopId))
        {
            MessageBox.Show("Dieser Auftrag ist bereits in der aktuellen Route vorhanden.", "Route", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _currentStops.Add(RouteStopDraft.FromPin(_selectedPin, _currentStops.Count + 1));
        RecalculateRouteSchedule();
        RefreshAll();
        NotifyRouteChanged();
        SelectRouteStop(stopId);
    }

    private void ShowSelectedPinTour()
    {
        if (_selectedPin is null)
        {
            return;
        }

        var stopId = BuildStopIdFromPin(_selectedPin);
        if (_tourByStopId.TryGetValue(stopId, out var tour))
        {
            var target = (_tourSelector.ItemsSource as IEnumerable<TourSelectionItem>)?.FirstOrDefault(item => item.Tour.Id == tour.Id);
            if (target is not null)
            {
                _tourSelector.SelectedItem = target;
                return;
            }

            LoadTour(tour);
        }
        else
        {
            MessageBox.Show("Für diesen Auftrag wurde aktuell keine gespeicherte Tour gefunden.", "Tour", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void ExportRouteToClipboard()
    {
        var rows = BuildRouteRows();
        if (rows.Count == 0)
        {
            MessageBox.Show("Es sind noch keine Stopps in der aktuellen Route vorhanden.", "Routenexport", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var export = string.Join(Environment.NewLine, rows.Select(row =>
            $"{row.Order}. {row.Name} | {row.Address} | ETA {row.PlannedArrival} | ETD {row.PlannedDeparture} | Fenster {row.TimeWindow} | Gewicht {row.Weight}"));
        Clipboard.SetText(export);
        MessageBox.Show("Die aktuelle Route wurde in die Zwischenablage kopiert.", "Routenexport", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void RecalculateRouteSchedule()
    {
        for (var index = 0; index < _currentStops.Count; index++)
        {
            _currentStops[index].Order = index + 1;
        }

        if (_currentStops.Count == 0)
        {
            return;
        }

        var segments = new List<int?>();
        RouteStopDraft? previous = null;
        foreach (var stop in _currentStops)
        {
            segments.Add(EstimateTravelMinutes(previous, stop));
            previous = stop;
        }

        segments.Add(0);

        var schedule = SchedulePlanner.Compute(
            _currentStops.Select(stop => new ScheduleStop(stop.Name, stop.TimeWindowStart, stop.TimeWindowEnd, stop.ServiceMinutes, stop.Weight, stop.OrderNumber, stop.Address)).ToList(),
            segments,
            _startTimeTextBox.Text);

        for (var index = 0; index < _currentStops.Count; index++)
        {
            var scheduled = schedule.Stops[index];
            _currentStops[index].PlannedArrival = scheduled.PlannedArrival;
            _currentStops[index].PlannedDeparture = scheduled.PlannedDeparture;
            _currentStops[index].ScheduleConflict = scheduled.ScheduleConflict;
            _currentStops[index].ScheduleConflictText = scheduled.ScheduleConflictText;
            _currentStops[index].WaitMinutes = scheduled.WaitMinutes;
        }
    }

    private int EstimateTravelMinutes(RouteStopDraft? from, RouteStopDraft to)
    {
        var fromId = from?.Id ?? "depot_start";
        var cacheKey = $"{fromId}->{to.Id}";
        if (_travelTimeCache.TryGetValue(cacheKey, out var cached))
        {
            return Math.Max(1, cached);
        }

        if (from is null || from.Latitude is null || from.Longitude is null || to.Latitude is null || to.Longitude is null)
        {
            return 3;
        }

        var distance = EstimateDistanceKm(from.Latitude.Value, from.Longitude.Value, to.Latitude.Value, to.Longitude.Value);
        return Math.Max(3, (int)Math.Round((distance / 42.0) * 60.0));
    }

    private void RefreshAll()
    {
        RefreshRouteStats();
        RefreshRouteDetails();
        RefreshRouteGrids();
        RefreshPinData();
        RefreshPinDetails();
        RefreshMarkerCanvas();
    }

    private void RefreshRouteStats()
    {
        _routeStatsHost.Children.Clear();

        var totalWeightKg = _currentStops.Sum(stop => ParseWeightKg(stop.Weight));
        var totalService = _currentStops.Sum(stop => Math.Max(0, stop.ServiceMinutes));
        var totalWait = _currentStops.Sum(stop => Math.Max(0, stop.WaitMinutes));
        var conflicts = _currentStops.Count(stop => stop.ScheduleConflict);
        var startTime = _startTimeTextBox.Text;
        var endTime = _currentStops.LastOrDefault()?.PlannedDeparture;

        _routeStatsHost.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Stopps", _currentStops.Count.ToString(), "Aktive Stopps in der Route"),
            new PageStatViewModel("Start / Ende", $"{startTime} → {(string.IsNullOrWhiteSpace(endTime) ? "-" : endTime)}", "Zeitplanung der aktuellen Route"),
            new PageStatViewModel("Service", $"{totalService} Min.", "Summe Aufenthaltszeiten"),
            new PageStatViewModel("Gewicht", totalWeightKg <= 0 ? "-" : $"{totalWeightKg:0} kg", "Gesamtgewicht der Route"),
            new PageStatViewModel("Wartezeit", $"{totalWait} Min.", "Zeitfensterbedingte Wartezeit"),
            new PageStatViewModel("Konflikte", conflicts.ToString(), conflicts == 0 ? "Keine erkannten Konflikte" : "Zeitfensterkonflikte vorhanden"),
        }, _panelBrush, _textBrush, _subTextBrush));
    }

    private void RefreshRouteDetails()
    {
        _routeTitleText.Text = _currentTour is null ? "Manuelle Route" : BuildTourLabel(_currentTour);
        _routeTitleText.Foreground = _textBrush;

        var routeSequence = _currentStops.Count == 0
            ? "Noch keine Stopps geplant."
            : $"Start → {string.Join(" → ", _currentStops.Take(5).Select(stop => stop.Name))}{(_currentStops.Count > 5 ? " → …" : string.Empty)} → Ende";
        _routeSummaryText.Text = string.Join(Environment.NewLine, new[]
        {
            $"Modus: {_currentRouteMode}",
            $"Reihenfolge: {routeSequence}",
            $"Erster Stopp: {_currentStops.FirstOrDefault()?.Address ?? "-"}",
            $"Letzter Stopp: {_currentStops.LastOrDefault()?.Address ?? "-"}",
        });
        _routeSummaryText.Foreground = _subTextBrush;

        if (_currentTour is null)
        {
            _resourceSummaryText.Text = "Kein gespeicherter Tourdatensatz geladen. Aufträge können rechts ausgewählt und zur Route hinzugefügt werden.";
        }
        else
        {
            _resourceSummaryText.Text = BuildResourceSummary(_currentTour);
        }
        _resourceSummaryText.Foreground = _subTextBrush;
    }

    private void RefreshRouteGrids()
    {
        var selectedStopId = (_routeGrid.SelectedItem as RouteStopRow)?.StopId;
        _routeGrid.ItemsSource = BuildRouteRows();
        _segmentGrid.ItemsSource = BuildSegmentRows();
        if (!string.IsNullOrWhiteSpace(selectedStopId))
        {
            SelectRouteStop(selectedStopId);
        }
    }

    private List<RouteStopRow> BuildRouteRows()
        => _currentStops.Select(stop => new RouteStopRow(
            stop.Id,
            stop.Order,
            string.IsNullOrWhiteSpace(stop.OrderNumber) ? "-" : stop.OrderNumber,
            stop.Name,
            stop.Address,
            FormatTimeWindow(stop.TimeWindowStart, stop.TimeWindowEnd),
            string.IsNullOrWhiteSpace(stop.PlannedArrival) ? "-" : stop.PlannedArrival,
            string.IsNullOrWhiteSpace(stop.PlannedDeparture) ? "-" : stop.PlannedDeparture,
            $"{Math.Max(0, stop.ServiceMinutes)} Min.",
            $"{Math.Max(0, stop.WaitMinutes)} Min.",
            string.IsNullOrWhiteSpace(stop.Weight) ? "-" : stop.Weight,
            stop.ScheduleConflict ? FirstNonEmpty(stop.ScheduleConflictText, "Konflikt") : "OK"))
            .ToList();

    private List<SegmentRow> BuildSegmentRows()
    {
        if (_currentStops.Count == 0)
        {
            return [];
        }

        var rows = new List<SegmentRow>();
        RouteStopDraft? previous = null;
        var previousLabel = "Start";
        foreach (var stop in _currentStops)
        {
            var travel = EstimateTravelMinutes(previous, stop);
            rows.Add(new SegmentRow(
                previousLabel,
                stop.Name,
                travel,
                previous?.PlannedDeparture ?? _startTimeTextBox.Text,
                stop.PlannedArrival,
                stop.WaitMinutes,
                stop.ScheduleConflict ? FirstNonEmpty(stop.ScheduleConflictText, "Konflikt") : string.Empty));
            previous = stop;
            previousLabel = stop.Name;
        }

        rows.Add(new SegmentRow(previousLabel, "Ende", 0, previous?.PlannedDeparture ?? _startTimeTextBox.Text, previous?.PlannedDeparture ?? string.Empty, 0, string.Empty));
        return rows;
    }

    private void RefreshPinData()
    {
        var selectedKey = BuildStopIdFromPin(_selectedPin);
        var filterValue = (_statusFilterSelector.SelectedItem as FilterOption)?.Value;
        var pins = _snapshot.Pins
            .Where(pin => filterValue is null || string.Equals(CleanStatus(pin.Status), filterValue, StringComparison.OrdinalIgnoreCase))
            .Select((pin, index) => BuildPinRow(pin, index + 1))
            .ToList();

        _pinGrid.ItemsSource = pins;
        if (!string.IsNullOrWhiteSpace(selectedKey))
        {
            var selected = pins.FirstOrDefault(pin => pin.StopId == selectedKey);
            if (selected is not null)
            {
                _pinGrid.SelectedItem = selected;
            }
        }
    }

    private PinRow BuildPinRow(PinRecord pin, int order)
    {
        var stopId = BuildStopIdFromPin(pin);
        var data = pin.Data;
        var address = BuildAddress(data);
        var inCurrentRoute = _currentStops.Any(stop => stop.Id == stopId);
        _tourByStopId.TryGetValue(stopId, out var assignedTour);
        return new PinRow(
            stopId,
            order,
            data.TryGetValue("Auftragsnummer", out var orderNumber) ? orderNumber ?? string.Empty : string.Empty,
            data.TryGetValue("Name", out var name) ? name ?? string.Empty : string.Empty,
            address,
            CleanStatus(pin.Status),
            data.TryGetValue("Gewicht", out var weight) ? weight ?? string.Empty : string.Empty,
            inCurrentRoute ? "Aktuelle Route" : assignedTour is null ? "Frei" : BuildTourLabel(assignedTour));
    }

    private void RefreshPinDetails()
    {
        if (_selectedPin is null)
        {
            _pinDetailText.Text = "Kein Auftrag ausgewählt.";
            _pinDetailText.Foreground = _subTextBrush;
            _showTourButton.IsEnabled = false;
            _addPinToRouteButton.IsEnabled = false;
            return;
        }

        var data = _selectedPin.Data;
        var stopId = BuildStopIdFromPin(_selectedPin);
        _tourByStopId.TryGetValue(stopId, out var tour);
        _pinDetailText.Text = string.Join(Environment.NewLine, new[]
        {
            $"Auftrag: {FirstNonEmpty(data.TryGetValue("Auftragsnummer", out var orderNumber) ? orderNumber : null, "-")}",
            $"Name: {FirstNonEmpty(data.TryGetValue("Name", out var name) ? name : null, "-")}",
            $"Adresse: {BuildAddress(data)}",
            $"Status: {CleanStatus(_selectedPin.Status)}",
            $"E-Mail: {FirstNonEmpty(data.TryGetValue("Email", out var email) ? email : null, "-")}",
            $"Telefon: {FirstNonEmpty(data.TryGetValue("Telefon", out var phone) ? phone : null, "-")}",
            $"Gewicht: {FirstNonEmpty(data.TryGetValue("Gewicht", out var weight) ? weight : null, "-")}",
            $"Tourzuordnung: {(tour is null ? "Keine gespeicherte Tour" : BuildTourLabel(tour))}",
        });
        _pinDetailText.Foreground = _subTextBrush;
        _showTourButton.IsEnabled = tour is not null;
        _addPinToRouteButton.IsEnabled = !_currentStops.Any(stop => stop.Id == stopId);
    }

    private void RefreshMarkerCanvas()
    {
        _markerCanvas.Children.Clear();
        if (_snapshot.Pins.Count == 0)
        {
            return;
        }

        const double width = 420;
        const double height = 280;
        var minLat = _snapshot.Pins.Min(pin => pin.Latitude);
        var maxLat = _snapshot.Pins.Max(pin => pin.Latitude);
        var minLon = _snapshot.Pins.Min(pin => pin.Longitude);
        var maxLon = _snapshot.Pins.Max(pin => pin.Longitude);
        var latRange = Math.Max(0.0001, maxLat - minLat);
        var lonRange = Math.Max(0.0001, maxLon - minLon);

        var routePoints = _currentStops
            .Where(stop => stop.Latitude is not null && stop.Longitude is not null)
            .Select(stop => ToCanvasPoint(stop.Latitude!.Value, stop.Longitude!.Value, minLat, latRange, minLon, lonRange, width, height))
            .ToList();
        if (routePoints.Count > 1)
        {
            _markerCanvas.Children.Add(new Polyline
            {
                Stroke = new SolidColorBrush(Color.FromRgb(99, 179, 237)),
                StrokeThickness = 3,
                Points = new PointCollection(routePoints),
            });
        }

        foreach (var pin in _snapshot.Pins)
        {
            var stopId = BuildStopIdFromPin(pin);
            var point = ToCanvasPoint(pin.Latitude, pin.Longitude, minLat, latRange, minLon, lonRange, width, height);
            var marker = new Button
            {
                Width = 18,
                Height = 18,
                ToolTip = $"{FirstNonEmpty(pin.Data.TryGetValue("Name", out var name) ? name : null, "Auftrag")}\n{BuildAddress(pin.Data)}",
                Background = GetStatusBrush(pin.Status),
                BorderBrush = _currentStops.Any(stop => stop.Id == stopId) ? Brushes.White : Brushes.Transparent,
                BorderThickness = _currentStops.Any(stop => stop.Id == stopId) ? new Thickness(2) : new Thickness(0),
                Tag = stopId,
            };
            marker.Click += (_, _) =>
            {
                _selectedPin = _snapshot.Pins.FirstOrDefault(candidate => BuildStopIdFromPin(candidate) == stopId);
                RefreshPinData();
                RefreshPinDetails();
                RefreshMarkerCanvas();
            };
            Canvas.SetLeft(marker, point.X - (marker.Width / 2));
            Canvas.SetTop(marker, point.Y - (marker.Height / 2));
            _markerCanvas.Children.Add(marker);
        }
    }

    private Border CreateLegendCard()
    {
        var counts = _snapshot.Pins
            .GroupBy(pin => CleanStatus(pin.Status), StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .ToList();

        var wrap = new WrapPanel { Margin = new Thickness(0, 12, 0, 0) };
        foreach (var group in counts)
        {
            var item = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(40, 255, 255, 255)),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(10, 6, 10, 6),
                Margin = new Thickness(0, 0, 8, 8),
                Child = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Children =
                    {
                        new Ellipse { Width = 10, Height = 10, Fill = GetStatusBrush(group.Key), Margin = new Thickness(0, 4, 8, 0) },
                        ViewStyling.Text($"{group.Key}: {group.Count()}", 12, FontWeights.SemiBold, _textBrush),
                    },
                },
            };
            wrap.Children.Add(item);
        }

        return new Border
        {
            Margin = new Thickness(0, 14, 0, 0),
            Background = new SolidColorBrush(Color.FromArgb(28, 255, 255, 255)),
            CornerRadius = new CornerRadius(12),
            Padding = new Thickness(12),
            Child = wrap,
        };
    }

    private Border CreateInfoCard(string title, TextBlock header, TextBlock content)
    {
        header.Text = title;
        header.Foreground = _textBrush;
        content.Foreground = _subTextBrush;
        return new Border
        {
            Background = new SolidColorBrush(Color.FromArgb(22, 255, 255, 255)),
            CornerRadius = new CornerRadius(12),
            Padding = new Thickness(14),
            Margin = new Thickness(0, 0, 12, 12),
            Child = new StackPanel
            {
                Children =
                {
                    header,
                    content,
                },
            },
        };
    }

    private Button CreateActionButton(string text, RoutedEventHandler onClick, Thickness? margin = null)
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
        button.Click += onClick;
        return button;
    }

    private TextBox CreateDialogTextBox(string value)
        => new()
        {
            Text = value,
            Padding = new Thickness(8, 6, 8, 6),
            Margin = new Thickness(0, 6, 0, 0),
        };

    private UIElement CreateDialogField(string label, TextBox textBox, int row)
    {
        var panel = new StackPanel();
        panel.Children.Add(ViewStyling.Text(label, 12, FontWeights.SemiBold, Brushes.White));
        panel.Children.Add(textBox);
        Grid.SetRow(panel, row);
        return panel;
    }

    private void ConfigureRouteGrid()
    {
        _routeGrid.MinHeight = 200;
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "#", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.Order)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Auftrag", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.OrderNumber)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.Name)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Adresse", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.Address)), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Zeitfenster", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.TimeWindow)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "ETA", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.PlannedArrival)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "ETD", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.PlannedDeparture)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Service", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.Service)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Warten", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.Wait)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Gewicht", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.Weight)) });
        _routeGrid.Columns.Add(new DataGridTextColumn { Header = "Status", Binding = new System.Windows.Data.Binding(nameof(RouteStopRow.Status)) });
    }

    private void ConfigureSegmentGrid()
    {
        _segmentGrid.MinHeight = 160;
        _segmentGrid.Columns.Add(new DataGridTextColumn { Header = "Von", Binding = new System.Windows.Data.Binding(nameof(SegmentRow.From)) });
        _segmentGrid.Columns.Add(new DataGridTextColumn { Header = "Nach", Binding = new System.Windows.Data.Binding(nameof(SegmentRow.To)) });
        _segmentGrid.Columns.Add(new DataGridTextColumn { Header = "Fahrt", Binding = new System.Windows.Data.Binding(nameof(SegmentRow.TravelLabel)) });
        _segmentGrid.Columns.Add(new DataGridTextColumn { Header = "Abfahrt", Binding = new System.Windows.Data.Binding(nameof(SegmentRow.Departure)) });
        _segmentGrid.Columns.Add(new DataGridTextColumn { Header = "Ankunft", Binding = new System.Windows.Data.Binding(nameof(SegmentRow.Arrival)) });
        _segmentGrid.Columns.Add(new DataGridTextColumn { Header = "Warten", Binding = new System.Windows.Data.Binding(nameof(SegmentRow.WaitLabel)) });
        _segmentGrid.Columns.Add(new DataGridTextColumn { Header = "Hinweis", Binding = new System.Windows.Data.Binding(nameof(SegmentRow.Note)), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
    }

    private void ConfigurePinGrid()
    {
        _pinGrid.MinHeight = 220;
        _pinGrid.SelectionChanged += (_, _) =>
        {
            if (_pinGrid.SelectedItem is PinRow row)
            {
                _selectedPin = _snapshot.Pins.FirstOrDefault(pin => BuildStopIdFromPin(pin) == row.StopId);
                RefreshPinDetails();
                RefreshMarkerCanvas();
            }
        };
        _pinGrid.Columns.Add(new DataGridTextColumn { Header = "#", Binding = new System.Windows.Data.Binding(nameof(PinRow.Order)) });
        _pinGrid.Columns.Add(new DataGridTextColumn { Header = "Auftrag", Binding = new System.Windows.Data.Binding(nameof(PinRow.OrderNumber)) });
        _pinGrid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(PinRow.Name)) });
        _pinGrid.Columns.Add(new DataGridTextColumn { Header = "Adresse", Binding = new System.Windows.Data.Binding(nameof(PinRow.Address)), Width = new DataGridLength(2, DataGridLengthUnitType.Star) });
        _pinGrid.Columns.Add(new DataGridTextColumn { Header = "Status", Binding = new System.Windows.Data.Binding(nameof(PinRow.Status)) });
        _pinGrid.Columns.Add(new DataGridTextColumn { Header = "Gewicht", Binding = new System.Windows.Data.Binding(nameof(PinRow.Weight)) });
        _pinGrid.Columns.Add(new DataGridTextColumn { Header = "Zuordnung", Binding = new System.Windows.Data.Binding(nameof(PinRow.Assignment)), Width = new DataGridLength(1.5, DataGridLengthUnitType.Star) });
    }

    private void SelectRouteStop(string stopId)
    {
        if (_routeGrid.ItemsSource is not IEnumerable<RouteStopRow> rows)
        {
            return;
        }

        var row = rows.FirstOrDefault(item => item.StopId == stopId);
        if (row is not null)
        {
            _routeGrid.SelectedItem = row;
            _routeGrid.ScrollIntoView(row);
        }
    }

    private void ApplyRouteSnapshot(RouteWorkspaceSnapshot snapshot)
    {
        _currentTour = snapshot.TourId is int tourId ? _snapshot.Tours.FirstOrDefault(tour => tour.Id == tourId) : null;
        _currentRouteMode = string.IsNullOrWhiteSpace(snapshot.RouteMode) ? "car" : snapshot.RouteMode;
        _startTimeTextBox.Text = string.IsNullOrWhiteSpace(snapshot.StartTime) ? "08:00" : snapshot.StartTime;
        _travelTimeCache = new Dictionary<string, int>(snapshot.TravelTimeCache);
        _currentEmployeeIds = snapshot.EmployeeIds;
        _currentVehicleId = snapshot.VehicleId;
        _currentTrailerId = snapshot.TrailerId;
        _currentStops.Clear();
        _currentStops.AddRange(snapshot.Stops.OrderBy(stop => stop.Order).Select(stop => RouteStopDraft.FromTourStop(stop)));
        RecalculateRouteSchedule();
    }

    private void NotifyRouteChanged()
    {
        _onRouteChanged?.Invoke(new RouteWorkspaceSnapshot(
            _currentTour?.Id,
            _startTimeTextBox.Text,
            _currentRouteMode,
            _currentStops.Select(stop => stop.ToTourStop()).ToList(),
            _currentEmployeeIds,
            _currentVehicleId,
            _currentTrailerId,
            new Dictionary<string, int>(_travelTimeCache)));
    }

    private string BuildResourceSummary(Tour tour)
    {
        var employees = tour.EmployeeIds.Select(id => _employeeLookup.TryGetValue(id, out var employee) ? employee.Name : id).ToList();
        var vehicleText = _vehicleLookup.TryGetValue(tour.VehicleId, out var vehicle)
            ? $"{vehicle.Name} ({vehicle.LicensePlate})"
            : "-";
        var trailerText = !string.IsNullOrWhiteSpace(tour.TrailerId) && _trailerLookup.TryGetValue(tour.TrailerId, out var trailer)
            ? $"{trailer.Name} ({trailer.LicensePlate})"
            : "Kein Anhänger";
        return string.Join(Environment.NewLine, new[]
        {
            $"Mitarbeiter: {(employees.Count == 0 ? "-" : string.Join(", ", employees))}",
            $"Fahrzeug: {vehicleText}",
            $"Anhänger: {trailerText}",
            $"Tourdatum: {tour.Date}",
            $"Startzeit: {tour.StartTime}",
        });
    }

    private static string BuildAddress(IReadOnlyDictionary<string, string?> data)
    {
        var street = FirstNonEmpty(data.TryGetValue("Strasse", out var streetValue) ? streetValue : null, string.Empty);
        var zip = FirstNonEmpty(data.TryGetValue("PLZ", out var zipValue) ? zipValue : null, string.Empty);
        var city = FirstNonEmpty(data.TryGetValue("Ort", out var cityValue) ? cityValue : null, string.Empty);
        var zipCity = string.Join(" ", new[] { zip, city }.Where(value => !string.IsNullOrWhiteSpace(value)));
        return string.Join(", ", new[] { street, zipCity }.Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static string BuildTourLabel(Tour tour)
    {
        var label = $"{tour.Date} · #{tour.Id}";
        if (!string.IsNullOrWhiteSpace(tour.Name))
        {
            label += $" · {tour.Name}";
        }

        label += $" · {tour.Stops.Count} Stopps";
        return label;
    }

    private static string BuildStopIdFromPin(PinRecord? pin)
    {
        if (pin is null)
        {
            return string.Empty;
        }

        if (pin.Data.TryGetValue("Auftragsnummer", out var orderNumber) && !string.IsNullOrWhiteSpace(orderNumber))
        {
            return orderNumber.Trim();
        }

        return $"{pin.Latitude:0.000000}:{pin.Longitude:0.000000}";
    }

    private static Dictionary<string, Tour> BuildTourStopLookup(IEnumerable<Tour> tours)
    {
        var result = new Dictionary<string, Tour>();
        foreach (var tour in tours)
        {
            foreach (var stop in tour.Stops)
            {
                var key = stop.Id;
                if (!string.IsNullOrWhiteSpace(key))
                {
                    result[key] = tour;
                }

                if (!string.IsNullOrWhiteSpace(stop.OrderNumber))
                {
                    result[stop.OrderNumber] = tour;
                }
            }
        }

        return result;
    }

    private static string FormatTimeWindow(string? start, string? end)
    {
        var cleanStart = string.IsNullOrWhiteSpace(start) ? string.Empty : start.Trim();
        var cleanEnd = string.IsNullOrWhiteSpace(end) ? string.Empty : end.Trim();
        return string.IsNullOrWhiteSpace(cleanStart + cleanEnd)
            ? "-"
            : $"{(string.IsNullOrWhiteSpace(cleanStart) ? "--:--" : cleanStart)}–{(string.IsNullOrWhiteSpace(cleanEnd) ? "--:--" : cleanEnd)}";
    }

    private static bool ValidateTimeWindow(string? start, string? end, out string error)
    {
        var startText = (start ?? string.Empty).Trim();
        var endText = (end ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(startText) && string.IsNullOrWhiteSpace(endText))
        {
            error = string.Empty;
            return true;
        }

        var startMinutes = string.IsNullOrWhiteSpace(startText) ? null : TimeParser.ToMinutes(startText);
        var endMinutes = string.IsNullOrWhiteSpace(endText) ? null : TimeParser.ToMinutes(endText);
        if (!string.IsNullOrWhiteSpace(startText) && startMinutes is null)
        {
            error = "Ungültige Startzeit. Bitte HH:MM verwenden.";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(endText) && endMinutes is null)
        {
            error = "Ungültige Endzeit. Bitte HH:MM verwenden.";
            return false;
        }

        if (startMinutes is not null && endMinutes is not null && startMinutes > endMinutes)
        {
            error = "Das Zeitfenster-Ende muss nach dem Start liegen.";
            return false;
        }

        error = string.Empty;
        return true;
    }

    private static string CleanStatus(string? status)
        => string.IsNullOrWhiteSpace(status) ? "nicht festgelegt" : status.Trim();

    private static string FirstNonEmpty(params string?[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;

    private static DateTime ParseTourDate(string? value)
        => DateTime.TryParseExact(value, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed
            : DateTime.MinValue;

    private static double ParseWeightKg(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }

        var cleaned = new string(value.Where(ch => char.IsDigit(ch) || ch is '.' or ',').ToArray()).Replace(',', '.');
        return double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) ? result : 0;
    }

    private static double EstimateDistanceKm(double latA, double lonA, double latB, double lonB)
    {
        const double radius = 6371.0;
        var dLat = DegreesToRadians(latB - latA);
        var dLon = DegreesToRadians(lonB - lonA);
        var h = Math.Pow(Math.Sin(dLat / 2.0), 2)
            + Math.Cos(DegreesToRadians(latA))
            * Math.Cos(DegreesToRadians(latB))
            * Math.Pow(Math.Sin(dLon / 2.0), 2);
        return radius * (2.0 * Math.Atan2(Math.Sqrt(h), Math.Sqrt(Math.Max(0.0, 1.0 - h))));
    }

    private static double DegreesToRadians(double value) => value * Math.PI / 180.0;

    private static Point ToCanvasPoint(double latitude, double longitude, double minLat, double latRange, double minLon, double lonRange, double width, double height)
    {
        var x = ((longitude - minLon) / lonRange) * (width - 30) + 15;
        var y = (1.0 - ((latitude - minLat) / latRange)) * (height - 30) + 15;
        return new Point(x, y);
    }

    private static Brush GetStatusBrush(string? status)
        => CleanStatus(status).ToLowerInvariant() switch
        {
            "auf dem weg" => new SolidColorBrush(Color.FromRgb(72, 187, 120)),
            "erledigt" => new SolidColorBrush(Color.FromRgb(66, 153, 225)),
            "storniert" => new SolidColorBrush(Color.FromRgb(245, 101, 101)),
            _ => new SolidColorBrush(Color.FromRgb(237, 137, 54)),
        };

    private sealed record TourSelectionItem(Tour Tour, string Label);
    private sealed record FilterOption(string Label, string? Value);
    private sealed record RouteStopRow(string StopId, int Order, string OrderNumber, string Name, string Address, string TimeWindow, string PlannedArrival, string PlannedDeparture, string Service, string Wait, string Weight, string Status);
    private sealed record SegmentRow(string From, string To, int TravelMinutes, string Departure, string Arrival, int WaitMinutes, string Note)
    {
        public string TravelLabel => $"{TravelMinutes} Min.";
        public string WaitLabel => $"{WaitMinutes} Min.";
    }

    private sealed record PinRow(string StopId, int Order, string OrderNumber, string Name, string Address, string Status, string Weight, string Assignment);

    private sealed class RouteStopDraft
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public double? Latitude { get; set; }
        public double? Longitude { get; set; }
        public int Order { get; set; }
        public string OrderNumber { get; set; } = string.Empty;
        public string Weight { get; set; } = string.Empty;
        public string TimeWindowStart { get; set; } = string.Empty;
        public string TimeWindowEnd { get; set; } = string.Empty;
        public int ServiceMinutes { get; set; }
        public string PlannedArrival { get; set; } = string.Empty;
        public string PlannedDeparture { get; set; } = string.Empty;
        public bool ScheduleConflict { get; set; }
        public string ScheduleConflictText { get; set; } = string.Empty;
        public int WaitMinutes { get; set; }

        public static RouteStopDraft FromTourStop(TourStop stop)
            => new()
            {
                Id = stop.Id,
                Name = stop.Name,
                Address = stop.Address,
                Latitude = stop.Latitude,
                Longitude = stop.Longitude,
                Order = stop.Order,
                OrderNumber = stop.OrderNumber,
                Weight = stop.Weight,
                TimeWindowStart = stop.TimeWindowStart,
                TimeWindowEnd = stop.TimeWindowEnd,
                ServiceMinutes = stop.ServiceMinutes,
                PlannedArrival = stop.PlannedArrival,
                PlannedDeparture = stop.PlannedDeparture,
                ScheduleConflict = stop.ScheduleConflict,
                ScheduleConflictText = stop.ScheduleConflictText,
                WaitMinutes = stop.WaitMinutes,
            };

        public static RouteStopDraft FromPin(PinRecord pin, int order)
            => new()
            {
                Id = BuildStopIdFromPin(pin),
                Name = FirstNonEmpty(pin.Data.TryGetValue("Name", out var name) ? name : null, "Auftrag"),
                Address = BuildAddress(pin.Data),
                Latitude = pin.Latitude,
                Longitude = pin.Longitude,
                Order = order,
                OrderNumber = FirstNonEmpty(pin.Data.TryGetValue("Auftragsnummer", out var orderNumber) ? orderNumber : null, string.Empty),
                Weight = FirstNonEmpty(pin.Data.TryGetValue("Gewicht", out var weight) ? weight : null, string.Empty),
                ServiceMinutes = 0,
            };

        public RouteNode ToRouteNode()
            => new(Id, Name, Latitude, Longitude, TimeWindowStart, TimeWindowEnd, ServiceMinutes);

        public TourStop ToTourStop()
            => new()
            {
                Id = Id,
                Name = Name,
                Address = Address,
                Latitude = Latitude,
                Longitude = Longitude,
                Order = Order,
                OrderNumber = OrderNumber,
                Weight = Weight,
                TimeWindowStart = TimeWindowStart,
                TimeWindowEnd = TimeWindowEnd,
                ServiceMinutes = ServiceMinutes,
                PlannedArrival = PlannedArrival,
                PlannedDeparture = PlannedDeparture,
                ScheduleConflict = ScheduleConflict,
                ScheduleConflictText = ScheduleConflictText,
                WaitMinutes = WaitMinutes,
            };
    }
}

using Microsoft.Win32;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.App.Views;
using Tourenplaner.CSharp.Application.Models;
using Tourenplaner.CSharp.Application.Services;
using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Domain.Enums;
using Tourenplaner.CSharp.Domain.Services;
using Tourenplaner.CSharp.Infrastructure.Backups;
using Tourenplaner.CSharp.Infrastructure.Repositories;

namespace Tourenplaner.CSharp.App;

public partial class MainWindow : Window
{
    private readonly MainWindowViewModel _viewModel = new();
    private readonly AppSnapshotService _snapshotService;
    private readonly AppPathProvider _pathProvider;
    private readonly JsonEmployeesRepository _employeesRepository;
    private readonly JsonVehicleCatalogRepository _vehicleCatalogRepository;
    private readonly JsonToursRepository _toursRepository;
    private readonly JsonSettingsRepository _settingsRepository;
    private AppSnapshot? _snapshot;
    private MapPageView.RouteWorkspaceSnapshot? _currentRouteWorkspace;
    private int? _preferredMapTourId;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
        NavigationList.ItemsSource = _viewModel.Items;
        NavigationList.SelectedIndex = 0;

        var repositoryRoot = RepositoryRootLocator.Find(AppContext.BaseDirectory);
        _pathProvider = new AppPathProvider(repositoryRoot);
        var fileStore = new Tourenplaner.CSharp.Infrastructure.Json.JsonFileStore();
        _employeesRepository = new JsonEmployeesRepository(fileStore, _pathProvider.EmployeesPath);
        _vehicleCatalogRepository = new JsonVehicleCatalogRepository(fileStore, _pathProvider.VehiclesPath);
        _toursRepository = new JsonToursRepository(fileStore, _pathProvider.ToursPath);
        _settingsRepository = new JsonSettingsRepository(fileStore, _pathProvider.SettingsPath, Path.Combine(_pathProvider.ConfigDirectory, "backups"));
        _snapshotService = new AppSnapshotService(new JsonAppDataBootstrapper(_pathProvider));
        Loaded += OnLoadedAsync;
    }

    private async void OnLoadedAsync(object sender, RoutedEventArgs e)
    {
        try
        {
            _snapshot = await _snapshotService.LoadAsync();
            await MaybeRunAutoBackupAsync();
            RenderSidebarQuickAccess();
            if (NavigationList.SelectedItem is NavigationItemViewModel item)
            {
                RenderPage(item);
            }
        }
        catch (Exception exc)
        {
            ContentHost.Children.Clear();
            ContentHost.Children.Add(CreateTextBlock("Fehler beim Laden der Anwendungsdaten", 20, FontWeights.SemiBold, "TextBrush"));
            ContentHost.Children.Add(CreateTextBlock(exc.Message, 13, FontWeights.Normal, "SubTextBrush", new Thickness(0, 10, 0, 0)));
        }
    }

    private void NavigationList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavigationList.SelectedItem is not NavigationItemViewModel item)
        {
            return;
        }

        _viewModel.SelectedItem = item;
        PageTitleText.Text = item.Title;
        PageDescriptionText.Text = item.Description;
        RenderPage(item);
    }

    private void RenderPage(NavigationItemViewModel item)
    {
        ContentHost.Children.Clear();
        switch (item.Key)
        {
            case "menu":
                RenderStartPage();
                break;
            case "calendar":
                RenderCalendarPage();
                break;
            case "map":
                RenderMapPage();
                break;
            case "gps":
                RenderGpsPage();
                break;
            case "list":
                RenderOrdersPage();
                break;
            case "nonmap":
                RenderNonMapOrdersPage();
                break;
            case "tours":
                RenderToursPage();
                break;
            case "employees":
                RenderEmployeesPage();
                break;
            case "vehicles":
                RenderVehiclesPage();
                break;
            case "settings":
                RenderSettingsPage();
                break;
            case "update":
                RenderUpdatesPage();
                break;
            default:
                RenderGenericPage(item, "Bereich vorbereitet.");
                break;
        }
    }

    private void RenderStartPage()
    {
        ContentHost.Children.Add(CreateTextBlock("Startseite", 22, FontWeights.SemiBold, "TextBrush"));
        ContentHost.Children.Add(CreateTextBlock(
            "Diese erste datengetriebene Startseite spiegelt die bekannte Bereichsstruktur wider und zeigt Live-Zahlen aus den vorhandenen JSON-Dateien des Python-Projekts.",
            13,
            FontWeights.Normal,
            "SubTextBrush",
            new Thickness(0, 10, 0, 0)));

        var quickAccessPanel = BuildQuickAccessPanel(includeDescription: true);
        if (quickAccessPanel is not null)
        {
            ContentHost.Children.Add(quickAccessPanel);
        }

        var launcherGrid = new UniformGrid
        {
            Columns = 2,
            Margin = new Thickness(0, 16, 0, 0),
        };

        foreach (var navItem in _viewModel.Items.Where(x => x.Key != "menu"))
        {
            var button = new Button
            {
                Content = navItem.Title,
                Margin = new Thickness(0, 0, 12, 12),
                Padding = new Thickness(16),
                FontSize = 16,
                Background = (Brush)FindResource("PanelBrush"),
                Foreground = (Brush)FindResource("TextBrush"),
                BorderBrush = (Brush)FindResource("AccentBrush"),
            };
            button.Click += (_, _) => NavigationList.SelectedItem = navItem;
            launcherGrid.Children.Add(button);
        }

        ContentHost.Children.Add(launcherGrid);
        ContentHost.Children.Add(CreateStatPanel(BuildOverviewStats()));
    }

    private void RenderCalendarPage()
    {
        ContentHost.Children.Add(new CalendarPageView(
            _snapshot?.Tours ?? Array.Empty<Tour>(),
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush")));
    }

    private void RenderMapPage()
    {
        ContentHost.Children.Add(new MapPageView(
            new MapPageView.AppSnapshotLike(
                _snapshot?.Pins.Count ?? 0,
                _snapshot?.Tours.Count ?? 0,
                _snapshot?.Employees.Count ?? 0,
                _snapshot?.VehicleCatalog.Vehicles.Count ?? 0,
                _snapshot?.Tours ?? Array.Empty<Tour>(),
                _snapshot?.Employees ?? Array.Empty<Employee>(),
                _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty,
                _snapshot?.Pins ?? Array.Empty<PinRecord>()),
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush"),
            _currentRouteWorkspace,
            snapshot => _currentRouteWorkspace = snapshot,
            _preferredMapTourId));
        _preferredMapTourId = null;
    }

    private void RenderGpsPage()
    {
        ContentHost.Children.Add(new GpsPageView(
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush")));
    }

    private void RenderOrdersPage()
    {
        ContentHost.Children.Add(new OrdersPageView(
            _snapshot?.Pins ?? Array.Empty<PinRecord>(),
            _snapshot?.Tours ?? Array.Empty<Tour>(),
            _snapshot?.SqlImportWorkspace ?? SqlImportWorkspace.Empty,
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush")));
    }

    private void RenderNonMapOrdersPage()
    {
        ContentHost.Children.Add(new NonMapOrdersPageView(
            _snapshot?.SqlImportWorkspace ?? SqlImportWorkspace.Empty,
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush")));
    }

    private void RenderToursPage()
    {
        ContentHost.Children.Add(new ToursPageView(
            _snapshot?.Tours ?? Array.Empty<Tour>(),
            _snapshot?.Employees ?? Array.Empty<Employee>(),
            _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty,
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush"),
            async () => await AddTourAsync(),
            async () => await SaveCurrentRouteAsTourAsync(),
            tour => ShowTourOnMap(tour),
            async tour => await EditTourAsync(tour),
            async tour => await DeleteTourAsync(tour)));
    }

    private void RenderEmployeesPage()
    {
        ContentHost.Children.Add(new EmployeesPageView(
            _snapshot?.Employees ?? Array.Empty<Employee>(),
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush"),
            async () => await AddEmployeeAsync(),
            async employee => await EditEmployeeAsync(employee),
            async employee => await DeleteEmployeeAsync(employee)));
    }

    private void RenderVehiclesPage()
    {
        ContentHost.Children.Add(new VehiclesPageView(
            _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty,
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush"),
            async () => await AddVehicleAsync(),
            async vehicle => await EditVehicleAsync(vehicle),
            async vehicle => await DeleteVehicleAsync(vehicle),
            async () => await AddTrailerAsync(),
            async trailer => await EditTrailerAsync(trailer),
            async trailer => await DeleteTrailerAsync(trailer)));
    }

    private void RenderSettingsPage()
    {
        var settings = _snapshot?.Settings ?? new AppSettings();
        var latestBackup = new BackupManager("GAWELA Tourenplaner", _pathProvider.ConfigDirectory, _pathProvider.DataDirectory, settings.BackupDir).FindLatestBackup();
        ContentHost.Children.Add(new SettingsPageView(
            settings,
            BuildQuickAccessOptions(),
            Path.GetFileName(latestBackup ?? string.Empty),
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush"),
            async mode => await SaveAppearanceAsync(mode),
            async (dataDir, server, database) => await SaveSqlSettingsAsync(dataDir, server, database),
            async items => await SaveQuickAccessSettingsAsync(items),
            async (mode, enabled, autoEnabled, backupDir, retentionDays, intervalDays) => await SaveBackupSettingsAsync(mode, enabled, autoEnabled, backupDir, retentionDays, intervalDays),
            async mode => await CreateBackupAsync(mode),
            async groups => await RestoreBackupAsync(groups)));
    }

    private void RenderUpdatesPage()
    {
        ContentHost.Children.Add(new UpdatesPageView(
            ReadVersionText(),
            Path.Combine(_pathProvider.BaseDirectory, "installer-dist", "GAWELA-Tourenplaner.appinstaller"),
            (Brush)FindResource("PanelBrush"),
            (Brush)FindResource("TextBrush"),
            (Brush)FindResource("SubTextBrush")));
    }

    private void RenderGenericPage(NavigationItemViewModel item, string detail)
    {
        ContentHost.Children.Add(CreateTextBlock(item.Title, 22, FontWeights.SemiBold, "TextBrush"));
        ContentHost.Children.Add(CreateTextBlock(detail, 13, FontWeights.Normal, "SubTextBrush", new Thickness(0, 10, 0, 0)));
    }


    private string ReadVersionText()
    {
        try
        {
            var versionPath = Path.Combine(_pathProvider.BaseDirectory, "version.txt");
            return File.Exists(versionPath) ? File.ReadAllText(versionPath).Trim() : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private IEnumerable<PageStatViewModel> BuildOverviewStats()
    {
        return new[]
        {
            new PageStatViewModel("Pins", (_snapshot?.Pins.Count ?? 0).ToString(), "Kartierte/geocodierte Aufträge"),
            new PageStatViewModel("Pending SQL", (_snapshot?.SqlImportWorkspace.PendingOrders.Count ?? 0).ToString(), "Offene SQL-Aufträge ohne Koordinaten"),
            new PageStatViewModel("Nicht-Karte", (_snapshot?.SqlImportWorkspace.NonMapOrders.Count ?? 0).ToString(), "Separierte Lieferarten aus dem SQL-Import"),
            new PageStatViewModel("Touren", (_snapshot?.Tours.Count ?? 0).ToString(), "Gespeicherte Liefertouren"),
            new PageStatViewModel("Mitarbeiter", (_snapshot?.Employees.Count ?? 0).ToString(), "Mitarbeiterstammdaten"),
            new PageStatViewModel("Fahrzeuge", (_snapshot?.VehicleCatalog.Vehicles.Count ?? 0).ToString(), "Zugfahrzeuge"),
            new PageStatViewModel("Anhänger", (_snapshot?.VehicleCatalog.Trailers.Count ?? 0).ToString(), "Trailer im System"),
            new PageStatViewModel("Geocode-Cache", (_snapshot?.SqlImportWorkspace.GeocodeCacheEntries ?? 0).ToString(), "Zwischengespeicherte Geocoding-Treffer"),
        };
    }

    private UIElement CreateStatPanel(IEnumerable<PageStatViewModel> stats)
    {
        var grid = new UniformGrid
        {
            Columns = 2,
            Margin = new Thickness(0, 20, 0, 0),
        };

        foreach (var stat in stats)
        {
            var card = new Border
            {
                Margin = new Thickness(0, 0, 12, 12),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(16),
                Background = (Brush)FindResource("PanelBrush"),
                Child = new StackPanel
                {
                    Children =
                    {
                        CreateTextBlock(stat.Title, 13, FontWeights.SemiBold, "SubTextBrush"),
                        CreateTextBlock(stat.Value, 24, FontWeights.Bold, "TextBrush", new Thickness(0, 6, 0, 0)),
                        CreateTextBlock(stat.Description, 12, FontWeights.Normal, "SubTextBrush", new Thickness(0, 8, 0, 0)),
                    },
                },
            };

            grid.Children.Add(card);
        }

        return grid;
    }

    private void ShowTourOnMap(Tour tour)
    {
        _preferredMapTourId = tour.Id;
        _currentRouteWorkspace = new MapPageView.RouteWorkspaceSnapshot(
            tour.Id,
            tour.StartTime,
            tour.RouteMode,
            tour.Stops.ToList(),
            tour.EmployeeIds,
            tour.VehicleId,
            tour.TrailerId,
            new Dictionary<string, int>(tour.TravelTimeCache));
        var mapItem = _viewModel.Items.FirstOrDefault(item => item.Key == "map");
        if (mapItem is not null)
        {
            NavigationList.SelectedItem = mapItem;
        }
    }

    private async Task SaveCurrentRouteAsTourAsync()
    {
        if (_currentRouteWorkspace is null || _currentRouteWorkspace.Stops.Count == 0)
        {
            MessageBox.Show(this, "Es ist aktuell keine Route mit Stopps aus dem Kartenbereich vorhanden.", "Liefertouren", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var draft = new Tour
        {
            Date = DateTime.Today.ToString("dd-MM-yyyy"),
            Name = _currentRouteWorkspace.TourId is int tourId ? $"Kopie Tour #{tourId}" : "Neue Route",
            StartTime = _currentRouteWorkspace.StartTime,
            RouteMode = _currentRouteWorkspace.RouteMode,
            Stops = _currentRouteWorkspace.Stops,
            EmployeeIds = _currentRouteWorkspace.EmployeeIds,
            VehicleId = _currentRouteWorkspace.VehicleId,
            TrailerId = _currentRouteWorkspace.TrailerId,
            TravelTimeCache = _currentRouteWorkspace.TravelTimeCache,
        };

        var dialog = new TourEditorWindow(draft, _snapshot?.Employees ?? Array.Empty<Employee>(), _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedTour is null)
        {
            return;
        }

        var tours = (_snapshot?.Tours ?? Array.Empty<Tour>()).ToList();
        var nextId = tours.Count == 0 ? 1 : tours.Max(tour => tour.Id) + 1;
        tours.Add(dialog.EditedTour with
        {
            Id = nextId,
            Stops = _currentRouteWorkspace.Stops,
            RouteMode = _currentRouteWorkspace.RouteMode,
            TravelTimeCache = _currentRouteWorkspace.TravelTimeCache,
        });
        await SaveToursAsync(tours);
    }

    private async Task SaveToursAsync(IReadOnlyList<Tour> tours)
    {
        await _toursRepository.SaveAsync(tours);
        await ReloadSnapshotAsync();
    }

    private async Task AddTourAsync()
    {
        var dialog = new TourEditorWindow(null, _snapshot?.Employees ?? Array.Empty<Employee>(), _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedTour is null)
        {
            return;
        }

        var tours = (_snapshot?.Tours ?? Array.Empty<Tour>()).ToList();
        var nextId = tours.Count == 0 ? 1 : tours.Max(tour => tour.Id) + 1;
        tours.Add(dialog.EditedTour with { Id = nextId });
        await SaveToursAsync(tours);
    }

    private async Task EditTourAsync(Tour tour)
    {
        var dialog = new TourEditorWindow(tour, _snapshot?.Employees ?? Array.Empty<Employee>(), _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedTour is null)
        {
            return;
        }

        var tours = (_snapshot?.Tours ?? Array.Empty<Tour>()).Select(item => item.Id == tour.Id ? dialog.EditedTour with { Id = tour.Id, Stops = item.Stops, TravelTimeCache = item.TravelTimeCache } : item).ToList();
        await SaveToursAsync(tours);
    }

    private async Task DeleteTourAsync(Tour tour)
    {
        var result = MessageBox.Show(this, $"Tour {tour.Name} wirklich löschen?", "Touren", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        var tours = (_snapshot?.Tours ?? Array.Empty<Tour>()).Where(item => item.Id != tour.Id).ToList();
        await SaveToursAsync(tours);
    }

    private async Task SaveVehicleCatalogAsync(VehicleCatalog catalog)
    {
        await _vehicleCatalogRepository.SaveAsync(catalog);
        await ReloadSnapshotAsync();
    }

    private async Task AddVehicleAsync()
    {
        var dialog = new VehicleEditorWindow { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedVehicle is null)
        {
            return;
        }

        var catalog = _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty;
        await SaveVehicleCatalogAsync(new VehicleCatalog(catalog.Vehicles.Append(dialog.EditedVehicle).ToList(), catalog.Trailers));
    }

    private async Task EditVehicleAsync(Vehicle vehicle)
    {
        var dialog = new VehicleEditorWindow(vehicle) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedVehicle is null)
        {
            return;
        }

        var catalog = _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty;
        var vehicles = catalog.Vehicles.Select(item => item.Id == vehicle.Id ? dialog.EditedVehicle : item).ToList();
        await SaveVehicleCatalogAsync(new VehicleCatalog(vehicles, catalog.Trailers));
    }

    private async Task DeleteVehicleAsync(Vehicle vehicle)
    {
        var result = MessageBox.Show(this, $"{vehicle.Name} wirklich löschen?", "Fahrzeuge", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        var catalog = _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty;
        await SaveVehicleCatalogAsync(new VehicleCatalog(catalog.Vehicles.Where(item => item.Id != vehicle.Id).ToList(), catalog.Trailers));
    }

    private async Task AddTrailerAsync()
    {
        var dialog = new VehicleEditorWindow((Trailer?)null) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedTrailer is null)
        {
            return;
        }

        var catalog = _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty;
        await SaveVehicleCatalogAsync(new VehicleCatalog(catalog.Vehicles, catalog.Trailers.Append(dialog.EditedTrailer).ToList()));
    }

    private async Task EditTrailerAsync(Trailer trailer)
    {
        var dialog = new VehicleEditorWindow(trailer) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedTrailer is null)
        {
            return;
        }

        var catalog = _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty;
        var trailers = catalog.Trailers.Select(item => item.Id == trailer.Id ? dialog.EditedTrailer : item).ToList();
        await SaveVehicleCatalogAsync(new VehicleCatalog(catalog.Vehicles, trailers));
    }

    private async Task DeleteTrailerAsync(Trailer trailer)
    {
        var result = MessageBox.Show(this, $"{trailer.Name} wirklich löschen?", "Fahrzeuge", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        var catalog = _snapshot?.VehicleCatalog ?? VehicleCatalog.Empty;
        await SaveVehicleCatalogAsync(new VehicleCatalog(catalog.Vehicles, catalog.Trailers.Where(item => item.Id != trailer.Id).ToList()));
    }

    private IReadOnlyList<SettingsPageView.QuickAccessOption> BuildQuickAccessOptions()
        => new[]
        {
            new SettingsPageView.QuickAccessOption(string.Empty, "Kein Eintrag"),
            new SettingsPageView.QuickAccessOption("action:export_route", "Route exportieren"),
            new SettingsPageView.QuickAccessOption("action:import_sql", "SQL importieren"),
            new SettingsPageView.QuickAccessOption("action:select_sql_dir", "SQL-Datenordner wählen"),
            new SettingsPageView.QuickAccessOption("page:map", "Karte"),
            new SettingsPageView.QuickAccessOption("page:tours", "Liefertouren"),
            new SettingsPageView.QuickAccessOption("page:list", "Auftragsliste"),
            new SettingsPageView.QuickAccessOption("page:vehicles", "Fahrzeuge"),
            new SettingsPageView.QuickAccessOption("page:employees", "Mitarbeiter"),
            new SettingsPageView.QuickAccessOption("page:settings", "Einstellungen"),
            new SettingsPageView.QuickAccessOption("page:update", "Updates"),
        };

    private UIElement? BuildQuickAccessPanel(bool includeDescription)
    {
        var settings = _snapshot?.Settings ?? new AppSettings();
        var options = BuildQuickAccessOptions()
            .Where(option => !string.IsNullOrWhiteSpace(option.Id))
            .ToDictionary(option => option.Id, option => option.Label, StringComparer.OrdinalIgnoreCase);
        var configured = (settings.QuickAccessItems ?? Array.Empty<string>())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(4)
            .ToList();

        if (configured.Count == 0)
        {
            return null;
        }

        var shell = new StackPanel { Margin = includeDescription ? new Thickness(0, 16, 0, 0) : new Thickness(0) };
        shell.Children.Add(CreateTextBlock("Schnellzugriff", includeDescription ? 18 : 15, FontWeights.SemiBold, "TextBrush"));
        if (includeDescription)
        {
            shell.Children.Add(CreateTextBlock(
                "Die in settings.json konfigurierten Slots werden direkt in der WPF-Shell verwendet und lösen Seitenwechsel oder Aktionen ohne Umweg über die Einstellungen aus.",
                12,
                FontWeights.Normal,
                "SubTextBrush",
                new Thickness(0, 8, 0, 0)));
        }

        var panel = new WrapPanel { Margin = new Thickness(0, 12, 0, 0) };
        foreach (var item in configured)
        {
            var label = options.TryGetValue(item, out var resolvedLabel) ? resolvedLabel : item;
            panel.Children.Add(CreateQuickAccessButton(label, item, includeDescription));
        }

        shell.Children.Add(panel);
        return shell;
    }

    private Button CreateQuickAccessButton(string label, string quickAccessId, bool prominent)
    {
        var button = new Button
        {
            Content = label,
            Margin = prominent ? new Thickness(0, 0, 12, 12) : new Thickness(0, 0, 0, 8),
            Padding = prominent ? new Thickness(14, 10, 14, 10) : new Thickness(12, 8, 12, 8),
            MinWidth = prominent ? 190 : 0,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Background = prominent
                ? new SolidColorBrush(Color.FromRgb(49, 130, 206))
                : new SolidColorBrush(Color.FromRgb(37, 99, 235)),
            Foreground = Brushes.White,
            BorderBrush = Brushes.Transparent,
        };
        button.Click += async (_, _) => await ExecuteQuickAccessAsync(quickAccessId);
        return button;
    }

    private async Task ExecuteQuickAccessAsync(string quickAccessId)
    {
        if (string.IsNullOrWhiteSpace(quickAccessId))
        {
            return;
        }

        if (quickAccessId.StartsWith("page:", StringComparison.OrdinalIgnoreCase))
        {
            NavigateToPage(quickAccessId["page:".Length..]);
            return;
        }

        switch (quickAccessId)
        {
            case "action:export_route":
                ExportCurrentRouteToClipboard();
                return;
            case "action:select_sql_dir":
            {
                var sqlDir = (_snapshot?.Settings ?? new AppSettings()).SqlDataDir;
                if (string.IsNullOrWhiteSpace(sqlDir))
                {
                    MessageBox.Show(this, "Es ist aktuell kein SQL-Datenordner konfiguriert.", "Schnellzugriff", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                OpenPath(sqlDir);
                return;
            }
            case "action:import_sql":
            {
                await ReloadSnapshotAsync();
                NavigateToPage("list");
                var workspace = _snapshot?.SqlImportWorkspace ?? SqlImportWorkspace.Empty;
                var reportText = string.IsNullOrWhiteSpace(workspace.LatestGeocodeFailureReport)
                    ? "Kein Geocoding-Fehlerreport gefunden."
                    : $"Letzter Report: {Path.GetFileName(workspace.LatestGeocodeFailureReport)}";
                MessageBox.Show(
                    this,
                    "Die WPF-Shell lädt die aktuellen SQL-Arbeitsdateien neu.\n\n"
                    + $"Pending-Aufträge: {workspace.PendingOrders.Count}\n"
                    + $"Nicht-Karten-Aufträge: {workspace.NonMapOrders.Count}\n"
                    + $"Geocode-Cache: {workspace.GeocodeCacheEntries}\n"
                    + reportText,
                    "SQL-Arbeitsstand",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }
        }
    }

    private void NavigateToPage(string key)
    {
        var target = _viewModel.Items.FirstOrDefault(item => string.Equals(item.Key, key, StringComparison.OrdinalIgnoreCase));
        if (target is not null)
        {
            NavigationList.SelectedItem = target;
        }
    }

    private void ExportCurrentRouteToClipboard()
    {
        var workspace = _currentRouteWorkspace;
        if (workspace is null || workspace.Stops.Count == 0)
        {
            var currentPage = NavigationList.SelectedItem as NavigationItemViewModel;
            if (!string.Equals(currentPage?.Key, "map", StringComparison.OrdinalIgnoreCase))
            {
                NavigateToPage("map");
            }

            MessageBox.Show(this, "Es ist aktuell keine Route mit Stopps geladen. Öffne eine Tour oder arbeite zuerst im Kartenbereich.", "Routenexport", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var rows = workspace.Stops
            .OrderBy(stop => stop.Order)
            .Select(stop => $"{stop.Order}. {stop.Name} | {stop.Address} | ETA {stop.PlannedArrival} | ETD {stop.PlannedDeparture} | Fenster {stop.TimeWindowStart}-{stop.TimeWindowEnd} | Gewicht {stop.Weight}")
            .ToList();

        Clipboard.SetText(string.Join(Environment.NewLine, rows));
        MessageBox.Show(this, "Die aktuelle Route wurde aus dem Shell-Schnellzugriff in die Zwischenablage kopiert.", "Routenexport", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private async Task SaveAppearanceAsync(AppearanceMode mode)
    {
        await SaveSettingsAsync((_snapshot?.Settings ?? new AppSettings()) with { AppearanceMode = mode });
    }

    private async Task SaveSqlSettingsAsync(string dataDir, string serverInstance, string database)
    {
        var existing = _snapshot?.Settings ?? new AppSettings();
        var resolvedDatabase = string.IsNullOrWhiteSpace(database) ? SqlDatabaseNameInference.InferFromDataDirectory(dataDir) : database.Trim();
        await SaveSettingsAsync(existing with
        {
            SqlDataDir = dataDir,
            SqlServerInstance = serverInstance,
            SqlDatabase = resolvedDatabase,
        });
    }

    private async Task SaveQuickAccessSettingsAsync(IReadOnlyList<string> items)
    {
        await SaveSettingsAsync((_snapshot?.Settings ?? new AppSettings()) with { QuickAccessItems = items.ToList() });
    }

    private async Task SaveBackupSettingsAsync(BackupMode mode, bool backupsEnabled, bool autoBackupEnabled, string backupDir, int retentionDays, int intervalDays)
    {
        var resolvedBackupDir = string.IsNullOrWhiteSpace(backupDir) ? Path.Combine(_pathProvider.ConfigDirectory, "backups") : backupDir.Trim();
        Directory.CreateDirectory(resolvedBackupDir);
        await SaveSettingsAsync((_snapshot?.Settings ?? new AppSettings()) with
        {
            BackupsEnabled = backupsEnabled,
            AutoBackupEnabled = autoBackupEnabled,
            BackupDir = resolvedBackupDir,
            BackupModeDefault = mode,
            BackupRetentionDays = retentionDays,
            AutoBackupIntervalDays = intervalDays,
        });
    }

    private async Task CreateBackupAsync(BackupMode mode)
    {
        if (_snapshot is null)
        {
            return;
        }

        var settings = _snapshot.Settings;
        if (!settings.BackupsEnabled)
        {
            var proceed = MessageBox.Show(this, "Backups sind deaktiviert. Trotzdem jetzt ein Backup erstellen?", "Backups", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (proceed != MessageBoxResult.Yes)
            {
                return;
            }
        }

        var manager = new BackupManager("GAWELA Tourenplaner", _pathProvider.ConfigDirectory, _pathProvider.DataDirectory, settings.BackupDir);
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            var backupPath = await manager.CreateBackupAsync(mode);
            manager.CleanupOldBackups(settings.BackupRetentionDays);
            await SaveSettingsAsync(settings with { LastBackupIso = DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture) }, reloadPage: false);
            MessageBox.Show(this, $"Backup erfolgreich erstellt:\n{backupPath}", "Backups", MessageBoxButton.OK, MessageBoxImage.Information);
            await ReloadSnapshotAsync();
        }
        finally
        {
            Mouse.OverrideCursor = null;
        }
    }

    private async Task RestoreBackupAsync(IReadOnlyList<string> groups)
    {
        var settings = _snapshot?.Settings ?? new AppSettings();
        var dialog = new OpenFileDialog
        {
            Title = "Backup auswählen",
            Filter = "Backup-Dateien (*.bak)|*.bak|Alle Dateien (*.*)|*.*",
            InitialDirectory = Directory.Exists(settings.BackupDir) ? settings.BackupDir : _pathProvider.ConfigDirectory,
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var confirmed = MessageBox.Show(this, "Die aktuellen Daten werden mit dem ausgewählten Backup überschrieben. Backup jetzt wiederherstellen?", "Backups", MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (confirmed != MessageBoxResult.Yes)
        {
            return;
        }

        var manager = new BackupManager("GAWELA Tourenplaner", _pathProvider.ConfigDirectory, _pathProvider.DataDirectory, settings.BackupDir);
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            await manager.RestoreBackupAsync(dialog.FileName, _pathProvider.DataDirectory, _pathProvider.ConfigDirectory, groups);
            await ReloadSnapshotAsync();
            var labels = groups.Select(group => BackupManager.RestoreLabels.TryGetValue(group, out var label) ? label : group);
            MessageBox.Show(this, $"Backup erfolgreich wiederhergestellt:\n{Path.GetFileName(dialog.FileName)}\n\nWiederhergestellt: {string.Join(", ", labels)}", "Backups", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        finally
        {
            Mouse.OverrideCursor = null;
        }
    }

    private async Task SaveSettingsAsync(AppSettings settings, bool reloadPage = true)
    {
        await _settingsRepository.SaveAsync(settings);
        _snapshot = await _snapshotService.LoadAsync();
        RenderSidebarQuickAccess();
        if (reloadPage && NavigationList.SelectedItem is NavigationItemViewModel item)
        {
            RenderPage(item);
        }
    }

    private async Task MaybeRunAutoBackupAsync()
    {
        var settings = _snapshot?.Settings;
        if (settings is null || !settings.BackupsEnabled || !settings.AutoBackupEnabled)
        {
            return;
        }

        var due = true;
        if (DateTimeOffset.TryParse(settings.LastBackupIso, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var lastBackup))
        {
            due = DateTimeOffset.UtcNow - lastBackup.ToUniversalTime() >= TimeSpan.FromDays(Math.Max(1, settings.AutoBackupIntervalDays));
        }

        if (!due)
        {
            return;
        }

        var manager = new BackupManager("GAWELA Tourenplaner", _pathProvider.ConfigDirectory, _pathProvider.DataDirectory, settings.BackupDir);
        try
        {
            await manager.CreateBackupAsync(settings.BackupModeDefault);
            manager.CleanupOldBackups(settings.BackupRetentionDays);
            await _settingsRepository.SaveAsync(settings with { LastBackupIso = DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture) });
            _snapshot = await _snapshotService.LoadAsync();
        }
        catch
        {
        }
    }

    private async Task ReloadSnapshotAsync()
    {
        _snapshot = await _snapshotService.LoadAsync();
        RenderSidebarQuickAccess();
        if (NavigationList.SelectedItem is NavigationItemViewModel item)
        {
            RenderPage(item);
        }
    }

    private void RenderSidebarQuickAccess()
    {
        SidebarQuickAccessHost.Children.Clear();
        var panel = BuildQuickAccessPanel(includeDescription: false);
        if (panel is not null)
        {
            SidebarQuickAccessHost.Children.Add(panel);
        }
    }

    private async Task AddEmployeeAsync()
    {
        var dialog = new EmployeeEditorWindow { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedEmployee is null)
        {
            return;
        }

        var employees = (_snapshot?.Employees ?? Array.Empty<Employee>()).ToList();
        employees.Add(dialog.EditedEmployee);
        await _employeesRepository.SaveAsync(employees);
        await ReloadSnapshotAsync();
    }

    private async Task EditEmployeeAsync(Employee employee)
    {
        var dialog = new EmployeeEditorWindow(employee) { Owner = this };
        if (dialog.ShowDialog() != true || dialog.EditedEmployee is null)
        {
            return;
        }

        var employees = (_snapshot?.Employees ?? Array.Empty<Employee>()).ToList();
        var index = employees.FindIndex(item => item.Id == employee.Id);
        if (index >= 0)
        {
            employees[index] = dialog.EditedEmployee;
            await _employeesRepository.SaveAsync(employees);
            await ReloadSnapshotAsync();
        }
    }

    private async Task DeleteEmployeeAsync(Employee employee)
    {
        var result = MessageBox.Show(this, $"{employee.Name} wirklich löschen?", "Mitarbeiter", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        var employees = (_snapshot?.Employees ?? Array.Empty<Employee>()).Where(item => item.Id != employee.Id).ToList();
        await _employeesRepository.SaveAsync(employees);
        await ReloadSnapshotAsync();
    }

    private TextBlock CreateTextBlock(string text, double size, FontWeight weight, string brushKey, Thickness? margin = null)
    {
        return new TextBlock
        {
            Text = text,
            FontSize = size,
            FontWeight = weight,
            Foreground = (Brush)FindResource(brushKey),
            TextWrapping = TextWrapping.Wrap,
            Margin = margin ?? new Thickness(0),
        };
    }
}

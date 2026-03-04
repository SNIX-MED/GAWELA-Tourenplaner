using System.Windows;
using GAWELA.Tourenplaner2.Models;
using GAWELA.Tourenplaner2.Services;

namespace GAWELA.Tourenplaner2;

public partial class MainWindow : Window
{
    private readonly string _dataDir = Path.Combine(AppContext.BaseDirectory, "data");
    private readonly string _employeesPath;
    private readonly string _vehiclesPath;

    private List<Employee> _employees = [];
    private VehiclePayload _vehiclePayload = new();
    private List<TourStop> _sampleStops = [];

    public MainWindow()
    {
        InitializeComponent();
        _employeesPath = Path.Combine(_dataDir, "employees.json");
        _vehiclesPath = Path.Combine(_dataDir, "vehicles.json");
        Loaded += OnLoaded;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(_dataDir);

        _employees = await EmployeeStorageService.LoadEmployeesAsync(_employeesPath);
        if (_employees.Count == 0)
        {
            _employees =
            [
                new Employee { Name = "Max Mustermann", Short = "MM", Phone = "079 000 00 00", Active = true },
                new Employee { Name = "Erika Muster", Short = "EM", Phone = "079 111 11 11", Active = true },
            ];
            _employees = await EmployeeStorageService.SaveEmployeesAsync(_employeesPath, _employees);
        }

        _vehiclePayload = await VehicleStorageService.LoadVehiclesAsync(_vehiclesPath);
        if (_vehiclePayload.Vehicles.Count == 0)
        {
            _vehiclePayload.Vehicles.Add(new Vehicle { Name = "LKW 1", LicensePlate = "ZH12345", Type = "truck", MaxPayloadKg = 1200, Active = true });
            _vehiclePayload = await VehicleStorageService.SaveVehiclesAsync(_vehiclesPath, _vehiclePayload);
        }

        EmployeesGrid.ItemsSource = _employees;
        VehiclesGrid.ItemsSource = _vehiclePayload.Vehicles;

        _sampleStops =
        [
            new TourStop { Order = 1, Name = "Depot", ServiceMinutes = 10, TimeWindowStart = "08:00", TimeWindowEnd = "09:00" },
            new TourStop { Order = 2, Name = "Kunde A", ServiceMinutes = 15, TimeWindowStart = "08:30", TimeWindowEnd = "10:00" },
            new TourStop { Order = 3, Name = "Kunde B", ServiceMinutes = 20, TimeWindowStart = "09:00", TimeWindowEnd = "11:00" },
        ];
        TourStopsGrid.ItemsSource = _sampleStops;
    }

    private void OnGoStart(object sender, RoutedEventArgs e) => MainTabControl.SelectedIndex = 0;
    private void OnGoEmployees(object sender, RoutedEventArgs e) => MainTabControl.SelectedIndex = 1;
    private void OnGoVehicles(object sender, RoutedEventArgs e) => MainTabControl.SelectedIndex = 2;
    private void OnGoTours(object sender, RoutedEventArgs e) => MainTabControl.SelectedIndex = 3;

    private void OnComputeSchedule(object sender, RoutedEventArgs e)
    {
        var travelSegments = new List<int?> { 20, 18, 14, 25 };
        var result = SchedulePlanner.ComputeSchedule(_sampleStops, travelSegments, StartTimeTextBox.Text);

        _sampleStops = result.Stops;
        TourStopsGrid.ItemsSource = null;
        TourStopsGrid.ItemsSource = _sampleStops;
        ScheduleSummaryText.Text = $"Fahrt: {result.TotalTravelMinutes} min | Service: {result.TotalServiceMinutes} min | Warten: {result.TotalWaitMinutes} min | Ende: {result.EndTime}";
    }
}

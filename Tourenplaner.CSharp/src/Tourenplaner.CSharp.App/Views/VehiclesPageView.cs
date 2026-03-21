using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class VehiclesPageView : ScrollViewer
{
    public VehiclesPageView(
        VehicleCatalog catalog,
        Brush panelBrush,
        Brush textBrush,
        Brush subTextBrush,
        Action? onAddVehicle,
        Action<Vehicle>? onEditVehicle,
        Action<Vehicle>? onDeleteVehicle,
        Action? onAddTrailer,
        Action<Trailer>? onEditTrailer,
        Action<Trailer>? onDeleteTrailer)
    {
        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build(catalog, panelBrush, textBrush, subTextBrush, onAddVehicle, onEditVehicle, onDeleteVehicle, onAddTrailer, onEditTrailer, onDeleteTrailer);
    }

    private static UIElement Build(
        VehicleCatalog catalog,
        Brush panelBrush,
        Brush textBrush,
        Brush subTextBrush,
        Action? onAddVehicle,
        Action<Vehicle>? onEditVehicle,
        Action<Vehicle>? onDeleteVehicle,
        Action? onAddTrailer,
        Action<Trailer>? onEditTrailer,
        Action<Trailer>? onDeleteTrailer)
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Fahrzeugverwaltung", 22, FontWeights.SemiBold, textBrush));
        shell.Children.Add(ViewStyling.Text("Zugfahrzeuge und Anhänger werden bereits getrennt aus dem Python-Datenmodell angezeigt und können nun grundlegend bearbeitet werden.", 13, FontWeights.Normal, subTextBrush, new Thickness(0, 10, 0, 0)));
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Zugfahrzeuge", catalog.Vehicles.Count.ToString(), "Fahrzeuge im Bestand"),
            new PageStatViewModel("Anhänger", catalog.Trailers.Count.ToString(), "Trailer im Bestand"),
        }, panelBrush, textBrush, subTextBrush));

        shell.Children.Add(ViewStyling.Text("Zugfahrzeuge", 18, FontWeights.SemiBold, textBrush, new Thickness(0, 20, 0, 0)));
        var vehicleActions = new UniformGrid { Columns = 3, Margin = new Thickness(0, 16, 0, 0) };
        var addVehicleButton = new Button { Content = "Fahrzeug hinzufügen", Margin = new Thickness(0, 0, 8, 0), Padding = new Thickness(12, 8, 12, 8) };
        var editVehicleButton = new Button { Content = "Bearbeiten", Margin = new Thickness(4, 0, 4, 0), Padding = new Thickness(12, 8, 12, 8) };
        var deleteVehicleButton = new Button { Content = "Löschen", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(12, 8, 12, 8) };
        vehicleActions.Children.Add(addVehicleButton);
        vehicleActions.Children.Add(editVehicleButton);
        vehicleActions.Children.Add(deleteVehicleButton);
        shell.Children.Add(vehicleActions);

        var vehicleGrid = ViewStyling.CreateReadOnlyGrid();
        vehicleGrid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(Vehicle.Name)) });
        vehicleGrid.Columns.Add(new DataGridTextColumn { Header = "Kennzeichen", Binding = new System.Windows.Data.Binding(nameof(Vehicle.LicensePlate)) });
        vehicleGrid.Columns.Add(new DataGridTextColumn { Header = "Nutzlast", Binding = new System.Windows.Data.Binding(nameof(Vehicle.MaxPayloadKg)) });
        vehicleGrid.Columns.Add(new DataGridTextColumn { Header = "Anhängelast", Binding = new System.Windows.Data.Binding(nameof(Vehicle.MaxTrailerLoadKg)) });
        vehicleGrid.Columns.Add(new DataGridCheckBoxColumn { Header = "Aktiv", Binding = new System.Windows.Data.Binding(nameof(Vehicle.Active)) });
        vehicleGrid.ItemsSource = catalog.Vehicles;
        shell.Children.Add(vehicleGrid);

        shell.Children.Add(ViewStyling.Text("Anhänger", 18, FontWeights.SemiBold, textBrush, new Thickness(0, 20, 0, 0)));
        var trailerActions = new UniformGrid { Columns = 3, Margin = new Thickness(0, 16, 0, 0) };
        var addTrailerButton = new Button { Content = "Anhänger hinzufügen", Margin = new Thickness(0, 0, 8, 0), Padding = new Thickness(12, 8, 12, 8) };
        var editTrailerButton = new Button { Content = "Bearbeiten", Margin = new Thickness(4, 0, 4, 0), Padding = new Thickness(12, 8, 12, 8) };
        var deleteTrailerButton = new Button { Content = "Löschen", Margin = new Thickness(8, 0, 0, 0), Padding = new Thickness(12, 8, 12, 8) };
        trailerActions.Children.Add(addTrailerButton);
        trailerActions.Children.Add(editTrailerButton);
        trailerActions.Children.Add(deleteTrailerButton);
        shell.Children.Add(trailerActions);

        var trailerGrid = ViewStyling.CreateReadOnlyGrid();
        trailerGrid.Columns.Add(new DataGridTextColumn { Header = "Name", Binding = new System.Windows.Data.Binding(nameof(Trailer.Name)) });
        trailerGrid.Columns.Add(new DataGridTextColumn { Header = "Kennzeichen", Binding = new System.Windows.Data.Binding(nameof(Trailer.LicensePlate)) });
        trailerGrid.Columns.Add(new DataGridTextColumn { Header = "Nutzlast", Binding = new System.Windows.Data.Binding(nameof(Trailer.MaxPayloadKg)) });
        trailerGrid.Columns.Add(new DataGridCheckBoxColumn { Header = "Aktiv", Binding = new System.Windows.Data.Binding(nameof(Trailer.Active)) });
        trailerGrid.ItemsSource = catalog.Trailers;
        shell.Children.Add(trailerGrid);

        addVehicleButton.Click += (_, _) => onAddVehicle?.Invoke();
        editVehicleButton.Click += (_, _) => { if (vehicleGrid.SelectedItem is Vehicle vehicle) onEditVehicle?.Invoke(vehicle); };
        deleteVehicleButton.Click += (_, _) => { if (vehicleGrid.SelectedItem is Vehicle vehicle) onDeleteVehicle?.Invoke(vehicle); };
        addTrailerButton.Click += (_, _) => onAddTrailer?.Invoke();
        editTrailerButton.Click += (_, _) => { if (trailerGrid.SelectedItem is Trailer trailer) onEditTrailer?.Invoke(trailer); };
        deleteTrailerButton.Click += (_, _) => { if (trailerGrid.SelectedItem is Trailer trailer) onDeleteTrailer?.Invoke(trailer); };

        return shell;
    }
}

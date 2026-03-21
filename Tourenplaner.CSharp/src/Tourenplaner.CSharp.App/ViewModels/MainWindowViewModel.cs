using System.Collections.ObjectModel;

namespace Tourenplaner.CSharp.App.ViewModels;

public sealed class MainWindowViewModel
{
    public ObservableCollection<NavigationItemViewModel> Items { get; } =
    [
        new("menu", "Start", "Direktzugriff auf alle Bereiche"),
        new("calendar", "Kalender", "Übersicht geplanter Touren"),
        new("map", "Karte", "Kartenansicht mit Routenpanel"),
        new("gps", "GPS", "Eingebettete GPS-Webansicht"),
        new("list", "Auftragsliste", "Kartierbare Aufträge"),
        new("nonmap", "Nicht-Karten-Aufträge", "Separater Bereich nach Lieferart"),
        new("tours", "Liefertouren", "Gespeicherte Touren"),
        new("employees", "Mitarbeiter", "Mitarbeiterverwaltung"),
        new("vehicles", "Fahrzeuge", "Zugfahrzeuge und Anhänger"),
        new("settings", "Einstellungen", "Konfiguration, Backups, SQL"),
        new("update", "Updates", "Installations- und Updatestatus"),
    ];

    public NavigationItemViewModel? SelectedItem { get; set; }
}

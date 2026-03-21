using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;
using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.App.Views;

public sealed class CalendarPageView : ScrollViewer
{
    private readonly IReadOnlyList<Tour> _tours;
    private readonly Brush _panelBrush;
    private readonly Brush _textBrush;
    private readonly Brush _subTextBrush;
    private readonly Calendar _calendar = new() { Margin = new Thickness(0, 16, 0, 0) };
    private readonly StackPanel _detailHost = new() { Margin = new Thickness(0, 16, 0, 0) };
    private readonly ListBox _upcomingList = new() { Margin = new Thickness(0, 12, 0, 0), MinHeight = 140 };

    public CalendarPageView(IReadOnlyList<Tour> tours, Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        _tours = tours;
        _panelBrush = panelBrush;
        _textBrush = textBrush;
        _subTextBrush = subTextBrush;

        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build();
        RefreshSelectedDate(_calendar.SelectedDate ?? DateTime.Today);
    }

    private UIElement Build()
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Kalender", 22, FontWeights.SemiBold, _textBrush));
        shell.Children.Add(ViewStyling.Text(
            "Die Kalenderseite zeigt jetzt echte Tourdaten pro Tag, Mitarbeitereinsätze und eine Vorschau der nächsten geplanten Touren.",
            13,
            FontWeights.Normal,
            _subTextBrush,
            new Thickness(0, 10, 0, 0)));

        var todayKey = NormalizeDate(DateTime.Today);
        var todayTours = _tours.Where(tour => string.Equals(tour.Date, todayKey, StringComparison.Ordinal)).ToList();
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Touren gesamt", _tours.Count.ToString(), "Geladene Liefertouren im Kalender"),
            new PageStatViewModel("Heute", todayTours.Count.ToString(), $"Touren am {todayKey}"),
            new PageStatViewModel("Mitarbeitereinsätze", _tours.Sum(tour => tour.EmployeeIds.Count).ToString(), "Zugeordnete Einsätze über alle Touren"),
            new PageStatViewModel("Nächste 10 Tage", CountUpcomingTours(10).ToString(), "Geplante Touren in der Vorschau"),
        }, _panelBrush, _textBrush, _subTextBrush));

        var layout = new Grid { Margin = new Thickness(0, 20, 0, 0) };
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.05, GridUnitType.Star) });
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.95, GridUnitType.Star) });

        var calendarCard = new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(18),
            Margin = new Thickness(0, 0, 12, 0),
        };
        var calendarStack = new StackPanel();
        calendarStack.Children.Add(ViewStyling.Text("Monatsübersicht", 16, FontWeights.SemiBold, _textBrush));
        calendarStack.Children.Add(ViewStyling.Text("Wähle ein Datum aus, um Touren, Stopps und Einsätze dieses Tages zu sehen.", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 6, 0, 0)));
        _calendar.DisplayDate = DateTime.Today;
        _calendar.SelectedDate = DateTime.Today;
        _calendar.SelectionMode = CalendarSelectionMode.SingleDate;
        _calendar.SelectedDatesChanged += (_, _) => RefreshSelectedDate(_calendar.SelectedDate ?? DateTime.Today);
        calendarStack.Children.Add(_calendar);
        calendarCard.Child = calendarStack;
        Grid.SetColumn(calendarCard, 0);
        layout.Children.Add(calendarCard);

        var detailCard = new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(18),
            Margin = new Thickness(12, 0, 0, 0),
        };
        var detailStack = new StackPanel();
        detailStack.Children.Add(ViewStyling.Text("Tagesdetails", 16, FontWeights.SemiBold, _textBrush));
        detailStack.Children.Add(_detailHost);
        detailStack.Children.Add(ViewStyling.Text("Nächste 10 Tage", 16, FontWeights.SemiBold, _textBrush, new Thickness(0, 18, 0, 0)));
        detailStack.Children.Add(ViewStyling.Text("Chronologische Vorschau geplanter Touren ab heute.", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 6, 0, 0)));
        _upcomingList.Foreground = _textBrush;
        _upcomingList.Background = new SolidColorBrush(Color.FromArgb(24, 255, 255, 255));
        _upcomingList.BorderBrush = Brushes.Transparent;
        _upcomingList.ItemsSource = BuildUpcomingTourItems(10);
        detailStack.Children.Add(_upcomingList);
        detailCard.Child = detailStack;
        Grid.SetColumn(detailCard, 1);
        layout.Children.Add(detailCard);

        shell.Children.Add(layout);
        return shell;
    }

    private void RefreshSelectedDate(DateTime date)
    {
        _detailHost.Children.Clear();
        var key = NormalizeDate(date);
        var toursForDate = _tours.Where(tour => string.Equals(tour.Date, key, StringComparison.Ordinal)).OrderBy(tour => tour.StartTime).ToList();
        var employeeAssignments = toursForDate.Sum(tour => tour.EmployeeIds.Count);
        var stopCount = toursForDate.Sum(tour => tour.Stops.Count);

        _detailHost.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Datum", key, "Ausgewählter Kalendertag"),
            new PageStatViewModel("Touren", toursForDate.Count.ToString(), "Geplante Touren an diesem Tag"),
            new PageStatViewModel("Einsätze", employeeAssignments.ToString(), "Zugeordnete Mitarbeiter"),
            new PageStatViewModel("Stopps", stopCount.ToString(), "Stopps über alle Tagestouren"),
        }, _panelBrush, _textBrush, _subTextBrush));

        if (toursForDate.Count == 0)
        {
            _detailHost.Children.Add(ViewStyling.Text("Für dieses Datum sind aktuell keine Touren gespeichert.", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 12, 0, 0)));
            return;
        }

        foreach (var tour in toursForDate)
        {
            var card = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(22, 255, 255, 255)),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 12, 0, 0),
                Child = new StackPanel
                {
                    Children =
                    {
                        ViewStyling.Text($"{tour.StartTime} · {string.IsNullOrWhiteSpace(tour.Name) ? $"Tour #{tour.Id}" : tour.Name}", 13, FontWeights.SemiBold, _textBrush),
                        ViewStyling.Text($"Mitarbeiter: {tour.EmployeeIds.Count} · Stopps: {tour.Stops.Count} · Fahrzeug: {tour.VehicleId}", 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 6, 0, 0)),
                        ViewStyling.Text(BuildRoutePreview(tour), 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 6, 0, 0)),
                    },
                },
            };
            _detailHost.Children.Add(card);
        }
    }

    private IEnumerable<string> BuildUpcomingTourItems(int days)
    {
        var today = DateTime.Today;
        var end = today.AddDays(days);
        return _tours
            .Select(tour => new { Tour = tour, Date = ParseTourDate(tour.Date) })
            .Where(item => item.Date is not null && item.Date.Value >= today && item.Date.Value <= end)
            .OrderBy(item => item.Date)
            .ThenBy(item => item.Tour.StartTime)
            .Select(item => $"{item.Tour.Date} · {item.Tour.StartTime} · {(string.IsNullOrWhiteSpace(item.Tour.Name) ? $"Tour #{item.Tour.Id}" : item.Tour.Name)} · {item.Tour.Stops.Count} Stopps")
            .ToList();
    }

    private int CountUpcomingTours(int days) => BuildUpcomingTourItems(days).Count();

    private static string BuildRoutePreview(Tour tour)
    {
        var stops = tour.Stops.OrderBy(stop => stop.Order).Take(3).Select(stop => stop.Name).Where(name => !string.IsNullOrWhiteSpace(name)).ToList();
        if (stops.Count == 0)
        {
            return "Route: noch keine Stopps.";
        }

        return $"Route: {string.Join(" → ", stops)}{(tour.Stops.Count > 3 ? " → …" : string.Empty)}";
    }

    private static DateTime? ParseTourDate(string? value)
        => DateTime.TryParseExact(value, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed
            : null;

    private static string NormalizeDate(DateTime value) => value.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture);
}

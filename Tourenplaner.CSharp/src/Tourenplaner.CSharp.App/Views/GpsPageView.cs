using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;

namespace Tourenplaner.CSharp.App.Views;

public sealed class GpsPageView : ScrollViewer
{
    private const string GpsUrl = "https://map.ktrac.ch/";

    public GpsPageView(Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build(panelBrush, textBrush, subTextBrush);
    }

    private static UIElement Build(Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        var browser = new WebBrowser();
        try
        {
            browser.Navigate(GpsUrl);
        }
        catch
        {
            // Ignore and leave fallback text/buttons available.
        }

        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("GPS", 22, FontWeights.SemiBold, textBrush));
        shell.Children.Add(ViewStyling.Text("Die GPS-Seite bindet jetzt das bekannte Zielportal direkt ein und bietet Browser-/Copy-/Reload-Aktionen als Fallbacks.", 13, FontWeights.Normal, subTextBrush, new Thickness(0, 10, 0, 0)));
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Zielportal", "map.ktrac.ch", "Eingebettete GPS-Webansicht"),
            new PageStatViewModel("Fallback", "Browser", "Externes Öffnen bei Problemen möglich"),
        }, panelBrush, textBrush, subTextBrush));

        var actionBar = new WrapPanel { Margin = new Thickness(0, 16, 0, 0) };
        actionBar.Children.Add(CreateButton("Neu laden", () =>
        {
            try { browser.Refresh(); } catch { }
        }));
        actionBar.Children.Add(CreateButton("Im Browser öffnen", () => OpenUri(GpsUrl), new Thickness(8, 0, 0, 0)));
        actionBar.Children.Add(CreateButton("URL kopieren", () => Clipboard.SetText(GpsUrl), new Thickness(8, 0, 0, 0)));
        shell.Children.Add(actionBar);

        shell.Children.Add(new Border
        {
            Background = panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(12),
            Margin = new Thickness(0, 16, 0, 0),
            Child = browser,
            MinHeight = 560,
        });

        shell.Children.Add(ViewStyling.Text("Hinweis: Wenn Webinhalte lokal blockiert sind, kann das Portal jederzeit extern im Standardbrowser geöffnet werden.", 12, FontWeights.Normal, subTextBrush, new Thickness(0, 12, 0, 0)));
        return shell;
    }

    private static Button CreateButton(string text, Action action, Thickness? margin = null)
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
        button.Click += (_, _) => action();
        return button;
    }

    private static void OpenUri(string uri)
    {
        try
        {
            Process.Start(new ProcessStartInfo(uri) { UseShellExecute = true });
        }
        catch
        {
            Clipboard.SetText(uri);
        }
    }
}

using System.Diagnostics;
using System.Xml.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;

namespace Tourenplaner.CSharp.App.Views;

public sealed class UpdatesPageView : ScrollViewer
{
    public UpdatesPageView(string currentVersion, string appInstallerPath, Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build(currentVersion, appInstallerPath, panelBrush, textBrush, subTextBrush);
    }

    private static UIElement Build(string currentVersion, string appInstallerPath, Brush panelBrush, Brush textBrush, Brush subTextBrush)
    {
        var info = ReadAppInstallerInfo(appInstallerPath);
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Updates", 22, FontWeights.SemiBold, textBrush));
        shell.Children.Add(ViewStyling.Text("Die Update-Seite liest jetzt reale Versions- und AppInstaller-Informationen aus dem Repository und bietet direkte Copy/Open-Aktionen.", 13, FontWeights.Normal, subTextBrush, new Thickness(0, 10, 0, 0)));
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new PageStatViewModel("Version lokal", string.IsNullOrWhiteSpace(currentVersion) ? "-" : currentVersion, "Wert aus version.txt"),
            new PageStatViewModel("AppInstaller", info.AppInstallerVersion, "Version aus der .appinstaller-Datei"),
            new PageStatViewModel("Update-Checks", info.UpdateCheckMode, "Ausgelesene OnLaunch-Einstellung"),
        }, panelBrush, textBrush, subTextBrush));

        shell.Children.Add(CreateCard("AppInstaller-URL", info.AppInstallerUri, textBrush, subTextBrush, () => CopyText(info.AppInstallerUri), () => OpenUri(info.AppInstallerUri)));
        shell.Children.Add(CreateCard("MSIX-Paket", info.MainPackageUri, textBrush, subTextBrush, () => CopyText(info.MainPackageUri), () => OpenUri(info.MainPackageUri)));
        shell.Children.Add(CreateCard("Lokale Datei", appInstallerPath, textBrush, subTextBrush, () => CopyText(appInstallerPath), () => OpenUri(appInstallerPath)));
        return shell;
    }

    private static Border CreateCard(string title, string value, Brush textBrush, Brush subTextBrush, Action copyAction, Action openAction)
    {
        var actions = new WrapPanel { Margin = new Thickness(0, 10, 0, 0) };
        actions.Children.Add(CreateButton("Kopieren", copyAction));
        actions.Children.Add(CreateButton("Öffnen", openAction, new Thickness(8, 0, 0, 0)));

        return new Border
        {
            Background = new SolidColorBrush(Color.FromArgb(22, 255, 255, 255)),
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(16),
            Margin = new Thickness(0, 16, 0, 0),
            Child = new StackPanel
            {
                Children =
                {
                    ViewStyling.Text(title, 16, FontWeights.SemiBold, textBrush),
                    ViewStyling.Text(string.IsNullOrWhiteSpace(value) ? "-" : value, 12, FontWeights.Normal, subTextBrush, new Thickness(0, 8, 0, 0)),
                    actions,
                },
            },
        };
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

    private static void CopyText(string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            Clipboard.SetText(value);
        }
    }

    private static void OpenUri(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo(value) { UseShellExecute = true });
        }
        catch
        {
            CopyText(value);
        }
    }

    private static AppInstallerInfo ReadAppInstallerInfo(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return AppInstallerInfo.Empty;
            }

            var document = XDocument.Load(path);
            XNamespace ns = document.Root?.GetDefaultNamespace() ?? XNamespace.None;
            var root = document.Root;
            var mainPackage = root?.Element(ns + "MainPackage");
            var onLaunch = root?.Element(ns + "UpdateSettings")?.Element(ns + "OnLaunch");
            return new AppInstallerInfo(
                root?.Attribute("Uri")?.Value ?? string.Empty,
                root?.Attribute("Version")?.Value ?? "-",
                mainPackage?.Attribute("Uri")?.Value ?? string.Empty,
                onLaunch?.Attribute("HoursBetweenUpdateChecks")?.Value is string hours ? $"alle {hours}h" : "-"
            );
        }
        catch
        {
            return AppInstallerInfo.Empty;
        }
    }

    private sealed record AppInstallerInfo(string AppInstallerUri, string AppInstallerVersion, string MainPackageUri, string UpdateCheckMode)
    {
        public static AppInstallerInfo Empty { get; } = new(string.Empty, "-", string.Empty, "-");
    }
}

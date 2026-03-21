using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.App.ViewModels;

namespace Tourenplaner.CSharp.App.Views;

internal static class ViewStyling
{
    public static TextBlock Text(string text, double size, FontWeight weight, Brush foreground, Thickness? margin = null)
        => new()
        {
            Text = text,
            FontSize = size,
            FontWeight = weight,
            Foreground = foreground,
            TextWrapping = TextWrapping.Wrap,
            Margin = margin ?? new Thickness(0),
        };

    public static Border StatCard(PageStatViewModel stat, Brush background, Brush textBrush, Brush subTextBrush)
        => new()
        {
            Margin = new Thickness(0, 0, 12, 12),
            CornerRadius = new CornerRadius(12),
            Padding = new Thickness(16),
            Background = background,
            Child = new StackPanel
            {
                Children =
                {
                    Text(stat.Title, 13, FontWeights.SemiBold, subTextBrush),
                    Text(stat.Value, 24, FontWeights.Bold, textBrush, new Thickness(0, 6, 0, 0)),
                    Text(stat.Description, 12, FontWeights.Normal, subTextBrush, new Thickness(0, 8, 0, 0)),
                },
            },
        };

    public static UniformGrid Stats(IEnumerable<PageStatViewModel> stats, Brush background, Brush textBrush, Brush subTextBrush)
    {
        var grid = new UniformGrid
        {
            Columns = 2,
            Margin = new Thickness(0, 20, 0, 0),
        };

        foreach (var stat in stats)
        {
            grid.Children.Add(StatCard(stat, background, textBrush, subTextBrush));
        }

        return grid;
    }

    public static DataGrid CreateReadOnlyGrid()
        => new()
        {
            AutoGenerateColumns = false,
            IsReadOnly = true,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            Margin = new Thickness(0, 16, 0, 0),
            HeadersVisibility = DataGridHeadersVisibility.Column,
            GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
            RowHeaderWidth = 0,
            MinHeight = 220,
        };
}

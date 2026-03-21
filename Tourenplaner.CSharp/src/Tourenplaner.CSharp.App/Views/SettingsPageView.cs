using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Domain.Enums;
using Tourenplaner.CSharp.Domain.Services;

namespace Tourenplaner.CSharp.App.Views;

public sealed class SettingsPageView : ScrollViewer
{
    private readonly Brush _panelBrush;
    private readonly Brush _textBrush;
    private readonly Brush _subTextBrush;
    private readonly Func<AppearanceMode, Task> _saveAppearanceAsync;
    private readonly Func<string, string, string, Task> _saveSqlSettingsAsync;
    private readonly Func<IReadOnlyList<string>, Task> _saveQuickAccessAsync;
    private readonly Func<BackupMode, bool, bool, string, int, int, Task> _saveBackupSettingsAsync;
    private readonly Func<BackupMode, Task> _createBackupAsync;
    private readonly Func<IReadOnlyList<string>, Task> _restoreBackupAsync;

    private readonly TextBox _sqlDataDirBox = new() { Padding = new Thickness(8, 6, 8, 6), MinWidth = 260 };
    private readonly TextBox _sqlServerBox = new() { Padding = new Thickness(8, 6, 8, 6), MinWidth = 220 };
    private readonly TextBox _sqlDatabaseBox = new() { Padding = new Thickness(8, 6, 8, 6), MinWidth = 220 };
    private readonly ComboBox[] _quickAccessBoxes = Enumerable.Range(0, 4).Select(_ => new ComboBox { MinWidth = 220, Margin = new Thickness(0, 6, 0, 0) }).ToArray();
    private readonly CheckBox _backupsEnabledCheck = new() { Content = "Backups aktivieren", Margin = new Thickness(0, 12, 0, 0) };
    private readonly CheckBox _autoBackupsEnabledCheck = new() { Content = "Automatische Backups aktivieren", Margin = new Thickness(0, 8, 0, 0) };
    private readonly TextBox _backupDirBox = new() { Padding = new Thickness(8, 6, 8, 6), MinWidth = 260 };
    private readonly ComboBox _backupModeBox = new() { MinWidth = 180 };
    private readonly TextBox _retentionBox = new() { Padding = new Thickness(8, 6, 8, 6), MinWidth = 120 };
    private readonly TextBox _autoIntervalBox = new() { Padding = new Thickness(8, 6, 8, 6), MinWidth = 120 };
    private readonly TextBlock _backupStatusText = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _lastBackupText = new() { TextWrapping = TextWrapping.Wrap };

    public SettingsPageView(
        AppSettings settings,
        IReadOnlyList<QuickAccessOption> quickAccessOptions,
        string latestBackupName,
        Brush panelBrush,
        Brush textBrush,
        Brush subTextBrush,
        Func<AppearanceMode, Task> saveAppearanceAsync,
        Func<string, string, string, Task> saveSqlSettingsAsync,
        Func<IReadOnlyList<string>, Task> saveQuickAccessAsync,
        Func<BackupMode, bool, bool, string, int, int, Task> saveBackupSettingsAsync,
        Func<BackupMode, Task> createBackupAsync,
        Func<IReadOnlyList<string>, Task> restoreBackupAsync)
    {
        _panelBrush = panelBrush;
        _textBrush = textBrush;
        _subTextBrush = subTextBrush;
        _saveAppearanceAsync = saveAppearanceAsync;
        _saveSqlSettingsAsync = saveSqlSettingsAsync;
        _saveQuickAccessAsync = saveQuickAccessAsync;
        _saveBackupSettingsAsync = saveBackupSettingsAsync;
        _createBackupAsync = createBackupAsync;
        _restoreBackupAsync = restoreBackupAsync;

        VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
        Content = Build(settings, quickAccessOptions, latestBackupName);
    }

    public sealed record QuickAccessOption(string Id, string Label);

    private UIElement Build(AppSettings settings, IReadOnlyList<QuickAccessOption> quickAccessOptions, string latestBackupName)
    {
        var shell = new StackPanel();
        shell.Children.Add(ViewStyling.Text("Einstellungen", 22, FontWeights.SemiBold, _textBrush));
        shell.Children.Add(ViewStyling.Text("Die WPF-Settings-Seite bildet jetzt Appearance, Schnellzugriffe, SQL-Konfiguration sowie Backup/Restore inklusive gruppierter Wiederherstellung direkt auf Basis der produktiven JSON-Daten nach.", 13, FontWeights.Normal, _subTextBrush, new Thickness(0, 10, 0, 0)));
        shell.Children.Add(ViewStyling.Stats(new[]
        {
            new ViewModels.PageStatViewModel("Theme", settings.AppearanceMode.ToString(), "Persistierte Darstellungspräferenz"),
            new ViewModels.PageStatViewModel("SQL-Datenbank", string.IsNullOrWhiteSpace(settings.SqlDatabase) ? "-" : settings.SqlDatabase, "Aktuell konfigurierter DB-Name"),
            new ViewModels.PageStatViewModel("Backupmodus", settings.BackupModeDefault.ToString(), "Standardmodus für Sicherungen"),
            new ViewModels.PageStatViewModel("Letztes Backup", string.IsNullOrWhiteSpace(latestBackupName) ? "-" : latestBackupName, "Neueste gefundene .bak-Datei"),
        }, _panelBrush, _textBrush, _subTextBrush));

        shell.Children.Add(BuildAppearanceCard(settings));
        shell.Children.Add(BuildSqlCard(settings));
        shell.Children.Add(BuildQuickAccessCard(settings, quickAccessOptions));
        shell.Children.Add(BuildBackupCard(settings, latestBackupName));
        return shell;
    }

    private UIElement BuildAppearanceCard(AppSettings settings)
    {
        var card = CreateCard("Darstellung", "System, Hell und Dunkel werden jetzt direkt in settings.json persistiert.");
        var row = new UniformGrid { Columns = 3, Margin = new Thickness(0, 14, 0, 0) };
        row.Children.Add(CreateAppearanceButton("System", AppearanceMode.System, settings.AppearanceMode));
        row.Children.Add(CreateAppearanceButton("Hell", AppearanceMode.Light, settings.AppearanceMode, new Thickness(8, 0, 8, 0)));
        row.Children.Add(CreateAppearanceButton("Dunkel", AppearanceMode.Dark, settings.AppearanceMode));
        ((StackPanel)card.Child).Children.Add(row);
        return card;
    }

    private UIElement BuildSqlCard(AppSettings settings)
    {
        var card = CreateCard("SQL-Konfiguration", "SQL-Datenordner, Serverinstanz und Datenbankname können bearbeitet und gespeichert werden. Der Datenbankname lässt sich wie im Python-System aus *.mdf-Dateien ableiten.");
        _sqlDataDirBox.Text = settings.SqlDataDir;
        _sqlServerBox.Text = settings.SqlServerInstance;
        _sqlDatabaseBox.Text = settings.SqlDatabase;

        var grid = new Grid { Margin = new Thickness(0, 14, 0, 0) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        for (var i = 0; i < 4; i++)
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }

        AddFormRow(grid, 0, "SQL-Datenordner", _sqlDataDirBox);
        AddFormRow(grid, 1, "SQL-Serverinstanz", _sqlServerBox);
        AddFormRow(grid, 2, "SQL-Datenbank", _sqlDatabaseBox);

        var actions = new WrapPanel { Margin = new Thickness(0, 14, 0, 0) };
        actions.Children.Add(CreateButton("Datenbank ableiten", async (_, _) => await RunSafeAsync(async () =>
        {
            _sqlDatabaseBox.Text = SqlDatabaseNameInference.InferFromDataDirectory(_sqlDataDirBox.Text);
            await SaveSqlAsync();
        })));
        actions.Children.Add(CreateButton("SQL speichern", async (_, _) => await RunSafeAsync(SaveSqlAsync), new Thickness(8, 0, 0, 0)));
        actions.Children.Add(CreateButton("Ordner öffnen", (_, _) => OpenPath(_sqlDataDirBox.Text), new Thickness(8, 0, 0, 0)));
        Grid.SetRow(actions, 3);
        Grid.SetColumn(actions, 1);
        grid.Children.Add(actions);

        ((StackPanel)card.Child).Children.Add(grid);
        return card;
    }

    private UIElement BuildQuickAccessCard(AppSettings settings, IReadOnlyList<QuickAccessOption> quickAccessOptions)
    {
        var card = CreateCard("Schnellzugriffe", "Vier Schnellzugriffs-Slots können wie im Python-System persistiert werden. Doppelte Einträge werden serverseitig bereinigt.");
        var grid = new UniformGrid { Columns = 2, Margin = new Thickness(0, 14, 0, 0) };
        var labelById = quickAccessOptions.ToDictionary(option => option.Id, option => option.Label, StringComparer.OrdinalIgnoreCase);
        var options = quickAccessOptions.Select(option => option.Label).ToList();

        for (var index = 0; index < _quickAccessBoxes.Length; index++)
        {
            var host = new StackPanel { Margin = new Thickness(index % 2 == 0 ? 0 : 8, 0, index % 2 == 0 ? 8 : 0, 12) };
            host.Children.Add(ViewStyling.Text($"Schnellzugriff {index + 1}", 12, FontWeights.SemiBold, _textBrush));
            foreach (var option in options)
            {
                _quickAccessBoxes[index].Items.Add(option);
            }

            var currentId = settings.QuickAccessItems.ElementAtOrDefault(index) ?? string.Empty;
            _quickAccessBoxes[index].SelectedItem = labelById.TryGetValue(currentId, out var label) ? label : options.FirstOrDefault();
            host.Children.Add(_quickAccessBoxes[index]);
            grid.Children.Add(host);
        }

        ((StackPanel)card.Child).Children.Add(grid);
        ((StackPanel)card.Child).Children.Add(CreateButton("Schnellzugriffe speichern", async (_, _) => await RunSafeAsync(async () =>
        {
            var selectedIds = _quickAccessBoxes
                .Select(box => quickAccessOptions.FirstOrDefault(option => string.Equals(option.Label, box.SelectedItem as string, StringComparison.Ordinal))?.Id ?? string.Empty)
                .ToList();
            await _saveQuickAccessAsync(selectedIds);
        }), new Thickness(0, 2, 0, 0)));
        return card;
    }

    private UIElement BuildBackupCard(AppSettings settings, string latestBackupName)
    {
        var card = CreateCard("Backups & Restore", "ZIP-basierte .bak-Backups unterstützen Voll-/Inkrementell-Modus, automatische Sicherungen, Retention und gruppierte Wiederherstellung der fachlich relevanten Datengruppen.");
        _backupsEnabledCheck.Foreground = _textBrush;
        _backupsEnabledCheck.IsChecked = settings.BackupsEnabled;
        _autoBackupsEnabledCheck.Foreground = _textBrush;
        _autoBackupsEnabledCheck.IsChecked = settings.AutoBackupEnabled;
        _backupDirBox.Text = settings.BackupDir;
        _backupModeBox.Items.Add("Vollbackup");
        _backupModeBox.Items.Add("Teilbackup");
        _backupModeBox.SelectedIndex = settings.BackupModeDefault == BackupMode.Incremental ? 1 : 0;
        _retentionBox.Text = settings.BackupRetentionDays.ToString(CultureInfo.InvariantCulture);
        _autoIntervalBox.Text = settings.AutoBackupIntervalDays.ToString(CultureInfo.InvariantCulture);
        _lastBackupText.Foreground = _subTextBrush;
        _lastBackupText.Text = $"Letztes Backup laut Settings: {FormatLastBackup(settings.LastBackupIso)}";
        _backupStatusText.Foreground = _subTextBrush;
        _backupStatusText.Text = string.IsNullOrWhiteSpace(latestBackupName) ? "Kein Backup gefunden." : $"Neueste Backup-Datei: {latestBackupName}";

        var host = (StackPanel)card.Child;
        host.Children.Add(_backupsEnabledCheck);
        host.Children.Add(_autoBackupsEnabledCheck);

        var form = new Grid { Margin = new Thickness(0, 14, 0, 0) };
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
        form.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        for (var i = 0; i < 4; i++)
        {
            form.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }

        AddFormRow(form, 0, "Backup-Ordner", _backupDirBox);
        AddFormRow(form, 1, "Standardmodus", _backupModeBox);
        AddFormRow(form, 2, "Aufbewahrungstage", _retentionBox);
        AddFormRow(form, 3, "Auto-Backup-Intervall", _autoIntervalBox);
        host.Children.Add(form);
        host.Children.Add(_lastBackupText);
        host.Children.Add(_backupStatusText);

        var actions = new WrapPanel { Margin = new Thickness(0, 14, 0, 0) };
        actions.Children.Add(CreateButton("Backup-Einstellungen speichern", async (_, _) => await RunSafeAsync(SaveBackupSettingsAsync)));
        actions.Children.Add(CreateButton("Backup jetzt ausführen", async (_, _) => await RunSafeAsync(CreateBackupNowAsync), new Thickness(8, 0, 0, 0)));
        actions.Children.Add(CreateButton("Backup wiederherstellen", async (_, _) => await RunSafeAsync(RestoreBackupAsync), new Thickness(8, 0, 0, 0)));
        actions.Children.Add(CreateButton("Ordner öffnen", (_, _) => OpenPath(_backupDirBox.Text), new Thickness(8, 0, 0, 0)));
        host.Children.Add(actions);
        return card;
    }

    private Border CreateCard(string title, string description)
    {
        return new Border
        {
            Background = _panelBrush,
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(16),
            Margin = new Thickness(0, 16, 0, 0),
            Child = new StackPanel
            {
                Children =
                {
                    ViewStyling.Text(title, 16, FontWeights.SemiBold, _textBrush),
                    ViewStyling.Text(description, 12, FontWeights.Normal, _subTextBrush, new Thickness(0, 8, 0, 0)),
                },
            },
        };
    }

    private Button CreateAppearanceButton(string text, AppearanceMode mode, AppearanceMode selectedMode, Thickness? margin = null)
    {
        var isSelected = mode == selectedMode;
        return CreateButton(text, async (_, _) => await RunSafeAsync(() => _saveAppearanceAsync(mode)), margin, isSelected);
    }

    private Button CreateButton(string text, RoutedEventHandler handler, Thickness? margin = null, bool selected = false)
    {
        var button = new Button
        {
            Content = text,
            Margin = margin ?? new Thickness(0),
            Padding = new Thickness(12, 8, 12, 8),
            Background = selected ? new SolidColorBrush(Color.FromRgb(37, 99, 235)) : new SolidColorBrush(Color.FromRgb(49, 130, 206)),
            Foreground = Brushes.White,
            BorderBrush = Brushes.Transparent,
        };
        button.Click += handler;
        return button;
    }

    private void AddFormRow(Grid grid, int row, string label, Control control)
    {
        var labelBlock = ViewStyling.Text(label, 12, FontWeights.SemiBold, _textBrush, new Thickness(0, row == 0 ? 0 : 10, 12, 0));
        Grid.SetRow(labelBlock, row);
        Grid.SetColumn(labelBlock, 0);
        grid.Children.Add(labelBlock);
        control.Margin = new Thickness(0, row == 0 ? 0 : 10, 0, 0);
        Grid.SetRow(control, row);
        Grid.SetColumn(control, 1);
        grid.Children.Add(control);
    }

    private async Task SaveSqlAsync()
    {
        await _saveSqlSettingsAsync(_sqlDataDirBox.Text.Trim(), _sqlServerBox.Text.Trim(), _sqlDatabaseBox.Text.Trim());
    }

    private async Task SaveBackupSettingsAsync()
    {
        await _saveBackupSettingsAsync(
            GetBackupMode(),
            _backupsEnabledCheck.IsChecked ?? false,
            _autoBackupsEnabledCheck.IsChecked ?? false,
            _backupDirBox.Text.Trim(),
            ParsePositive(_retentionBox.Text, "Aufbewahrungstage"),
            ParsePositive(_autoIntervalBox.Text, "Auto-Backup-Intervall"));
    }

    private async Task CreateBackupNowAsync()
    {
        await SaveBackupSettingsAsync();
        await _createBackupAsync(GetBackupMode());
    }

    private async Task RestoreBackupAsync()
    {
        var selection = new RestoreSelectionWindow { Owner = Window.GetWindow(this) };
        if (selection.ShowDialog() != true || selection.SelectedGroups is null)
        {
            return;
        }

        await SaveBackupSettingsAsync();
        await _restoreBackupAsync(selection.SelectedGroups);
    }

    private BackupMode GetBackupMode()
        => _backupModeBox.SelectedIndex == 1 ? BackupMode.Incremental : BackupMode.Full;

    private static int ParsePositive(string value, string field)
    {
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) || parsed < 1 || parsed > 365)
        {
            throw new InvalidOperationException($"{field} muss zwischen 1 und 365 liegen.");
        }

        return parsed;
    }

    private static string FormatLastBackup(string? isoValue)
    {
        if (string.IsNullOrWhiteSpace(isoValue))
        {
            return "Noch kein Backup gespeichert";
        }

        return DateTimeOffset.TryParse(isoValue, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed)
            ? parsed.ToLocalTime().ToString("dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture)
            : isoValue.Trim();
    }

    private async Task RunSafeAsync(Func<Task> action)
    {
        try
        {
            await action();
        }
        catch (Exception exc)
        {
            var owner = Window.GetWindow(this);
            if (owner is not null)
            {
                MessageBox.Show(owner, exc.Message, "Einstellungen", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                MessageBox.Show(exc.Message, "Einstellungen", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }

    private static void OpenPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
        {
        }
    }
}

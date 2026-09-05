using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Microsoft.Win32;

namespace KelsoftReportExport;

/// <summary>A financial year with its tick state in the list.</summary>
public sealed class YearSelection(FinancialYear year, bool isSelected) : INotifyPropertyChanged
{
    private bool _isSelected = isSelected;

    public FinancialYear Year { get; } = year;

    public string Label => $"FY{Year.Label}";

    public string Period => $"{Year.Start:dd/MM/yyyy}  –  {Year.End:dd/MM/yyyy}";

    public string EntryText => Year.EntryCount > 0 ? $"{Year.EntryCount:N0} entries" : "no entries";

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = value;
            OnPropertyChanged();
            SelectionChanged?.Invoke();
        }
    }

    public static event Action? SelectionChanged;

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public partial class MainWindow : Window
{
    // Below either of these the roomy spacing no longer fits, so the layout tightens:
    // the header loses its strapline, the report cards lose their descriptions, and every
    // gap closes up. Both sit just above the natural size of the roomy layout, so the
    // switch happens before anything would have to scroll. It is display scaling that
    // makes this necessary — at 150% a 1080p desktop is only about 690 units tall.
    private const double CompactBelowWidth = 880;
    private const double CompactBelowHeight = 860;

    private readonly ObservableCollection<YearSelection> _years = [];
    private string _clientName = "";

    public MainWindow()
    {
        InitializeComponent();
        YearsList.ItemsSource = _years;
        YearSelection.SelectionChanged += UpdateExportState;
        Closed += (_, _) => YearSelection.SelectionChanged -= UpdateExportState;
        StateChanged += (_, _) => UpdateMaximiseGlyph();

        // Seeded from the requested size so the first frame is already right, then kept in
        // step as the window is resized — including the shrink WindowSizing may apply
        // before the window is ever shown.
        IsCompact = Width < CompactBelowWidth || Height < CompactBelowHeight;
        SizeChanged += (_, _) =>
            IsCompact = ActualWidth < CompactBelowWidth || ActualHeight < CompactBelowHeight;

        BodyGrid.LayoutUpdated += (_, _) => UpdateBodyMinimum();

        WindowSizing.Apply(this);
        UpdateExportState();

        // Allow a data file to be passed on the command line, or dropped on the .exe.
        var startupFile = Environment.GetCommandLineArgs().Skip(1).FirstOrDefault();
        if (startupFile is not null && System.IO.File.Exists(startupFile))
            Loaded += (_, _) => LoadDataFile(startupFile);
    }

    /// <summary>True while the window is too small for the roomy spacing; the styles read it.</summary>
    public bool IsCompact
    {
        get => (bool)GetValue(IsCompactProperty);
        private set => SetValue(IsCompactProperty, value);
    }

    public static readonly DependencyProperty IsCompactProperty =
        DependencyProperty.Register(nameof(IsCompact), typeof(bool), typeof(MainWindow),
            new PropertyMetadata(false));

    /// <summary>
    /// Holds the body to the shortest height that shows every step: the four steps as they
    /// currently stand, plus the year list's own minimum. Below that the grid overflows the
    /// card and the scrollbar takes over. Measuring the steps rather than naming a number
    /// keeps this true as they change size — the client chip appearing, the compact spacing
    /// coming in, a strapline wrapping to a second line.
    /// </summary>
    private void UpdateBodyMinimum()
    {
        // Row heights, not the grid's, because a grid squeezed below its content still
        // reports the height it was given rather than the height it needs.
        var steps = BodyGrid.RowDefinitions.Sum(row => row.ActualHeight) - YearsRow.ActualHeight;
        var minimum = steps + YearsShell.MinHeight;

        if (Math.Abs(minimum - BodyGrid.MinHeight) > 0.5)
            BodyGrid.MinHeight = minimum;
    }

    private void BrowseData_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Select a Kelsoft data file",
            Filter = "Kelsoft data files (*.mdb;*.accdb)|*.mdb;*.accdb|All files (*.*)|*.*",
            CheckFileExists = true,
        };

        if (dialog.ShowDialog(this) == true)
            LoadDataFile(dialog.FileName);
    }

    private void LoadDataFile(string path)
    {
        _years.Clear();
        ClientLabel.Text = "";
        ClientChipPanel.Visibility = Visibility.Collapsed;

        List<FinancialYear> years;
        try
        {
            using var file = new KelsoftDataFile(path);

            var missing = file.MissingTables();
            if (missing.Count > 0)
            {
                ShowWarning("That file does not look like a Kelsoft data file.\n\n" +
                            $"Missing table(s): {string.Join(", ", missing)}");
                return;
            }

            _clientName = file.ClientName();
            years = [.. file.FinancialYears()];
        }
        catch (Exception ex)
        {
            ShowWarning($"Could not read the data file.\n\n{ex.Message}");
            return;
        }

        DataFileBox.Text = path;
        ClientLabel.Text = _clientName;
        ClientChipPanel.Visibility = Visibility.Visible;
        RiseIn(ClientChipPanel);

        foreach (var year in years)
            _years.Add(new YearSelection(year, year.EntryCount > 0));

        EmptyYears.Visibility = _years.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        if (_years.Count == 0)
            EmptyYears.Text = "This file defines no financial years.";

        OutputBox.Text = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(path) ?? "",
            SuggestFileName(path));

        SetStatus("Ready to export.");
        UpdateExportState();
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "Save the monthly report workbook",
            Filter = "Excel workbook (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            AddExtension = true,
            FileName = System.IO.Path.GetFileName(OutputBox.Text),
            InitialDirectory = System.IO.Path.GetDirectoryName(OutputBox.Text) ?? "",
            OverwritePrompt = true,
        };

        if (dialog.ShowDialog(this) == true)
            OutputBox.Text = dialog.FileName;
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e) => SetAll(true);

    private void SelectNone_Click(object sender, RoutedEventArgs e) => SetAll(false);

    private void SetAll(bool value)
    {
        foreach (var year in _years) year.IsSelected = value;
        UpdateExportState();
    }

    private void ReportChoice_Changed(object sender, RoutedEventArgs e) => UpdateExportState();

    private List<YearSelection> SelectedYears() => [.. _years.Where(y => y.IsSelected)];

    private bool WantProfitAndLoss => ProfitAndLossCheck.IsChecked == true;

    private bool WantBalanceSheet => BalanceSheetCheck.IsChecked == true;

    private bool WantGeneralLedger => GeneralLedgerCheck.IsChecked == true;

    private void UpdateExportState()
    {
        if (ExportButton is null) return;

        var selected = _years.Count(y => y.IsSelected);

        ExportButton.IsEnabled =
            DataFileBox.Text.Length > 0 &&
            selected > 0 &&
            (WantProfitAndLoss || WantBalanceSheet || WantGeneralLedger);

        YearSummary.Text = _years.Count == 0
            ? ""
            : $"{selected} of {_years.Count} selected";
    }

    private string SuggestFileName(string dataFilePath)
    {
        var stem = System.IO.Path.GetFileNameWithoutExtension(dataFilePath);
        var labels = _years.Where(y => y.IsSelected).Select(y => y.Year.Label).OrderBy(l => l).ToList();
        var span = labels.Count switch
        {
            0 => "",
            1 => $" FY{labels[0]}",
            _ => $" FY{labels[0]}-FY{labels[^1]}",
        };
        return $"{stem} - Monthly Reports{span}.xlsx";
    }

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        var years = SelectedYears().Select(y => y.Year).OrderBy(y => y.Label).ToList();
        var dataFile = DataFileBox.Text;
        var outputPath = OutputBox.Text.Trim();
        var wantProfitAndLoss = WantProfitAndLoss;
        var wantBalanceSheet = WantBalanceSheet;
        var wantGeneralLedger = WantGeneralLedger;

        if (years.Count == 0) return;

        if (outputPath.Length == 0)
        {
            ShowWarning("Choose where to save the Excel file.");
            return;
        }

        if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            outputPath += ".xlsx";

        SetBusy(true);

        var unbalanced = new List<string>();
        var succeeded = false;

        try
        {
            var clientName = _clientName;

            await Task.Run(() =>
            {
                using var file = new KelsoftDataFile(dataFile);

                var exports = new List<YearExport>();
                foreach (var year in years)
                {
                    Report($"Building FY{year.Label}…");

                    var profitAndLoss = wantProfitAndLoss
                        ? StatementBuilder.Build(file, year, clientName)
                        : null;

                    BalanceSheet? balanceSheet = null;
                    if (wantBalanceSheet)
                    {
                        balanceSheet = BalanceSheetBuilder.Build(file, year, clientName);
                        if (!balanceSheet.Balances(out var difference, out var month))
                            unbalanced.Add($"FY{year.Label} — worst at {month:MMMM yyyy}, out by {difference:N2}");
                    }

                    GeneralLedger? generalLedger = null;
                    if (wantGeneralLedger)
                    {
                        generalLedger = GeneralLedgerBuilder.Build(file, year, clientName);

                        // The ledger and the balance sheet fail together on the same data,
                        // so only report it once.
                        if (!wantBalanceSheet && !generalLedger.Balances(out var difference, out var month))
                            unbalanced.Add($"FY{year.Label} — worst at {month:MMMM yyyy}, out by {difference:N2}");
                    }

                    exports.Add(new YearExport(year, profitAndLoss, balanceSheet, generalLedger));
                }

                Report("Writing workbook…");
                ExcelExporter.Export(exports, outputPath);
            });

            var sheets = years.Count * ((wantProfitAndLoss ? 1 : 0) + (wantBalanceSheet ? 1 : 0) +
                                        (wantGeneralLedger ? 1 : 0));
            SetStatus($"Wrote {sheets} worksheet(s).");
            succeeded = true;
        }
        catch (Exception ex)
        {
            SetStatus("Export failed.");
            ShowWarning($"The export did not finish.\n\n{ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }

        if (!succeeded) return;

        // Prompt only once the window is out of its working state, so the progress bar
        // isn't still running behind a modal dialog.
        if (unbalanced.Count > 0)
            ShowWarning("The balance sheet does not balance in every month for:\n\n" +
                        string.Join("\n", unbalanced) +
                        "\n\nThe workbook was still written. This points at the underlying data — " +
                        "usually an unbalanced journal, or bank entries whose allocation dates " +
                        "straddle a month end — rather than at the export.");

        if (OpenWhenDone.IsChecked == true)
            Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true });
        else
            MessageBox.Show(this, $"Saved to:\n{outputPath}", "Export complete",
                MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void Report(string message) => Dispatcher.Invoke(() => SetStatus(message));

    private void SetStatus(string message) => StatusLabel.Text = message;

    private void SetBusy(bool busy)
    {
        if (busy)
        {
            Progress.Opacity = 0;
            Progress.Visibility = Visibility.Visible;
            FadeTo(Progress, 1);
        }
        else
        {
            var fade = new DoubleAnimation(0, TimeSpan.FromMilliseconds(180));
            fade.Completed += (_, _) => Progress.Visibility = Visibility.Collapsed;
            Progress.BeginAnimation(OpacityProperty, fade);
        }

        Cursor = busy ? System.Windows.Input.Cursors.Wait : null;

        ExportButton.IsEnabled = !busy;
        BrowseDataButton.IsEnabled = !busy;
        BrowseOutputButton.IsEnabled = !busy;
        YearsList.IsEnabled = !busy;

        if (!busy) UpdateExportState();
    }

    // ------------------------------------------------------- animation

    /// <summary>Fades an element in while it settles upward into place.</summary>
    private static void RiseIn(FrameworkElement element, double fromY = 8, int milliseconds = 260)
    {
        var shift = new TranslateTransform(0, fromY);
        element.RenderTransform = shift;

        var duration = TimeSpan.FromMilliseconds(milliseconds);

        element.BeginAnimation(OpacityProperty,
            new DoubleAnimation(0, 1, duration));

        shift.BeginAnimation(TranslateTransform.YProperty,
            new DoubleAnimation(fromY, 0, duration)
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut },
            });
    }

    private static void FadeTo(UIElement element, double opacity, int milliseconds = 180) =>
        element.BeginAnimation(OpacityProperty,
            new DoubleAnimation(opacity, TimeSpan.FromMilliseconds(milliseconds)));

    // ---------------------------------------------------- window chrome

    // Segoe Fluent Icons chrome glyphs: U+E922 maximise, U+E923 restore. Named because
    // a bare private-use character mid-expression says nothing to a reader.
    private const string MaximiseGlyph = "";
    private const string RestoreGlyph = "";

    private void Header_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            ToggleMaximise();
            return;
        }

        if (e.ButtonState != System.Windows.Input.MouseButtonState.Pressed) return;

        try
        {
            DragMove();
        }
        catch (InvalidOperationException)
        {
            // The button was released before the drag started; nothing to do.
        }
    }

    private void Minimise_Click(object sender, RoutedEventArgs e) =>
        WindowState = WindowState.Minimized;

    private void Maximise_Click(object sender, RoutedEventArgs e) => ToggleMaximise();

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    private void ToggleMaximise() =>
        WindowState = WindowState == WindowState.Maximized
            ? WindowState.Normal
            : WindowState.Maximized;

    private void UpdateMaximiseGlyph()
    {
        var maximised = WindowState == WindowState.Maximized;
        var label = maximised ? "Restore" : "Maximise";

        MaximiseButton.Content = maximised ? RestoreGlyph : MaximiseGlyph;
        MaximiseButton.ToolTip = label;

        // The glyph carries no meaning to a screen reader, so keep the name in step.
        AutomationProperties.SetName(MaximiseButton, label);
    }

    private void ShowWarning(string message) =>
        MessageBox.Show(this, message, "Kelsoft Report Export",
            MessageBoxButton.OK, MessageBoxImage.Warning);
}

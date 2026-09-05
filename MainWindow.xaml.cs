using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace KelsoftReportExport;

/// <summary>A financial year with its tick state in the list.</summary>
public sealed class YearSelection(FinancialYear year, bool isSelected) : INotifyPropertyChanged
{
    private bool _isSelected = isSelected;

    public FinancialYear Year { get; } = year;

    public string Display =>
        $"FY{Year.Label}      {Year.Start:dd/MM/yyyy} – {Year.End:dd/MM/yyyy}      " +
        (Year.EntryCount > 0 ? $"{Year.EntryCount:N0} entries" : "no entries");

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
    private readonly ObservableCollection<YearSelection> _years = [];
    private string _clientName = "";

    public MainWindow()
    {
        InitializeComponent();
        YearsList.ItemsSource = _years;
        YearSelection.SelectionChanged += UpdateExportState;
        Closed += (_, _) => YearSelection.SelectionChanged -= UpdateExportState;
        UpdateExportState();

        // Allow a data file to be passed on the command line, or dropped on the .exe.
        var startupFile = Environment.GetCommandLineArgs().Skip(1).FirstOrDefault();
        if (startupFile is not null && System.IO.File.Exists(startupFile))
            Loaded += (_, _) => LoadDataFile(startupFile);
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
        ClientLabel.Text = $"Client:  {_clientName}";

        foreach (var year in years)
            _years.Add(new YearSelection(year, year.EntryCount > 0));

        if (_years.Count == 0)
            ClientLabel.Text += "      (no financial years defined in this file)";

        OutputBox.Text = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(path) ?? "",
            SuggestFileName(path));

        SetStatus($"{_years.Count} financial year(s) found.");
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

    private void UpdateExportState()
    {
        if (ExportButton is null) return;

        ExportButton.IsEnabled =
            DataFileBox.Text.Length > 0 &&
            _years.Any(y => y.IsSelected) &&
            (WantProfitAndLoss || WantBalanceSheet);
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

        if (years.Count == 0) return;

        if (outputPath.Length == 0)
        {
            ShowWarning("Choose where to save the Excel file.");
            return;
        }

        if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            outputPath += ".xlsx";

        SetBusy(true);

        try
        {
            var clientName = _clientName;
            var unbalanced = new List<string>();

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

                    exports.Add(new YearExport(year, profitAndLoss, balanceSheet));
                }

                Report("Writing workbook…");
                ExcelExporter.Export(exports, outputPath);
            });

            var sheets = years.Count * ((wantProfitAndLoss ? 1 : 0) + (wantBalanceSheet ? 1 : 0));
            SetStatus($"Wrote {sheets} worksheet(s).");

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
        catch (Exception ex)
        {
            SetStatus("Export failed.");
            ShowWarning($"The export did not finish.\n\n{ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void Report(string message) => Dispatcher.Invoke(() => SetStatus(message));

    private void SetStatus(string message) => StatusLabel.Text = message;

    private void SetBusy(bool busy)
    {
        Progress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        Cursor = busy ? System.Windows.Input.Cursors.Wait : null;

        ExportButton.IsEnabled = !busy;
        BrowseDataButton.IsEnabled = !busy;
        BrowseOutputButton.IsEnabled = !busy;
        YearsList.IsEnabled = !busy;

        if (!busy) UpdateExportState();
    }

    private void ShowWarning(string message) =>
        MessageBox.Show(this, message, "Kelsoft Report Export",
            MessageBoxButton.OK, MessageBoxImage.Warning);
}

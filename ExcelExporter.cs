using System.Globalization;
using ClosedXML.Excel;

namespace KelsoftReportExport;

/// <summary>What to write for one financial year — any combination of the reports.</summary>
public sealed record YearExport(
    FinancialYear Year,
    ProfitAndLossStatement? ProfitAndLoss,
    BalanceSheet? BalanceSheet,
    GeneralLedger? GeneralLedger);

/// <summary>
/// Writes a worksheet per report per financial year, laid out like the Access reports
/// but with a column per month.
/// </summary>
public static class ExcelExporter
{
    public static void Export(IEnumerable<YearExport> exports, string outputPath)
    {
        using var workbook = new XLWorkbook();

        foreach (var export in exports)
        {
            if (export.ProfitAndLoss is { } profitAndLoss)
                WriteProfitAndLoss(workbook, profitAndLoss);

            if (export.BalanceSheet is { } balanceSheet)
                WriteBalanceSheet(workbook, balanceSheet);

            if (export.GeneralLedger is { } generalLedger)
                WriteGeneralLedger(workbook, generalLedger);
        }

        workbook.SaveAs(outputPath);
    }

    // ------------------------------------------------------------ profit and loss

    private static void WriteProfitAndLoss(XLWorkbook workbook, ProfitAndLossStatement statement)
    {
        var sheet = workbook.Worksheets.Add($"{statement.Year.SheetName} P&L");
        var writer = new SheetWriter(sheet, statement.MonthCount, trailingHeader: "Total");

        writer.Title(
            statement.ClientName,
            "Profit and Loss Statement",
            $"For the period between {statement.Year.Start:dd/MM/yyyy} and {statement.Year.End:dd/MM/yyyy}");
        writer.Headings(statement.Months);

        if (!statement.Sales.IsEmpty) WriteSection(writer, statement.Sales, statement.MonthCount);
        if (!statement.CostOfSales.IsEmpty) WriteSection(writer, statement.CostOfSales, statement.MonthCount);

        writer.TotalRow("GROSS PROFIT/LOSS", statement.GrossProfit);
        writer.Blank();

        if (!statement.OtherIncome.IsEmpty) WriteSection(writer, statement.OtherIncome, statement.MonthCount);

        writer.TotalRow("TOTAL INCOME", statement.TotalIncome);
        writer.Blank();

        if (!statement.OtherExpenses.IsEmpty) WriteSection(writer, statement.OtherExpenses, statement.MonthCount);

        var netProfit = statement.NetProfit;
        writer.TotalRow(netProfit.Sum() >= 0 ? "NET PROFIT" : "NET LOSS", netProfit, emphasise: true);

        writer.Blank();
        writer.Note("Figures are GST-exclusive net movements, taken from ALLOCATIONS by allocation date. " +
                    "Year-end journals appear in the month they were posted.");
        writer.Finish();
    }

    private static void WriteSection(SheetWriter writer, StatementSection section, int months)
    {
        writer.SectionHeading(section.Heading);

        foreach (var line in section.Lines)
            writer.AccountRow(line.AccountId, line.AccountName, line.Monthly);

        writer.TotalRow(section.TotalLabel, section.Totals(months));
        writer.Blank();
    }

    // -------------------------------------------------------------- balance sheet

    private static void WriteBalanceSheet(XLWorkbook workbook, BalanceSheet sheetData)
    {
        var sheet = workbook.Worksheets.Add($"{sheetData.Year.SheetName} BS");

        // Balances are cumulative, so summing the months would be meaningless — no total column.
        var writer = new SheetWriter(sheet, sheetData.MonthCount);

        writer.Title(
            sheetData.ClientName,
            "Balance Sheet",
            $"As at each month end, financial year ending {sheetData.Year.End:dd/MM/yyyy}");
        writer.Headings(sheetData.Months);

        writer.SectionHeading("OWNERS EQUITY");
        WriteSection(writer, sheetData.MainCapital, sheetData.MonthCount);
        WriteSection(writer, sheetData.OtherCapital, sheetData.MonthCount);
        WriteSection(writer, sheetData.ProfitLoss, sheetData.MonthCount);
        writer.TotalRow("TOTAL OWNERS EQUITY", sheetData.TotalOwnersEquity, emphasise: true);
        writer.Blank();

        writer.SectionHeading("represented by");
        writer.Blank();

        WriteSection(writer, sheetData.CurrentAssets, sheetData.MonthCount);
        WriteSection(writer, sheetData.IntangibleAssets, sheetData.MonthCount);
        WriteSection(writer, sheetData.NonCurrentAssets, sheetData.MonthCount);
        writer.TotalRow("TOTAL ASSETS", sheetData.TotalAssets);
        writer.Blank();

        WriteSection(writer, sheetData.CurrentLiabilities, sheetData.MonthCount);
        WriteSection(writer, sheetData.IntangibleLiabilities, sheetData.MonthCount);
        WriteSection(writer, sheetData.NonCurrentLiabilities, sheetData.MonthCount);
        writer.TotalRow("TOTAL LIABILITIES", sheetData.TotalLiabilities);
        writer.Blank();

        writer.TotalRow("NET ASSETS", sheetData.NetAssets, emphasise: true);

        writer.Blank();
        writer.Note("Balances as at each month end: opening balances for the year, plus movements to that date. " +
                    "Net Assets equals Total Owners Equity in every month.");
        writer.Finish();
    }

    private static void WriteSection(SheetWriter writer, BalanceSheetSection section, int months)
    {
        if (section.IsEmpty) return;

        writer.SectionHeading(section.Heading);

        foreach (var line in section.Lines)
            writer.AccountRow(line.AccountId, line.AccountName, line.Monthly);

        writer.TotalRow(section.TotalLabel, section.Totals(months));
        writer.Blank();
    }

    // ------------------------------------------------------------ general ledger

    private static void WriteGeneralLedger(XLWorkbook workbook, GeneralLedger ledger)
    {
        var sheet = workbook.Worksheets.Add($"{ledger.Year.SheetName} GL");
        var months = ledger.MonthCount;

        var writer = new SheetWriter(sheet, months,
            leadingHeader: "Opening", trailingHeader: "Closing");

        writer.Title(
            ledger.ClientName,
            "General Ledger Summary",
            $"Movement by month, financial year ending {ledger.Year.End:dd/MM/yyyy}");
        writer.Headings(ledger.Months);

        foreach (var group in ledger.Groups)
        {
            writer.SectionHeading(group.Heading);

            foreach (var line in group.Lines)
                writer.AccountRow(line.AccountId, line.AccountName, line.Monthly,
                    leading: line.Opening, trailing: line.Closing);

            writer.TotalRow($"TOTAL {group.Heading.ToUpperInvariant()}",
                group.MonthlyTotals(months),
                leading: group.OpeningTotal, trailing: group.ClosingTotal);
            writer.Blank();
        }

        writer.TotalRow("TOTAL — ALL ACCOUNTS", ledger.MonthlyTotals(),
            leading: ledger.OpeningTotal, trailing: ledger.ClosingTotal, emphasise: true);

        writer.Blank();
        writer.Note("Debit positive, credit negative, on the raw ledger basis — no presentation sign flips. " +
                    "Because every entry balances, the all-accounts row is zero in a period that balances.");
        writer.Finish();
    }

    // ---------------------------------------------------------------- sheet layout

    private sealed class SheetWriter
    {
        private const string MoneyFormat = "#,##0.00;(#,##0.00);\"-\"";
        private const int CodeColumn = 1;
        private const int NameColumn = 2;
        private const int FirstValueColumn = 3;
        private const int HeaderRow = 5;

        private readonly IXLWorksheet _sheet;
        private readonly int _months;
        private readonly string? _leadingHeader;
        private readonly string? _trailingHeader;
        private readonly int _firstMonthColumn;

        private int _row = HeaderRow + 1;

        public SheetWriter(IXLWorksheet sheet, int months,
            string? leadingHeader = null, string? trailingHeader = null)
        {
            _sheet = sheet;
            _months = months;
            _leadingHeader = leadingHeader;
            _trailingHeader = trailingHeader;
            _firstMonthColumn = FirstValueColumn + (leadingHeader is null ? 0 : 1);
        }

        private int TrailingColumn => _firstMonthColumn + _months;

        private int LastColumn => _trailingHeader is null ? TrailingColumn - 1 : TrailingColumn;

        public void Title(string client, string reportName, string period)
        {
            Set(1, client, bold: true, size: 14);
            Set(2, reportName, bold: true, size: 12);
            Set(3, period, bold: false, size: 9.5);

            foreach (var row in new[] { 1, 2, 3 })
                _sheet.Range(row, CodeColumn, row, LastColumn).Merge();

            void Set(int row, string text, bool bold, double size)
            {
                var cell = _sheet.Cell(row, CodeColumn);
                cell.Value = text;
                cell.Style.Font.Bold = bold;
                cell.Style.Font.FontSize = size;
            }
        }

        public void Headings(IReadOnlyList<DateTime> monthList)
        {
            _sheet.Cell(HeaderRow, CodeColumn).Value = "Code";
            _sheet.Cell(HeaderRow, NameColumn).Value = "Account";

            if (_leadingHeader is not null)
                _sheet.Cell(HeaderRow, FirstValueColumn).Value = _leadingHeader;

            // Invariant culture keeps every abbreviation to three letters — current CLDR
            // data renders September as "Sept", which breaks the column rhythm.
            for (var i = 0; i < _months; i++)
                _sheet.Cell(HeaderRow, _firstMonthColumn + i).Value =
                    monthList[i].ToString("MMM yy", CultureInfo.InvariantCulture);

            if (_trailingHeader is not null)
                _sheet.Cell(HeaderRow, TrailingColumn).Value = _trailingHeader;

            var headings = _sheet.Range(HeaderRow, CodeColumn, HeaderRow, LastColumn);
            headings.Style.Font.Bold = true;
            headings.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            headings.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            _sheet.Cell(HeaderRow, CodeColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            _sheet.Cell(HeaderRow, NameColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }

        public void SectionHeading(string text)
        {
            var cell = _sheet.Cell(_row, NameColumn);
            cell.Value = text;
            cell.Style.Font.Bold = true;
            _row++;
        }

        public void AccountRow(double accountId, string name, decimal[] values,
            decimal? leading = null, decimal? trailing = null)
        {
            var code = _sheet.Cell(_row, CodeColumn);
            code.Value = accountId;
            code.Style.NumberFormat.Format = "0.00";

            _sheet.Cell(_row, NameColumn).Value = name;
            WriteValues(values, leading, trailing);
            _row++;
        }

        public void TotalRow(string label, decimal[] values,
            decimal? leading = null, decimal? trailing = null, bool emphasise = false)
        {
            _sheet.Cell(_row, NameColumn).Value = label;
            WriteValues(values, leading, trailing);

            var span = _sheet.Range(_row, NameColumn, _row, LastColumn);
            span.Style.Font.Bold = true;
            span.Style.Border.TopBorder = XLBorderStyleValues.Thin;

            if (emphasise)
                span.Style.Border.BottomBorder = XLBorderStyleValues.Double;

            _row++;
        }

        public void Blank() => _row++;

        public void Note(string text)
        {
            var cell = _sheet.Cell(_row, NameColumn);
            cell.Value = text;
            cell.Style.Font.Italic = true;
            cell.Style.Font.FontSize = 9;
            _row++;
        }

        public void Finish()
        {
            _sheet.Column(CodeColumn).Width = 9;
            _sheet.Column(NameColumn).Width = 40;

            if (_leadingHeader is not null)
                _sheet.Column(FirstValueColumn).Width = 15;

            for (var i = 0; i < _months; i++)
                _sheet.Column(_firstMonthColumn + i).Width = 13.5;

            if (_trailingHeader is not null)
                _sheet.Column(TrailingColumn).Width = 15;

            _sheet.SheetView.FreezeRows(HeaderRow);
            _sheet.SheetView.FreezeColumns(NameColumn);

            _sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            _sheet.PageSetup.FitToPages(1, 0);
        }

        private void WriteValues(decimal[] values, decimal? leading, decimal? trailing)
        {
            if (_leadingHeader is not null)
                Money(FirstValueColumn, leading ?? 0m, bold: true);

            for (var i = 0; i < _months; i++)
                Money(_firstMonthColumn + i, values[i], bold: false);

            if (_trailingHeader is not null)
                Money(TrailingColumn, trailing ?? values.Sum(), bold: true);

            void Money(int column, decimal value, bool bold)
            {
                var cell = _sheet.Cell(_row, column);
                cell.Value = value;
                cell.Style.NumberFormat.Format = MoneyFormat;
                if (bold) cell.Style.Font.Bold = true;
            }
        }
    }
}

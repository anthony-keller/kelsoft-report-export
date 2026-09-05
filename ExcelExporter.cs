using System.Globalization;
using ClosedXML.Excel;

namespace KelsoftReportExport;

/// <summary>What to write for one financial year — either report, or both.</summary>
public sealed record YearExport(
    FinancialYear Year,
    ProfitAndLossStatement? ProfitAndLoss,
    BalanceSheet? BalanceSheet);

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
        }

        workbook.SaveAs(outputPath);
    }

    // ------------------------------------------------------------ profit and loss

    private static void WriteProfitAndLoss(XLWorkbook workbook, ProfitAndLossStatement statement)
    {
        var sheet = workbook.Worksheets.Add($"{statement.Year.SheetName} P&L");
        var writer = new SheetWriter(sheet, statement.MonthCount, withTotalColumn: true);

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
        var writer = new SheetWriter(sheet, sheetData.MonthCount, withTotalColumn: false);

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

    // ---------------------------------------------------------------- sheet layout

    private sealed class SheetWriter(IXLWorksheet sheet, int months, bool withTotalColumn)
    {
        private const string MoneyFormat = "#,##0.00;(#,##0.00);\"-\"";
        private const int CodeColumn = 1;
        private const int NameColumn = 2;
        private const int FirstMonthColumn = 3;
        private const int HeaderRow = 5;

        private int _row = HeaderRow + 1;

        private int TotalColumn => FirstMonthColumn + months;
        private int LastColumn => withTotalColumn ? TotalColumn : TotalColumn - 1;

        public void Title(string client, string reportName, string period)
        {
            Set(1, client, bold: true, size: 14);
            Set(2, reportName, bold: true, size: 12);
            Set(3, period, bold: false, size: 9.5);

            foreach (var row in new[] { 1, 2, 3 })
                sheet.Range(row, CodeColumn, row, LastColumn).Merge();

            void Set(int row, string text, bool bold, double size)
            {
                var cell = sheet.Cell(row, CodeColumn);
                cell.Value = text;
                cell.Style.Font.Bold = bold;
                cell.Style.Font.FontSize = size;
            }
        }

        public void Headings(IReadOnlyList<DateTime> monthList)
        {
            sheet.Cell(HeaderRow, CodeColumn).Value = "Code";
            sheet.Cell(HeaderRow, NameColumn).Value = "Account";

            // Invariant culture keeps every abbreviation to three letters — current CLDR
            // data renders September as "Sept", which breaks the column rhythm.
            for (var i = 0; i < months; i++)
                sheet.Cell(HeaderRow, FirstMonthColumn + i).Value =
                    monthList[i].ToString("MMM yy", CultureInfo.InvariantCulture);

            if (withTotalColumn)
                sheet.Cell(HeaderRow, TotalColumn).Value = "Total";

            var headings = sheet.Range(HeaderRow, CodeColumn, HeaderRow, LastColumn);
            headings.Style.Font.Bold = true;
            headings.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            headings.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            sheet.Cell(HeaderRow, CodeColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Cell(HeaderRow, NameColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }

        public void SectionHeading(string text)
        {
            var cell = sheet.Cell(_row, NameColumn);
            cell.Value = text;
            cell.Style.Font.Bold = true;
            _row++;
        }

        public void AccountRow(double accountId, string name, decimal[] values)
        {
            var code = sheet.Cell(_row, CodeColumn);
            code.Value = accountId;
            code.Style.NumberFormat.Format = "0.00";

            sheet.Cell(_row, NameColumn).Value = name;
            WriteValues(values);
            _row++;
        }

        public void TotalRow(string label, decimal[] values, bool emphasise = false)
        {
            sheet.Cell(_row, NameColumn).Value = label;
            WriteValues(values);

            var span = sheet.Range(_row, NameColumn, _row, LastColumn);
            span.Style.Font.Bold = true;
            span.Style.Border.TopBorder = XLBorderStyleValues.Thin;

            if (emphasise)
                span.Style.Border.BottomBorder = XLBorderStyleValues.Double;

            _row++;
        }

        public void Blank() => _row++;

        public void Note(string text)
        {
            var cell = sheet.Cell(_row, NameColumn);
            cell.Value = text;
            cell.Style.Font.Italic = true;
            cell.Style.Font.FontSize = 9;
            _row++;
        }

        public void Finish()
        {
            sheet.Column(CodeColumn).Width = 9;
            sheet.Column(NameColumn).Width = 40;
            for (var i = 0; i < months; i++)
                sheet.Column(FirstMonthColumn + i).Width = 13.5;
            if (withTotalColumn)
                sheet.Column(TotalColumn).Width = 15;

            sheet.SheetView.FreezeRows(HeaderRow);
            sheet.SheetView.FreezeColumns(NameColumn);

            sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            sheet.PageSetup.FitToPages(1, 0);
        }

        private void WriteValues(decimal[] values)
        {
            for (var i = 0; i < months; i++)
            {
                var cell = sheet.Cell(_row, FirstMonthColumn + i);
                cell.Value = values[i];
                cell.Style.NumberFormat.Format = MoneyFormat;
            }

            if (!withTotalColumn) return;

            var total = sheet.Cell(_row, TotalColumn);
            total.Value = values.Sum();
            total.Style.NumberFormat.Format = MoneyFormat;
            total.Style.Font.Bold = true;
        }
    }
}

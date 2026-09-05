namespace KelsoftReportExport;

/// <summary>
/// The ACCOUNT_TYPES lookup lives in the Kelsoft front-end (Kelsoft.accdb), not in the
/// client data file, so the map is carried here. Ids and names match that table exactly —
/// the names matter because the original report's Net expression switches on them.
/// </summary>
public static class AccountTypes
{
    public const int CurrentAsset = 1;
    public const int NonCurrentAsset = 2;
    public const int CurrentLiability = 3;
    public const int NonCurrentLiability = 4;
    public const int MainCapital = 5;
    public const int OtherExpenses = 6;
    public const int OtherIncome = 7;
    public const int Sales = 9;
    public const int CostOfSales = 10;
    public const int IntangibleAsset = 12;
    public const int IntangibleLiability = 13;
    public const int ProfitOrLoss = 14;
    public const int OtherCapital = 15;

    /// <summary>Accounts the profit or loss for the year is posted to on the balance sheet.</summary>
    public const double NetProfitAccount = 9998;
    public const double NetLossAccount = 9999;

    /// <summary>GST accrued on allocations is carried against these two accounts.</summary>
    public const int GstInputCredits = 1000;
    public const int GstCollected = 1001;

    public static readonly IReadOnlyDictionary<int, string> Names = new Dictionary<int, string>
    {
        [1] = "Current Asset",
        [2] = "Non-Current Asset",
        [3] = "Current Liability",
        [4] = "Non-Current Liability",
        [5] = "Main Capital",
        [6] = "Other Expenses",
        [7] = "Other Income",
        [9] = "Sales",
        [10] = "Cost of Sales",
        [12] = "Intangible Asset",
        [13] = "Intangible Liability",
        [14] = "Profit/Loss",
        [15] = "Other Capital",
    };

    /// <summary>The four types the Profit and Loss Statement reports on.</summary>
    public static readonly int[] ProfitAndLoss = [Sales, CostOfSales, OtherIncome, OtherExpenses];

    /// <summary>
    /// Every type in the order a general ledger presents them: assets, liabilities,
    /// capital, then the trading accounts.
    /// </summary>
    public static readonly int[] LedgerOrder =
    [
        CurrentAsset, NonCurrentAsset, IntangibleAsset,
        CurrentLiability, NonCurrentLiability, IntangibleLiability,
        MainCapital, OtherCapital, ProfitOrLoss,
        Sales, CostOfSales, OtherIncome, OtherExpenses,
    ];

    /// <summary>
    /// Revenue nets credit-positive, expenditure nets debit-positive — mirrors the
    /// IIf chain in qryCreateCombinedTransactionsAll.
    /// </summary>
    public static bool IsCreditPositive(int accountType) =>
        accountType is Sales or OtherIncome;
}

public sealed record FinancialYear(string Label, DateTime Start, DateTime End, int EntryCount)
{
    public string SheetName => $"FY{Label}";

    public override string ToString() =>
        $"FY{Label}   {Start:dd/MM/yyyy} - {End:dd/MM/yyyy}   ({EntryCount:N0} entries)";

    /// <summary>The month buckets spanned by this year, in order.</summary>
    public IReadOnlyList<DateTime> Months()
    {
        var months = new List<DateTime>();
        var cursor = new DateTime(Start.Year, Start.Month, 1);
        var last = new DateTime(End.Year, End.Month, 1);
        while (cursor <= last)
        {
            months.Add(cursor);
            cursor = cursor.AddMonths(1);
        }
        return months;
    }
}

public sealed record Account(double Id, string Name, int AccountType);

/// <summary>One account's line on the statement: its net movement in each month.</summary>
public sealed class StatementLine
{
    public required double AccountId { get; init; }
    public required string AccountName { get; init; }
    public required decimal[] Monthly { get; init; }

    public decimal Total => Monthly.Sum();

    /// <summary>The report suppresses rows that round to zero; a monthly row is kept if any month moved.</summary>
    public bool HasMovement => Monthly.Any(m => Math.Round(m, 2) != 0m) || Math.Round(Total, 2) != 0m;
}

/// <summary>A block of the statement — the account rows for one account type, plus its total.</summary>
public sealed class StatementSection
{
    public required string Heading { get; init; }
    public required string TotalLabel { get; init; }
    public required List<StatementLine> Lines { get; init; }

    public bool IsEmpty => Lines.Count == 0;

    public decimal[] Totals(int monthCount)
    {
        var totals = new decimal[monthCount];
        foreach (var line in Lines)
            for (var i = 0; i < monthCount; i++)
                totals[i] += line.Monthly[i];
        return totals;
    }
}

/// <summary>A full Profit and Loss Statement for one financial year, laid out by month.</summary>
public sealed class ProfitAndLossStatement
{
    public required string ClientName { get; init; }
    public required FinancialYear Year { get; init; }
    public required IReadOnlyList<DateTime> Months { get; init; }
    public required StatementSection Sales { get; init; }
    public required StatementSection CostOfSales { get; init; }
    public required StatementSection OtherIncome { get; init; }
    public required StatementSection OtherExpenses { get; init; }

    public int MonthCount => Months.Count;

    /// <summary>Gross Profit = Total Sales - Total Cost of Sales.</summary>
    public decimal[] GrossProfit => Subtract(Sales.Totals(MonthCount), CostOfSales.Totals(MonthCount));

    /// <summary>Total Income = Gross Profit + Other Income.</summary>
    public decimal[] TotalIncome => Add(GrossProfit, OtherIncome.Totals(MonthCount));

    /// <summary>Net Profit = Gross Profit + Other Income - Other Expenses.</summary>
    public decimal[] NetProfit => Subtract(TotalIncome, OtherExpenses.Totals(MonthCount));

    private static decimal[] Add(decimal[] a, decimal[] b) =>
        [.. a.Select((v, i) => v + b[i])];

    private static decimal[] Subtract(decimal[] a, decimal[] b) =>
        [.. a.Select((v, i) => v - b[i])];
}

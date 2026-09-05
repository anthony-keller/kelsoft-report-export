namespace KelsoftReportExport;

/// <summary>
/// One balance sheet line: an account and its closing balance at each month end,
/// already flipped to the side it is presented on (assets debit-positive,
/// liabilities and capital credit-positive).
/// </summary>
public sealed class BalanceSheetLine
{
    public required double AccountId { get; init; }
    public required string AccountName { get; init; }
    public required decimal[] Monthly { get; init; }

    /// <summary>The report suppresses rows where round(Net,2) = 0 in every month.</summary>
    public bool HasBalance => Monthly.Any(v => Math.Round(v, 2) != 0m);
}

public sealed class BalanceSheetSection
{
    public required string Heading { get; init; }
    public required string TotalLabel { get; init; }
    public required List<BalanceSheetLine> Lines { get; init; }

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

/// <summary>
/// A Balance Sheet for one financial year, showing the position as at each month end.
/// Unlike the P&amp;L these are cumulative balances, so there is no meaningful annual total —
/// the final month is the year-end position.
/// </summary>
public sealed class BalanceSheet
{
    public required string ClientName { get; init; }
    public required FinancialYear Year { get; init; }
    public required IReadOnlyList<DateTime> Months { get; init; }

    public required BalanceSheetSection CurrentAssets { get; init; }
    public required BalanceSheetSection IntangibleAssets { get; init; }
    public required BalanceSheetSection NonCurrentAssets { get; init; }
    public required BalanceSheetSection CurrentLiabilities { get; init; }
    public required BalanceSheetSection IntangibleLiabilities { get; init; }
    public required BalanceSheetSection NonCurrentLiabilities { get; init; }
    public required BalanceSheetSection MainCapital { get; init; }
    public required BalanceSheetSection OtherCapital { get; init; }
    public required BalanceSheetSection ProfitLoss { get; init; }

    public int MonthCount => Months.Count;

    public decimal[] TotalAssets => Sum(
        CurrentAssets.Totals(MonthCount),
        IntangibleAssets.Totals(MonthCount),
        NonCurrentAssets.Totals(MonthCount));

    public decimal[] TotalLiabilities => Sum(
        CurrentLiabilities.Totals(MonthCount),
        IntangibleLiabilities.Totals(MonthCount),
        NonCurrentLiabilities.Totals(MonthCount));

    /// <summary>Total Assets − Total Liabilities.</summary>
    public decimal[] NetAssets
    {
        get
        {
            var assets = TotalAssets;
            var liabilities = TotalLiabilities;
            return [.. assets.Select((v, i) => v - liabilities[i])];
        }
    }

    /// <summary>Main Capital + Other Capital + Profit/Loss, all on a credit basis.</summary>
    public decimal[] TotalOwnersEquity => Sum(
        MainCapital.Totals(MonthCount),
        OtherCapital.Totals(MonthCount),
        ProfitLoss.Totals(MonthCount));

    /// <summary>
    /// The accounting identity. Every entry balances, so Net Assets must equal Owners Equity
    /// in every month; a failure means the underlying data does not balance.
    /// </summary>
    public bool Balances(out decimal worstDifference, out DateTime? worstMonth)
    {
        var net = NetAssets;
        var equity = TotalOwnersEquity;

        worstDifference = 0m;
        worstMonth = null;

        for (var i = 0; i < MonthCount; i++)
        {
            var difference = Math.Abs(Math.Round(net[i] - equity[i], 2));
            if (difference <= worstDifference) continue;

            worstDifference = difference;
            worstMonth = Months[i];
        }

        return worstDifference <= 0.01m;
    }

    private decimal[] Sum(params decimal[][] parts)
    {
        var totals = new decimal[MonthCount];
        foreach (var part in parts)
            for (var i = 0; i < MonthCount; i++)
                totals[i] += part[i];
        return totals;
    }
}

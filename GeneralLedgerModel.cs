namespace KelsoftReportExport;

/// <summary>
/// One account's line in the general ledger summary: where it started the year,
/// what it moved each month, and where it finished. Movements are on the raw ledger
/// basis — debit positive, credit negative — for every account type alike.
/// </summary>
public sealed class GeneralLedgerLine
{
    public required double AccountId { get; init; }
    public required string AccountName { get; init; }
    public required decimal Opening { get; init; }
    public required decimal[] Monthly { get; init; }

    public decimal Closing => Opening + Monthly.Sum();

    public bool HasActivity =>
        Math.Round(Opening, 2) != 0m ||
        Math.Round(Closing, 2) != 0m ||
        Monthly.Any(v => Math.Round(v, 2) != 0m);
}

/// <summary>Accounts of one type, in the order the ledger presents them.</summary>
public sealed class GeneralLedgerGroup
{
    public required string Heading { get; init; }
    public required List<GeneralLedgerLine> Lines { get; init; }

    public bool IsEmpty => Lines.Count == 0;

    public decimal OpeningTotal => Lines.Sum(l => l.Opening);

    public decimal ClosingTotal => Lines.Sum(l => l.Closing);

    public decimal[] MonthlyTotals(int monthCount)
    {
        var totals = new decimal[monthCount];
        foreach (var line in Lines)
            for (var i = 0; i < monthCount; i++)
                totals[i] += line.Monthly[i];
        return totals;
    }
}

/// <summary>
/// A general ledger summary for one financial year — every account, grouped by type,
/// with a column per month. Effectively a trial balance carried across twelve periods.
/// </summary>
public sealed class GeneralLedger
{
    public required string ClientName { get; init; }
    public required FinancialYear Year { get; init; }
    public required IReadOnlyList<DateTime> Months { get; init; }
    public required List<GeneralLedgerGroup> Groups { get; init; }

    public int MonthCount => Months.Count;

    public decimal OpeningTotal => Groups.Sum(g => g.OpeningTotal);

    public decimal ClosingTotal => Groups.Sum(g => g.ClosingTotal);

    public decimal[] MonthlyTotals()
    {
        var totals = new decimal[MonthCount];
        foreach (var group in Groups)
        {
            var groupTotals = group.MonthlyTotals(MonthCount);
            for (var i = 0; i < MonthCount; i++)
                totals[i] += groupTotals[i];
        }
        return totals;
    }

    /// <summary>
    /// Debits equal credits, so the cumulative position across every account must be
    /// zero at each month end. A non-zero figure means the source entries do not balance.
    /// </summary>
    public bool Balances(out decimal worstDifference, out DateTime? worstMonth)
    {
        var monthly = MonthlyTotals();
        var running = OpeningTotal;

        worstDifference = 0m;
        worstMonth = null;

        for (var i = 0; i < MonthCount; i++)
        {
            running += monthly[i];

            var difference = Math.Abs(Math.Round(running, 2));
            if (difference <= worstDifference) continue;

            worstDifference = difference;
            worstMonth = Months[i];
        }

        return worstDifference <= 0.01m;
    }
}

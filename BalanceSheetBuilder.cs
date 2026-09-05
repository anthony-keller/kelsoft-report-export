namespace KelsoftReportExport;

/// <summary>
/// Builds a Balance Sheet showing the position as at each month end of a financial year.
/// <para>
/// The Access report's <c>qryBsTransactions</c> is an eleven-way union because it splits a
/// period from the prior part of the year. Run with the period starting on the first day of
/// the financial year — which is what a cumulative monthly column needs — every "prior part
/// of the year" branch is empty and it reduces to: opening balances, plus balance sheet
/// movements to date, plus the profit or loss for the year to date, plus the two GST accruals.
/// </para>
/// </summary>
public static class BalanceSheetBuilder
{
    private static readonly int[] BalanceSheetTypes =
    [
        AccountTypes.CurrentAsset, AccountTypes.NonCurrentAsset, AccountTypes.IntangibleAsset,
        AccountTypes.CurrentLiability, AccountTypes.NonCurrentLiability, AccountTypes.IntangibleLiability,
        AccountTypes.MainCapital, AccountTypes.OtherCapital,
    ];

    public static BalanceSheet Build(KelsoftDataFile file, FinancialYear year, string clientName)
    {
        var accounts = file.Accounts();
        var months = year.Months();
        var count = months.Count;

        var monthIndex = months
            .Select((m, i) => (Key: m.Year * 100 + m.Month, Index: i))
            .ToDictionary(x => x.Key, x => x.Index);

        // Movement per account per month, as a net amount (debit positive).
        var deltas = new Dictionary<double, decimal[]>();
        var profitDelta = new decimal[count];

        decimal[] Series(double accountId)
        {
            if (!deltas.TryGetValue(accountId, out var series))
                deltas[accountId] = series = new decimal[count];
            return series;
        }

        foreach (var (accountId, y, m, debit, credit) in file.MonthlyMovements(year.Start, year.End))
        {
            if (!accounts.TryGetValue(accountId, out var account)) continue;
            if (!monthIndex.TryGetValue(y * 100 + m, out var index)) continue;

            var net = debit - credit;

            if (BalanceSheetTypes.Contains(account.AccountType))
                Series(accountId)[index] += net;
            else if (AccountTypes.ProfitAndLoss.Contains(account.AccountType))
                profitDelta[index] += net;
        }

        // GST accrued on allocation rows: input credits are a debit, GST collected a credit.
        foreach (var (gstAccount, y, m, amount) in file.MonthlyGst(year.Start, year.End))
        {
            if (!monthIndex.TryGetValue(y * 100 + m, out var index)) continue;

            Series(gstAccount)[index] += gstAccount == AccountTypes.GstInputCredits ? amount : -amount;
        }

        // Opening balances apply to every month of the year.
        var opening = file.OpeningBalances(year.Label);
        foreach (var accountId in opening.Keys)
            Series(accountId);

        var balances = new Dictionary<double, decimal[]>();
        foreach (var (accountId, series) in deltas)
        {
            var running = opening.TryGetValue(accountId, out var start) ? start : 0m;
            var cumulative = new decimal[count];
            for (var i = 0; i < count; i++)
            {
                running += series[i];
                cumulative[i] = running;
            }
            balances[accountId] = cumulative;
        }

        var profitLoss = ProfitLossLines(profitDelta, accounts, count);

        return new BalanceSheet
        {
            ClientName = clientName,
            Year = year,
            Months = months,
            CurrentAssets = Section("Current Assets", "TOTAL CURRENT ASSETS",
                AccountTypes.CurrentAsset, debitBasis: true, accounts, balances, count),
            IntangibleAssets = Section("Intangible Assets", "TOTAL INTANGIBLE ASSETS",
                AccountTypes.IntangibleAsset, debitBasis: true, accounts, balances, count),
            NonCurrentAssets = Section("Non-Current Assets", "TOTAL NON-CURRENT ASSETS",
                AccountTypes.NonCurrentAsset, debitBasis: true, accounts, balances, count),
            CurrentLiabilities = Section("Current Liabilities", "TOTAL CURRENT LIABILITIES",
                AccountTypes.CurrentLiability, debitBasis: false, accounts, balances, count),
            IntangibleLiabilities = Section("Intangible Liabilities", "TOTAL INTANGIBLE LIABILITIES",
                AccountTypes.IntangibleLiability, debitBasis: false, accounts, balances, count),
            NonCurrentLiabilities = Section("Non-Current Liabilities", "TOTAL NON-CURRENT LIABILITIES",
                AccountTypes.NonCurrentLiability, debitBasis: false, accounts, balances, count),
            MainCapital = Section("Capital", "TOTAL CAPITAL",
                AccountTypes.MainCapital, debitBasis: false, accounts, balances, count),
            OtherCapital = Section("Other Capital", "TOTAL OTHER CAPITAL",
                AccountTypes.OtherCapital, debitBasis: false, accounts, balances, count),
            ProfitLoss = profitLoss,
        };
    }

    /// <summary>
    /// The year-to-date trading result, carried onto the balance sheet against account 9998
    /// (Net Profit) or 9999 (Net Loss) exactly as the report does — which one depends on the
    /// sign in that month, so a year that crosses over populates both rows.
    /// </summary>
    private static BalanceSheetSection ProfitLossLines(
        decimal[] profitDelta, IReadOnlyDictionary<double, Account> accounts, int count)
    {
        var profit = new decimal[count];
        var loss = new decimal[count];

        var running = 0m;
        for (var i = 0; i < count; i++)
        {
            running += profitDelta[i];

            // Net is debit positive, so a positive balance on profit and loss accounts is a loss.
            if (running < 0) profit[i] = -running;
            else loss[i] = -running;
        }

        var lines = new List<BalanceSheetLine>();

        foreach (var (accountId, values) in
                 new[] { (AccountTypes.NetProfitAccount, profit), (AccountTypes.NetLossAccount, loss) })
        {
            var line = new BalanceSheetLine
            {
                AccountId = accountId,
                AccountName = accounts.TryGetValue(accountId, out var account)
                    ? account.Name
                    : accountId == AccountTypes.NetProfitAccount ? "Net Profit" : "Net Loss",
                Monthly = values,
            };

            if (line.HasBalance) lines.Add(line);
        }

        return new BalanceSheetSection
        {
            Heading = "Profit / Loss",
            TotalLabel = "TOTAL PROFIT / LOSS",
            Lines = lines,
        };
    }

    private static BalanceSheetSection Section(
        string heading,
        string totalLabel,
        int accountType,
        bool debitBasis,
        IReadOnlyDictionary<double, Account> accounts,
        IReadOnlyDictionary<double, decimal[]> balances,
        int count)
    {
        var lines = new List<BalanceSheetLine>();

        foreach (var (accountId, cumulative) in balances)
        {
            if (!accounts.TryGetValue(accountId, out var account)) continue;
            if (account.AccountType != accountType) continue;

            // Assets are shown as debit balances, liabilities and capital as credit balances.
            var presented = debitBasis
                ? cumulative
                : cumulative.Select(v => -v).ToArray();

            var line = new BalanceSheetLine
            {
                AccountId = accountId,
                AccountName = account.Name,
                Monthly = presented,
            };

            if (line.HasBalance) lines.Add(line);
        }

        lines.Sort((a, b) => a.AccountId.CompareTo(b.AccountId));

        return new BalanceSheetSection
        {
            Heading = heading,
            TotalLabel = totalLabel,
            Lines = lines,
        };
    }
}

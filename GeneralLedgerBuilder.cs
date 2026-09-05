namespace KelsoftReportExport;

/// <summary>
/// Builds a monthly general ledger summary — every account, its opening balance, its
/// movement in each month and its closing balance.
/// <para>
/// The Access "Ledger Enquiry" lists each allocation individually beneath an opening
/// balance row. This is the same ledger summarised to one column per month, which makes
/// it a trial balance carried across the year: the closing column for balance sheet
/// accounts matches the Balance Sheet, and the movements on trading accounts match the
/// Profit and Loss (before its presentation sign flips).
/// </para>
/// </summary>
public static class GeneralLedgerBuilder
{
    public static GeneralLedger Build(KelsoftDataFile file, FinancialYear year, string clientName)
    {
        var accounts = file.Accounts();
        var months = year.Months();
        var count = months.Count;

        var monthIndex = months
            .Select((m, i) => (Key: m.Year * 100 + m.Month, Index: i))
            .ToDictionary(x => x.Key, x => x.Index);

        var movements = new Dictionary<double, decimal[]>();

        decimal[] Series(double accountId)
        {
            if (!movements.TryGetValue(accountId, out var series))
                movements[accountId] = series = new decimal[count];
            return series;
        }

        foreach (var (accountId, y, m, debit, credit) in file.MonthlyMovements(year.Start, year.End))
        {
            if (!accounts.ContainsKey(accountId)) continue;
            if (!monthIndex.TryGetValue(y * 100 + m, out var index)) continue;

            Series(accountId)[index] += debit - credit;
        }

        // GST accrued on allocation rows rather than posted directly — the same two
        // accounts the balance sheet carries, so the ledger agrees with it.
        foreach (var (gstAccount, y, m, amount) in file.MonthlyGst(year.Start, year.End))
        {
            if (!monthIndex.TryGetValue(y * 100 + m, out var index)) continue;

            Series(gstAccount)[index] += gstAccount == AccountTypes.GstInputCredits ? amount : -amount;
        }

        var opening = file.OpeningBalances(year.Label);
        foreach (var accountId in opening.Keys)
            if (accounts.ContainsKey(accountId))
                Series(accountId);

        var groups = new List<GeneralLedgerGroup>();

        foreach (var accountType in AccountTypes.LedgerOrder)
        {
            var lines = new List<GeneralLedgerLine>();

            foreach (var (accountId, monthly) in movements)
            {
                var account = accounts[accountId];
                if (account.AccountType != accountType) continue;

                var line = new GeneralLedgerLine
                {
                    AccountId = accountId,
                    AccountName = account.Name,
                    Opening = opening.TryGetValue(accountId, out var start) ? start : 0m,
                    Monthly = monthly,
                };

                if (line.HasActivity) lines.Add(line);
            }

            if (lines.Count == 0) continue;

            lines.Sort((a, b) => a.AccountId.CompareTo(b.AccountId));

            groups.Add(new GeneralLedgerGroup
            {
                Heading = AccountTypes.Names.TryGetValue(accountType, out var name)
                    ? name
                    : $"Account type {accountType}",
                Lines = lines,
            });
        }

        return new GeneralLedger
        {
            ClientName = clientName,
            Year = year,
            Months = months,
            Groups = groups,
        };
    }
}

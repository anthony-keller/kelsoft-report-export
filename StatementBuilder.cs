namespace KelsoftReportExport;

/// <summary>
/// Builds a monthly Profit and Loss Statement using the same rules as the Access report:
/// net movement per account over the period, signed by account type, zero rows suppressed.
/// </summary>
public static class StatementBuilder
{
    public static ProfitAndLossStatement Build(KelsoftDataFile file, FinancialYear year, string clientName)
    {
        var accounts = file.Accounts();
        var months = year.Months();
        var monthIndex = months
            .Select((m, i) => (Key: m.Year * 100 + m.Month, Index: i))
            .ToDictionary(x => x.Key, x => x.Index);

        // account id -> monthly net movement
        var nets = new Dictionary<double, decimal[]>();

        foreach (var (accountId, y, m, debit, credit) in file.MonthlyMovements(year.Start, year.End))
        {
            if (!accounts.TryGetValue(accountId, out var account)) continue;
            if (!AccountTypes.ProfitAndLoss.Contains(account.AccountType)) continue;
            if (!monthIndex.TryGetValue(y * 100 + m, out var index)) continue;

            var net = AccountTypes.IsCreditPositive(account.AccountType)
                ? credit - debit
                : debit - credit;

            if (!nets.TryGetValue(accountId, out var monthly))
                nets[accountId] = monthly = new decimal[months.Count];

            monthly[index] += net;
        }

        return new ProfitAndLossStatement
        {
            ClientName = clientName,
            Year = year,
            Months = months,
            Sales = Section("TRADING INCOME", "TOTAL SALES", AccountTypes.Sales, accounts, nets, months.Count),
            CostOfSales = Section("less Cost of Sales", "TOTAL COST OF SALES", AccountTypes.CostOfSales, accounts, nets, months.Count),
            OtherIncome = Section("Income", "INCOME", AccountTypes.OtherIncome, accounts, nets, months.Count),
            OtherExpenses = Section("Expenditure", "TOTAL EXPENSES", AccountTypes.OtherExpenses, accounts, nets, months.Count),
        };
    }

    private static StatementSection Section(
        string heading,
        string totalLabel,
        int accountType,
        IReadOnlyDictionary<double, Account> accounts,
        IReadOnlyDictionary<double, decimal[]> nets,
        int monthCount)
    {
        var lines = new List<StatementLine>();

        foreach (var (accountId, monthly) in nets)
        {
            var account = accounts[accountId];
            if (account.AccountType != accountType) continue;

            var line = new StatementLine
            {
                AccountId = accountId,
                AccountName = account.Name,
                Monthly = monthly,
            };

            if (line.HasMovement) lines.Add(line);
        }

        lines.Sort((a, b) => a.AccountId.CompareTo(b.AccountId));

        return new StatementSection
        {
            Heading = heading,
            TotalLabel = totalLabel,
            Lines = lines,
        };
    }
}

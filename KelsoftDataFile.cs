using System.Data;
using System.Data.OleDb;

namespace KelsoftReportExport;

/// <summary>
/// Reads a Kelsoft client data file (the back-end .mdb the front-end links to).
/// </summary>
public sealed class KelsoftDataFile : IDisposable
{
    private static readonly string[] Providers =
        ["Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"];

    private readonly OleDbConnection _connection;

    public string Path { get; }

    public KelsoftDataFile(string path)
    {
        Path = path;
        _connection = Connect(path);
    }

    private static OleDbConnection Connect(string path)
    {
        Exception? last = null;
        foreach (var provider in Providers)
        {
            var connection = new OleDbConnection($"Provider={provider};Data Source={path};");
            try
            {
                connection.Open();
                return connection;
            }
            catch (Exception ex)
            {
                last = ex;
                connection.Dispose();
            }
        }

        // The provider is a system-registered COM component and cannot be shipped with the
        // app, so report the bitness this build actually needs — an x86 build cannot load
        // the 64-bit provider, or vice versa.
        var bitness = Environment.Is64BitProcess ? "64-bit" : "32-bit";

        throw new InvalidOperationException(
            $"Could not open the data file. The {bitness} Microsoft Access Database Engine " +
            $"(ACE OLEDB) must be installed on this machine. This is the {bitness} build, so " +
            "it needs the provider that comes with " +
            $"{bitness} Office — or the matching Access Database Engine redistributable.", last);
    }

    /// <summary>Confirms the file carries the tables a Kelsoft data file is expected to have.</summary>
    public IReadOnlyList<string> MissingTables()
    {
        var expected = new[] { "ACCOUNTS", "ALLOCATIONS", "FINANCIAL_YEAR", "CLIENT_DETAILS" };
        var present = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        using var schema = _connection.GetOleDbSchemaTable(
            OleDbSchemaGuid.Tables, [null, null, null, "TABLE"]);
        if (schema is not null)
            foreach (DataRow row in schema.Rows)
                present.Add((string)row["TABLE_NAME"]);

        return [.. expected.Where(t => !present.Contains(t))];
    }

    public string ClientName()
    {
        using var command = new OleDbCommand("SELECT TOP 1 name FROM CLIENT_DETAILS", _connection);
        return command.ExecuteScalar() as string ?? "(client name not set)";
    }

    /// <summary>Financial years defined in the file, each with the count of allocations falling inside it.</summary>
    public IReadOnlyList<FinancialYear> FinancialYears()
    {
        var years = new List<(string Label, DateTime Start, DateTime End)>();

        using (var command = new OleDbCommand(
            "SELECT financial_year, start_date, end_date FROM FINANCIAL_YEAR ORDER BY financial_year",
            _connection))
        using (var reader = command.ExecuteReader())
        {
            while (reader.Read())
            {
                if (reader.IsDBNull(1) || reader.IsDBNull(2)) continue;
                years.Add((reader.GetString(0), reader.GetDateTime(1), reader.GetDateTime(2)));
            }
        }

        var result = new List<FinancialYear>();
        foreach (var (label, start, end) in years)
            result.Add(new FinancialYear(label, start, end, CountAllocations(start, end)));

        return result;
    }

    private int CountAllocations(DateTime start, DateTime end)
    {
        using var command = new OleDbCommand(
            "SELECT COUNT(*) FROM ALLOCATIONS WHERE allocation_date BETWEEN ? AND ?", _connection);
        command.Parameters.Add("start", OleDbType.Date).Value = start;
        command.Parameters.Add("end", OleDbType.Date).Value = end;
        return Convert.ToInt32(command.ExecuteScalar());
    }

    /// <summary>Accounts keyed by id, rounded to 2dp because account_id is stored as a Single.</summary>
    public IReadOnlyDictionary<double, Account> Accounts()
    {
        var accounts = new Dictionary<double, Account>();

        using var command = new OleDbCommand(
            "SELECT account_id, name, account_type FROM ACCOUNTS", _connection);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (reader.IsDBNull(0)) continue;
            var id = Math.Round(Convert.ToDouble(reader.GetValue(0)), 2);
            var name = reader.IsDBNull(1) ? $"Account {id}" : reader.GetString(1);
            var type = reader.IsDBNull(2) ? 0 : Convert.ToInt32(reader.GetValue(2));
            accounts[id] = new Account(id, name, type);
        }

        return accounts;
    }

    /// <summary>
    /// Debit and credit totals per account per calendar month across the period.
    /// Uses BETWEEN to match the original report exactly; allocation_date carries no time component.
    /// </summary>
    public IReadOnlyList<(double AccountId, int Year, int Month, decimal Debit, decimal Credit)>
        MonthlyMovements(DateTime start, DateTime end)
    {
        const string sql = """
            SELECT account_id,
                   Year(allocation_date)  AS yr,
                   Month(allocation_date) AS mth,
                   Sum(debit_amount)      AS dr,
                   Sum(credit_amount)     AS cr
            FROM ALLOCATIONS
            WHERE allocation_date BETWEEN ? AND ?
            GROUP BY account_id, Year(allocation_date), Month(allocation_date)
            """;

        using var command = new OleDbCommand(sql, _connection);
        command.Parameters.Add("start", OleDbType.Date).Value = start;
        command.Parameters.Add("end", OleDbType.Date).Value = end;

        var rows = new List<(double, int, int, decimal, decimal)>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (reader.IsDBNull(0)) continue;
            rows.Add((
                Math.Round(Convert.ToDouble(reader.GetValue(0)), 2),
                Convert.ToInt32(reader.GetValue(1)),
                Convert.ToInt32(reader.GetValue(2)),
                reader.IsDBNull(3) ? 0m : Convert.ToDecimal(reader.GetValue(3)),
                reader.IsDBNull(4) ? 0m : Convert.ToDecimal(reader.GetValue(4))));
        }

        return rows;
    }

    /// <summary>
    /// Opening balances for a financial year as a net amount per account (debit positive).
    /// The balance sheet reads these directly rather than deriving them from the prior year.
    /// </summary>
    public IReadOnlyDictionary<double, decimal> OpeningBalances(string financialYear)
    {
        const string sql = """
            SELECT account_id, Sum(debit_balance - credit_balance) AS net
            FROM OPENING_BALANCES
            WHERE financial_year = ?
            GROUP BY account_id
            """;

        using var command = new OleDbCommand(sql, _connection);
        command.Parameters.Add("year", OleDbType.VarWChar, 4).Value = financialYear;

        var balances = new Dictionary<double, decimal>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (reader.IsDBNull(0)) continue;
            balances[Math.Round(Convert.ToDouble(reader.GetValue(0)), 2)] =
                reader.IsDBNull(1) ? 0m : Convert.ToDecimal(reader.GetValue(1));
        }

        return balances;
    }

    /// <summary>
    /// GST accrued on allocation rows, per month, for the two accounts the balance sheet
    /// carries: 1000 (input credits, posted as a debit) and 1001 (collected, posted as a credit).
    /// These offset the BAS clearing journals posted directly to the same accounts.
    /// </summary>
    public IReadOnlyList<(int GstAccount, int Year, int Month, decimal Amount)>
        MonthlyGst(DateTime start, DateTime end)
    {
        const string sql = """
            SELECT gstaccount,
                   Year(allocation_date)  AS yr,
                   Month(allocation_date) AS mth,
                   Sum(Round(gstamount, 2)) AS amt
            FROM ALLOCATIONS
            WHERE allocation_date BETWEEN ? AND ?
              AND gstaccount IN (1000, 1001)
            GROUP BY gstaccount, Year(allocation_date), Month(allocation_date)
            """;

        using var command = new OleDbCommand(sql, _connection);
        command.Parameters.Add("start", OleDbType.Date).Value = start;
        command.Parameters.Add("end", OleDbType.Date).Value = end;

        var rows = new List<(int, int, int, decimal)>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (reader.IsDBNull(0)) continue;
            rows.Add((
                Convert.ToInt32(reader.GetValue(0)),
                Convert.ToInt32(reader.GetValue(1)),
                Convert.ToInt32(reader.GetValue(2)),
                reader.IsDBNull(3) ? 0m : Convert.ToDecimal(reader.GetValue(3))));
        }

        return rows;
    }

    public void Dispose() => _connection.Dispose();
}

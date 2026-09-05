# Kelsoft — Monthly Report Export

A Windows app (.NET 10, WPF) that reads a Kelsoft **client data file** — the back-end `.mdb`
the Kelsoft front-end links to — and writes an Excel workbook with **a worksheet per report
per financial year**, each broken down **by month**.

The Access reports produce a single column for a chosen date range. This produces the same
statements with a column per month.

| Report | Sheet | Columns |
|---|---|---|
| Profit and Loss Statement | `FY2025 P&L` | 12 months + Total |
| Balance Sheet | `FY2025 BS` | 12 month-end positions (no total — see below) |

## Running it

```
KelsoftReportExport\publish\KelsoftReportExport.exe ["path\to\data.mdb"]
```

1. **Browse…** to a client data file (`.mdb` or `.accdb`), or pass one as an argument.
2. Tick which reports you want, and which financial years. Years with no allocations are
   listed but unticked.
3. Choose where to save the `.xlsx`.
4. **Export.**

### Requirements

- Windows, .NET 10 desktop runtime.
- The **64-bit** Microsoft Access Database Engine (ACE OLEDB). It ships with 64-bit Office;
  otherwise install the Access Database Engine redistributable. The app is built x64 for
  this reason — a 32-bit build cannot load the 64-bit provider.

WPF was chosen over WinForms deliberately: it lays out in device-independent units, so the
window is correct on high-DPI displays without any per-control scaling work.

## Where the figures come from

Read from the data file directly — the front-end `.accdb` is not involved.

| Input | Source |
|---|---|
| Account names and types | `ACCOUNTS` |
| Movements | `ALLOCATIONS`, by `allocation_date` |
| Opening balances | `OPENING_BALANCES`, for the year being reported |
| GST accruals | `ALLOCATIONS.gstamount` / `gstaccount` |
| Year boundaries | `FINANCIAL_YEAR` |
| Heading | `CLIENT_DETAILS.name` |

`ACCOUNT_TYPES` lives in the front-end rather than the data file, so that 13-row map is
carried in `Model.cs`.

### Profit and Loss

Net movement is signed by account type, mirroring the `IIf` chain in
`qryCreateCombinedTransactionsAll`:

| Account type | Id | Net |
|---|---|---|
| Sales | 9 | credit − debit |
| Cost of Sales | 10 | debit − credit |
| Other Income | 7 | credit − debit |
| Other Expenses | 6 | debit − credit |

Then the report's own formulas:

- `Gross Profit = Total Sales − Total Cost of Sales`
- `Total Income = Gross Profit + Other Income`
- `Net Profit = Gross Profit + Other Income − Other Expenses`

Sections follow the report order: TRADING INCOME → less Cost of Sales → GROSS PROFIT/LOSS →
Income → TOTAL INCOME → Expenditure → NET PROFIT/LOSS.

### Balance Sheet

The report's `qryBsTransactions` is an eleven-way `UNION ALL` because it separates a report
period from the earlier part of the financial year, and treats bank (posting) accounts by
`transaction_date` rather than `allocation_date` in that earlier part.

A monthly balance sheet always runs from the first day of the financial year, so every
"earlier part of the year" branch is empty and the union reduces to four things:

1. `OPENING_BALANCES` for the year,
2. movements on balance sheet accounts (types 1, 2, 3, 4, 5, 12, 13, 15) to that month end,
3. the year-to-date trading result, posted to account **9998 Net Profit** or **9999 Net Loss**,
4. GST accrued on allocations — account **1000** as a debit, **1001** as a credit.

Point 4 is not optional. GST is carried on the allocation row (`gstamount`), not as its own
posting, so the allocations alone are short by exactly the GST. The BAS clearing journals
post directly to 1000/1001 and offset the accrual, which is why both accounts net to zero at
year end.

Presentation: assets on a debit basis, liabilities and capital on a credit basis. Layout
follows the report — OWNERS EQUITY (Capital + Other Capital + Profit/Loss), then
"represented by" with Assets, Liabilities and NET ASSETS.

**There is no total column.** Balances are cumulative, so summing the months is meaningless;
the last column is the year-end position.

### Differences from the Access reports

- **Zero-row suppression.** The reports drop an account when `round(Net,2) = 0` for the
  period. Here a row is kept if **any month** is non-zero, so an account that nets to zero
  over the year still shows the movement behind it.
- **Amounts are GST-exclusive**, using `debit_amount` / `credit_amount` — the same fields the
  reports use, not the `gross_*` columns.
- **`gstaccount` 1002** (GST Adjustments) is ignored, exactly as the report ignores it. In the
  sample file it sums to zero.

### Year-end journals

Journals stay in the month they were posted, so the columns tie to the annual statement.
June therefore carries most year-end adjustments and looks heavy. General journals are
identifiable as `ALLOCATIONS.transaction_id IS NULL`, so splitting them into their own
column later is a contained change to `StatementBuilder`.

## When the balance sheet does not balance

Every entry balances, so Net Assets must equal Total Owners Equity in every month. The app
checks this per month and warns, naming the worst month, but still writes the workbook —
a failure points at the data, not the export. The usual causes are an unbalanced general
journal, or bank entries whose allocation dates straddle a month end.

In the sample file, 83 of 84 months balance at exactly 0.00. August 2024 is out by
86,970.75 and corrects itself in September:

| Source | Aug 2024 | Sep 2024 |
|---|---|---|
| A bank (posting) account | +92,669.01 | −98,487.02 |
| One payment, legs split across the month end | −4,752.00 | +4,320.00 |
| General journals | −46,878.41 | 0.00 |

## Code layout

| File | Role |
|---|---|
| `Model.cs` | Account-type map, financial year, P&L statement types and formulas |
| `BalanceSheetModel.cs` | Balance sheet types, totals and the balance check |
| `KelsoftDataFile.cs` | OLEDB reads; provider fallback ACE 16 → 12 |
| `StatementBuilder.cs` | Monthly movements → Profit and Loss |
| `BalanceSheetBuilder.cs` | Monthly movements + opening balances → Balance Sheet |
| `ExcelExporter.cs` | Worksheet writing (ClosedXML) |
| `App.xaml`, `MainWindow.xaml(.cs)` | The window |

## Verifying a change

The reference data file should produce, for FY2025:

| Sales | Cost of Sales | Other Income | Other Expenses | Gross Profit | Net Profit |
|---|---|---|---|---|---|
| 929,179.68 | 390,435.68 | 1,714,064.58 | 2,085,669.29 | 538,744.00 | 167,139.29 |

and a balance sheet at 30/06/2025 of Total Assets 5,844,276.86, Total Liabilities
5,764,732.86, Net Assets **79,544.00** = Total Owners Equity.

Two invariants worth keeping in any test:

- Net Assets equals Total Owners Equity in every month (except the August 2024 data issue).
- The balance sheet's Profit/Loss line equals the cumulative P&L Net Profit, month for month.
  This ties the two reports to each other and holds in all seven years.

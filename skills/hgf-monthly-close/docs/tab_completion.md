# HGF Consolidated — Tab Completion Instructions

Per-tab procedures for completing the consolidated workbook. Each section below captures one tab: what data it needs, where that data comes from, and the step-by-step actions to populate it. Follow these in order each month.

> Status: instructions are being captured live. Tabs not yet documented here will be added as we work through them.

---

## Workbook target

`HGF_CONSOLIDATED_<MONTH>_<YEAR>_GENERATED.xlsx` (in the workspace folder).

All 33 tabs should be **visible** (unhidden) before populating. If they're hidden, run:

```python
import openpyxl
wb = openpyxl.load_workbook(path)
for s in wb.sheetnames:
    wb[s].sheet_state = "visible"
wb.save(path)
```

## Tabs in this workbook

| # | Tab name | Status |
|---|---|---|
| 1 | READ ME | not yet documented |
| 2 | Links | not yet documented |
| 3 | Returns By Channel | not yet documented |
| 4 | Accts Rec | not yet documented |
| 5 | Accts Pay | not yet documented |
| 6 | Debt Schedule | not yet documented |
| 7 | Balance Sheet | not yet documented |
| 8 | Consulting Breakdown | not yet documented |
| 9 | Revenue Per Employee | **documented** (see below) |
| 10 | High Level Recap | not yet documented |
| 11 | Department Insights | not yet documented |
| 12 | Dept Profit Ratio | not yet documented |
| 13 | Chargeback Recap | **documented** (see below) |
| 14 | Chargeback Detail | not yet documented |
| 15 | Trailing 12 Months | **documented** (see below) |
| 16 | Mkt Spnd vs Rev | **documented** (see below) |
| 17 | SGA and NP vs Rev | **documented** (see below) |
| 18 | DTC Payment Method | not yet documented |
| 19 | Period FULL sheet | populated by writer (see Adjustments Playbook §3) |
| 20 | Period SIMPLE sheet | populated by writer (see Adjustments Playbook §3) |
| 21 | YTD | **documented** (see below) |
| 22 | YTD By Department | **documented** (see below) |
| 23 | Monthly YoY | **documented** (see below) |
| 24 | Online MASS | not yet documented |
| 25 | OG DTC | not yet documented |
| 26 | APA | not yet documented |
| 27 | Trend House | not yet documented |
| 28 | RAW DATA_Master File | populated by writer |
| 29 | RAW DATA_COGS & Freight | populated by writer |
| 30 | RAW DATA_Payroll | populated by writer |
| 31 | Online Returns Accrual v Actual | **documented** (see below) |
| 32 | DPCache_Shopify (10.1.25-12.31.) | not yet documented |
| 33 | DPCache_Shopify (7.1.25-9.30.25) | not yet documented |

---

## Tab procedures

<!-- Each tab will be documented in this section as we work through them. Format:

### <Tab name>

**Data needed:**
- ...

**Sources:**
- ...

**Steps:**
1. ...

**Validation:**
- ...
-->

### Revenue Per Employee

This tab tracks headcount and revenue per employee, broken into three year-blocks (2024 in rows 1–9, 2025 in rows 12–20, 2026 in rows 24–32). Each block has the same shape: P/R Employees row, Remote Team row, Total row (formula), Revenue row, Revenue per Employee row (formula), with monthly columns B–M (Jan–Dec) and Average / Total columns N / O.

For the current year only the three "input" rows need monthly entries each month:

- **P/R Employees** (row 27 in 2026): the count of employees with a positive gross-pay amount in the Payroll Journal for the month. For March 2026 = **43**.
- **Remote Team** (row 28 in 2026): the count of remote team members (per the user-supplied roster maintained in rows 70–85 of this tab). For March 2026 = **19**.
- **Revenue** (row 31 in 2026): **gross revenue** for the month — pull from the SIMPLE tab row 8 (Sales) column U TOTAL$. NOT row 10 Net Sales. For March 2026 = **$1,950,439.45**.

The Total row (29) and Revenue per Employee row (32) are formulas that auto-derive once the input rows are filled.

**Steps each month:**

1. Open the Payroll Journal for the period. Count the rows on the `Payroll` sheet that have a name in column B and a positive amount in column C (exclude blanks and total rows). The base skill's payroll extractor already produces this count — `extractor_outputs/payroll_journal/employee_count.json` if present, otherwise count manually.
2. Get the Remote Team count from the user (or from the Remote Team -<YEAR> roster maintained in rows 70+ of this same tab).
3. Open the consolidated workbook, go to the `Revenue Per Employee` tab.
4. In the current year-block, write the count into the P/R Employees row at the right month column (e.g. for March 2026 → cell **D27**), and the Remote Team count into the Remote Team row (e.g. **D28**).
5. **Update the column-N "Average" formulas** for the current year so the divisor matches the count of months populated. Example: when populating March (the third month), N27 / N28 / N31 each go from `=SUM(...)/2` to `=SUM(...)/3`. April will be `/4`, etc. Once December is populated the formulas should be `/12`, matching the prior-year blocks.
6. Recalculate the workbook with LibreOffice headless so the totals refresh.

**March 2026 actuals applied:**

| Cell | Value | Note |
|---|---|---|
| D27 | 43 | Payroll Journal employees with positive gross pay |
| D28 | 19 | Remote Team count (per user instruction) |
| D31 | 1,950,439.45 | Gross Sales from SIMPLE tab row 8 column U |
| N27 | `=SUM(B27:M27)/3` | was `/2` — bumped to `/3` for three months populated |
| N28 | `=SUM(B28:M28)/3` | was `/2` |
| N31 | `=SUM(B31:M31)/3` | was `/2` |

**Totals row formulas — column N (Average) and column O (TOTAL):**

The base year-block layout already includes these and they derive correctly once N27/N28/N31 use the right divisor. Do not need to be edited manually each month:

- D29, etc. = `=+D27+D28` — monthly total employees
- N29 = `=+N27+N28` — average total employees YTD
- N32 = `=+N31/N29` — average revenue / average employees
- O31 = `=SUM(B31:M31)` — YTD revenue (unfilled months are 0, so this is YTD until December)
- O32 = `=+O31/N29` — YTD revenue / average employees

**Validation after recalc (March 2026):**

- D29 (Total Mar) = 43 + 19 = 62 ✓
- D32 (Rev/Employee Mar) = $1,950,439.45 / 62 = **$31,458.70**
- N27 (Avg P/R YTD) = (43 + 43 + 43) / 3 = 43.00 ✓
- N28 (Avg Remote YTD) = (17 + 19 + 19) / 3 = 18.33 ✓
- N29 (Avg Total YTD) = 43 + 18.33 = 61.33 ✓
- N31 (Avg Revenue YTD) = (2,714,637 + 1,512,526.94 + 1,950,439.45) / 3 = **$2,059,201.13**
- O31 (YTD Revenue) = sum of three months = **$6,177,603.39**
- O32 (YTD Rev / Avg Employee) = $6,177,603.39 / 61.33 = **$100,721.79**
- Sanity: 2025 Mar Rev/Employee was $29,944.54 with 53 total employees; 2026 Mar is $31,458.70 with 62 — both look reasonable.

### Chargeback Recap

This tab is one continuous monthly chargeback log running from 2022 through 2026, in stacked year-blocks (2023 starts row 1, 2024 starts row 27, 2025 starts row 40, 2026 starts row 53). Each year-block has the same five "input" columns plus auto-derived percentages and totals.

**Year-block layout (row offsets within each block — 2026 example shown):**

| Col | Header | Type | 2026-March cell |
|---|---|---|---|
| A | Period label (`YYYY \| N (Month)`) | text | A56 |
| B | Contractual Allowance | input | B56 |
| C | % of Total | formula `=B/L` | C56 |
| D | Penalties | input | D56 |
| E | % of Total | formula `=D/L` | E56 |
| F | Amazon Holdback Provision | input | F56 |
| G | % of Total | formula `=F/L` | G56 |
| H | Returns | input | H56 |
| I | % of Total | formula `=H/L` | I56 |
| J | Software Fees | input | J56 |
| K | % of Total | formula `=J/L` | K56 |
| L | Total Chargebacks | formula `=B+D+F+H+J` | L56 |
| M | Monthly Sales | input | M56 |
| N | % of Sales | formula `=L/M` | N56 |

**Steps each month:**

1. Open the OG chargeback PDF for the period (`- OG _ Chargeback Report - <MM>. <MONTH> <YEAR>.pdf` in the workspace folder).
2. Find the per-month summary table on the first or second page. The line for the current month looks like:
   `<YEAR> | <MM>. (<Month>) -$<Allowance> XX.X% -$<Penalty> X.X% -$<Provision> X.X% -$<Returns> XX.X% -$<Software> X.X% -$<Total>`
   The base skill's chargeback extractor parses this into `extractor_outputs/chargeback_pdf/chargeback_monthly_summary.csv`. Filter to the current `(year, month_num)` and pull the five `amount` rows for categories `allowance`, `penalty`, `provision`, `return`, `software_fees` — those map 1-to-1 onto columns B / D / F / H / J.
3. Get **Monthly Sales** (column M) — this is the **gross** sales figure for the month, same number you put in the Revenue Per Employee tab row 31 column for the current month, sourced from the SIMPLE tab row 8 column U.
4. Write the six values into the current month's row in the 2026 block:
   - `B<row>` Contractual Allowance (negative)
   - `D<row>` Penalties (negative)
   - `F<row>` Amazon Holdback Provision (negative)
   - `H<row>` Returns (negative)
   - `J<row>` Software Fees (negative)
   - `M<row>` Monthly Sales (positive — gross revenue)
5. Recalculate. The L (Total Chargebacks), C/E/G/I/K (% of Total), and N (% of Sales) cells auto-derive.

**Allowance vs Returns — historical reclassification note:**

In Jan and Feb 2026, the prior values in the workbook do NOT match the PDF figures cleanly — about $15K–$20K of Allowance was reclassified into Returns (totals match, but the split was changed). This was a manual accountant adjustment, not a documented standing rule. **Default behavior going forward (per user direction May 2026):** apply PDF figures as-is. If a reclassification needs to happen for a specific month, do it manually with a note in the Chat Archive.

**March 2026 actuals applied (row 56):**

| Cell | Value | Source |
|---|---|---|
| B56 Contractual Allowance | −65,092 | PDF allowance |
| D56 Penalties | −7,821 | PDF penalty |
| F56 Amazon Holdback Provision | −15,661 | PDF provision |
| H56 Returns | −15,596 | PDF return |
| J56 Software Fees | −34 | PDF software_fees |
| M56 Monthly Sales | 1,950,439.45 | SIMPLE row 8 col U gross sales |

**Validation after recalc:**

- L56 Total Chargebacks = −104,204 (PDF total = −104,205, $1 rounding) ✓
- C56 % Allowance = 62.5% ✓ (matches PDF)
- E56 % Penalties = 7.5% ✓
- G56 % Provision = 15.0% ✓
- I56 % Returns = 15.0% ✓
- N56 % of Sales = −5.3%
- Sanity: prior month L55 (Feb) = −123,522 with N55 = −8.2%. March is lower in dollars and % — reasonable given March's higher gross sales and lower returns.

### Trailing 12 Months

This tab maintains a rolling 12-month P&L view. Each month occupies two columns: a value column and a percentage column (e.g., G/H = April 2025, I/J = May 2025, … AC/AD = March 2026). The TTM TOTAL column AE sums all 12 month-value columns.

**Layout for the 2026 close** (positions of the value columns; the column to the right of each is the % column, formula-driven):

| Column | Period |
|---|---|
| G | April 2025 |
| I | May 2025 |
| K | June 2025 |
| M | July 2025 |
| O | August 2025 |
| Q | September 2025 |
| S | October 2025 |
| U | November 2025 |
| W | December 2025 |
| Y | January 2026 |
| AA | February 2026 |
| **AC** | **March 2026** ← current period |
| AE | TTM TOTAL (formula `=G+I+K+M+…+AC`) |

Headers are in row 7. Row numbering matches the SIMPLE tab exactly (row 8 = Sales, row 9 = Returns, …, row 77 = Operating Income).

**Steps each month:**

1. Identify the column for the current period in row 7 (look for `<Month> <YEAR>`).
2. For every row from 8 through 77, copy the value in **`MARCH 2026_SIMPLE!U<row>`** (column U is "TOTAL$") into **`Trailing 12 Months!<period_col><row>`**. This is a row-for-row paste of static numeric values, not formulas. Skip rows where SIMPLE column U is blank or non-numeric.
3. Each month, the prior month's column needs no edits — only the current period's column gets populated.
4. Recalculate. The percentage column to the right of the current period (e.g., AD for March 2026) auto-derives, and the TTM TOTAL column AE refreshes.
5. **Each new month, before populating, shift all month-blocks one position to the left** so the tab always shows the trailing 12 months ending at the current period. The oldest month ages out, the newest month gets the rightmost slot.

   **Shift rule:** for each pair of adjacent month-blocks, copy the right block's value column over the left block's value column. Move row 7 headers the same way. The percentage columns (H, J, L, …, AD) are formulas of the form `=<value_col><row>/<value_col><row>` and update automatically once values are in place — no need to edit them. After shifting, the rightmost value column (AC) is empty, ready for the new month's data.

   **Concrete column moves for the April 2026 close** (run this *before* copying SIMPLE → AC):

   | Step | From | To | Period that lands in "To" |
   |---|---|---|---|
   | 1 | I rows 7–77 | G rows 7–77 | May 2025 |
   | 2 | K | I | June 2025 |
   | 3 | M | K | July 2025 |
   | 4 | O | M | August 2025 |
   | 5 | Q | O | September 2025 |
   | 6 | S | Q | October 2025 |
   | 7 | U | S | November 2025 |
   | 8 | W | U | December 2025 |
   | 9 | Y | W | January 2026 |
   | 10 | AA | Y | February 2026 |
   | 11 | AC | AA | March 2026 |
   | 12 | (clear AC) | AC | (becomes April 2026 slot) |

   April 2025 (the original column G data) gets overwritten in step 1 — that's the month aging out of the trailing window.

   Then update row 7 headers: G7=`May 2025`, I7=`June 2025`, …, AA7=`March 2026`, AC7=`April 2026`. Then proceed with step 2 of the monthly procedure (copy SIMPLE column U → AC).

   **Reference shift script** (paste into a python tool call against the workbook before running the April close):

   ```python
   import openpyxl
   wb = openpyxl.load_workbook(path)
   ttm = wb['Trailing 12 Months']
   value_cols = ['G','I','K','M','O','Q','S','U','W','Y','AA','AC']
   # shift each row 7..77 one block to the left
   for r in range(7, 78):
       for i in range(len(value_cols)-1):
           src = value_cols[i+1]
           dst = value_cols[i]
           ttm[f'{dst}{r}'] = ttm[f'{src}{r}'].value
       # clear the rightmost (will be overwritten with new month's SIMPLE U)
       ttm[f'{value_cols[-1]}{r}'] = 0 if r != 7 else None
   # set the new rightmost header
   ttm[f'{value_cols[-1]}7'] = '<New Month> <YEAR>'   # e.g. 'April 2026'
   wb.save(path)
   ```

   **Validation after shift, before populating:** spot-check that AA7 reads the just-completed month's label (e.g. `March 2026` for the April close) and AA8/AA22/AA77 hold March's values. AE TTM TOTAL will temporarily be wrong until SIMPLE U is copied into AC.

### Mkt Spnd vs Rev

A small trailing-12-months tab tracking marketing spend as a percentage of consumer-facing gross revenue. The title in cell A1 is the explicit definition: "Marketing Spend as a % of Gross Revenue Net of Trend House."

Two input rows, one formula row:

| Cell | Field | Source |
|---|---|---|
| Row 5 | Gross Revenue (net of Trend House) | SIMPLE row 8 column U (TOTAL) **minus** SIMPLE row 8 column G (Trend House channel) |
| Row 6 | Marketing | SIMPLE row 37 column U (TOTAL Marketing across all channels) |
| Row 7 | % | formula `=row6/row5` — auto-derives |

**Important:** the gross revenue figure here is **NOT** the total gross sales (row 8 col U). It is total gross sales **minus** Trend House sales (row 8 col G). Trend House is the wholesale-to-retailers channel and doesn't drive consumer marketing spend, so this metric isolates consumer-facing revenue. (Per user direction May 2026.)

The Marketing row uses total marketing across all channels (SIMPLE U37) — TH marketing is left in (it's typically negligible; March 2026 TH marketing was $427.69 of $111,146.39).

**Steps each month:**

1. Open the SIMPLE tab. Read row 8: column U (total Sales) and column G (Trend House Sales).
2. Compute Gross Revenue net of TH = U8 − G8.
3. Read row 37 column U for total Marketing.
4. Find the column for the current period in row 4 of `Mkt Spnd vs Rev` (e.g., March 2026 = column M).
5. Write the net-of-TH revenue into row 5, marketing into row 6.
6. Recalculate. Row 7 (%), N5/N6 (Total), N7 (% of total), O5/O6 (Monthly Avg), O7 (% of avg) all auto-derive.

**Column shift each new month:** same as Trailing 12 Months — shift each column one position to the left (C→B, D→C, …, M→L) before populating, drop the oldest, add the new month at column M. Headers in row 4 shift the same way. The percentage formulas in row 7 reference adjacent rows so they continue to work after the shift.

**March 2026 actuals applied:**

| Cell | Value | Calc |
|---|---|---|
| M5 Gross Revenue (net TH) | $708,046.05 | 1,950,439.45 − 1,242,393.40 |
| M6 Marketing | $111,146.39 | SIMPLE U37 |
| M7 % | 15.70% | formula |

**Validation:**

- 12-month total: Gross Rev $8.50M / Marketing $1.37M / 16.11% blended ratio
- Monthly trend has been 12–20%; March's 15.70% is right in the middle — no flag
- Lowest months historically were October 2025 (12.29%) and May 2025 (13.32%); highest November 2025 (19.77%)

### SGA and NP vs Rev

A small trailing-12-months tab showing G&A expenses and Net Income each as a percentage of total Gross Revenue. Title in A1: "G&A Expenses and Net Income(Loss) as a % of Total Gross Revenue."

Three input rows, two formula rows (% rows are derived):

| Cell | Field | Source |
|---|---|---|
| Row 5 | Gross Revenue | SIMPLE row 8 column U (TOTAL Sales — *includes* Trend House, unlike the Mkt Spnd vs Rev tab) |
| Row 6 | G&A Expenses | SIMPLE row 38 column U + SIMPLE row 75 column U (Total Controllable Expenses + Total Expenses (General Corporate)) |
| Row 7 | Net Income | SIMPLE row 77 column U (Operating Income) — the workbook treats Operating Income as Net Income; there is no row below 77 for taxes / other items in the consolidated SIMPLE view |
| Row 8 | G&A Percentage | formula `=row6/row5` |
| Row 9 | Net Income Percentage | formula `=row7/row5` |

**Steps each month:**

1. Open the SIMPLE tab. Read U8 (gross sales total), U38 (Total Controllable Expenses), U75 (Total Expenses — General Corporate), and U77 (Operating Income).
2. Compute G&A = U38 + U75.
3. Find the column for the current period in row 4 of the `SGA and NP vs Rev` tab.
4. Plug Gross Revenue → row 5, G&A → row 6, Net Income → row 7.
5. Recalculate. Rows 8 and 9 (% formulas) and N5/N6/N7 (Total) and O5/O6/O7 (Monthly Avg) auto-derive.

**Column shift each new month:** same pattern as Trailing 12 Months and Mkt Spnd vs Rev — shift each value column one position to the left before populating, drop the oldest, add the new month at column M. Headers in row 4 shift the same way.

**March 2026 actuals applied:**

| Cell | Value | Calc |
|---|---|---|
| M5 Gross Revenue | $1,950,439.45 | SIMPLE U8 |
| M6 G&A Expenses | $609,959.65 | $256,060.51 + $353,899.14 |
| M7 Net Income | $227,075.41 | SIMPLE U77 (Operating Income) |
| M8 G&A % | 31.27% | formula |
| M9 Net Income % | 11.64% | formula |

**Validation:**

- 12-month blended G&A ratio: 29.01% (range 20.26% July → 40.37% November)
- 12-month blended Net Income ratio: 16.07% (range 4.58% May → 23.78% July)
- March 2026's G&A 31.27% is on the high side of the range but in line with Feb's 39.36% — reasonable for a consumer-revenue-heavy month
- March's Net Income 11.64% is right around the median; well above Feb's 6.27%

### YTD

This tab is the year-to-date P&L: each month of the current calendar year occupies one column, plus a YTD column (S) with formulas that sum across all 12 monthly columns, and a % column (T). Headers are in row 7. Row numbering matches the SIMPLE tab (rows 8–77), with extra rows 78–82 for Non Recurring Expenses / Net Business Income that don't exist on SIMPLE.

**Layout (current = March 2026):**

| Column | Period |
|---|---|
| G | January 2026 |
| H | February 2026 |
| **I** | **March 2026** ← current period |
| J | April 2026 |
| K | May 2026 |
| L–R | June 2026 – December 2026 |
| S | YTD (formula `=SUM(G:R)`) |
| T | % of Sales (formula `=S<row>/S8`) |

Past months hold static values. Summary rows (10 Net Sales, 22 Total COGS, 24 Gross Profit, 38 Total Controllable, 40 Contribution Margin, 75 Total General Corp, 77 Operating Income) are also stored as static numbers in the monthly columns (they do *not* sum from their component rows — only the YTD column S derives via formula).

**Steps each month:**

1. Open the SIMPLE tab. For every row from 8 through 77 where SIMPLE column U has a numeric value, copy that value into the YTD tab's current-period column at the same row.
   - Includes the summary rows (10, 22, 24, 38, 40, 75, 77) — their YTD-column formula picks up the per-month static values.
2. Open the Addbacks PDF for the period. The total is on the cover ("$23,195 for March 2026"). Plug that total into the **`Operating Expenses`** row 80 of the YTD tab at the current-period column. *(The "Non Recurring Expenses" label is the section header on row 79; row 80 is where the dollar value goes.)*
3. Recalculate. The following auto-derive:
   - Row 82 Net Business Income (=row77+row78+row80 monthly; YTD same)
   - Column S YTD totals
   - Column T % of Sales for both monthly and YTD

**Column does not shift each month** — this tab spans the full calendar year (Jan–Dec) and grows through it, so for the April close you simply populate column J, etc. At the start of a new calendar year, copy the YTD tab and reset all monthly value columns to 0; the headers in row 7 advance by 12 months.

**March 2026 actuals applied:**

| Cell | Value | Source |
|---|---|---|
| I8 Sales | $1,950,439.45 | SIMPLE U8 |
| I9 Returns | −$95,520.73 | SIMPLE U9 |
| I10 Net Sales | $1,854,918.73 | SIMPLE U10 |
| I22 Total COGS | $1,017,883.67 | SIMPLE U22 |
| I24 Gross Profit | $837,035.06 | SIMPLE U24 |
| I38 Total Controllable | $256,060.51 | SIMPLE U38 |
| I40 Contribution Margin | $580,974.55 | SIMPLE U40 |
| I75 Total General Corp | $353,899.14 | SIMPLE U75 |
| I77 Operating Income | $227,075.41 | SIMPLE U77 |
| **I80 Operating Expenses (Addbacks)** | **$23,195** | **March Addbacks PDF** |
| I82 Net Business Income | $250,270.41 | formula |

61 rows total copied from SIMPLE U. Plus row 80 manual.

**Validation after recalc:**

- I82 Net Business Income = I77 + I78 + I80 = 227,075 + 0 + 23,195 = $250,270 ✓
- YTD column S totals: Sales $6,177,604, Operating Income $837,950, Net Business Income $883,549 — three months populated
- Column T (% of Sales): row 22 (COGS) = 53.24% YTD; row 75 (General Corp) = 17.90% YTD; row 77 (Operating Income) = 14.23% YTD
- Sanity: Jan addbacks were $14,702, Feb $7,703, March $23,195 — March is highest of the three but in a normal range

### Monthly YoY

This tab places each month of the current year side-by-side with the same month of the prior year, showing dollar diff and % of revenue for each. Each month occupies a 6-column "block": current-year value, current-year %, prior-year value, prior-year %, Diff, plus a spacer.

**Block layout for the 12 months of the year:**

| Month | Cur-yr value | Cur-yr % | Prior-yr value | Prior-yr % | Diff | Spacer |
|---|---|---|---|---|---|---|
| Jan | G | H | I | J | K | L |
| Feb | M | N | O | P | Q | R |
| **Mar** | **S** | **T** | **U** | **V** | **W** | X |
| Apr | Y | Z | AA | AB | AC | AD |
| May | AE | AF | AG | AH | AI | AJ |
| Jun | AK | AL | AM | AN | AO | AP |
| Jul | AQ | AR | AS | AT | AU | AV |
| Aug | AW | AX | AY | AZ | BA | BB |
| Sep | BC | BD | BE | BF | BG | BH |
| Oct | BI | BJ | BK | BL | BM | BN |
| Nov | BO | BP | BQ | BR | BS | BT |
| Dec | BU | BV | BW | BX | BY | BZ |

Headers go in row 7. Row numbering matches SIMPLE / FULL (rows 8–77). All values are **static** (not formulas) — the tab needs to be repopulated each month.

**Hidden columns:** the tab ships with all not-yet-populated month-blocks hidden. Unhide them all before populating so the layout is visible. (For the March 2026 close, 29 columns were hidden, all in the April→December blocks plus the rightmost columns of August.)

**Steps each month:**

1. **Unhide all hidden columns** on the Monthly YoY tab so the entire month grid is visible:

   ```python
   import openpyxl
   wb = openpyxl.load_workbook(path)
   ws = wb['Monthly YoY']
   for col, dim in ws.column_dimensions.items():
       if dim.hidden:
           dim.hidden = False
   wb.save(path)
   ```

2. Identify the 6-column block for the current month (e.g., March = S/T/U/V/W/X). Set the block headers in row 7:
   - Cur-year value column header (e.g., `S7`) = `<MONTH> <CURRENT_YEAR>` (e.g., `MARCH 2026`)
   - Prior-year value column header (e.g., `U7`) = `<MONTH> <PRIOR_YEAR>` (e.g., `MARCH 2025`)

3. Open the **`<MONTH> <YEAR> FULL`** tab. The two columns you need are:
   - `DQ` = TOTAL - ACTUAL (current month, all channels combined)
   - `DS` = TOTAL PRIOR YEAR (same row labels, prior year)

4. For each row from 8 through 77 with numeric values in DQ or DS, write five static values into the current month's block:

   | YoY column | Value |
   |---|---|
   | `S<r>` | `FULL!DQ<r>` (current-year value) |
   | `U<r>` | `FULL!DS<r>` (prior-year value) |
   | `T<r>` | `S<r> / S8` (current-year % of Sales) |
   | `V<r>` | `U<r> / U8` (prior-year % of Sales) |
   | `W<r>` | `S<r> - U<r>` (Diff) |

   For row 8 (Sales), `T8` and `V8` both = 1 (100% of itself).

5. No recalculation needed — every value is static. Save and done.

**Column structure does not shift each month** — this tab spans the full calendar year (Jan–Dec). Each month, populate the next block; existing earlier-month blocks stay as-is. New calendar year requires a fresh copy with all blocks reset and headers advanced.

**March 2026 actuals applied** (60 rows populated in S/T/U/V/W; rows 18–19 / 30 / 32 / 66 / 72 are zeros):

| Row | Label | Mar 2026 (S) | Mar 2025 (U) | Diff (W) | YoY % |
|---|---|---:|---:|---:|---:|
| 8 | Sales | 1,950,439.45 | 1,567,337.66 | +383,101.79 | +24.4% |
| 22 | Total COGS | 1,017,883.67 | 776,321.96 | +241,561.71 | +31.1% |
| 24 | Gross Profit | 837,035.06 | 702,868.70 | +134,166.36 | +19.1% |
| 38 | Total Controllable | 256,060.51 | 184,968.40 | +71,092.11 | +38.4% |
| 40 | Contribution Margin | 580,974.55 | 517,900.29 | +63,074.26 | +12.2% |
| 75 | Total General Corp | 353,899.14 | 252,424.74 | +101,474.41 | +40.2% |
| 77 | Operating Income | 227,075.41 | 265,475.56 | −38,400.15 | −14.5% |

**Validation:**

- Sales grew 24%, but COGS grew 31% (margin compressed by 266 bps) and General Corporate grew 40% — both pressures
- Net effect: Operating Income declined $38K YoY despite revenue growth — worth flagging in the close report
- Largest dollar movers: Finished Goods +$198K, Tariffs +$50K, Marketing +$42K, Travel - GC +$17K, Repairs −$17K, Shipping −$18K

### YTD By Department

This tab is a year-to-date P&L broken out by sales channel. It's a one-period snapshot (not month-by-month) — every close, the prior month's data gets overwritten with the new YTD-through-current-month figures.

The data flows in **two stages**:

1. **Per-month channel breakdown** lives in a separate workbook: `YTD By Dept_Data_2026.xlsx` (in the workspace folder). That file has 12 monthly tabs (JAN, FEB, MAR, …, DEC), a YTD 2026 tab that sums them via formulas, and Q1/Q2/Q3/Q4 tabs.
2. The consolidated workbook's **YTD By Department** tab is populated by reading the **YTD 2026** sheet from that data file.

**Channel layout in the data file's monthly sheets** (e.g. MAR sheet):

| Col | Channel |
|---|---|
| G | Trend House (Brick & Mortar - China) |
| H | OG Specialty (RH, ZG, DreamMaker) |
| I | Online Lux |
| J | Online Mass / Online |
| K | OG-DTC |
| L | All Pop Art |
| M | Ink |
| N | TOTAL (formula `=SUM(G:M)`) |

**Channel layout in the data file's YTD 2026 sheet** (the consolidation sheet):

| Col | Channel | Mapped from monthly col |
|---|---|---|
| G | Trend House | G |
| I | OG Specialty | H |
| K | Online Lux | I |
| M | Online Mass | J |
| O | OG-DTC | K |
| Q | All Pop Art | L |
| S | Ink | M |
| U | TOTAL | sum of all 7 |

(The H/J/L/N/P/R/T/V columns are `% of channel sales` formulas.)

**Channel layout in the consolidated workbook's YTD By Department tab** (4 channels + TOTAL — the channels with actual revenue; OG Spec/Online Lux/Ink were deleted as "empty columns" in prior periods and stay deleted):

| Col | Channel |
|---|---|
| G | Trend House |
| I | Online |
| K | OG-DTC |
| M | All Pop Art |
| O | TOTAL (includes OG Specialty allocated expenses, even though OG Spec has no separate column) |

The H/J/L/N/P columns are static `% of channel sales` values.

**Steps each month:**

1. **Populate the current-month sheet in the data file.** Open `YTD By Dept_Data_2026.xlsx` and go to the matching tab (e.g., MAR for March). Update D4 to the correct period label (e.g., `March 2026` — the template often ships with the wrong year).
2. **Copy SIMPLE channel splits into the data file's month sheet.** For each row 8–77 from the consolidated workbook's SIMPLE tab, map the channel columns:

   | SIMPLE col | Channel | → Data file col |
   |---|---|---|
   | G | Trend House | G |
   | I | OG Specialty | H |
   | K | Online LUX | I |
   | M | Online | J |
   | O | OG-DTC | K |
   | Q | All Pop Art | L |
   | S | Ink | M |

   The data file's column N (TOTAL) is a formula and auto-derives.

3. **Recalc the data file** with LibreOffice headless so the YTD 2026 sheet's `=JAN!G8+FEB!G8+MAR!G8+…` cross-sheet formulas resolve to numbers.

4. **Read the YTD 2026 sheet** and write static values into the consolidated workbook's YTD By Department tab. Mapping:

   | Data file YTD 2026 col | Consolidated YTD By Department col |
   |---|---|
   | G (Trend House) | G |
   | M (Online Mass) | I |
   | O (OG-DTC) | K |
   | Q (All Pop Art) | M |
   | U (TOTAL) | O |

   For each row, also write the % column to the right: `H<r> = G<r> / G8`, `J<r> = I<r> / I8`, etc. For row 8 (Sales) every % = 100%.

5. **Update the period header** in the consolidated tab: `B4 = "YTD TOTALS - Through <MONTH> <YEAR>"`.

6. Recalc the consolidated workbook.

**Empty columns:** the existing layout already has only the 4 active-revenue channels. OG Specialty / Online Lux / Ink columns were deleted in prior periods — their values (when they exist, like OG Specialty expense allocations in March) roll up into the TOTAL column via the data file's YTD 2026 sum formula. Don't add them back unless they get revenue.

**March 2026 actuals applied** (channel YTD totals through March 2026):

| Row | Label | Trend House | Online | OG-DTC | All Pop Art | TOTAL |
|---|---|---:|---:|---:|---:|---:|
| 8 | Sales | 4,013,140 | 1,564,794 | 593,882 | 5,788 | 6,177,604 |
| 22 | Total COGS | 2,334,408 | 630,066 | 168,098 | 1,014 | 3,133,586 |
| 24 | Gross Profit | 1,670,557 | 702,082 | 378,412 | 3,720 | 2,754,771 |
| 38 | Total Controllable | 275,514 | 294,431 | 291,631 | 0 | 861,577 |
| 40 | Contribution Margin | 1,395,042 | 407,650 | 86,781 | 3,720 | 1,893,194 |
| 75 | Total General Corp | 574,395 | 295,319 | 173,152 | 1,254 | 1,055,243 |
| 77 | Operating Income | 820,647 | 112,331 | **−86,371** | 2,467 | 837,951 |

Operating margin by channel: Trend House 20.4%, Online 7.2%, OG-DTC **−14.5%** (losing money), All Pop Art 42.6%, blended TOTAL 13.6%.

**Validation:**

- Sales TOTAL ($6,177,604) ties to YTD tab S8 ($6,177,604) ✓
- Operating Income TOTAL ($837,951) ties to YTD tab S77 ($837,950, rounding $1) ✓
- DTC operating loss YTD ($86K) is worth flagging — driven by high allocated controllable expenses ($291K vs $107K Online; mostly Marketing at $136K vs $222K and Consulting at $111K)

### Online Returns Accrual v Actual

A small reconciliation tab that compares the monthly Online channel returns *accrual* (an estimate, calculated as 15.3% of Online sales) against the *actual* returns dollars from the chargeback PDF. The Diff column shows over/under-accrual; Running Balance is the cumulative position over time.

The tab has two stacked year-blocks (2025 in rows 3–16, 2026 in rows 20–33). Each block has the same shape:

| Col | Field | Type |
|---|---|---|
| A | Month label | text |
| B | Accrual | input — 15.3% × Online sales |
| C | Actual | input — actual chargebacks against Online |
| D | Diff | formula `=B−C` |
| E | Running Balance | formula `=prev_E + D` |

**Steps each month:**

1. **Get the Accrual** from the consolidated workbook's `RAW DATA_Master File` tab cell **B80**. The formula there is `=-B71*0.153` (negative of 15.3% × Online Sales row 71). Take the **absolute value** and enter it as a positive number — the workbook convention has all entries in column B positive. For March 2026: B71 = $542,805 → accrual = $83,049.17.

2. **Get the Actual** from the OG chargeback PDF — specifically, the customer-detail section's **"Online Total"** line. The chargeback extractor produces this in `extractor_outputs/chargeback_pdf/chargeback_customer_detail.csv`; filter rows where `department == "Online"` and `is_total_row == true`. Take the absolute value (chargebacks are reported negative; the workbook's column C is positive). For March 2026: Online Total from PDF = −$86,532 → enter $86,532.

   Note: "Online Total" excludes both the B&M Total (Brick & Mortar customers like Burlington / Wal-Mart Stores PR) and the Amazon Holdback Provision (which is a separate financial item, not a customer-driven return).

3. Find the month row for the current period in the 2026 block (e.g., March = row 24). Plug the Accrual into column B and the Actual into column C.

4. Recalc. Diff (D) and Running Balance (E) auto-derive.

**Column-shift not required** — this tab spans the calendar year, with each month occupying its own row. The 2025 block stays as historical reference.

**March 2026 actuals applied (row 24):**

| Cell | Value | Source |
|---|---|---|
| B24 Accrual | $83,049.17 | RAW DATA_Master File!B80 (= 542,805 × 0.153) |
| C24 Actual | $86,532.00 | Chargeback PDF customer-detail "Online Total" |
| D24 Diff | −$3,482.84 | formula |
| E24 Running Balance | −$92,054.65 | formula (Feb running −$88,571.81 + Mar diff −$3,483) |

**Validation:**

- March under-accrued by $3,483 (modest — actuals slightly higher than the 15.3% estimate)
- Cumulative under-accrual through March is −$92,054 (running balance drifts negative each month — the 15.3% rate has been low all year; Jan and Feb each under-accrued by ~$48K and ~$41K)
- If the running balance keeps drifting more negative, this would suggest the 15.3% rate may need to be revised upward in next year's plan

### HGF Dept Recap 2026 — "Diego Format" (separate workbook)

This is a separate workbook in the workspace folder: `HGF Dept Recap 2026_Diego Format.xlsx`. It tracks per-channel monthly P&L for the four primary channels in a different layout than the consolidated workbook. Each channel has its own sheet:

| Sheet | Channel | SIMPLE source column |
|---|---|---|
| ONLINE | Online Mass | M |
| DTC | OG-DTC | O |
| TH | Trend House | G |
| ONLINE LUX | Online LUX | K |

**Column layout (same on all 4 sheets):** each month is a 2-column block — value column + % column.

| Month | Value col | % col |
|---|---|---|
| JAN | C | D |
| FEB | E | F |
| **MAR** | **G** | **H** |
| APR | I | J |
| MAY | K | L |
| JUN | M | N |
| JUL | O | P |
| AUG | Q | R |
| SEP | S | T |
| OCT | U | V |
| NOV | W | X |
| DEC | Y | Z |
| TOTAL 2026 | AA | AB |
| Q1 | AD | AE |
| Q2 | AG | AH |
| Q3 | AJ | AK |
| Q4 | AM | AN |

The % columns and the TOTAL/Q1–Q4 columns are formula-driven and update automatically.

**Row layout (Diego rows ↔ SIMPLE rows for the value side):**

| Diego row | Label | SIMPLE row |
|---|---|---|
| 3 | Sales | 8 |
| 4 | Returns & Allowances | 9 |
| 5 | NET SALES | (formula `=row3+row4`) |
| 8 | Cost of Product & Labor | 14 |
| 9 | Finished Goods | 15 |
| 10 | Shipping | 16 |
| 11 | Tariffs | 17 |
| 12 | Fulfillment | 18 |
| 13 | Temporary Staffing | 19 |
| 14 | Warehouse Rent-Curci | 20 |
| 15 | Royalty Expense | 21 |
| 16 | Total COGS | (formula `=SUM(row8:row15)`) |
| 18 | GROSS PROFIT | (formula `=row5−row16`) |
| 21 | Payroll - Sales | 27 |
| 22 | Payroll Accrual | 28 |
| 23 | Payroll Tax | 29 |
| 24 | Commissions | 30 |
| 25 | Travel | 31 |
| 26 | LOC Interest | 32 |
| 27 | Meals & Entertainment | 33 |
| 28 | Software & Web Services | 34 |
| 29 | Samples | 35 |
| 30 | Consulting | 36 |
| 31 | Marketing | 37 |
| 32 | Total Controllable Expenses | (formula `=SUM(row21:row31)`) |
| 34 | CONTRIBUTION MARGIN | (formula `=row18−row32`) |
| 36 | Total Expenses (General Corporate) | 75 |
| 38 | OPERATING INCOME | (formula `=row34−row36`) |

So Diego rows 3–4 are offset by +5 from SIMPLE rows 8–9; Diego rows 8–15 are offset by +6 from SIMPLE rows 14–21; Diego rows 21–31 are offset by +6 from SIMPLE rows 27–37. Row 36 is a special single-cell mapping to SIMPLE row 75.

**Steps each month:**

1. Open `HGF Dept Recap 2026_Diego Format.xlsx`.
2. For each of the 4 sheets (ONLINE, DTC, TH, ONLINE LUX), identify the SIMPLE column for that channel (per the table above).
3. For each Diego input row in the table above, write `SIMPLE!<channel_col><simple_row>` into `Diego!<sheet>!<period_col><diego_row>`. For March that means writing into column G of each sheet across 22 rows.
4. Skip the formula rows (5, 16, 18, 32, 34, 38) — they auto-derive.
5. The % columns to the right of each value column already hold the right formulas (e.g., row 4 is `=C4/C3`, etc.) — they auto-recalc.
6. Recalc the Diego file with LibreOffice headless so the formula rows / TOTAL / Q1–Q4 columns refresh.

**Channels not in this file:** OG Specialty and All Pop Art are not tracked in Diego Format — only the four primary revenue channels. When reconciling totals back to SIMPLE row 77 Operating Income, add OG Specialty (often a small loss from allocated payroll) and APA back in.

**March 2026 actuals applied:**

| Sheet | Sales (G3) | Net Sales (G5) | Total COGS (G16) | Gross Profit (G18) | Total Controllable (G32) | Contrib Margin (G34) | Total Gen Corp (G36) | **Operating Income (G38)** |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| TH | 1,242,393 | 1,240,382 | 760,093 | 480,289 | 65,188 | 415,102 | 195,303 | **219,798** |
| ONLINE | 542,805 | 459,756 | 213,728 | 246,028 | 99,511 | 146,518 | 101,171 | **45,346** |
| DTC | 163,408 | 153,293 | 43,760 | 109,533 | 91,362 | 18,171 | 45,802 | **−27,630** |
| ONLINE LUX | 0 | 0 | 0 | 0 | 0 | 0 | 0 | **0** |
| **Diego Subtotal** | 1,948,606 | 1,853,431 | 1,017,581 | 835,851 | 256,061 | 579,791 | 342,276 | **237,514** |
| SIMPLE U77 | 1,950,439 | 1,854,919 | 1,017,884 | 837,035 | 256,061 | 580,975 | 353,899 | **227,075** |
| Reconciling diff | 1,833 | 1,488 | 303 | 1,184 | 0 | 1,184 | 11,623 | −10,439 |

Reconciling diff = OG Specialty (−$11,124 OI) + APA (+$684 OI) ≈ −$10,439, plus pennies of channel-allocation rounding.

**Validation:**

- TH OI $219,798 matches SIMPLE channel split in YTD By Department exercise ✓
- DTC OI −$27,630 matches ✓ (DTC is losing money this month)
- Online OI $45,346 matches ✓
- ONLINE LUX zeros across the board are correct — no LUX activity in March 2026 ✓

### Department Reports folder — per-channel files (Tracker + MARCH tabs)

The `Department Reports/` folder contains one workbook per channel each month — for March 2026 there are four:

| File | Channel | Consolidated workbook source tab | SIMPLE source column |
|---|---|---|---|
| `HGF <MONTH> <YEAR> TREND HOUSE.xlsx` | Trend House | Trend House | G |
| `HGF <MONTH> <YEAR> ONLINE.xlsx` | Online Mass | Online MASS | M |
| `HGF <MONTH> <YEAR> OG DTC.xlsx` | OG-DTC | OG DTC | O |
| `HGF <MONTH> <YEAR> OG DTC - JG.xlsx` | OG-DTC (JG copy) | OG DTC | O |

Each file has up to 5 tabs (`MARCH`, `Q1 RECAP`, `Ledgers`, `Impulse Buy`, `Tracker`). Only **Tracker** and **MARCH** are populated by this procedure — the others are filled in by separate processes / different owners.

The procedure for each file is identical, just substitute the channel-specific source per the table above.

The file has 5 sheets — `MARCH`, `Q1 RECAP`, `Ledgers`, `Impulse Buy`, and **`Tracker`**. Only the **Tracker** tab is populated by this procedure (the others are filled in by separate processes / different owners).

**Tracker tab structure:** identical to the Diego Format file's per-channel sheets, but for Trend House only and stops at row 34 (CONTRIBUTION MARGIN — no Total Expenses or Operating Income rows).

**Column layout** (same monthly value/% pattern as the Diego file): MAR value column = **G**.

**Row layout (Tracker rows ↔ SIMPLE rows):** same as the Diego Format file mapping for rows 3–31, **without rows 36/38** (this Tracker doesn't extend to General Corp / Operating Income).

| Tracker row | Label | SIMPLE row |
|---|---|---|
| 3 | Sales | 8 |
| 4 | Returns & Allowances | 9 |
| 5 | NET SALES | (formula) |
| 8 | Cost of Product & Labor | 14 |
| 9 | Finished Goods | 15 |
| 10 | Shipping | 16 |
| 11 | Tariffs | 17 |
| 12 | Fulfillment | 18 |
| 13 | Temporary Staffing | 19 |
| 14 | Warehouse Rent-Curci | 20 |
| 15 | Royalty Expense | 21 |
| 16 | Total COGS | (formula) |
| 18 | GROSS PROFIT | (formula) |
| 21 | Payroll - Sales | 27 |
| 22 | Payroll Accrual | 28 |
| 23 | Payroll Tax | 29 |
| 24 | Commissions | 30 |
| 25 | Travel | 31 |
| 26 | LOC Interest | 32 |
| 27 | Meals & Entertainment | 33 |
| 28 | Software & Web Services | 34 |
| 29 | Samples | 35 |
| 30 | Consulting | 36 |
| 31 | Marketing | 37 |
| 32 | Total Controllable Expenses | (formula) |
| 34 | CONTRIBUTION MARGIN | (formula) |

**Steps each month (per file):**

1. Find the destination file in `Department Reports/`.
2. Open the **Tracker** tab.
3. For each input row in the table above, write the value from the consolidated workbook's SIMPLE tab at the matching SIMPLE row in the channel's column (G for Trend House, M for Online, O for OG-DTC, K for Online LUX), into the Tracker's current-month column (column **G** for March, **I** for April, etc.).
4. Skip rows 5, 16, 18, 32, 34 — they're formulas.
5. Recalc the destination file. The "Total" / formula rows pick up the new values automatically.

**Note on the consolidated channel tabs:** the consolidated workbook has sheets called `Trend House`, `Online MASS`, `OG DTC`, and `APA` — these are all fully formula-driven (every cell pulls from `<MONTH> <YEAR> FULL`!E/G/H/I columns automatically) and need no manual update. They serve as the source for the per-file MARCH-tab snapshot (next section).

**March 2026 actuals applied (Tracker column G):**

| Row | Label | Value |
|---|---|---:|
| 3 | Sales | 1,242,393.40 |
| 4 | Returns & Allowances | −2,011.00 |
| 5 | NET SALES (formula) | 1,240,382.40 |
| 9 | Finished Goods | 607,710.08 |
| 10 | Shipping | 69,720.58 |
| 11 | Tariffs | 73,451.30 |
| 14 | Warehouse Rent-Curci | 9,001.00 |
| 15 | Royalty Expense | 210.00 |
| 16 | Total COGS (formula) | 760,092.96 |
| 18 | GROSS PROFIT (formula) | 480,289.44 |
| 21 | Payroll - Sales | 36,909.05 |
| 22 | Payroll Accrual | 3,571.84 |
| 23 | Payroll Tax | 3,177.75 |
| 25 | Travel | 9,734.66 |
| 27 | Meals & Entertainment | 3,006.68 |
| 30 | Consulting | 8,360.04 |
| 31 | Marketing | 427.69 |
| 32 | Total Controllable (formula) | 65,187.70 |
| 34 | CONTRIBUTION MARGIN (formula) | **415,101.74** |

**Validation:**

- Sales $1,242,393 matches SIMPLE G8 ✓
- Total COGS $760,093 matches Diego Format TH sheet G16 ✓
- Contribution Margin $415,102 matches Diego Format TH sheet G34 ✓
- TH is the most profitable channel — Contribution Margin 33.5% of Net Sales, well above any other channel

#### MARCH tab (in each per-channel file)

Each per-channel file in `Department Reports/` has a **MARCH** tab (or a tab named for whatever month the file represents) — a one-page snapshot showing current month vs prior year vs budget for that channel. It ships empty as a template each month.

**Source:** the consolidated workbook's matching channel tab (Trend House → `Trend House`, Online → `Online MASS`, OG-DTC → `OG DTC`). Those tabs are fully formula-driven, pulling from `<MONTH> <YEAR> FULL`!E (current), G (prior year), J (budget) for each row. After the consolidated workbook is recalc'd, those cached values represent exactly what the snapshot needs.

**Column layout on the consolidated Trend House tab (and in the MARCH tab once populated):**

| Col | Field |
|---|---|
| A | Section header (Revenue / COGS / etc.) |
| B | Line label |
| G | Trend House (current month value) |
| H | % of revenue, current |
| I | Prior Year value |
| J | % of revenue, prior year |
| K | Diff (current − prior) |
| L | Budget |
| M | % of revenue, budget |
| N | Diff (current − budget) |

Headers in row 7 with data in rows 8–40 (Sales through Contribution Margin). Rows 42 (Total Expenses - General Corporate) and 44 (Operating Income) are intentionally **excluded** from the per-channel snapshot — General Corp is allocated centrally rather than directly attributable, so the per-channel view ends at Contribution Margin.

**Steps each month (per file):**

1. Open the consolidated workbook's matching channel tab (per the file→source map at the top of this section). Confirm it shows the current period in cell E4 (e.g., "March 2026"). If not, recalc the workbook first.
2. Open the destination file's current-month-named tab.
3. Copy every populated cell from the consolidated channel tab into the destination month tab at the same row/column **EXCEPT row 42 (Total Expenses - General Corporate) and row 44 (Operating Income)** — those two rows are intentionally excluded from the per-channel snapshot per user direction (May 2026). The snapshot stops at row 40 Contribution Margin.
4. **Copy values, not formulas** — the consolidated formulas reference cells that don't exist in the destination file's standalone workbook, so copy-as-values is required.
5. Save. No recalc needed since everything is now static values.

**Quick Python recipe:**

```python
import openpyxl
# Per-file mapping: dest filename → (consolidated source tab, SIMPLE channel column)
files = [
    ('HGF MARCH 2026 TREND HOUSE.xlsx', 'Trend House', 'G'),
    ('HGF MARCH 2026 ONLINE.xlsx',      'Online MASS', 'M'),
    ('HGF MARCH 2026 OG DTC.xlsx',      'OG DTC',      'O'),
    ('HGF MARCH 2026 OG DTC - JG.xlsx', 'OG DTC',      'O'),
]
SKIP_ROWS = {42, 44}  # Total Expenses (General Corp) + Operating Income — excluded per channel snapshot convention
wb_src = openpyxl.load_workbook(consolidated_path, data_only=True)
for fname, src_tab, simple_col in files:
    src = wb_src[src_tab]
    dst_wb = openpyxl.load_workbook(dest_dir + fname)
    march = dst_wb['MARCH']
    for r in range(1, 45):
        if r in SKIP_ROWS: continue
        for c in range(1, 15):
            v = src.cell(row=r, column=c).value
            if v is not None:
                march.cell(row=r, column=c).value = v
    # Tracker: write SIMPLE channel column to Tracker col G (March), per row_map
    dst_wb.save(dest_dir + fname)
```

**March 2026 actuals applied — Contribution Margin row 40 highlights, by file:**

| File | Sales | vs PY | vs Budget | Total COGS | Gross Profit | Contrib Margin (row 40) |
|---|---:|---:|---:|---:|---:|---:|
| TREND HOUSE | 1,242,393 | +$294,501 | −$31,469 | 760,093 | 480,289 | **415,102** |
| ONLINE | 542,805 | +$55,985 | +$69,005 | 213,728 | 246,028 | **146,518** |
| OG DTC | 163,408 | +$33,276 | −$95,994 | 43,760 | 109,533 | **18,171** |
| OG DTC - JG | 163,408 | +$33,276 | −$95,994 | 43,760 | 109,533 | **18,171** (same as OG DTC) |

**Tracker derived totals (row 34 CONTRIBUTION MARGIN, column G), by file:**

| File | NET SALES | Total COGS | Gross Profit | Total Controllable | Contrib Margin |
|---|---:|---:|---:|---:|---:|
| TREND HOUSE | 1,240,382 | 760,093 | 480,289 | 65,188 | **415,102** |
| ONLINE | 459,756 | 213,728 | 246,028 | 99,511 | **146,518** |
| OG DTC | 153,293 | 43,760 | 109,533 | 91,362 | **18,171** |
| OG DTC - JG | 153,293 | 43,760 | 109,533 | 91,362 | **18,171** |

**Validation:**

- TREND HOUSE: Sales beat prior year by 31%, missed budget by 2%; Contribution Margin $415K — short of budget $39K but $40K above prior year
- ONLINE: Sales up 11.5% YoY and beat budget by $69K; Contribution Margin $147K
- OG DTC: Sales up 26% YoY but materially below budget ($259K target → $163K actual, −$96K shortfall); Contribution Margin only $18K (was $9K short of break-even before allocations)
- All four files have rows 42 (Total Gen Corp) and 44 (Operating Income) intentionally blank in the MARCH tab

#### Online Returns Accrual v Actual tab (ONLINE file only)

The `HGF <MONTH> <YEAR> ONLINE.xlsx` file has an extra tab — **`Online Returns Accrual v Actual`** — that mirrors the same-named tab in the consolidated workbook. The other three per-channel files do not have this tab.

**Source:** the consolidated workbook's `Online Returns Accrual v Actual` tab. The values for the current month are already populated there (see the Online Returns Accrual v Actual section earlier in this document).

**Important — row offset:** the ONLINE file's 2026 block starts one row earlier than the main file's:

| Period | Main file row | ONLINE file row |
|---|---|---|
| 2026 January | 22 | 21 |
| 2026 February | 23 | 22 |
| **2026 March** | **24** | **23** |
| 2026 April | 25 | 24 |
| ... | ... | ... |

(The 2025 block at rows 5–16 is the same in both files.)

**Steps each month:**

1. Open both the consolidated workbook and the ONLINE file at the `Online Returns Accrual v Actual` tab.
2. Find the current period's row in each file using the offset table above.
3. Copy the **Accrual** (column B) and **Actual** (column C) values from the main file into the ONLINE file. Do NOT copy column D (Diff) or column E (Running Balance) — those are formulas that auto-derive.
4. Recalc the ONLINE file. The running balance picks up correctly because it's `=prev_E + D<this_row>`.

**March 2026 actuals applied:**

| Cell | Value |
|---|---|
| B23 Accrual | 83,049.165 |
| C23 Actual | 86,532 |
| D23 Diff (formula) | −3,482.84 |
| E23 Running Balance (formula) | −92,054.65 |

The ONLINE file's running balance now matches the consolidated workbook (−$92,054.65 through March 2026) exactly.

**Row-label mismatches to flag (March 2026):**

Two rows have different account labels between SIMPLE column B and TTM column B, but their row numbers align. The user directive is to copy by row number, which means the SIMPLE-side account flows into the TTM-side row regardless of label:

| Row | SIMPLE label | TTM label | March 2026 value copied |
|---|---|---|---|
| 68 | Equipment Leasing | Professional Fees - IT | $12,172.00 |
| 73 | LOC Interest | Misc. Expense | $3,170.00 |

These mismatches existed before the March close and are likely just stale labeling in TTM column B. Worth a separate review pass to fix the labels in TTM, but they don't break the close.

**March 2026 actuals applied (column AC):**

61 rows copied row-for-row from `MARCH 2026_SIMPLE!U<r>`. Key results:

| Row | Label | March 2026 (AC) | TTM Total (AE) |
|---|---|---:|---:|
| 8 | Sales | 1,950,439.45 | 23,959,458.39 |
| 9 | Returns | −95,520.73 | −1,057,885.79 |
| 10 | Net Sales | 1,854,918.73 | 22,901,572.60 |
| 22 | Total COGS | 1,017,883.67 | 12,108,918.84 |
| 24 | Gross Profit | 837,035.06 | 10,792,653.77 |
| 38 | Total Controllable | 256,060.51 | 2,930,670.33 |
| 40 | Contribution Margin | 580,974.55 | 7,861,983.44 |
| 75 | Total General Corp | 353,899.14 | 4,010,605.73 |
| 77 | Operating Income | 227,075.41 | 3,851,377.71 |

**Validation:**

- AC8 ($1.95M) matches gross sales used in Revenue Per Employee D31 ✓
- AE row 77 (TTM Operating Income $3.85M) is consistent with 12-month run-rate at current contribution margin
- Each AD<row> percentage now displays correctly (no `#DIV/0!` for revenue rows)


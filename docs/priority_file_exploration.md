# Priority Close Files Exploration

Explored on 2026-04-30 against the March 2026 sample close package.

## Files Covered

1. `Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx`
2. `Workpapers MARCH/DATA/TH March 2026 Revenue Report.xlsx`
3. `Workpapers MARCH/DATA/- OG _ Chargeback Report - 03. March 2026.pdf`
4. `PAYROLL/Payroll Journal_March 2026.xlsx`
5. `BR Info.xlsx`
6. `Workpapers MARCH/DATA/DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx`
7. `Workpapers MARCH/DATA/INTERNAL - Division COGS 2019 - Current (26).xlsx`
8. `Workpapers MARCH/DATA/March Addbacks_$23,195.pdf`
9. `Workpapers MARCH/DATA/HGF CONSOLIDATED_MARCH 2026.xlsx`

## Overall Extractor Notes

- Clean operational tabs should be read with `polars.read_excel`.
- Presentation and report tabs need `openpyxl`-based extraction because they have title rows, merged cells, formulas, `#DIV/0!`, and multi-block layouts.
- Use `data_only=True` for business values and `data_only=False` for dependency/formula auditing.
- Styles matter. Red/yellow fills in GL/payroll-related files are business signals, not decoration.
- PDF email exports are extractable with `pdfplumber`, but tables need custom cleanup.

## 1. Profit and Loss By Dept.xlsx

Shape:
- 1 sheet: `Profit and Loss by Department`
- 55 rows x 13 columns
- 220 formulas
- 4 merged ranges

Layout:
- Rows 1-3 are report title/date.
- Row 5 contains department columns:
  - `All Pop Art`
  - `Brick Mortar - China`
  - `OG-DTC`
  - `Online`
  - `Z-COMPANY`
  - `Art Dept`
  - `Corporate Dept`
  - `IT Dept`
  - `Operations Dept`
  - `Production Dept`
  - `Total Z-COMPANY`
  - `Total`
- Row labels are in column A.
- Major lines include `Total Sales`, `Total Income`, `Total Cost of Goods Sold`, `Gross Profit`.

Extractor:
- Treat as a matrix P&L.
- Header row: 5.
- Data rows: 6 through about 52.
- Output long format: `line_item`, `department`, `amount`.
- Use `data_only=True`.

## 2. TH March 2026 Revenue Report.xlsx

Shape:
- 3 sheets: `Summary `, `Details `, `USA Stock `
- Summary: 14 x 9, 27 formulas
- Details: 39 x 11, 140 formulas
- USA Stock: 14 x 11, 63 formulas

Key totals:
- Summary total revenue: `1,242,393.40`
- Summary total cost: `750,574.496`
- Summary GM percentage: `39.5864%`

Extractor:
- Read all three tabs with `polars.read_excel`.
- Normalize headers by trimming whitespace.
- `Summary ` is account-level.
- `Details ` and `USA Stock ` are PO-level.
- Key fields:
  - `Internal PO`
  - `Account`
  - `Revenue`
  - `Production Cost`
  - `Shipping cost`
  - `Tariff`
  - `Total Cost`
  - `GM %`
  - `GM $`

## 3. OG Chargeback Report PDF

Shape:
- 2 pages
- Page 1 has one large extracted table.
- Page 2 has summary lines and smaller extracted tables.

Key March 2026 line from text:
- `2026 | 03. (March)`
- Contractual allowance: `-$65,092`
- Penalties: `-$7,821`
- Amazon holdback provision: `-$15,661`
- Returns: `-$15,596`
- Software fees: `-$34`
- Total chargebacks: `-$104,205`

Other useful text:
- March B&M Burlington PO: `-$1,884`
- March total by customer table text: `-$88,543`
- Note says miscellaneous reversals of chargebacks/remittances explain the difference.
- Walmart Marketplace reversal example: `-711.37`, `-647.62`, `-63.75`; reversals are ignored.

Extractor:
- Use `pdfplumber.extract_text()` for headline/month totals first.
- Use regex for `YYYY | NN. (Month)` rows.
- Use table extraction secondarily for category/customer detail.
- Preserve source page/line references for review.

## 4. Payroll Journal_March 2026.xlsx

Shape:
- 2 sheets: `Payroll`, `Payroll Distribution`
- Payroll: 70 x 14, 77 formulas
- Payroll Distribution: 29 x 2, 15 formulas

Style signals:
- Red fill `FFFF0000` appears on allocation/problem cells.
- Yellow fill `FFFFFF00` appears on a comment/allocation note.

Key distribution values:
- Trend House: `36,909.048`
- OG Specialty USA: `0`
- Online Lux: `0`
- Online: `1,846.152`
- DTC: `6,153.84`
- Corp: `82,809.24`
- Art: `32,961.56`
- IT: `15,846.15`
- Production: `64,635.26`
- Total: `241,161.25`

Lital allocation block:
- DTC: `772`
- Online: `6,176`
- TH: `1,544`
- Corp: `6,948`
- Total: `15,440`

Extractor:
- `Payroll` tab is a no-header allocation grid. Use `openpyxl`, not default `polars`.
- `Payroll Distribution` is the primary extraction target.
- Output categories as `payroll_bucket`, `amount`, `section`.
- Also collect styled-cell notes for review.

## 5. BR Info.xlsx

Shape:
- 1 sheet: `2026`
- 10 x 13
- No formulas

Layout:
- Row 1: year.
- Row 2: months.
- Rows 3-10: override/account lines.

March values:
- AllPopArt Sales: `1,833`
- AllPopArt Returns and Allowances: `-346`
- Employee Benefits: `10,898`
- Bank Fees: `799`
- Merchant Account Fees: `4,619`
- License & Tax: `2,052`
- Equipment Leasing: `12,172`
- LOC Interest: `3,170`

Extractor:
- Read as override table.
- Convert wide months to long format: `year`, `month`, `line_item`, `amount`.
- This file should override or supplement the consolidated raw-data lines.

## 6. DTC & WS Monthly Revenue Report.xlsx

Shape:
- 4 sheets: `March-Revenue `, `Shopify`, `Refunds`, `Coupons`
- `Shopify`: 160 x 12
- `Refunds`: 22 x 22
- `Coupons`: 20 x 10

Key summary values:
- DTC net sales: `146,599`
- WS net sales: `16,809.05`
- Total net sales: `163,408.05`
- OG-DTC refunds: `9,940.81`
- OG-WS refunds: `173.75`
- Total refunds: `10,114.56`

Extractor:
- `Shopify`, `Refunds`, and `Coupons` are clean `polars.read_excel` tabs.
- `March-Revenue ` is a small pivot/summary tab; parse row blocks:
  - `REVENUE`
  - `REFUNDS`
- Output separate tables:
  - `dtc_ws_revenue_summary`
  - `shopify_orders`
  - `refunds`
  - `coupons`

## 7. INTERNAL - Division COGS 2019 - Current.xlsx

Shape:
- 14 sheets
- Year tabs: `2018`, `2019`, `2020`, `2021`, `2022`, `2023`, `2024`, `2025`, `2026`
- Partner detail tabs for 2022-2026.
- Many tabs report 1,000 rows but actual used data is much smaller.

Important shape:
- Year tabs are month/type matrices.
- In 2020+ tabs, columns start with `Month`, `Type`, then business channels.
- 2026 header includes:
  - `Brick & Mortar - China`
  - `Specialty USA: (Prev. B&M USA & Trade)`
  - `Speciality: USA SAMPLES`
  - `Online`
  - `Online-Reworks`
  - `Online - UK`
  - `Online - Canada`
  - `Online - USA`
  - `Online Textiles - MWW`
  - `Online- Simple Canvas`
  - `Online- IMPORTED BEDDING & CH CURCI`
  - `Online - SAMPLES`

Extractor:
- Treat year tabs as semi-clean matrices.
- Use header row 1 for 2020 onward, row 2 for 2018/2019.
- Forward-fill `Month` down for sub-rows like `Material Cogs`, `Labor Cogs`, `SHIPPING`.
- Output long format: `year`, `month`, `type`, `channel`, `amount`.
- Partner detail tabs need a separate unpivot:
  - Month headers span groups of 3 measures: COGS, Material Cost, Labor Cost.

## 8. March Addbacks_$23,195.pdf

Shape:
- 2 pages
- Email thread export
- No extracted tables

Key text:
- Subject/thread: `HGF General Ledger for Review`
- Instruction: ignore allocation journal entries from the preliminary P&L.
- Attachment/reference: `HGF GL_March_Sent April 13.xlsx`
- Addbacks: `Addbacks (red) $23,195`
- Attachment/reference: `HGF GL_March_Sent April 13_DONE.xlsx`

Extractor:
- Use text extraction only.
- Capture directives and dollar amounts.
- Connect this PDF to the reviewed GL workbook by filename.
- The actual row-level addbacks are likely the red-filled rows in the DONE GL, not in this PDF.

## 9. HGF CONSOLIDATED_MARCH 2026.xlsx

Shape:
- 31 sheets
- Central workbook, but this copy appears to still contain many zero or `#DIV/0!` formula outputs.

Important tabs:
- `READ ME`
- `High Level Recap`
- `Department Insights`
- `Dept Profit Ratio`
- `Chargeback Recap`
- `Chargeback Detail`
- `MARCH 2026 FULL `
- `MARCH 2026_SIMPLE`
- `YTD`
- `YTD By Department`
- `Monthly YoY`
- `Online MASS`
- `OG DTC`
- `APA`
- `Trend House`
- `RAW DATA_Master File`
- `RAW DATA_COGS & Freight`
- `RAW DATA_Payroll`
- `Online Returns Accrual v Actual`

Style and formula notes:
- `MARCH 2026 FULL ` has 6,284 formulas.
- Heavy fill-color use marks comparison blocks:
  - blue: prior year
  - green: budget
  - pale yellow: prior quarter / reference
  - yellow: manual/review cells
- `RAW DATA_Master File` contains red, green, and yellow fills and should be a primary source for raw mappings.

Extractor:
- Do not rely on `polars.read_excel` for presentation tabs because `#DIV/0!` causes calamine dtype errors.
- Use `openpyxl` with `data_only=True` for values.
- Use `openpyxl` with `data_only=False` for formula dependencies.
- Build tab-specific extractors:
  - `readme_tabs`
  - `high_level_recap`
  - `department_insights`
  - `chargeback_recap`
  - `chargeback_detail`
  - `monthly_full_pnl`
  - `monthly_simple_pnl`
  - `raw_master_file`
  - `raw_cogs_freight`
  - `raw_payroll`

## Recommended Extractor Order

1. `BR Info.xlsx`: simple override extractor.
2. `TH March 2026 Revenue Report.xlsx`: clean revenue tables.
3. `DTC & WS Monthly Revenue Report.xlsx`: clean order/refund/coupon tabs plus summary.
4. `Payroll Journal_March 2026.xlsx`: payroll distribution plus styled review cells.
5. `INTERNAL - Division COGS 2019 - Current.xlsx`: COGS unpivot.
6. `Profit and Loss By Dept.xlsx`: matrix P&L unpivot.
7. `March Addbacks_$23,195.pdf`: directive extraction.
8. `OG Chargeback Report PDF`: regex/table hybrid extraction.
9. `HGF CONSOLIDATED_MARCH 2026.xlsx`: central workbook/reference extractor.

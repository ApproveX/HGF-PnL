---
name: HGF Close - Mapping Rules Learned from March 2026
description: Critical mapping rules for the HGF monthly P&L close pipeline, discovered by comparing our PRELIM against the accountant's FINAL for March 2026. These rules should be applied to all future close cycles.
type: project
originSessionId: 11748c49-d881-41ed-bc7e-91de4503cc75
---
## Corporate Department = "Total Z-COMPANY" (not "Corporate Dept")
When mapping Corp-level GL values (Software & Web Corp, Consulting Corp, Travel Corp, HR Recruiting, etc.), use the **"Total Z-COMPANY"** column from P&L by Department, NOT the narrower "Corporate Dept" column. Total Z-COMPANY excludes BM-China but includes all other departments, which is what the accountant expects for Corp-level line items.

**Why:** The Corporate Dept column only captures expenses coded directly to the corporate cost center. The accountant wants the consolidated company total (minus BM-China) for these line items.

**How to apply:** In consolidated_values.json, any `raw_master.gl.*` or `raw_master.consulting.corp`, `raw_master.software_web.corp`, `raw_master.travel.corp`, `raw_master.meals.corp` should pull from "Total Z-COMPANY" column.

## BR Info Overrides: Replace vs Supplement
For certain GL line items, when BR Info provides an override value, it **replaces** the GL base value entirely (the GL cell goes to 0 or empty). This applies to:
- **Bank Fees**: BR Info value replaces GL; GL cell = 0
- **Merchant Account Fees**: BR Info value replaces GL; GL cell shows only the small remainder
- **Equipment Lease**: BR Info value replaces GL; GL cell = empty

For other items (LOC Interest, Licenses & Taxes, etc.), BR Info supplements the GL value normally.

**Why:** The accountant treats certain BR Info items as the authoritative source, zeroing out the GL-derived figure.

**How to apply:** Check which BR Info items the accountant treats as replacements vs supplements. For replacement items, set the GL base cell to 0 and put the full amount in the adjustment cell.

## BR Info Cents Precision
BR Info values should preserve full decimal precision from the source file, not round to whole dollars:
- LOC Interest: 3,169.56 (not 3,170)
- Licenses & Taxes: 2,052.41 (not 2,052)
- APA Sales: 1,832.74 (not 1,833)

**Why:** The accountant's FINAL uses exact cents from BR Info.

**How to apply:** When extracting BR Info, preserve original float values without rounding.

## Chargeback Returns → TH Returns
The Trend House Returns cell in RAW DATA_Master File should include the **B&M chargeback total** from the chargeback report PDF (e.g., -$2,011 for March 2026).

**Why:** Chargebacks against B&M (brick & mortar) are treated as TH returns.

**How to apply:** Extract chargeback B&M total and map to `raw_master.returns.trend_house`.

## DTC Returns = Full DTC + WS Refund Total
DTC Returns should use the **full DTC+WS refund total** from the Monthly Revenue report, not just the DTC portion. For March 2026 this was -$10,115.17 (not -$9,940.81).

**Why:** Wholesale refunds processed through Shopify are rolled into the DTC returns line.

**How to apply:** Sum all refund line items (DTC refunds + WS refunds) for the DTC returns value.

## Online COGS: Merge Standalone "Online" into "Online-USA"
In Division COGS, standalone "Online" rows should be **merged into "Online-USA"** totals:
- Online COGS = Online-USA COGS + standalone Online COGS
- Online Material = Online-USA Material + standalone Online Material
- Online Labor = Online-USA Labor + standalone Online Labor

**Why:** The accountant doesn't distinguish standalone Online from Online-USA in the consolidated P&L; they're one line.

**How to apply:** When building `raw_cogs.current_month.cogs.online_usa`, add standalone online values.

## Tariffs Must Be Mapped
Tariffs from Division COGS must be extracted and mapped to `raw_cogs.tariffs`. For March 2026 this was $73,451.30.

**Why:** Tariffs were initially missed (set to 0) because the extractor output wasn't properly mapped.

**How to apply:** Look for tariff/duty rows in Division COGS matrix and sum them into `raw_cogs.tariffs`.

## TH Shipping for Samples
The TH Shipping (Samples) cell should include sample shipping costs. For March 2026 this was $307.46.

**Why:** The accountant includes sample-related shipping as a separate line item.

**How to apply:** Extract sample shipping from the appropriate source and map to the TH shipping samples cell.

## Label Conventions
- "DTC" in certain Master File cells should be labeled "OG Specialty Trade" to match the accountant's convention
- COGS column headers D1/E1 need correct labeling

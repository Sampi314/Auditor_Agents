# Audit Report - Brewery Financial Model 10Y

## Style Convention Summary

*Detection method: Statistical Inference*
*Confidence: High*

| Cell Type | Font Colour | Fill Colour | Font Style | Source | Confidence |
|---|---|---|---|---|---|
| Input | Blue (#0000FF) | No Fill | Normal | Inferred (95%) | ✅ High |
| Formula | Black (#000000) | No Fill | Normal | Inferred (98%) | ✅ High |
| Header | Black (#000000) | No Fill | Bold | Inferred (90%) | ✅ High |

---

## Audit Findings

| Sheet Name | Cell Reference | Description of the Location | Short Error Category | Long Description of Error |
|---|---|---|---|---|
| Assum_Pricing | C27:C30 | Calculated Excise per HL section | 🔴 MAJOR Logic Error | **HIGH:** Australian Excise Duty calculation does not subtract the 1.15% alcohol threshold. Current formula `ABV * Rate * 100` overstates excise costs significantly compared to the required `(ABV - 1.15%) * Rate * 100`. |
| Calc_Volume | B4:M120, B5:M120 | Period headers (Date and Year rows) | Hardcoded Date/Header | **HIGH:** Dates and Year numbers are hardcoded values instead of linking to the master `Timeline` sheet. This breaks model dynamicism if the start date changes. |
| SKU_Matrix | A4:A27 | SKU_ID column | Redundancy / Logic Gap | **MEDIUM:** The entire SKU Matrix (24 products) is unused in the model's calculations. Revenue and COGS are aggregated at Brand/Format level, ignoring the granular SKU data provided. |
| Calc_Revenue | B34:M36 | Key Metrics section (Gross/Net Revenue per HL) | Number Format Mismatch | **LOW:** These metrics are formatted as 'General' but should be Currency or Number with decimals for readability. |
| Cover | B13:B20 | Contents List | Contents Mismatch | **LOW:** 'Annual_Summary' sheet is missing from the table of contents. Detailed Assumption/Calculation sheets are grouped into single lines, reducing navigability. |
| Global | N/A | Entire Workbook | Scalability Flaw | **MEDIUM:** The workbook uses no Named Cell Styles (all cells are 'Normal'). This makes it difficult to enforce formatting standards or perform automated audits. |
| Assum_Costs | C6:C25 | COGS Assumptions | Logic Gap | **MEDIUM:** COGS are calculated using high-level averages per HL instead of utilizing the detailed Bill of Materials (BOM) for each SKU as per industry best practice. |
| IS, CF, BS | N/A | Financial Statements | Missing Business Logic | **MEDIUM:** 10% GST is not explicitly modeled in the financial statements or working capital calculations, potentially impacting cash flow accuracy. |
| Assum_Capex | F6 | Maintenance Capex - Year 4 | Non-Standard Input | **LOW:** Hard-coded value '0' is formatted with black font instead of the model's blue input style. |

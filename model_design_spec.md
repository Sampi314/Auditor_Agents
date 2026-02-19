# Master Brewery Model Specification v3.0

## Model Architecture
This model is a 12-sheet, 3-way integrated financial forecast for a regional Australian brewery. It is designed for maximum granularity, audibility, and scalability.

### Sheet Breakdown
1.  **01_Control**: Central scenario toggle (Base, High Cost, Aggressive Expansion) and global statutory rates (GST, Tax, Super, Excise).
2.  **02_Inputs_Sales**: Detailed SKU matrix (9 products) with ABV, multi-channel pricing (On-trade/Off-trade), and 12-month seasonality curves.
3.  **03_Inputs_Ops**: Granular Bill of Materials (BOM) including Malt, Hops, Packaging, and Utilities. Departmental headcount and salary tables.
4.  **04_Calcs_Volume**: Month-by-month sales and production volume (HL) flow with compounded growth logic.
5.  **05_Calcs_Revenue**: Auditable revenue waterfall. Gross Sales -> Australian Volumetric Excise (with 1.15% exemption) -> Net Revenue.
6.  **06_Calcs_COGS**: BOM-driven cost calculations. Dynamic malt/hops usage linked to production volume.
7.  **07_Calcs_OpEx**: Detailed labor costs (Salaries + Super + Payroll Tax) and variable S&M/G&A expenses.
8.  **08_Calcs_WC_Inv**: (Placeholder for future sub-ledger expansion) Tracks AR and Inventory days.
9.  **09_Calcs_Capex_Tax**: Straight-line depreciation and Australian corporate tax (30%) provisions.
10. **10_Outputs_3Way**: Fully integrated Income Statement, Balance Sheet, and Cash Flow. Format: 12 months (Y1) then 4 annual summaries.
11. **11_Dashboard**: North Star KPIs: Revenue/HL, EBITDA Margin %, ROIC, and Free Cash Flow.
12. **12_Audit_Checks**: Professional quality control sheet with Balance Sheet integrity and cash covenant triggers.

## Key Formulas (Auditable Logic)
*   **Australian Excise (Packaged)**: `=MAX(0, HL * 100 * (ABV - 1.15%) * Rate_Packaged)`
*   **BOM Usage**: `=SKU_Volume * Ingredients_per_HL * Loss_Factor`
*   **Labor Cost**: `=Headcount * Avg_Salary / 12 * (1 + Super_Rate + Payroll_Tax_Rate)`

## Scenario Logic
*   **Base**: 0.5% monthly volume growth.
*   **High Cost**: Input costs inflated by 20% from Month 13.
*   **Aggressive**: 1.5% monthly growth + capacity expansion CAPEX at Month 18.

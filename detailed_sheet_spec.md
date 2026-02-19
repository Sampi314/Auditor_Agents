# Exhaustive Brewery Model Specification (Step 5 Style)

## 00_Cover
*   **Objective**: Model navigation and metadata.
*   **Layout**: TOC with hyperlinks to sheets.
*   **Key Inputs**: Model Date, Version, Modeller Name.

## 01_Control
*   **Objective**: Centralize scenario and global driver selection.
*   **Layout**: Scenario Toggle (Index 1-3), Global Tax Rates (GST, Corp Tax), Inflation, Australian Excise Rates.
*   **Named Ranges**: `Scen_Index`, `Rate_GST`, `Rate_CorpTax`, `Rate_Excise_Pack`, `Rate_Excise_Dr`.

## 02_Inputs_Sales
*   **Objective**: Detailed SKU and Channel assumptions.
*   **Layout**:
    - SKU Matrix: Name, ABV, HL/Unit, Packaging Type.
    - Pricing: Gross Price/HL (On-trade vs Off-trade).
    - Trade Spend: Promotional % of Gross Rev per channel.
    - Seasonality: 12-month indices.
*   **Formulas**: `Unit_Sales = HL_Vol / (HL_per_Unit)`.

## 03_Inputs_Costs
*   **Objective**: Granular COGS and Inventory assumptions.
*   **Layout**:
    - BOM: Malt (kg/HL), Hops (g/HL), Yeast, Water, Adjuncts.
    - Packaging: Unit cost for bottles, cans, kegs, labels, cardboard.
    - Logistics: Freight-In (per tonne of Malt/Hops), Freight-Out (per HL/Pallet).
    - Inventory: Target Days on Hand (DOH) for RM and FG.

## 04_Inputs_OpEx_Capex
*   **Objective**: Fixed costs and asset assumptions.
*   **Layout**:
    - Labor: Dept-wise headcount, Avg Salary, Superannuation (11.5%), Payroll Tax.
    - Capex: List of assets, useful lives (years), expansion triggers.

## 05_Calcs_Volume
*   **Objective**: Calculate production and sales volume flow.
*   **Layout**:
    - Sales HL: Base HL * Seasonality * Growth.
    - Yield Loss: (Sales HL) / (1 - Packaging Loss %) / (1 - Brewing Loss %).
    - Production Plan: Month-by-month HL requirements.

## 06_Calcs_Revenue
*   **Objective**: Auditable revenue waterfall.
*   - Gross Sales = HL * Price/HL.
*   - Excise = LAL * Excise_Rate (LAL = HL * 100 * (ABV - 1.15%)).
*   - Trade Spend = Gross Sales * Promo %.
*   - Net Revenue = Gross Sales - Excise - Trade Spend.

## 07_Calcs_Inventory
*   **Objective**: Detailed sub-ledger for working capital.
*   - Opening Balance + Purchases - Usage = Closing Balance.
*   - RM Value = (Usage_m + Usage_m+1) / 2 * DOH.

## 08_Calcs_OpEx_Capex
*   - Monthly Salaries = HC * Salary / 12 * (1 + Benefits %).
*   - Depreciation = (NBV_Opening + Capex) / Remaining Life.

## 09_Outputs_3Way
*   - **IS**: Net Revenue down to NPAT.
*   - **BS**: Cash, AR, Inventory, Assets vs. AP, Debt, Equity.
*   - **CF**: Operating, Investing, Financing (Indirect Method).

## 10_Dashboard
*   - **KPIs**: Revenue/HL, Gross Margin %, EBITDA %, FCF, ROIC, Working Capital Days.
*   - **Covenants**: Debt/EBITDA (< 3.0x), Interest Cover (> 4.0x).

## 11_Checks
*   - Balance Sheet Balance (Assets - Liab - Eq = 0).
*   - Cash Check (Ending CF = BS Cash).
*   - Capacity Check (Production < Max HL).

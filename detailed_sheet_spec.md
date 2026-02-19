# Exhaustive Sheet-by-Sheet Specification (Master Model v3)

This document provides a cell-level design intent for the 12-sheet Master Brewery Model.

## Sheet 01: Control
- **Named Ranges**: `Scenario_Select`, `Stat_GST`, `Stat_CorpTax`, `Excise_Pack`, `Excise_Draught`.
- **Logic**: All calculations in sheets 04-09 must reference `Scenario_Select` to toggle input drivers.

## Sheet 02: Inputs_Sales
- **Data Structure**: Table of 9 SKUs.
- **Seasonality**: 1x12 row of multipliers. Month 12 (December) peaks at 1.45x.
- **Formulas**: `Target_Vol_m = Base_Vol * Seasonality_Index[m] * (1 + Growth_Rate)^m`.

## Sheet 03: Inputs_Ops
- **BOM Matrix**: 9x4 table linking SKU IDs to Malt (kg), Hops (g), Packaging (units), and Utilities ($).
- **Headcount Table**: Split by Production, Packaging, Quality, S&M, and Admin.

## Sheet 04: Calcs_Volume
- **Objective**: Detailed volume flow per SKU.
- **Integrity**: Sum of individual SKUs must equal `Total_HL`.

## Sheet 05: Calcs_Revenue
- **Waterfall**:
    1. Gross Sales (On-Trade + Off-Trade).
    2. Excise Duty: Logic `HL * 100 * (ABV - 0.0115) * Rate`.
    3. Net Revenue (The core top-line for IS).

## Sheet 06: Calcs_COGS
- **Malt Usage**: `SUMPRODUCT(SKU_Vols, SKU_Malt_BOM) * Malt_Price`.
- **Hops Usage**: `SUMPRODUCT(SKU_Vols, SKU_Hops_BOM/1000) * Hops_Price`.
- **Total COGS**: Materials + Packaging + Variable Logistics.

## Sheet 07: Calcs_OpEx
- **Labor Formula**: `HC * Avg_Salary / 12 * (1 + Super + PayrollTax)`.
- **S&M**: `Net_Revenue * 12%` (Variable component).

## Sheet 10: Outputs_3Way
- **Layout**: Columns B-M (Year 1 Monthly), Columns N-Q (Years 2-5 Annual).
- **Integration**: `BS_Cash = CF_Ending_Cash`. `RE_Closing = RE_Opening + NPAT`.

## Sheet 12: Audit_Checks
- **Checks**:
    1. `Assets - (Liabilities + Equity) = 0`.
    2. `Production_Volume <= Brew_Capacity`.
    3. `Min_Cash_Reserve > $500,000`.

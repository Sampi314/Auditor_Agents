# Builder Blueprint: Brewery Financial Model 🍻

> *"Detail in every drop. SKU-level drivers for precision."*

## Overview

This blueprint is designed for manufacturing and FMCG (Fast-Moving Consumer Goods) models, specifically breweries. It focuses on granular production drivers, complex Cost of Goods Sold (COGS) calculations, and distinct modules for Volume, Pricing, and Costs.

## Structure & Architecture

Based on `Brewery_Financial_Model_10Y 1.xlsx`, the model structure follows a strict Input -> Calculation -> Output flow.

### 1. Inputs (Assumptions) 📥

These sheets contain all hard-coded assumptions (Blue font).

- **`SKU_Matrix`**: Defines the product portfolio (Beer types, Kegs vs. Bottles vs. Cans).
- **`Assum_Volume`**: Sales volume projections by SKU/Channel.
- **`Assum_Pricing`**: Price per unit (Hectoliter/Case) by SKU/Channel.
- **`Assum_Costs`**:
    - **Variable Costs**: Raw materials (Malt, Hops, Water), Packaging.
    - **Fixed Costs**: Labor, Rent, Utilities, Marketing.
- **`Assum_Capex`**: Capital expenditures for brewery equipment, tanks, bottling lines.
- **`Assum_WC`**: Working Capital days (Receivables, Payables, Inventory).
- **`Assum_Finance`**: Debt terms, Equity injections, Tax rates.

### 2. Calculations (Processing) ⚙️

These sheets contain the calculation engines. No hard-coding allowed here.

- **`Calc_Volume`**:
    - **Drivers**: Seasonality adjustments, growth rates applied to base volumes.
    - **Production**: Sales Volume + Target Inventory Change = Production Volume.
- **`Calc_Revenue`**:
    - `Volume (HL)` × `Price/HL` = `Gross Revenue`.
    - Less: Excise Duty, Discounts = `Net Revenue`.
- **`Calc_COGS`**:
    - **BOM (Bill of Materials)** logic: Usage per HL × Cost per Unit.
    - Direct Labor + Overhead allocation.
- **`Calc_Opex`**: Staff costs, SG&A, Marketing spend.
- **`Calc_Capex`**: Depreciation schedules (Straight-line).
- **`Calc_WC`**:
    - Debtors (Revenue × Days/365).
    - Creditors (COGS × Days/365).
    - Inventory (Raw Mat + Fin Goods).
- **`Calc_Debt`**: Interest expense and Principal repayment schedules.

### 3. Outputs (Reporting) 📊

These sheets present the results.

- **`IS` (Income Statement)**:
    - Revenue -> Gross Profit -> EBITDA -> EBIT -> Net Income.
- **`BS` (Balance Sheet)**:
    - Assets (Cash, AR, Inventory, PPE).
    - Liabilities (AP, Debt).
    - Equity (Retained Earnings).
- **`CF` (Cash Flow)**:
    - Operating, Investing, Financing activities.
    - **Check**: Ending Cash matches BS Cash.
- **`KPIs`**:
    - Volume (HL Sold).
    - Gross Margin %.
    - EBITDA Margin %.
    - Revenue per HL.
- **`Checks`**:
    - Balance Sheet Balance (Assets - Liabs - Equity = 0).
    - Cash minimum check.
- **`Annual_Summary`**: High-level annual view for executives.

## Key Logic & Formulas

- **Volume-Driven**: Nearly all revenue and variable costs are functions of the `Calc_Volume` sheet.
- **Unit Conversions**: Careful handling of Units (e.g., Hectoliters vs. Cases vs. Pints).
- **Excise Duty**: Specific calculation often based on ABV (Alcohol by Volume) or volume produced.

## Implementation Steps

1.  **SKU Setup**: Define the `SKU_Matrix` first. This drives the rows in all Assumption and Calc sheets.
2.  **Volume & Price**: Build `Assum_Volume` and `Assum_Pricing` linked to SKUs.
3.  **Revenue Engine**: Calculate Gross Revenue.
4.  **COGS Build**: Link specific raw material costs to production volumes.
5.  **Opex & Fixed Costs**: Layer in fixed overheads.
6.  **Financial Statements**: Link the engines to IS/BS/CF.
7.  **Dashboards**: Build `KPIs` and `Annual_Summary`.

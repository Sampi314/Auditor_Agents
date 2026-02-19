# Brewery Financial Model Design Specification

## Overview
This document outlines the architecture and logic for the 3-Way Integrated Financial Model for a regional Australian craft brewery.

## 1. Sheet Structure (SumProduct Best Practice)
- **00_Cover**: Title, Version, Navigation.
- **01_Control**: Scenario toggle (Base, High Cost, Channel Shift), GST (10%), Corp Tax (30%), Inflation assumptions.
- **02_Inputs_Sales**:
    - SKUs: Lager, Pale Ale, IPA, Stout, Seasonal (Summer), Seasonal (Winter).
    - ABV%: 4.2% to 7.5%.
    - Pricing: Wholesale prices per HL for On-trade vs Off-trade.
    - Seasonality: Monthly weights peaking in Dec/Jan.
- **03_Inputs_Costs**:
    - BOM: Malt (kg/HL), Hops (g/HL), Yeast, Water.
    - Packaging: Bottles/Cans/Kegs unit costs.
    - Logistics: Variable cost per HL.
- **04_Inputs_OpEx_Capex**:
    - Headcount: Brewing, Sales, Admin.
    - Capex: Brewhouse, Fermenters, Canning Line.
- **05_Calcs_Volume**: Volume roll-up by SKU, Channel, and Time.
- **06_Calcs_Revenue**:
    - Gross Sales.
    - **Excise Duty**: Volumetric calculation (L Litres * (ABV - 1.15%) * Rate).
    - Trade Spend (Rebates/Promotions).
- **07_Calcs_COGS**: Raw materials + Direct Labor + Energy.
- **08_Calcs_OpEx**: Salaries + S&M + G&A.
- **09_Calcs_Capex_Dep**: Asset sub-ledgers and depreciation.
- **10_Calcs_Finance**: Term Loan, Overdraft, Tax (GST/Income Tax).
- **11-13_Outputs**: IS, BS, CF.
- **14_Dashboard**: KPIs (Gross Margin %, EBITDA %, FCF, ROIC).
- **15_Checks**: Balance Sheet check, Capacity check, Cash check.

## 2. Key Industry Logic
- **Hectolitres (HL)**: All production and sales are tracked in HL (1 HL = 100 Litres).
- **Yield Loss**: 5% brewing loss, 2% packaging loss modeled.
- **Excise**: Australian ATO rates (Feb 2024 approx: $60.22/LAL for packaged, $40.50/LAL for draught).
- **Capacity**: Maximum annual HL output. Capex trigger adds 50,000 HL capacity.

## 3. Financial Integration
- **Retained Earnings**: Linked to Net Income.
- **Cash**: Ending Cash on BS linked to CF Statement.
- **Working Capital**: Days Sales Outstanding (DSO) and Days Payable Outstanding (DPO) logic.

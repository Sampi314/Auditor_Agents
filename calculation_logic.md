# Detailed Calculation Logic

## 1. Volume and Revenue
- **Volume Calculation**:
    - `Base_Volume_Mth = (Annual_Plan_HL / 12) * Growth_Factor`
    - `Seasonal_HL = Base_Volume_Mth * Seasonality_Index[m]`
    - `Channel_HL = Seasonal_HL * Channel_Split%`
- **Excise Duty (AU)**:
    - `Litre_of_Alcohol (LAL) = HL * 100 * (ABV - 1.15%)`
    - `Excise_Amount = LAL * [Relevant_Rate]`
- **Net Revenue**:
    - `Net_Rev = (HL * Price_per_HL) - Excise_Amount - (Gross_Sales * Trade_Spend%)`

## 2. Cost of Goods Sold (COGS)
- **BOM Materials**:
    - `Malt_Cost = HL * Malt_kg_per_HL * Price_per_kg`
    - `Hops_Cost = HL * Hops_kg_per_HL * Price_per_kg`
- **Packaging**:
    - `Pack_Cost = HL * (1 / HL_per_Unit) * Unit_Cost`
- **Variable Utilities**:
    - `Energy_Cost = HL * kWh_per_HL * Rate_per_kWh`

## 3. Operating Expenses (OpEx)
- **Labor**: `Headcount[Dept] * Avg_Salary[Dept] / 12`
- **S&M**: `Marketing_Fixed + (Net_Revenue * Marketing_Var%)`

## 4. Asset and Debt Schedules
- **Depreciation**: `Straight-line monthly based on Useful Life (Months)`
- **Debt Service**:
    - `Interest = Opening_Balance * (Rate / 12)`
    - `Principal = Fixed_Monthly_Repayment`

## 5. Working Capital (3-Way Integration)
- **Accounts Receivable**: `Net_Revenue[current_month] * (DSO / 30)`
- **Accounts Payable**: `(COGS + OpEx_Cash) * (DPO / 30)`
- **Inventory**: `Raw_Materials_Cost * (Inventory_Days / 30)`

## 6. Financial Statements
- **NPAT**: Links to Retained Earnings on Balance Sheet.
- **Cashflow**: Ending Cash links to Cash on Balance Sheet.
- **Balance Sheet Check**: `Total Assets - (Total Liabilities + Equity) = 0`

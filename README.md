# 3-Statement Financial Model with Control Accounts

## Overview

This is a fully integrated 3-statement financial model built to professional financial modeling standards. The model features comprehensive control accounts for every balance sheet line item, ensuring complete audit trail and transparency.

## Model Structure

### 1. **Inputs Sheet**
Contains all model assumptions and drivers:
- Revenue assumptions (starting revenue, monthly growth rate)
- Cost of Goods Sold (% of revenue)
- Operating expenses (Sales & Marketing, G&A)
- Working capital assumptions (AR days, Inventory days, AP days, etc.)
- Fixed asset parameters (opening balances, capex, depreciation rate)
- Debt parameters (opening balance, interest rate, repayment schedule)
- Equity opening balances
- Tax rate

**Color Coding:** All input cells are highlighted in **YELLOW** to clearly distinguish assumptions from calculations.

### 2. **Workings Sheet - Control Accounts**

This sheet contains the core control account methodology. Every balance sheet item is driven by its own control account following the structure:

```
Opening Balance
+ Increases (movements that increase the account)
- Decreases (movements that reduce the account)
= Closing Balance
```

#### Control Accounts Included:

**Current Assets:**
- **Accounts Receivable Control**
  - Opening Balance
  - + Credit Sales
  - - Cash Collections (based on collection period)
  - = Closing Balance

- **Inventory Control**
  - Opening Balance
  - + Purchases
  - - COGS Withdrawal
  - = Closing Balance

- **Prepayments Control**
  - Opening Balance
  - + Cash Prepayments
  - - Expensed to P&L
  - = Closing Balance

**Current Liabilities:**
- **Accounts Payable Control**
  - Opening Balance
  - + Credit Purchases
  - - Cash Payments
  - = Closing Balance

- **Accrued Expenses Control**
  - Opening Balance
  - + Expenses Accrued
  - - Cash Payments
  - = Closing Balance

- **Deferred Revenue Control**
  - Opening Balance
  - + Cash Received in Advance
  - - Revenue Recognized
  - = Closing Balance

**Fixed Assets:**
- **PP&E Gross Control**
  - Opening Balance
  - + Capital Expenditure
  - - Disposals
  - = Closing Balance

- **Accumulated Depreciation Control**
  - Opening Balance
  - + Depreciation Expense
  - - Disposals
  - = Closing Balance

**Long-term Liabilities:**
- **Debt Control**
  - Opening Balance
  - + Drawdowns
  - - Principal Repayments
  - = Closing Balance
  - Interest calculation for P&L

**Equity:**
- **Share Capital Control**
  - Opening Balance
  - + New Issuances
  - - Buybacks
  - = Closing Balance

- **Retained Earnings Control**
  - Opening Balance
  - + Net Income (from P&L)
  - - Dividends
  - = Closing Balance

- **Cash Control**
  - Opening Balance
  - = Closing Balance (from Cash Flow Statement)

### 3. **Income Statement**

Monthly P&L statement showing:
- Revenue (linked to revenue drivers)
- Cost of Goods Sold (% of revenue)
- Gross Profit
- Operating Expenses:
  - Sales & Marketing
  - General & Administrative
  - Depreciation (from control account)
- EBIT
- Interest Expense (from debt control account)
- Earnings Before Tax
- Income Tax (based on tax rate)
- **Net Income**

### 4. **Balance Sheet**

Monthly balance sheet completely driven by control account closing balances:

**Assets:**
- Cash (from Cash Control)
- Accounts Receivable (from AR Control)
- Inventory (from Inventory Control)
- Prepayments (from Prepayments Control)
- PP&E Gross (from PP&E Gross Control)
- Less: Accumulated Depreciation (from Acc. Dep. Control)
- **Total Assets**

**Liabilities:**
- Accounts Payable (from AP Control)
- Accrued Expenses (from Accrued Control)
- Deferred Revenue (from Deferred Revenue Control)
- Long-term Debt (from Debt Control)
- **Total Liabilities**

**Equity:**
- Share Capital (from Share Capital Control)
- Retained Earnings (from Retained Earnings Control)
- **Total Equity**

**Balance Check:** Assets = Liabilities + Equity (verification row shows 0 if balanced)

### 5. **Cash Flow Statement**

Built using the **indirect method** - reconciling net income to cash:

**Operating Activities:**
- Net Income
- Add back: Depreciation (non-cash expense)
- Changes in Working Capital:
  - Increase/Decrease in AR
  - Increase/Decrease in Inventory
  - Increase/Decrease in Prepayments
  - Increase/Decrease in AP
  - Increase/Decrease in Accrued Expenses
  - Increase/Decrease in Deferred Revenue
- **Cash from Operating Activities**

**Investing Activities:**
- Capital Expenditure
- **Cash from Investing Activities**

**Financing Activities:**
- Debt Repayments
- Dividends Paid
- **Cash from Financing Activities**

**Net Change in Cash**
- Opening Cash
- **Closing Cash**

**Reconciliation Check:** Closing cash matches Balance Sheet (verification row shows 0)

## Key Features

### Financial Modeling Best Practices

1. **No Hardcodes in Calculation Blocks**
   - All calculations reference inputs or other cells
   - Easy to audit and trace formulas

2. **Clear Separation of Inputs, Workings, and Outputs**
   - Inputs: Yellow highlighted cells
   - Workings: Control account calculations
   - Outputs: Three financial statements

3. **Transparent Formulas**
   - All formulas are visible and traceable
   - No hidden calculations

4. **Control Account Methodology**
   - Every balance sheet item has a control account
   - Full audit trail of all movements
   - Opening + Increases - Decreases = Closing

5. **Complete Integration**
   - Income Statement ↔ Balance Sheet via Retained Earnings
   - Balance Sheet ↔ Cash Flow via balance changes
   - Cash Flow ↔ Balance Sheet via closing cash
   - All three statements fully reconcile

6. **Monthly Timeline**
   - 12-month projection period
   - Month-over-month analysis capability

7. **No Circular References**
   - Model is structured to avoid circularity
   - Clean calculation flow

## How to Use the Model

### Modifying Assumptions

1. Go to the **Inputs** sheet
2. Change any **yellow highlighted cells** to test different scenarios
3. All three statements will automatically update

### Key Assumptions to Consider

- **Revenue Growth:** Adjust monthly growth rate to model different scenarios
- **Working Capital:** Change collection/payment periods to see cash impact
- **Capex:** Modify monthly capital expenditure
- **Debt:** Adjust repayment schedule or interest rate
- **Operating Expenses:** Update fixed costs

### Analyzing Results

1. **Profitability:** Review Income Statement for margins and profitability trends
2. **Liquidity:** Check Balance Sheet cash position and working capital
3. **Cash Generation:** Analyze Cash Flow Statement to understand cash drivers
4. **Balance Sheet Strength:** Monitor debt levels and equity position

### Verification

The model includes two key verification checks:

1. **Balance Sheet Check (Row 40):** Should always equal 0
   - Confirms Assets = Liabilities + Equity

2. **Cash Flow Check (Row 33):** Should always equal 0
   - Confirms Cash Flow Statement closing cash = Balance Sheet cash

## Model Architecture

```
INPUTS
   ↓
WORKINGS (Control Accounts)
   ↓
   ├─→ INCOME STATEMENT
   ├─→ BALANCE SHEET
   └─→ CASH FLOW STATEMENT
```

## Technical Specifications

- **File Format:** Excel (.xlsx)
- **Number of Sheets:** 5
- **Timeline:** 12 months (Jan 2025 - Dec 2025)
- **Calculation Method:** Formula-based (no macros)
- **Control Accounts:** 11 comprehensive control accounts

## Files Included

1. `Financial_Model_3_Statement.xlsx` - The complete financial model
2. `build_financial_model.py` - Python script used to generate the model
3. `README.md` - This documentation file

## Future Enhancements

Potential additions to the model:
- Sensitivity analysis tables
- Scenario comparison (Base/Best/Worst case)
- Extended projection period (24 or 36 months)
- Additional financial ratios and KPIs
- Charts and visualizations
- Tax loss carryforward tracking
- Multi-currency support

## Notes

- This model follows best-practice financial modeling standards
- All balance sheet movements are fully traceable through control accounts
- The indirect cash flow method provides reconciliation to net income
- Model is designed for easy auditing and stakeholder review

---

**Created:** November 2025
**Version:** 1.0
**Model Type:** Integrated 3-Statement Financial Model with Control Accounts

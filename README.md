# 3-Statement Financial Model with Control Accounts

## Overview

This is a fully integrated 3-statement financial model built to professional financial modeling standards. The model includes:
- **Income Statement** (P&L)
- **Balance Sheet**
- **Cash Flow Statement** (Indirect Method)

All three statements are fully reconciled and linked through comprehensive control accounts.

## Key Features

### 1. **Control Account Architecture**
Every Balance Sheet line item is driven by its own control account that tracks:
- **Opening Balance**: Starting position for the period
- **Increases**: All movements that increase the account
- **Decreases**: All movements that reduce the account
- **Closing Balance**: Opening + Increases - Decreases

### 2. **No Hard-Codes in Calculations**
- All inputs are centralized in the Inputs sheet
- Calculation blocks contain only formulas
- Clear separation of inputs, workings, and outputs

### 3. **Monthly Timeline**
- 12-month projection period
- Month-over-month tracking
- Easy to extend timeline if needed

### 4. **Full Integration**
- Income Statement flows to Retained Earnings
- Balance Sheet balances to zero
- Cash Flow reconciles opening to closing cash
- All statements linked through control accounts

## File Structure

### **1. Inputs Sheet**
Contains all assumptions and starting balances:

**Revenue Assumptions:**
- Monthly Revenue (with 5% growth)
- Revenue Growth %

**Cost Assumptions:**
- COGS % of Revenue (40%)
- Operating Expenses (Fixed at $25,000/month)

**Working Capital Assumptions:**
- AR Days (30 days)
- Inventory Days (45 days)
- AP Days (30 days)
- Prepayment amounts and amortization periods

**Capex & Fixed Assets:**
- Monthly Capex schedule
- Depreciation Rate (20% annual)
- Asset Disposals

**Debt & Financing:**
- Debt Drawdowns
- Debt Repayments
- Interest Rate (6% annual)

**Tax:**
- Tax Rate (25%)

**Starting Balances (Month 0):**
- Opening Cash: $50,000
- Opening Fixed Assets (Gross): $100,000
- Opening Accumulated Depreciation: $20,000
- Opening Share Capital: $100,000
- Opening Retained Earnings: $30,000
- All other accounts start at $0

### **2. Workings Sheet**
Contains all control accounts that drive the Balance Sheet:

#### **Asset Control Accounts:**

**Accounts Receivable:**
- Opening AR
- \+ Credit Sales (Revenue)
- \- Cash Collections
- = Closing AR

**Inventory:**
- Opening Inventory
- \+ Purchases
- \- COGS Withdrawal
- = Closing Inventory

**Prepayments:**
- Opening Prepayments
- \+ New Prepayments (Cash Paid)
- \- Amortization (Expense)
- = Closing Prepayments

**Fixed Assets (Gross):**
- Opening Fixed Assets
- \+ Capex Additions
- \- Disposals
- = Closing Fixed Assets (Gross)

**Accumulated Depreciation:**
- Opening Accumulated Depreciation
- \+ Depreciation Expense
- \- Disposal Depreciation
- = Closing Accumulated Depreciation

**Net Fixed Assets:**
- = Gross Fixed Assets - Accumulated Depreciation

#### **Liability Control Accounts:**

**Accounts Payable:**
- Opening AP
- \+ Purchases (Credit)
- \- Cash Payments
- = Closing AP

**Accrued Expenses:**
- Opening Accrued Expenses
- \+ Accruals (Expense Recognition)
- \- Cash Payments
- = Closing Accrued Expenses

**Deferred Revenue:**
- Opening Deferred Revenue
- \+ Cash Received in Advance
- \- Revenue Recognition
- = Closing Deferred Revenue

**Debt:**
- Opening Debt
- \+ Debt Drawdowns
- \- Debt Repayments
- = Closing Debt

### **3. Income Statement**
Standard P&L structure:

- **Revenue**: Linked from Inputs
- **COGS**: Linked from Inventory Control Account withdrawal
- **Gross Profit**: Revenue - COGS
- **Operating Expenses**:
  - Salaries & Wages
  - Other Operating Expenses
  - Prepayment Amortization
- **EBITDA**: Gross Profit - Operating Expenses
- **Depreciation**: Linked from Depreciation Control Account
- **EBIT**: EBITDA - Depreciation
- **Interest Expense**: Calculated from Debt Control Account
- **EBT**: EBIT - Interest
- **Income Tax**: 25% of positive EBT
- **Net Income**: EBT - Tax

### **4. Balance Sheet**
Standard Balance Sheet structure with all items linked to control accounts:

**Assets:**
- Cash (from Cash Flow Statement)
- Accounts Receivable (from AR Control)
- Inventory (from Inventory Control)
- Prepayments (from Prepayments Control)
- Fixed Assets (Net) (from Fixed Assets Control)

**Liabilities:**
- Accounts Payable (from AP Control)
- Accrued Expenses (from Accrued Expenses Control)
- Deferred Revenue (from Deferred Revenue Control)
- Long-term Debt (from Debt Control)

**Equity:**
- Share Capital (constant)
- Retained Earnings (opening + cumulative Net Income)

**Balance Check:**
- Assets = Liabilities + Equity
- Red highlighted cell shows difference (should be 0)

### **5. Cash Flow Statement**
Indirect method, fully reconciled:

**Operating Activities:**
- Net Income
- \+ Depreciation (non-cash)
- Changes in Working Capital:
  - (Increase)/Decrease in AR
  - (Increase)/Decrease in Inventory
  - (Increase)/Decrease in Prepayments
  - Increase/(Decrease) in AP
  - Increase/(Decrease) in Accrued Expenses
  - Increase/(Decrease) in Deferred Revenue

**Investing Activities:**
- Capital Expenditure
- Proceeds from Asset Disposals

**Financing Activities:**
- Debt Drawdowns
- Debt Repayments
- Interest Paid
- Equity Injections
- Dividends Paid

**Cash Reconciliation:**
- Opening Cash
- \+ Net Change in Cash
- = Closing Cash

## How the Model Works

### Integration Flow:

1. **Inputs Drive Everything**: All assumptions are in the Inputs sheet
2. **Control Accounts Calculate Movements**: Workings sheet tracks all balance sheet changes
3. **Income Statement**: Pulls from Inputs and Control Accounts
4. **Balance Sheet**: Every line item pulls from corresponding Control Account closing balance
5. **Cash Flow**: Calculated using indirect method from Net Income and balance sheet movements
6. **Closing Cash**: Feeds back to Balance Sheet to complete the circle

### Key Reconciliation Points:

1. **Balance Sheet Balances**: Total Assets = Total Liabilities + Equity
2. **Cash Flow Reconciles**: Opening Cash + Net Cash Flow = Closing Cash
3. **Retained Earnings Rolls Forward**: Opening RE + Net Income = Closing RE
4. **All Control Accounts Balance**: Opening + Increases - Decreases = Closing

## How to Use the Model

### Modifying Assumptions:

1. **Open the Inputs sheet**
2. **Change any yellow-highlighted input cells**
3. **All calculations update automatically**

### Common Modifications:

**To change revenue:**
- Update the "Monthly Revenue" row in Inputs sheet

**To change costs:**
- Update "COGS % of Revenue" or "Operating Expenses"

**To add Capex:**
- Enter amount in "Monthly Capex" for desired month

**To add/remove debt:**
- Enter "Debt Drawdowns" or "Debt Repayments" in desired month

**To change working capital:**
- Modify "AR Days", "Inventory Days", or "AP Days"

### Extending the Timeline:

To extend beyond 12 months:
1. Modify the `months = 12` variable in `build_financial_model.py`
2. Re-run the script: `python3 build_financial_model.py`
3. New workbook will be generated with extended timeline

## Model Standards Compliance

### Financial Modeling Best Practices:

✓ **Separation of Inputs/Workings/Outputs**
✓ **No hard-coded values in calculation blocks**
✓ **Transparent, auditable formulas**
✓ **Consistent formatting and structure**
✓ **Clear labeling and documentation**
✓ **Full integration and reconciliation**
✓ **Control account architecture**

### Formula Standards:

✓ **Absolute references for inputs**: `=$B$6`
✓ **Relative references for time series**: `=B10`
✓ **Named ranges not used** (for transparency)
✓ **No circular references** (unless required)
✓ **Consistent formula structure across periods**

## Regenerating the Model

If you need to rebuild the model:

```bash
python3 build_financial_model.py
```

This will create a fresh `Financial_Model_3_Statement.xlsx` file with all formulas and structure.

## Technical Details

**Built with:**
- Python 3.x
- openpyxl library

**File Format:**
- Excel .xlsx format
- Compatible with Excel 2007+, Google Sheets, LibreOffice Calc

**Color Coding:**
- Blue headers: Timeline periods
- Light blue: Section headers
- Green: Asset-related items
- Orange: Liability-related items
- Yellow: Key calculated totals
- Red: Balance check (should always be 0)

## Model Validation

### Check These Items:

1. **Balance Sheet Balance Check** (should be 0 every month)
2. **Cash Flow Reconciliation** (Opening + Change = Closing)
3. **Retained Earnings** (should roll forward correctly)
4. **All Control Accounts** (Opening + Increases - Decreases = Closing)

### Common Issues:

- If Balance Sheet doesn't balance: Check control account formulas
- If Cash doesn't reconcile: Check Cash Flow Statement links
- If results look wrong: Verify Inputs sheet assumptions

## License

This model is provided as-is for educational and business purposes.

## Support

For issues or questions, please refer to the model documentation or review the Python script that generates the model.

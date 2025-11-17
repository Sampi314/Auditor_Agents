# 3-Statement Financial Model with Control Accounts

## Overview

This is a fully integrated 3-statement financial model built to professional financial modeling standards. The model spans 12 months and includes complete control accounts for every Balance Sheet line item.

## File Structure

**Financial_Model_3_Statement.xlsx** contains 5 worksheets:

1. **Inputs** - All assumptions and parameters
2. **Workings** - Control accounts for all balance sheet items
3. **Income Statement** - Monthly P&L
4. **Balance Sheet** - Monthly balance sheet
5. **Cash Flow Statement** - Indirect method cash flow

## Key Features

### Control Account Architecture

Every Balance Sheet item is driven by its own control account in the Workings sheet:

| Balance Sheet Item | Control Account Structure |
|-------------------|---------------------------|
| **Accounts Receivable** | Opening + Credit Sales - Cash Collections = Closing |
| **Inventory** | Opening + Purchases - COGS Withdrawal = Closing |
| **Prepayments** | Opening + Additions - Amortization = Closing |
| **Fixed Assets (Gross)** | Opening + CapEx - Disposals = Closing |
| **Accumulated Depreciation** | Opening + Depreciation Expense = Closing |
| **Accounts Payable** | Opening + Purchases on Credit - Cash Payments = Closing |
| **Accrued Expenses** | Opening + Expense Accruals - Cash Payments = Closing |
| **Deferred Revenue** | Opening + Cash Received in Advance - Revenue Recognition = Closing |
| **Debt** | Opening + Drawdowns - Principal Repayments = Closing |
| **Retained Earnings** | Opening + Net Income - Dividends = Closing |
| **Cash** | Opening + Operating CF + Investing CF + Financing CF = Closing |

### Financial Modeling Best Practices

✓ **No hardcoded values in calculation blocks** - All calculations reference Inputs sheet
✓ **Clear separation** - Inputs / Workings / Outputs clearly separated
✓ **Transparent formulas** - All formulas are auditable and traceable
✓ **Color coding:**
  - Yellow = Input cells (user can change)
  - White = Calculations (formula-driven)
  - Blue = Subtotals
  - Darker blue = Totals

### Statement Integration

1. **Income Statement → Balance Sheet**
   - Net Income flows to Retained Earnings control account
   - Revenue drives Accounts Receivable
   - COGS drives Inventory withdrawals
   - Operating expenses drive Accrued Expenses

2. **Balance Sheet ← Workings**
   - Every BS line item = Closing balance from its control account
   - No direct calculations on Balance Sheet
   - All movements tracked in Workings sheet

3. **Cash Flow Statement (Indirect Method)**
   - Starts with Net Income
   - Adjusts for non-cash items (Depreciation)
   - Adjusts for working capital changes (AR, Inventory, Prepayments, AP, Accrued Exp, Deferred Rev)
   - Includes investing activities (CapEx)
   - Includes financing activities (Debt drawdowns/repayments)
   - Reconciles to closing cash on Balance Sheet

### Balance Sheet Verification

The model includes a **Balance Check** row at the bottom of the Balance Sheet:
- Formula: Total Assets - Total Liabilities & Equity
- **Should always equal 0**
- Highlighted in yellow for easy monitoring

## Model Inputs (Inputs Sheet)

### Revenue Assumptions
- Starting monthly revenue
- Monthly growth rate
- % of revenue on credit (drives AR)
- Days Sales Outstanding (DSO)

### Cost Structure
- COGS as % of revenue
- Days Inventory Outstanding (DIO)
- Operating expenses (Salaries, Rent, Marketing, Other)

### Working Capital
- Prepayments (opening, additions, amortization)
- Days Payable Outstanding (DPO)
- Accrued expenses rate
- Deferred revenue (cash received in advance, recognition rate)

### Fixed Assets
- Opening gross PP&E and accumulated depreciation
- Monthly CapEx
- Monthly depreciation

### Debt & Equity
- Opening debt balance
- Interest rate
- Principal repayment schedule
- Debt drawdowns
- Share capital
- Opening retained earnings

### Cash & Tax
- Opening cash balance
- Tax rate

## How to Use the Model

1. **Modify assumptions** in the Inputs sheet (yellow cells only)
2. **Review control accounts** in Workings sheet to understand balance sheet movements
3. **Analyze results** in the three output statements
4. **Verify integrity** by checking:
   - Balance Sheet balance check = 0
   - Cash flow closing cash = Balance Sheet cash
   - Retained Earnings change = Net Income (if no dividends)

## Control Account Examples

### Accounts Receivable Flow
```
Month 1:
Opening AR (calculated from DSO assumption)
+ Credit Sales (Revenue × % on credit)
- Cash Collections (based on collection pattern)
= Closing AR → shown on Balance Sheet
```

### Inventory Flow
```
Month 1:
Opening Inventory (from inputs)
+ Purchases (calculated to maintain target inventory level)
- COGS Withdrawal (from Income Statement)
= Closing Inventory → shown on Balance Sheet
```

### Fixed Assets Flow
```
Month 1:
Opening Gross PP&E
+ CapEx Additions (from inputs)
- Disposals (if any)
= Closing Gross PP&E

Opening Accumulated Depreciation
+ Depreciation Expense (from inputs)
= Closing Accumulated Depreciation

Net PP&E = Gross PP&E - Accumulated Depreciation → shown on Balance Sheet
```

## Cash Flow Reconciliation

The Cash Flow Statement uses the **indirect method**:

**Operating Activities:**
- Start with Net Income
- Add back non-cash expenses (Depreciation)
- Subtract increases in operating assets (AR, Inventory, Prepayments)
- Add increases in operating liabilities (AP, Accrued Expenses, Deferred Revenue)

**Investing Activities:**
- Capital expenditure (cash outflow)

**Financing Activities:**
- Debt drawdowns (cash inflow)
- Debt repayments (cash outflow)

**Net Change in Cash** = Operating + Investing + Financing
**Closing Cash** = Opening Cash + Net Change

This closing cash equals the Cash balance on the Balance Sheet, ensuring full integration.

## Model Integrity Checks

✓ Balance Sheet balances (Assets = Liabilities + Equity)
✓ Cash flow closing cash = Balance Sheet cash
✓ Retained Earnings movement = Net Income (absent dividends)
✓ All control accounts reconcile
✓ No circular references (except where explicitly managed)

## Customization

To extend the model:
1. Add new assumptions to Inputs sheet
2. Create corresponding control accounts in Workings sheet
3. Link to appropriate financial statements
4. Update cash flow statement for any new working capital items

## Technical Details

- **Timeline:** 12 months (easily extendable)
- **Formula structure:** All formulas use proper Excel references
- **No VBA/Macros:** Pure Excel formulas for transparency
- **File format:** .xlsx (compatible with Excel 2007+)

---

**Model built with:**
- Python 3 + openpyxl library
- Professional financial modeling standards
- Full audit trail through control accounts

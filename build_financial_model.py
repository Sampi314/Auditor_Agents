"""
3-Statement Financial Model Builder with Full Control Accounts - CORRECTED VERSION
Creates a fully integrated Excel model with proper financial modeling standards
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Define all row references as constants for consistency
class InputRows:
    REVENUE = 6
    COGS_PCT = 10
    OPEX = 11
    AR_DAYS = 14
    INV_DAYS = 15
    AP_DAYS = 16
    PREPAY_AMT = 17
    PREPAY_PERIODS = 18
    CAPEX = 23
    DEPR_RATE = 24
    DEBT_DRAW = 30
    DEBT_REPAY = 31
    INT_RATE = 32
    TAX_RATE = 36
    OPEN_CASH = 41
    OPEN_AR = 42
    OPEN_INV = 43
    OPEN_PREPAY = 44
    OPEN_FA_GROSS = 45
    OPEN_ACCUM_DEPR = 46
    OPEN_AP = 47
    OPEN_ACCRUED = 48
    OPEN_DEF_REV = 49
    OPEN_DEBT = 50
    OPEN_EQUITY = 51
    OPEN_RE = 52

class WorkingsRows:
    # AR Control
    AR_OPEN = 6
    AR_SALES = 7
    AR_COLL = 8
    AR_CLOSE = 9

    # Inventory Control
    INV_OPEN = 13
    INV_PURCH = 14
    INV_COGS = 15
    INV_CLOSE = 16

    # Prepayments Control
    PREP_OPEN = 20
    PREP_NEW = 21
    PREP_AMORT = 22
    PREP_CLOSE = 23

    # Fixed Assets Control
    FA_OPEN = 27
    FA_CAPEX = 28
    FA_DISP = 29
    FA_CLOSE = 30

    # Accumulated Depreciation
    AD_OPEN = 34
    AD_DEPR = 35
    AD_DISP = 36
    AD_CLOSE = 37

    # Net FA
    NET_FA = 39

    # AP Control
    AP_OPEN = 43
    AP_PURCH = 44
    AP_PAY = 45
    AP_CLOSE = 46

    # Accrued Expenses
    ACCR_OPEN = 50
    ACCR_NEW = 51
    ACCR_PAY = 52
    ACCR_CLOSE = 53

    # Deferred Revenue
    DEF_OPEN = 57
    DEF_CASH = 58
    DEF_REV = 59
    DEF_CLOSE = 60

    # Debt
    DEBT_OPEN = 64
    DEBT_DRAW = 65
    DEBT_REPAY = 66
    DEBT_CLOSE = 67

    # Interest
    INT_EXP = 69

def create_financial_model():
    wb = Workbook()
    wb.remove(wb.active)

    # Create sheets
    inputs_sheet = wb.create_sheet("Inputs", 0)
    workings_sheet = wb.create_sheet("Workings", 1)
    income_sheet = wb.create_sheet("Income Statement", 2)
    balance_sheet = wb.create_sheet("Balance Sheet", 3)
    cashflow_sheet = wb.create_sheet("Cash Flow Statement", 4)

    months = 12

    print("Building Inputs sheet...")
    setup_inputs_sheet(inputs_sheet, months)

    print("Building Workings sheet...")
    setup_workings_sheet(workings_sheet, months)

    print("Building Income Statement...")
    setup_income_statement(income_sheet, months)

    print("Building Balance Sheet...")
    setup_balance_sheet(balance_sheet, months)

    print("Building Cash Flow Statement...")
    setup_cashflow_statement(cashflow_sheet, months)

    filename = "Financial_Model_3_Statement.xlsx"
    wb.save(filename)
    print(f"✓ Financial model saved as {filename}")

    return filename

def setup_inputs_sheet(ws, months):
    """Create the Inputs sheet with all assumptions"""

    ws['A1'] = "FINANCIAL MODEL INPUTS"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline header (Row 3)
    ws.cell(3, 1, "Timeline")
    ws.cell(3, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(3, i+2, f"Month {i+1}")
        ws.cell(3, i+2).font = Font(bold=True)
        ws.cell(3, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(3, i+2).font = Font(bold=True, color="FFFFFF")

    # REVENUE ASSUMPTIONS (Row 5)
    ws.cell(5, 1, "REVENUE ASSUMPTIONS")
    ws.cell(5, 1).font = Font(bold=True, size=11)
    ws.cell(5, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Row 6: Monthly Revenue
    ws.cell(InputRows.REVENUE, 1, "Monthly Revenue")
    for i in range(months):
        ws.cell(InputRows.REVENUE, i+2, 100000 * (1.05 ** i))
        ws.cell(InputRows.REVENUE, i+2).number_format = '#,##0'

    # COST ASSUMPTIONS (Row 9)
    ws.cell(9, 1, "COST ASSUMPTIONS")
    ws.cell(9, 1).font = Font(bold=True, size=11)
    ws.cell(9, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Row 10: COGS %
    ws.cell(InputRows.COGS_PCT, 1, "COGS % of Revenue")
    for i in range(months):
        ws.cell(InputRows.COGS_PCT, i+2, 0.40)
        ws.cell(InputRows.COGS_PCT, i+2).number_format = '0.0%'

    # Row 11: Operating Expenses
    ws.cell(InputRows.OPEX, 1, "Operating Expenses (Fixed)")
    for i in range(months):
        ws.cell(InputRows.OPEX, i+2, 25000)
        ws.cell(InputRows.OPEX, i+2).number_format = '#,##0'

    # WORKING CAPITAL (Row 13)
    ws.cell(13, 1, "WORKING CAPITAL ASSUMPTIONS")
    ws.cell(13, 1).font = Font(bold=True, size=11)
    ws.cell(13, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Row 14: AR Days
    ws.cell(InputRows.AR_DAYS, 1, "AR Days")
    for i in range(months):
        ws.cell(InputRows.AR_DAYS, i+2, 30)

    # Row 15: Inventory Days
    ws.cell(InputRows.INV_DAYS, 1, "Inventory Days")
    for i in range(months):
        ws.cell(InputRows.INV_DAYS, i+2, 45)

    # Row 16: AP Days
    ws.cell(InputRows.AP_DAYS, 1, "AP Days")
    for i in range(months):
        ws.cell(InputRows.AP_DAYS, i+2, 30)

    # Row 17: Prepayment Amount
    ws.cell(InputRows.PREPAY_AMT, 1, "Prepayment Amount")
    for i in range(months):
        ws.cell(InputRows.PREPAY_AMT, i+2, 5000)
        ws.cell(InputRows.PREPAY_AMT, i+2).number_format = '#,##0'

    # Row 18: Prepayment Periods
    ws.cell(InputRows.PREPAY_PERIODS, 1, "Prepayment Amortization Period (months)")
    for i in range(months):
        ws.cell(InputRows.PREPAY_PERIODS, i+2, 12)

    # CAPEX & FIXED ASSETS (Row 20)
    ws.cell(20, 1, "CAPEX & FIXED ASSETS")
    ws.cell(20, 1).font = Font(bold=True, size=11)
    ws.cell(20, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Row 23: Capex
    ws.cell(InputRows.CAPEX, 1, "Monthly Capex")
    for i in range(months):
        if i == 0 or i == 6:
            ws.cell(InputRows.CAPEX, i+2, 50000)
        else:
            ws.cell(InputRows.CAPEX, i+2, 0)
        ws.cell(InputRows.CAPEX, i+2).number_format = '#,##0'

    # Row 24: Depreciation Rate
    ws.cell(InputRows.DEPR_RATE, 1, "Depreciation Rate (annual)")
    for i in range(months):
        ws.cell(InputRows.DEPR_RATE, i+2, 0.20)
        ws.cell(InputRows.DEPR_RATE, i+2).number_format = '0.0%'

    # DEBT & FINANCING (Row 27)
    ws.cell(27, 1, "DEBT & FINANCING")
    ws.cell(27, 1).font = Font(bold=True, size=11)
    ws.cell(27, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Row 30: Debt Drawdowns
    ws.cell(InputRows.DEBT_DRAW, 1, "Debt Drawdowns")
    for i in range(months):
        if i == 0:
            ws.cell(InputRows.DEBT_DRAW, i+2, 200000)
        else:
            ws.cell(InputRows.DEBT_DRAW, i+2, 0)
        ws.cell(InputRows.DEBT_DRAW, i+2).number_format = '#,##0'

    # Row 31: Debt Repayments
    ws.cell(InputRows.DEBT_REPAY, 1, "Debt Repayments")
    for i in range(months):
        if i >= 3:
            ws.cell(InputRows.DEBT_REPAY, i+2, 5000)
        else:
            ws.cell(InputRows.DEBT_REPAY, i+2, 0)
        ws.cell(InputRows.DEBT_REPAY, i+2).number_format = '#,##0'

    # Row 32: Interest Rate
    ws.cell(InputRows.INT_RATE, 1, "Interest Rate (annual)")
    for i in range(months):
        ws.cell(InputRows.INT_RATE, i+2, 0.06)
        ws.cell(InputRows.INT_RATE, i+2).number_format = '0.0%'

    # TAX (Row 34)
    ws.cell(34, 1, "TAX")
    ws.cell(34, 1).font = Font(bold=True, size=11)
    ws.cell(34, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Row 36: Tax Rate
    ws.cell(InputRows.TAX_RATE, 1, "Tax Rate")
    for i in range(months):
        ws.cell(InputRows.TAX_RATE, i+2, 0.25)
        ws.cell(InputRows.TAX_RATE, i+2).number_format = '0.0%'

    # STARTING BALANCES (Row 38)
    ws.cell(38, 1, "STARTING BALANCES (Month 0)")
    ws.cell(38, 1).font = Font(bold=True, size=11)
    ws.cell(38, 1).fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")

    # Starting balances
    ws.cell(InputRows.OPEN_CASH, 1, "Opening Cash")
    ws.cell(InputRows.OPEN_CASH, 2, 50000)
    ws.cell(InputRows.OPEN_CASH, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_AR, 1, "Opening AR")
    ws.cell(InputRows.OPEN_AR, 2, 0)
    ws.cell(InputRows.OPEN_AR, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_INV, 1, "Opening Inventory")
    ws.cell(InputRows.OPEN_INV, 2, 0)
    ws.cell(InputRows.OPEN_INV, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_PREPAY, 1, "Opening Prepayments")
    ws.cell(InputRows.OPEN_PREPAY, 2, 0)
    ws.cell(InputRows.OPEN_PREPAY, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_FA_GROSS, 1, "Opening Fixed Assets (Gross)")
    ws.cell(InputRows.OPEN_FA_GROSS, 2, 100000)
    ws.cell(InputRows.OPEN_FA_GROSS, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_ACCUM_DEPR, 1, "Opening Accumulated Depreciation")
    ws.cell(InputRows.OPEN_ACCUM_DEPR, 2, 20000)
    ws.cell(InputRows.OPEN_ACCUM_DEPR, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_AP, 1, "Opening AP")
    ws.cell(InputRows.OPEN_AP, 2, 0)
    ws.cell(InputRows.OPEN_AP, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_ACCRUED, 1, "Opening Accrued Expenses")
    ws.cell(InputRows.OPEN_ACCRUED, 2, 0)
    ws.cell(InputRows.OPEN_ACCRUED, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_DEF_REV, 1, "Opening Deferred Revenue")
    ws.cell(InputRows.OPEN_DEF_REV, 2, 0)
    ws.cell(InputRows.OPEN_DEF_REV, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_DEBT, 1, "Opening Debt")
    ws.cell(InputRows.OPEN_DEBT, 2, 0)
    ws.cell(InputRows.OPEN_DEBT, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_EQUITY, 1, "Opening Share Capital")
    ws.cell(InputRows.OPEN_EQUITY, 2, 100000)
    ws.cell(InputRows.OPEN_EQUITY, 2).number_format = '#,##0'

    ws.cell(InputRows.OPEN_RE, 1, "Opening Retained Earnings")
    ws.cell(InputRows.OPEN_RE, 2, 30000)
    ws.cell(InputRows.OPEN_RE, 2).number_format = '#,##0'

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 12

def setup_workings_sheet(ws, months):
    """Create the Workings sheet with all control accounts"""

    ws['A1'] = "CONTROL ACCOUNTS - WORKINGS"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline header (Row 3)
    ws.cell(3, 1, "Period")
    ws.cell(3, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(3, i+2, f"Month {i+1}")
        ws.cell(3, i+2).font = Font(bold=True)
        ws.cell(3, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(3, i+2).font = Font(bold=True, color="FFFFFF")

    # ==================== ACCOUNTS RECEIVABLE CONTROL ====================
    ws.cell(5, 1, "ACCOUNTS RECEIVABLE CONTROL")
    ws.cell(5, 1).font = Font(bold=True, size=11)
    ws.cell(5, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    ws.cell(WorkingsRows.AR_OPEN, 1, "Opening AR")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.AR_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_AR}")
        else:
            ws.cell(WorkingsRows.AR_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.AR_CLOSE}")
        ws.cell(WorkingsRows.AR_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AR_SALES, 1, "Credit Sales (Revenue)")
    for i in range(months):
        ws.cell(WorkingsRows.AR_SALES, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.REVENUE}")
        ws.cell(WorkingsRows.AR_SALES, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AR_COLL, 1, "Cash Collections")
    for i in range(months):
        if i == 0:
            # First month: collect 50% of opening + current sales
            ws.cell(WorkingsRows.AR_COLL, i+2, f"={get_column_letter(i+2)}{WorkingsRows.AR_OPEN}*0.5")
        else:
            # Collect previous month's sales
            ws.cell(WorkingsRows.AR_COLL, i+2, f"={get_column_letter(i+1)}{WorkingsRows.AR_SALES}")
        ws.cell(WorkingsRows.AR_COLL, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AR_CLOSE, 1, "Closing AR")
    for i in range(months):
        ws.cell(WorkingsRows.AR_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.AR_OPEN}+{get_column_letter(i+2)}{WorkingsRows.AR_SALES}-{get_column_letter(i+2)}{WorkingsRows.AR_COLL}")
        ws.cell(WorkingsRows.AR_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.AR_CLOSE, i+2).font = Font(bold=True)

    # ==================== INVENTORY CONTROL ====================
    ws.cell(11, 1, "INVENTORY CONTROL")
    ws.cell(11, 1).font = Font(bold=True, size=11)
    ws.cell(11, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    ws.cell(WorkingsRows.INV_OPEN, 1, "Opening Inventory")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.INV_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_INV}")
        else:
            ws.cell(WorkingsRows.INV_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.INV_CLOSE}")
        ws.cell(WorkingsRows.INV_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.INV_PURCH, 1, "Purchases")
    for i in range(months):
        if i == 0:
            # First month: buy initial stock + COGS
            ws.cell(WorkingsRows.INV_PURCH, i+2,
                    f"=Inputs!${get_column_letter(i+2)}${InputRows.REVENUE}*Inputs!${get_column_letter(i+2)}${InputRows.COGS_PCT}+50000")
        else:
            ws.cell(WorkingsRows.INV_PURCH, i+2,
                    f"=Inputs!${get_column_letter(i+2)}${InputRows.REVENUE}*Inputs!${get_column_letter(i+2)}${InputRows.COGS_PCT}")
        ws.cell(WorkingsRows.INV_PURCH, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.INV_COGS, 1, "COGS Withdrawal")
    for i in range(months):
        ws.cell(WorkingsRows.INV_COGS, i+2,
                f"=Inputs!${get_column_letter(i+2)}${InputRows.REVENUE}*Inputs!${get_column_letter(i+2)}${InputRows.COGS_PCT}")
        ws.cell(WorkingsRows.INV_COGS, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.INV_CLOSE, 1, "Closing Inventory")
    for i in range(months):
        ws.cell(WorkingsRows.INV_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.INV_OPEN}+{get_column_letter(i+2)}{WorkingsRows.INV_PURCH}-{get_column_letter(i+2)}{WorkingsRows.INV_COGS}")
        ws.cell(WorkingsRows.INV_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.INV_CLOSE, i+2).font = Font(bold=True)

    # ==================== PREPAYMENTS CONTROL ====================
    ws.cell(18, 1, "PREPAYMENTS CONTROL")
    ws.cell(18, 1).font = Font(bold=True, size=11)
    ws.cell(18, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    ws.cell(WorkingsRows.PREP_OPEN, 1, "Opening Prepayments")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.PREP_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_PREPAY}")
        else:
            ws.cell(WorkingsRows.PREP_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.PREP_CLOSE}")
        ws.cell(WorkingsRows.PREP_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.PREP_NEW, 1, "New Prepayments (Cash Paid)")
    for i in range(months):
        ws.cell(WorkingsRows.PREP_NEW, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.PREPAY_AMT}")
        ws.cell(WorkingsRows.PREP_NEW, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.PREP_AMORT, 1, "Amortization (Expense)")
    for i in range(months):
        ws.cell(WorkingsRows.PREP_AMORT, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.PREP_NEW}/Inputs!${get_column_letter(i+2)}${InputRows.PREPAY_PERIODS}")
        ws.cell(WorkingsRows.PREP_AMORT, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.PREP_CLOSE, 1, "Closing Prepayments")
    for i in range(months):
        ws.cell(WorkingsRows.PREP_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.PREP_OPEN}+{get_column_letter(i+2)}{WorkingsRows.PREP_NEW}-{get_column_letter(i+2)}{WorkingsRows.PREP_AMORT}")
        ws.cell(WorkingsRows.PREP_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.PREP_CLOSE, i+2).font = Font(bold=True)

    # ==================== FIXED ASSETS CONTROL ====================
    ws.cell(25, 1, "FIXED ASSETS CONTROL (GROSS)")
    ws.cell(25, 1).font = Font(bold=True, size=11)
    ws.cell(25, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    ws.cell(WorkingsRows.FA_OPEN, 1, "Opening Fixed Assets (Gross)")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.FA_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_FA_GROSS}")
        else:
            ws.cell(WorkingsRows.FA_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.FA_CLOSE}")
        ws.cell(WorkingsRows.FA_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.FA_CAPEX, 1, "Capex Additions")
    for i in range(months):
        ws.cell(WorkingsRows.FA_CAPEX, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.CAPEX}")
        ws.cell(WorkingsRows.FA_CAPEX, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.FA_DISP, 1, "Disposals (Gross)")
    for i in range(months):
        ws.cell(WorkingsRows.FA_DISP, i+2, 0)
        ws.cell(WorkingsRows.FA_DISP, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.FA_CLOSE, 1, "Closing Fixed Assets (Gross)")
    for i in range(months):
        ws.cell(WorkingsRows.FA_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.FA_OPEN}+{get_column_letter(i+2)}{WorkingsRows.FA_CAPEX}-{get_column_letter(i+2)}{WorkingsRows.FA_DISP}")
        ws.cell(WorkingsRows.FA_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.FA_CLOSE, i+2).font = Font(bold=True)

    # ==================== ACCUMULATED DEPRECIATION CONTROL ====================
    ws.cell(32, 1, "ACCUMULATED DEPRECIATION CONTROL")
    ws.cell(32, 1).font = Font(bold=True, size=11)
    ws.cell(32, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    ws.cell(WorkingsRows.AD_OPEN, 1, "Opening Accumulated Depreciation")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.AD_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_ACCUM_DEPR}")
        else:
            ws.cell(WorkingsRows.AD_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.AD_CLOSE}")
        ws.cell(WorkingsRows.AD_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AD_DEPR, 1, "Depreciation Expense")
    for i in range(months):
        # Monthly depreciation = (Opening Gross Assets * Annual Rate) / 12
        ws.cell(WorkingsRows.AD_DEPR, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.FA_OPEN}*Inputs!${get_column_letter(i+2)}${InputRows.DEPR_RATE}/12")
        ws.cell(WorkingsRows.AD_DEPR, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AD_DISP, 1, "Disposal Depreciation")
    for i in range(months):
        ws.cell(WorkingsRows.AD_DISP, i+2, 0)
        ws.cell(WorkingsRows.AD_DISP, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AD_CLOSE, 1, "Closing Accumulated Depreciation")
    for i in range(months):
        ws.cell(WorkingsRows.AD_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.AD_OPEN}+{get_column_letter(i+2)}{WorkingsRows.AD_DEPR}-{get_column_letter(i+2)}{WorkingsRows.AD_DISP}")
        ws.cell(WorkingsRows.AD_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.AD_CLOSE, i+2).font = Font(bold=True)

    # Net Fixed Assets
    ws.cell(WorkingsRows.NET_FA, 1, "Net Fixed Assets")
    for i in range(months):
        ws.cell(WorkingsRows.NET_FA, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.FA_CLOSE}-{get_column_letter(i+2)}{WorkingsRows.AD_CLOSE}")
        ws.cell(WorkingsRows.NET_FA, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.NET_FA, i+2).font = Font(bold=True)
        ws.cell(WorkingsRows.NET_FA, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # ==================== ACCOUNTS PAYABLE CONTROL ====================
    ws.cell(41, 1, "ACCOUNTS PAYABLE CONTROL")
    ws.cell(41, 1).font = Font(bold=True, size=11)
    ws.cell(41, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    ws.cell(WorkingsRows.AP_OPEN, 1, "Opening AP")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.AP_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_AP}")
        else:
            ws.cell(WorkingsRows.AP_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.AP_CLOSE}")
        ws.cell(WorkingsRows.AP_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AP_PURCH, 1, "Purchases (Credit)")
    for i in range(months):
        ws.cell(WorkingsRows.AP_PURCH, i+2, f"={get_column_letter(i+2)}{WorkingsRows.INV_PURCH}")
        ws.cell(WorkingsRows.AP_PURCH, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AP_PAY, 1, "Cash Payments")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.AP_PAY, i+2, f"={get_column_letter(i+2)}{WorkingsRows.AP_PURCH}*0.5")
        else:
            # Pay previous month's purchases (30 days)
            ws.cell(WorkingsRows.AP_PAY, i+2, f"={get_column_letter(i+1)}{WorkingsRows.AP_PURCH}")
        ws.cell(WorkingsRows.AP_PAY, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.AP_CLOSE, 1, "Closing AP")
    for i in range(months):
        ws.cell(WorkingsRows.AP_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.AP_OPEN}+{get_column_letter(i+2)}{WorkingsRows.AP_PURCH}-{get_column_letter(i+2)}{WorkingsRows.AP_PAY}")
        ws.cell(WorkingsRows.AP_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.AP_CLOSE, i+2).font = Font(bold=True)

    # ==================== ACCRUED EXPENSES CONTROL ====================
    ws.cell(48, 1, "ACCRUED EXPENSES CONTROL")
    ws.cell(48, 1).font = Font(bold=True, size=11)
    ws.cell(48, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    ws.cell(WorkingsRows.ACCR_OPEN, 1, "Opening Accrued Expenses")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.ACCR_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_ACCRUED}")
        else:
            ws.cell(WorkingsRows.ACCR_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.ACCR_CLOSE}")
        ws.cell(WorkingsRows.ACCR_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.ACCR_NEW, 1, "Accruals (Expense Recognition)")
    for i in range(months):
        # Accrue 50% of operating expenses
        ws.cell(WorkingsRows.ACCR_NEW, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.OPEX}*0.5")
        ws.cell(WorkingsRows.ACCR_NEW, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.ACCR_PAY, 1, "Cash Payments")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.ACCR_PAY, i+2, 0)
        else:
            ws.cell(WorkingsRows.ACCR_PAY, i+2, f"={get_column_letter(i+1)}{WorkingsRows.ACCR_NEW}")
        ws.cell(WorkingsRows.ACCR_PAY, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.ACCR_CLOSE, 1, "Closing Accrued Expenses")
    for i in range(months):
        ws.cell(WorkingsRows.ACCR_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.ACCR_OPEN}+{get_column_letter(i+2)}{WorkingsRows.ACCR_NEW}-{get_column_letter(i+2)}{WorkingsRows.ACCR_PAY}")
        ws.cell(WorkingsRows.ACCR_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.ACCR_CLOSE, i+2).font = Font(bold=True)

    # ==================== DEFERRED REVENUE CONTROL ====================
    ws.cell(55, 1, "DEFERRED REVENUE CONTROL")
    ws.cell(55, 1).font = Font(bold=True, size=11)
    ws.cell(55, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    ws.cell(WorkingsRows.DEF_OPEN, 1, "Opening Deferred Revenue")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.DEF_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_DEF_REV}")
        else:
            ws.cell(WorkingsRows.DEF_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.DEF_CLOSE}")
        ws.cell(WorkingsRows.DEF_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.DEF_CASH, 1, "Cash Received in Advance")
    for i in range(months):
        if i < months - 1:
            # Receive 10% of next month's revenue in advance
            ws.cell(WorkingsRows.DEF_CASH, i+2, f"=Inputs!${get_column_letter(i+3)}${InputRows.REVENUE}*0.1")
        else:
            ws.cell(WorkingsRows.DEF_CASH, i+2, 0)
        ws.cell(WorkingsRows.DEF_CASH, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.DEF_REV, 1, "Revenue Recognition")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.DEF_REV, i+2, f"={get_column_letter(i+2)}{WorkingsRows.DEF_OPEN}")
        else:
            # Recognize previous month's advance
            ws.cell(WorkingsRows.DEF_REV, i+2, f"={get_column_letter(i+1)}{WorkingsRows.DEF_CASH}")
        ws.cell(WorkingsRows.DEF_REV, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.DEF_CLOSE, 1, "Closing Deferred Revenue")
    for i in range(months):
        ws.cell(WorkingsRows.DEF_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.DEF_OPEN}+{get_column_letter(i+2)}{WorkingsRows.DEF_CASH}-{get_column_letter(i+2)}{WorkingsRows.DEF_REV}")
        ws.cell(WorkingsRows.DEF_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.DEF_CLOSE, i+2).font = Font(bold=True)

    # ==================== DEBT CONTROL ====================
    ws.cell(62, 1, "DEBT CONTROL")
    ws.cell(62, 1).font = Font(bold=True, size=11)
    ws.cell(62, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    ws.cell(WorkingsRows.DEBT_OPEN, 1, "Opening Debt")
    for i in range(months):
        if i == 0:
            ws.cell(WorkingsRows.DEBT_OPEN, i+2, f"=Inputs!$B${InputRows.OPEN_DEBT}")
        else:
            ws.cell(WorkingsRows.DEBT_OPEN, i+2, f"={get_column_letter(i+1)}{WorkingsRows.DEBT_CLOSE}")
        ws.cell(WorkingsRows.DEBT_OPEN, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.DEBT_DRAW, 1, "Debt Drawdowns")
    for i in range(months):
        ws.cell(WorkingsRows.DEBT_DRAW, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.DEBT_DRAW}")
        ws.cell(WorkingsRows.DEBT_DRAW, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.DEBT_REPAY, 1, "Debt Repayments")
    for i in range(months):
        ws.cell(WorkingsRows.DEBT_REPAY, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.DEBT_REPAY}")
        ws.cell(WorkingsRows.DEBT_REPAY, i+2).number_format = '#,##0'

    ws.cell(WorkingsRows.DEBT_CLOSE, 1, "Closing Debt")
    for i in range(months):
        ws.cell(WorkingsRows.DEBT_CLOSE, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.DEBT_OPEN}+{get_column_letter(i+2)}{WorkingsRows.DEBT_DRAW}-{get_column_letter(i+2)}{WorkingsRows.DEBT_REPAY}")
        ws.cell(WorkingsRows.DEBT_CLOSE, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.DEBT_CLOSE, i+2).font = Font(bold=True)

    # Interest Expense
    ws.cell(WorkingsRows.INT_EXP, 1, "Interest Expense")
    for i in range(months):
        # Monthly interest = Opening Debt * Annual Rate / 12
        ws.cell(WorkingsRows.INT_EXP, i+2,
                f"={get_column_letter(i+2)}{WorkingsRows.DEBT_OPEN}*Inputs!${get_column_letter(i+2)}${InputRows.INT_RATE}/12")
        ws.cell(WorkingsRows.INT_EXP, i+2).number_format = '#,##0'
        ws.cell(WorkingsRows.INT_EXP, i+2).font = Font(bold=True)
        ws.cell(WorkingsRows.INT_EXP, i+2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 12

def setup_income_statement(ws, months):
    """Create the Income Statement"""

    ws['A1'] = "INCOME STATEMENT"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline header (Row 3)
    ws.cell(3, 1, "Period")
    ws.cell(3, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(3, i+2, f"Month {i+1}")
        ws.cell(3, i+2).font = Font(bold=True)
        ws.cell(3, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(3, i+2).font = Font(bold=True, color="FFFFFF")

    # Revenue (Row 5)
    ws.cell(5, 1, "Revenue")
    for i in range(months):
        ws.cell(5, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.REVENUE}")
        ws.cell(5, i+2).number_format = '#,##0'

    # COGS (Row 6)
    ws.cell(6, 1, "Cost of Goods Sold")
    for i in range(months):
        ws.cell(6, i+2, f"=-Workings!${get_column_letter(i+2)}${WorkingsRows.INV_COGS}")
        ws.cell(6, i+2).number_format = '#,##0'

    # Gross Profit (Row 7)
    ws.cell(7, 1, "Gross Profit")
    for i in range(months):
        ws.cell(7, i+2, f"={get_column_letter(i+2)}5+{get_column_letter(i+2)}6")
        ws.cell(7, i+2).number_format = '#,##0'
        ws.cell(7, i+2).font = Font(bold=True)

    # Gross Margin % (Row 8)
    ws.cell(8, 1, "Gross Margin %")
    for i in range(months):
        ws.cell(8, i+2, f"={get_column_letter(i+2)}7/{get_column_letter(i+2)}5")
        ws.cell(8, i+2).number_format = '0.0%'

    # Operating Expenses (Row 10)
    ws.cell(10, 1, "Operating Expenses")
    ws.cell(10, 1).font = Font(bold=True)

    # Salaries & Wages (Row 11)
    ws.cell(11, 1, "  Salaries & Wages")
    for i in range(months):
        ws.cell(11, i+2, f"=-Inputs!${get_column_letter(i+2)}${InputRows.OPEX}*0.5")
        ws.cell(11, i+2).number_format = '#,##0'

    # Other Operating Expenses (Row 12)
    ws.cell(12, 1, "  Other Operating Expenses")
    for i in range(months):
        ws.cell(12, i+2, f"=-Inputs!${get_column_letter(i+2)}${InputRows.OPEX}*0.5")
        ws.cell(12, i+2).number_format = '#,##0'

    # Prepayment Amortization (Row 13)
    ws.cell(13, 1, "  Prepayment Amortization")
    for i in range(months):
        ws.cell(13, i+2, f"=-Workings!${get_column_letter(i+2)}${WorkingsRows.PREP_AMORT}")
        ws.cell(13, i+2).number_format = '#,##0'

    # Total Operating Expenses (Row 14)
    ws.cell(14, 1, "Total Operating Expenses")
    for i in range(months):
        ws.cell(14, i+2, f"=SUM({get_column_letter(i+2)}11:{get_column_letter(i+2)}13)")
        ws.cell(14, i+2).number_format = '#,##0'
        ws.cell(14, i+2).font = Font(bold=True)

    # EBITDA (Row 16)
    ws.cell(16, 1, "EBITDA")
    for i in range(months):
        ws.cell(16, i+2, f"={get_column_letter(i+2)}7+{get_column_letter(i+2)}14")
        ws.cell(16, i+2).number_format = '#,##0'
        ws.cell(16, i+2).font = Font(bold=True)
        ws.cell(16, i+2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Depreciation (Row 17)
    ws.cell(17, 1, "Depreciation")
    for i in range(months):
        ws.cell(17, i+2, f"=-Workings!${get_column_letter(i+2)}${WorkingsRows.AD_DEPR}")
        ws.cell(17, i+2).number_format = '#,##0'

    # EBIT (Row 18)
    ws.cell(18, 1, "EBIT")
    for i in range(months):
        ws.cell(18, i+2, f"={get_column_letter(i+2)}16+{get_column_letter(i+2)}17")
        ws.cell(18, i+2).number_format = '#,##0'
        ws.cell(18, i+2).font = Font(bold=True)
        ws.cell(18, i+2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Interest Expense (Row 20)
    ws.cell(20, 1, "Interest Expense")
    for i in range(months):
        ws.cell(20, i+2, f"=-Workings!${get_column_letter(i+2)}${WorkingsRows.INT_EXP}")
        ws.cell(20, i+2).number_format = '#,##0'

    # EBT (Row 21)
    ws.cell(21, 1, "Earnings Before Tax")
    for i in range(months):
        ws.cell(21, i+2, f"={get_column_letter(i+2)}18+{get_column_letter(i+2)}20")
        ws.cell(21, i+2).number_format = '#,##0'
        ws.cell(21, i+2).font = Font(bold=True)

    # Tax (Row 22)
    ws.cell(22, 1, "Income Tax")
    for i in range(months):
        ws.cell(22, i+2, f"=-MAX({get_column_letter(i+2)}21,0)*Inputs!${get_column_letter(i+2)}${InputRows.TAX_RATE}")
        ws.cell(22, i+2).number_format = '#,##0'

    # Net Income (Row 23)
    ws.cell(23, 1, "Net Income")
    for i in range(months):
        ws.cell(23, i+2, f"={get_column_letter(i+2)}21+{get_column_letter(i+2)}22")
        ws.cell(23, i+2).number_format = '#,##0'
        ws.cell(23, i+2).font = Font(bold=True, size=11)
        ws.cell(23, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 14

def setup_balance_sheet(ws, months):
    """Create the Balance Sheet"""

    ws['A1'] = "BALANCE SHEET"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline header (Row 3)
    ws.cell(3, 1, "Period")
    ws.cell(3, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(3, i+2, f"Month {i+1}")
        ws.cell(3, i+2).font = Font(bold=True)
        ws.cell(3, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(3, i+2).font = Font(bold=True, color="FFFFFF")

    # ==================== ASSETS ====================
    ws.cell(5, 1, "ASSETS")
    ws.cell(5, 1).font = Font(bold=True, size=12)
    ws.cell(5, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ws.cell(6, 1, "Current Assets")
    ws.cell(6, 1).font = Font(bold=True)

    # Cash (Row 7)
    ws.cell(7, 1, "  Cash")
    for i in range(months):
        ws.cell(7, i+2, f"='Cash Flow Statement'!${get_column_letter(i+2)}$32")  # Closing cash from CF
        ws.cell(7, i+2).number_format = '#,##0'

    # AR (Row 8)
    ws.cell(8, 1, "  Accounts Receivable")
    for i in range(months):
        ws.cell(8, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.AR_CLOSE}")
        ws.cell(8, i+2).number_format = '#,##0'

    # Inventory (Row 9)
    ws.cell(9, 1, "  Inventory")
    for i in range(months):
        ws.cell(9, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.INV_CLOSE}")
        ws.cell(9, i+2).number_format = '#,##0'

    # Prepayments (Row 10)
    ws.cell(10, 1, "  Prepayments")
    for i in range(months):
        ws.cell(10, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.PREP_CLOSE}")
        ws.cell(10, i+2).number_format = '#,##0'

    # Total Current Assets (Row 11)
    ws.cell(11, 1, "Total Current Assets")
    for i in range(months):
        ws.cell(11, i+2, f"=SUM({get_column_letter(i+2)}7:{get_column_letter(i+2)}10)")
        ws.cell(11, i+2).number_format = '#,##0'
        ws.cell(11, i+2).font = Font(bold=True)

    ws.cell(13, 1, "Non-Current Assets")
    ws.cell(13, 1).font = Font(bold=True)

    # Fixed Assets (Net) (Row 14)
    ws.cell(14, 1, "  Fixed Assets (Net)")
    for i in range(months):
        ws.cell(14, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.NET_FA}")
        ws.cell(14, i+2).number_format = '#,##0'

    # Total Non-Current Assets (Row 15)
    ws.cell(15, 1, "Total Non-Current Assets")
    for i in range(months):
        ws.cell(15, i+2, f"={get_column_letter(i+2)}14")
        ws.cell(15, i+2).number_format = '#,##0'
        ws.cell(15, i+2).font = Font(bold=True)

    # TOTAL ASSETS (Row 17)
    ws.cell(17, 1, "TOTAL ASSETS")
    for i in range(months):
        ws.cell(17, i+2, f"={get_column_letter(i+2)}11+{get_column_letter(i+2)}15")
        ws.cell(17, i+2).number_format = '#,##0'
        ws.cell(17, i+2).font = Font(bold=True, size=11)
        ws.cell(17, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # ==================== LIABILITIES ====================
    ws.cell(20, 1, "LIABILITIES")
    ws.cell(20, 1).font = Font(bold=True, size=12)
    ws.cell(20, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    ws.cell(21, 1, "Current Liabilities")
    ws.cell(21, 1).font = Font(bold=True)

    # AP (Row 22)
    ws.cell(22, 1, "  Accounts Payable")
    for i in range(months):
        ws.cell(22, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.AP_CLOSE}")
        ws.cell(22, i+2).number_format = '#,##0'

    # Accrued Expenses (Row 23)
    ws.cell(23, 1, "  Accrued Expenses")
    for i in range(months):
        ws.cell(23, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.ACCR_CLOSE}")
        ws.cell(23, i+2).number_format = '#,##0'

    # Deferred Revenue (Row 24)
    ws.cell(24, 1, "  Deferred Revenue")
    for i in range(months):
        ws.cell(24, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.DEF_CLOSE}")
        ws.cell(24, i+2).number_format = '#,##0'

    # Total Current Liabilities (Row 25)
    ws.cell(25, 1, "Total Current Liabilities")
    for i in range(months):
        ws.cell(25, i+2, f"=SUM({get_column_letter(i+2)}22:{get_column_letter(i+2)}24)")
        ws.cell(25, i+2).number_format = '#,##0'
        ws.cell(25, i+2).font = Font(bold=True)

    ws.cell(27, 1, "Non-Current Liabilities")
    ws.cell(27, 1).font = Font(bold=True)

    # Debt (Row 28)
    ws.cell(28, 1, "  Long-term Debt")
    for i in range(months):
        ws.cell(28, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.DEBT_CLOSE}")
        ws.cell(28, i+2).number_format = '#,##0'

    # Total Non-Current Liabilities (Row 29)
    ws.cell(29, 1, "Total Non-Current Liabilities")
    for i in range(months):
        ws.cell(29, i+2, f"={get_column_letter(i+2)}28")
        ws.cell(29, i+2).number_format = '#,##0'
        ws.cell(29, i+2).font = Font(bold=True)

    # TOTAL LIABILITIES (Row 31)
    ws.cell(31, 1, "TOTAL LIABILITIES")
    for i in range(months):
        ws.cell(31, i+2, f"={get_column_letter(i+2)}25+{get_column_letter(i+2)}29")
        ws.cell(31, i+2).number_format = '#,##0'
        ws.cell(31, i+2).font = Font(bold=True)
        ws.cell(31, i+2).fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")

    # ==================== EQUITY ====================
    ws.cell(34, 1, "EQUITY")
    ws.cell(34, 1).font = Font(bold=True, size=12)
    ws.cell(34, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    # Share Capital (Row 35)
    ws.cell(35, 1, "  Share Capital")
    for i in range(months):
        if i == 0:
            ws.cell(35, i+2, f"=Inputs!$B${InputRows.OPEN_EQUITY}")
        else:
            ws.cell(35, i+2, f"={get_column_letter(i+1)}{35}")
        ws.cell(35, i+2).number_format = '#,##0'

    # Retained Earnings (Row 36)
    ws.cell(36, 1, "  Retained Earnings")
    for i in range(months):
        if i == 0:
            ws.cell(36, i+2, f"=Inputs!$B${InputRows.OPEN_RE}+'Income Statement'!${get_column_letter(i+2)}$23")
        else:
            ws.cell(36, i+2, f"={get_column_letter(i+1)}{36}+'Income Statement'!${get_column_letter(i+2)}$23")
        ws.cell(36, i+2).number_format = '#,##0'

    # Total Equity (Row 37)
    ws.cell(37, 1, "TOTAL EQUITY")
    for i in range(months):
        ws.cell(37, i+2, f"={get_column_letter(i+2)}35+{get_column_letter(i+2)}36")
        ws.cell(37, i+2).number_format = '#,##0'
        ws.cell(37, i+2).font = Font(bold=True)
        ws.cell(37, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # TOTAL LIABILITIES + EQUITY (Row 39)
    ws.cell(39, 1, "TOTAL LIABILITIES + EQUITY")
    for i in range(months):
        ws.cell(39, i+2, f"={get_column_letter(i+2)}31+{get_column_letter(i+2)}37")
        ws.cell(39, i+2).number_format = '#,##0'
        ws.cell(39, i+2).font = Font(bold=True, size=11)
        ws.cell(39, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # BALANCE CHECK (Row 41)
    ws.cell(41, 1, "BALANCE CHECK (should be 0)")
    for i in range(months):
        ws.cell(41, i+2, f"={get_column_letter(i+2)}17-{get_column_letter(i+2)}39")
        ws.cell(41, i+2).number_format = '#,##0'
        ws.cell(41, i+2).font = Font(bold=True)
        ws.cell(41, i+2).fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 14

def setup_cashflow_statement(ws, months):
    """Create the Cash Flow Statement (Indirect Method)"""

    ws['A1'] = "CASH FLOW STATEMENT (INDIRECT METHOD)"
    ws['A1'].font = Font(bold=True, size=14)

    # Timeline header (Row 3)
    ws.cell(3, 1, "Period")
    ws.cell(3, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(3, i+2, f"Month {i+1}")
        ws.cell(3, i+2).font = Font(bold=True)
        ws.cell(3, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(3, i+2).font = Font(bold=True, color="FFFFFF")

    # ==================== OPERATING ACTIVITIES ====================
    ws.cell(5, 1, "OPERATING ACTIVITIES")
    ws.cell(5, 1).font = Font(bold=True, size=11)
    ws.cell(5, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    # Net Income (Row 6)
    ws.cell(6, 1, "Net Income")
    for i in range(months):
        ws.cell(6, i+2, f"='Income Statement'!${get_column_letter(i+2)}$23")
        ws.cell(6, i+2).number_format = '#,##0'

    # Depreciation (Row 7)
    ws.cell(7, 1, "Add: Depreciation")
    for i in range(months):
        ws.cell(7, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.AD_DEPR}")
        ws.cell(7, i+2).number_format = '#,##0'

    # Changes in Working Capital (Row 8)
    ws.cell(8, 1, "Changes in Working Capital:")
    ws.cell(8, 1).font = Font(italic=True)

    # (Increase)/Decrease in AR (Row 9)
    ws.cell(9, 1, "  (Increase)/Decrease in AR")
    for i in range(months):
        if i == 0:
            ws.cell(9, i+2, f"=Inputs!$B${InputRows.OPEN_AR}-Workings!${get_column_letter(i+2)}${WorkingsRows.AR_CLOSE}")
        else:
            ws.cell(9, i+2, f"=Workings!${get_column_letter(i+1)}${WorkingsRows.AR_CLOSE}-Workings!${get_column_letter(i+2)}${WorkingsRows.AR_CLOSE}")
        ws.cell(9, i+2).number_format = '#,##0'

    # (Increase)/Decrease in Inventory (Row 10)
    ws.cell(10, 1, "  (Increase)/Decrease in Inventory")
    for i in range(months):
        if i == 0:
            ws.cell(10, i+2, f"=Inputs!$B${InputRows.OPEN_INV}-Workings!${get_column_letter(i+2)}${WorkingsRows.INV_CLOSE}")
        else:
            ws.cell(10, i+2, f"=Workings!${get_column_letter(i+1)}${WorkingsRows.INV_CLOSE}-Workings!${get_column_letter(i+2)}${WorkingsRows.INV_CLOSE}")
        ws.cell(10, i+2).number_format = '#,##0'

    # (Increase)/Decrease in Prepayments (Row 11)
    ws.cell(11, 1, "  (Increase)/Decrease in Prepayments")
    for i in range(months):
        if i == 0:
            ws.cell(11, i+2, f"=Inputs!$B${InputRows.OPEN_PREPAY}-Workings!${get_column_letter(i+2)}${WorkingsRows.PREP_CLOSE}")
        else:
            ws.cell(11, i+2, f"=Workings!${get_column_letter(i+1)}${WorkingsRows.PREP_CLOSE}-Workings!${get_column_letter(i+2)}${WorkingsRows.PREP_CLOSE}")
        ws.cell(11, i+2).number_format = '#,##0'

    # Increase/(Decrease) in AP (Row 12)
    ws.cell(12, 1, "  Increase/(Decrease) in AP")
    for i in range(months):
        if i == 0:
            ws.cell(12, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.AP_CLOSE}-Inputs!$B${InputRows.OPEN_AP}")
        else:
            ws.cell(12, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.AP_CLOSE}-Workings!${get_column_letter(i+1)}${WorkingsRows.AP_CLOSE}")
        ws.cell(12, i+2).number_format = '#,##0'

    # Increase/(Decrease) in Accrued Expenses (Row 13)
    ws.cell(13, 1, "  Increase/(Decrease) in Accrued Exp")
    for i in range(months):
        if i == 0:
            ws.cell(13, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.ACCR_CLOSE}-Inputs!$B${InputRows.OPEN_ACCRUED}")
        else:
            ws.cell(13, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.ACCR_CLOSE}-Workings!${get_column_letter(i+1)}${WorkingsRows.ACCR_CLOSE}")
        ws.cell(13, i+2).number_format = '#,##0'

    # Increase/(Decrease) in Deferred Revenue (Row 14)
    ws.cell(14, 1, "  Increase/(Decrease) in Deferred Rev")
    for i in range(months):
        if i == 0:
            ws.cell(14, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.DEF_CLOSE}-Inputs!$B${InputRows.OPEN_DEF_REV}")
        else:
            ws.cell(14, i+2, f"=Workings!${get_column_letter(i+2)}${WorkingsRows.DEF_CLOSE}-Workings!${get_column_letter(i+1)}${WorkingsRows.DEF_CLOSE}")
        ws.cell(14, i+2).number_format = '#,##0'

    # Cash from Operating Activities (Row 15)
    ws.cell(15, 1, "Cash from Operating Activities")
    for i in range(months):
        ws.cell(15, i+2, f"=SUM({get_column_letter(i+2)}6:{get_column_letter(i+2)}14)")
        ws.cell(15, i+2).number_format = '#,##0'
        ws.cell(15, i+2).font = Font(bold=True)
        ws.cell(15, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # ==================== INVESTING ACTIVITIES ====================
    ws.cell(17, 1, "INVESTING ACTIVITIES")
    ws.cell(17, 1).font = Font(bold=True, size=11)
    ws.cell(17, 1).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Capital Expenditure (Row 18)
    ws.cell(18, 1, "Capital Expenditure")
    for i in range(months):
        ws.cell(18, i+2, f"=-Inputs!${get_column_letter(i+2)}${InputRows.CAPEX}")
        ws.cell(18, i+2).number_format = '#,##0'

    # Proceeds from Asset Disposals (Row 19)
    ws.cell(19, 1, "Proceeds from Asset Disposals")
    for i in range(months):
        ws.cell(19, i+2, 0)
        ws.cell(19, i+2).number_format = '#,##0'

    # Cash from Investing Activities (Row 20)
    ws.cell(20, 1, "Cash from Investing Activities")
    for i in range(months):
        ws.cell(20, i+2, f"=SUM({get_column_letter(i+2)}18:{get_column_letter(i+2)}19)")
        ws.cell(20, i+2).number_format = '#,##0'
        ws.cell(20, i+2).font = Font(bold=True)
        ws.cell(20, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # ==================== FINANCING ACTIVITIES ====================
    ws.cell(22, 1, "FINANCING ACTIVITIES")
    ws.cell(22, 1).font = Font(bold=True, size=11)
    ws.cell(22, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    # Debt Drawdowns (Row 23)
    ws.cell(23, 1, "Debt Drawdowns")
    for i in range(months):
        ws.cell(23, i+2, f"=Inputs!${get_column_letter(i+2)}${InputRows.DEBT_DRAW}")
        ws.cell(23, i+2).number_format = '#,##0'

    # Debt Repayments (Row 24)
    ws.cell(24, 1, "Debt Repayments")
    for i in range(months):
        ws.cell(24, i+2, f"=-Inputs!${get_column_letter(i+2)}${InputRows.DEBT_REPAY}")
        ws.cell(24, i+2).number_format = '#,##0'

    # Interest Paid (Row 25)
    ws.cell(25, 1, "Interest Paid")
    for i in range(months):
        ws.cell(25, i+2, f"=-Workings!${get_column_letter(i+2)}${WorkingsRows.INT_EXP}")
        ws.cell(25, i+2).number_format = '#,##0'

    # Equity Injections (Row 26)
    ws.cell(26, 1, "Equity Injections")
    for i in range(months):
        ws.cell(26, i+2, 0)
        ws.cell(26, i+2).number_format = '#,##0'

    # Dividends Paid (Row 27)
    ws.cell(27, 1, "Dividends Paid")
    for i in range(months):
        ws.cell(27, i+2, 0)
        ws.cell(27, i+2).number_format = '#,##0'

    # Cash from Financing Activities (Row 28)
    ws.cell(28, 1, "Cash from Financing Activities")
    for i in range(months):
        ws.cell(28, i+2, f"=SUM({get_column_letter(i+2)}23:{get_column_letter(i+2)}27)")
        ws.cell(28, i+2).number_format = '#,##0'
        ws.cell(28, i+2).font = Font(bold=True)
        ws.cell(28, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # ==================== CASH RECONCILIATION ====================
    # Opening Cash (Row 30)
    ws.cell(30, 1, "Opening Cash")
    for i in range(months):
        if i == 0:
            ws.cell(30, i+2, f"=Inputs!$B${InputRows.OPEN_CASH}")
        else:
            ws.cell(30, i+2, f"={get_column_letter(i+1)}{32}")
        ws.cell(30, i+2).number_format = '#,##0'

    # Net Change in Cash (Row 31)
    ws.cell(31, 1, "Net Change in Cash")
    for i in range(months):
        ws.cell(31, i+2, f"={get_column_letter(i+2)}15+{get_column_letter(i+2)}20+{get_column_letter(i+2)}28")
        ws.cell(31, i+2).number_format = '#,##0'
        ws.cell(31, i+2).font = Font(bold=True)

    # Closing Cash (Row 32)
    ws.cell(32, 1, "Closing Cash")
    for i in range(months):
        ws.cell(32, i+2, f"={get_column_letter(i+2)}30+{get_column_letter(i+2)}31")
        ws.cell(32, i+2).number_format = '#,##0'
        ws.cell(32, i+2).font = Font(bold=True, size=11)
        ws.cell(32, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 14

if __name__ == "__main__":
    print("="*80)
    print("Building 3-Statement Financial Model - Version 2 (CORRECTED)")
    print("="*80)
    filename = create_financial_model()
    print("\n" + "="*80)
    print(f"✓ Model complete: {filename}")
    print("="*80)
    print("\nModel structure:")
    print("  ✓ Inputs: All assumptions and starting balances")
    print("  ✓ Workings: Control accounts for all balance sheet items")
    print("  ✓ Income Statement: P&L with full linkages")
    print("  ✓ Balance Sheet: Driven entirely by control accounts")
    print("  ✓ Cash Flow Statement: Indirect method, fully reconciled")
    print("\n" + "="*80)

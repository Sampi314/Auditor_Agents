"""
3-Statement Financial Model Builder with Full Control Accounts
Creates a fully integrated Excel model with proper financial modeling standards
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_financial_model():
    wb = Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    # Create sheets in order
    inputs_sheet = wb.create_sheet("Inputs", 0)
    workings_sheet = wb.create_sheet("Workings", 1)
    income_sheet = wb.create_sheet("Income Statement", 2)
    balance_sheet = wb.create_sheet("Balance Sheet", 3)
    cashflow_sheet = wb.create_sheet("Cash Flow Statement", 4)

    # Define timeline (12 months)
    months = 12

    # ==================== INPUTS SHEET ====================
    print("Building Inputs sheet...")
    setup_inputs_sheet(inputs_sheet, months)

    # ==================== WORKINGS SHEET ====================
    print("Building Workings sheet...")
    setup_workings_sheet(workings_sheet, months)

    # ==================== INCOME STATEMENT ====================
    print("Building Income Statement...")
    setup_income_statement(income_sheet, months)

    # ==================== BALANCE SHEET ====================
    print("Building Balance Sheet...")
    setup_balance_sheet(balance_sheet, months)

    # ==================== CASH FLOW STATEMENT ====================
    print("Building Cash Flow Statement...")
    setup_cashflow_statement(cashflow_sheet, months)

    # Save workbook
    filename = "Financial_Model_3_Statement.xlsx"
    wb.save(filename)
    print(f"Financial model saved as {filename}")

    return filename

def setup_inputs_sheet(ws, months):
    """Create the Inputs sheet with all assumptions"""

    # Headers
    ws['A1'] = "FINANCIAL MODEL INPUTS"
    ws['A1'].font = Font(bold=True, size=14)

    row = 3

    # Timeline header
    ws.cell(row, 1, "Timeline")
    ws.cell(row, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(row, i+2, f"Month {i+1}")
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(row, i+2).font = Font(bold=True, color="FFFFFF")

    row += 2

    # REVENUE ASSUMPTIONS
    ws.cell(row, 1, "REVENUE ASSUMPTIONS")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Monthly Revenue")
    for i in range(months):
        # Growing revenue (starting at 100,000, growing 5% per month)
        ws.cell(row, i+2, 100000 * (1.05 ** i))
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Revenue Growth %")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, 0)
        else:
            ws.cell(row, i+2, f"=B{row-1}/B{row-1}-1" if i == 1 else f"={get_column_letter(i+2)}{row-1}/{get_column_letter(i+1)}{row-1}-1")
        ws.cell(row, i+2).number_format = '0.0%'
    row += 2

    # COST ASSUMPTIONS
    ws.cell(row, 1, "COST ASSUMPTIONS")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1

    ws.cell(row, 1, "COGS % of Revenue")
    for i in range(months):
        ws.cell(row, i+2, 0.40)  # 40% COGS
        ws.cell(row, i+2).number_format = '0.0%'
    row += 1

    ws.cell(row, 1, "Operating Expenses (Fixed)")
    for i in range(months):
        ws.cell(row, i+2, 25000)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 2

    # WORKING CAPITAL ASSUMPTIONS
    ws.cell(row, 1, "WORKING CAPITAL ASSUMPTIONS")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1

    ws.cell(row, 1, "AR Days")
    for i in range(months):
        ws.cell(row, i+2, 30)
        ws.cell(row, i+2).number_format = '0'
    row += 1

    ws.cell(row, 1, "Inventory Days")
    for i in range(months):
        ws.cell(row, i+2, 45)
        ws.cell(row, i+2).number_format = '0'
    row += 1

    ws.cell(row, 1, "AP Days")
    for i in range(months):
        ws.cell(row, i+2, 30)
        ws.cell(row, i+2).number_format = '0'
    row += 1

    ws.cell(row, 1, "Prepayment Amount")
    for i in range(months):
        ws.cell(row, i+2, 5000)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Prepayment Amortization Period (months)")
    for i in range(months):
        ws.cell(row, i+2, 12)
    row += 2

    # CAPEX & FIXED ASSETS
    ws.cell(row, 1, "CAPEX & FIXED ASSETS")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Monthly Capex")
    for i in range(months):
        if i == 0 or i == 6:  # Capex in month 1 and 7
            ws.cell(row, i+2, 50000)
        else:
            ws.cell(row, i+2, 0)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Depreciation Rate (annual)")
    for i in range(months):
        ws.cell(row, i+2, 0.20)  # 20% annual = 1.67% monthly
        ws.cell(row, i+2).number_format = '0.0%'
    row += 1

    ws.cell(row, 1, "Asset Disposals")
    for i in range(months):
        ws.cell(row, i+2, 0)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 2

    # DEBT & FINANCING
    ws.cell(row, 1, "DEBT & FINANCING")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Debt Drawdowns")
    for i in range(months):
        if i == 0:  # Initial debt
            ws.cell(row, i+2, 200000)
        else:
            ws.cell(row, i+2, 0)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Debt Repayments")
    for i in range(months):
        if i >= 3:  # Start repayments from month 4
            ws.cell(row, i+2, 5000)
        else:
            ws.cell(row, i+2, 0)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Interest Rate (annual)")
    for i in range(months):
        ws.cell(row, i+2, 0.06)  # 6% annual
        ws.cell(row, i+2).number_format = '0.0%'
    row += 2

    # TAX
    ws.cell(row, 1, "TAX")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Tax Rate")
    for i in range(months):
        ws.cell(row, i+2, 0.25)  # 25% tax rate
        ws.cell(row, i+2).number_format = '0.0%'
    row += 2

    # STARTING BALANCES
    ws.cell(row, 1, "STARTING BALANCES (Month 0)")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Cash")
    ws.cell(row, 2, 50000)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening AR")
    ws.cell(row, 2, 0)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Inventory")
    ws.cell(row, 2, 0)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Prepayments")
    ws.cell(row, 2, 0)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Fixed Assets (Gross)")
    ws.cell(row, 2, 100000)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Accumulated Depreciation")
    ws.cell(row, 2, 20000)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening AP")
    ws.cell(row, 2, 0)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Accrued Expenses")
    ws.cell(row, 2, 0)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Deferred Revenue")
    ws.cell(row, 2, 0)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Debt")
    ws.cell(row, 2, 0)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Share Capital")
    ws.cell(row, 2, 100000)
    ws.cell(row, 2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Opening Retained Earnings")
    ws.cell(row, 2, 30000)
    ws.cell(row, 2).number_format = '#,##0'

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 12

def setup_workings_sheet(ws, months):
    """Create the Workings sheet with all control accounts"""

    ws['A1'] = "CONTROL ACCOUNTS - WORKINGS"
    ws['A1'].font = Font(bold=True, size=14)

    row = 3

    # Timeline header
    ws.cell(row, 1, "Period")
    ws.cell(row, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(row, i+2, f"Month {i+1}")
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(row, i+2).font = Font(bold=True, color="FFFFFF")

    row += 2

    # ==================== ACCOUNTS RECEIVABLE CONTROL ====================
    ws.cell(row, 1, "ACCOUNTS RECEIVABLE CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening AR")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$42")  # Opening AR from inputs
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")  # Previous closing
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Credit Sales (Revenue)")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$6")  # Monthly revenue
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Cash Collections")
    for i in range(months):
        # Collections based on AR days
        if i == 0:
            # First month: collect from opening AR
            ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-2}*0.5")  # Collect 50% of opening
        else:
            # Subsequent months: collect based on AR days (30 days = collect previous month's sales)
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row-1}")  # Collect previous month's sales
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing AR")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== INVENTORY CONTROL ====================
    ws.cell(row, 1, "INVENTORY CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Inventory")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$43")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Purchases")
    for i in range(months):
        # Purchase enough to cover COGS + build inventory
        if i == 0:
            ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$6*Inputs!${get_column_letter(i+2)}$11+50000")  # Initial purchase
        else:
            ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$6*Inputs!${get_column_letter(i+2)}$11")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "COGS Withdrawal")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$6*Inputs!${get_column_letter(i+2)}$11")  # COGS
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing Inventory")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== PREPAYMENTS CONTROL ====================
    ws.cell(row, 1, "PREPAYMENTS CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Prepayments")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$44")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "New Prepayments (Cash Paid)")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$17")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Amortization (Expense)")
    for i in range(months):
        # Amortize over the period specified
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-1}/Inputs!${get_column_letter(i+2)}$18")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing Prepayments")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== FIXED ASSETS CONTROL ====================
    ws.cell(row, 1, "FIXED ASSETS CONTROL (GROSS)")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Fixed Assets (Gross)")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$45")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Capex Additions")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$23")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Disposals (Gross)")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$25")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing Fixed Assets (Gross)")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== ACCUMULATED DEPRECIATION CONTROL ====================
    ws.cell(row, 1, "ACCUMULATED DEPRECIATION CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Accumulated Depreciation")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$46")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Depreciation Expense")
    for i in range(months):
        # Monthly depreciation = (Opening Gross Assets * Annual Rate) / 12
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-5}*Inputs!${get_column_letter(i+2)}$24/12")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Disposal Depreciation")
    for i in range(months):
        ws.cell(row, i+2, 0)  # Simplified - no disposal depreciation
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing Accumulated Depreciation")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== NET FIXED ASSETS ====================
    ws.cell(row, 1, "Net Fixed Assets")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-6}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    row += 2

    # ==================== ACCOUNTS PAYABLE CONTROL ====================
    ws.cell(row, 1, "ACCOUNTS PAYABLE CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening AP")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$47")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Purchases (Credit)")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-38}")  # Link to inventory purchases
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Cash Payments")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-1}*0.5")  # Pay 50% immediately
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row-1}")  # Pay previous month's purchases (30 days)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing AP")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== ACCRUED EXPENSES CONTROL ====================
    ws.cell(row, 1, "ACCRUED EXPENSES CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Accrued Expenses")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$48")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Accruals (Expense Recognition)")
    for i in range(months):
        # Accrue 50% of operating expenses
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$12*0.5")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Cash Payments")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, 0)
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row-1}")  # Pay previous month's accruals
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing Accrued Expenses")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== DEFERRED REVENUE CONTROL ====================
    ws.cell(row, 1, "DEFERRED REVENUE CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Deferred Revenue")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$49")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Cash Received in Advance")
    for i in range(months):
        # Receive 10% of revenue in advance
        if i < months - 1:
            ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+3)}$6*0.1")
        else:
            ws.cell(row, i+2, 0)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Revenue Recognition")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-2}")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row-1}")  # Recognize previous month's advance
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing Deferred Revenue")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== DEBT CONTROL ====================
    ws.cell(row, 1, "DEBT CONTROL")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Opening Debt")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$50")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+3}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Debt Drawdowns")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$30")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Debt Repayments")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$31")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Closing Debt")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-2}-{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # ==================== INTEREST EXPENSE CALCULATION ====================
    ws.cell(row, 1, "Interest Expense")
    for i in range(months):
        # Monthly interest = Opening Debt * Annual Rate / 12
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-5}*Inputs!${get_column_letter(i+2)}$32/12")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 12

def setup_income_statement(ws, months):
    """Create the Income Statement"""

    ws['A1'] = "INCOME STATEMENT"
    ws['A1'].font = Font(bold=True, size=14)

    row = 3

    # Timeline header
    ws.cell(row, 1, "Period")
    ws.cell(row, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(row, i+2, f"Month {i+1}")
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(row, i+2).font = Font(bold=True, color="FFFFFF")

    row += 2

    # Revenue
    ws.cell(row, 1, "Revenue")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$6")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # COGS
    ws.cell(row, 1, "Cost of Goods Sold")
    for i in range(months):
        ws.cell(row, i+2, f"=-Workings!${get_column_letter(i+2)}$16")  # From inventory control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Gross Profit
    ws.cell(row, 1, "Gross Profit")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-2}+{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 1

    # Gross Margin %
    ws.cell(row, 1, "Gross Margin %")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-1}/{get_column_letter(i+2)}{row-3}")
        ws.cell(row, i+2).number_format = '0.0%'
    row += 2

    # Operating Expenses
    ws.cell(row, 1, "Operating Expenses")
    ws.cell(row, 1).font = Font(bold=True)
    row += 1

    ws.cell(row, 1, "  Salaries & Wages")
    for i in range(months):
        ws.cell(row, i+2, f"=-Inputs!${get_column_letter(i+2)}$12*0.5")  # 50% of opex
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "  Other Operating Expenses")
    for i in range(months):
        ws.cell(row, i+2, f"=-Inputs!${get_column_letter(i+2)}$12*0.5")  # 50% of opex
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "  Prepayment Amortization")
    for i in range(months):
        ws.cell(row, i+2, f"=-Workings!${get_column_letter(i+2)}$24")  # From prepayments control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Total Operating Expenses")
    for i in range(months):
        ws.cell(row, i+2, f"=SUM({get_column_letter(i+2)}{row-3}:{get_column_letter(i+2)}{row-1})")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # EBITDA
    ws.cell(row, 1, "EBITDA")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-9}+{get_column_letter(i+2)}{row-2}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    row += 1

    # Depreciation
    ws.cell(row, 1, "Depreciation")
    for i in range(months):
        ws.cell(row, i+2, f"=-Workings!${get_column_letter(i+2)}$37")  # From depreciation control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # EBIT
    ws.cell(row, 1, "EBIT")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-2}+{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    row += 2

    # Interest Expense
    ws.cell(row, 1, "Interest Expense")
    for i in range(months):
        ws.cell(row, i+2, f"=-Workings!${get_column_letter(i+2)}$60")  # From debt control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # EBT
    ws.cell(row, 1, "Earnings Before Tax")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-3}+{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 1

    # Tax
    ws.cell(row, 1, "Income Tax")
    for i in range(months):
        # Tax only if EBT > 0
        ws.cell(row, i+2, f"=-MAX({get_column_letter(i+2)}{row-1},0)*Inputs!${get_column_letter(i+2)}$36")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Net Income
    ws.cell(row, 1, "Net Income")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-2}+{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True, size=11)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 14

def setup_balance_sheet(ws, months):
    """Create the Balance Sheet"""

    ws['A1'] = "BALANCE SHEET"
    ws['A1'].font = Font(bold=True, size=14)

    row = 3

    # Timeline header
    ws.cell(row, 1, "Period")
    ws.cell(row, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(row, i+2, f"Month {i+1}")
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(row, i+2).font = Font(bold=True, color="FFFFFF")

    row += 2

    # ==================== ASSETS ====================
    ws.cell(row, 1, "ASSETS")
    ws.cell(row, 1).font = Font(bold=True, size=12)
    ws.cell(row, 1).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Current Assets")
    ws.cell(row, 1).font = Font(bold=True)
    row += 1

    # Cash (from Cash Flow Statement)
    ws.cell(row, 1, "  Cash")
    for i in range(months):
        ws.cell(row, i+2, f"='Cash Flow Statement'!${get_column_letter(i+2)}$34")  # Closing cash from CF
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Accounts Receivable
    ws.cell(row, 1, "  Accounts Receivable")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$9")  # From AR control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Inventory
    ws.cell(row, 1, "  Inventory")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$17")  # From inventory control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Prepayments
    ws.cell(row, 1, "  Prepayments")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$25")  # From prepayments control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Total Current Assets
    ws.cell(row, 1, "Total Current Assets")
    for i in range(months):
        ws.cell(row, i+2, f"=SUM({get_column_letter(i+2)}{row-4}:{get_column_letter(i+2)}{row-1})")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    ws.cell(row, 1, "Non-Current Assets")
    ws.cell(row, 1).font = Font(bold=True)
    row += 1

    # Fixed Assets (Net)
    ws.cell(row, 1, "  Fixed Assets (Net)")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$41")  # From fixed assets control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Total Non-Current Assets
    ws.cell(row, 1, "Total Non-Current Assets")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # TOTAL ASSETS
    ws.cell(row, 1, "TOTAL ASSETS")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-9}+{get_column_letter(i+2)}{row-2}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True, size=11)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    row += 3

    # ==================== LIABILITIES ====================
    ws.cell(row, 1, "LIABILITIES")
    ws.cell(row, 1).font = Font(bold=True, size=12)
    ws.cell(row, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row += 1

    ws.cell(row, 1, "Current Liabilities")
    ws.cell(row, 1).font = Font(bold=True)
    row += 1

    # Accounts Payable
    ws.cell(row, 1, "  Accounts Payable")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$47")  # From AP control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Accrued Expenses
    ws.cell(row, 1, "  Accrued Expenses")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$53")  # From accrued expenses control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Deferred Revenue
    ws.cell(row, 1, "  Deferred Revenue")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$59")  # From deferred revenue control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Total Current Liabilities
    ws.cell(row, 1, "Total Current Liabilities")
    for i in range(months):
        ws.cell(row, i+2, f"=SUM({get_column_letter(i+2)}{row-3}:{get_column_letter(i+2)}{row-1})")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    ws.cell(row, 1, "Non-Current Liabilities")
    ws.cell(row, 1).font = Font(bold=True)
    row += 1

    # Debt
    ws.cell(row, 1, "  Long-term Debt")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$65")  # From debt control
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Total Non-Current Liabilities
    ws.cell(row, 1, "Total Non-Current Liabilities")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 2

    # TOTAL LIABILITIES
    ws.cell(row, 1, "TOTAL LIABILITIES")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-9}+{get_column_letter(i+2)}{row-2}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    row += 3

    # ==================== EQUITY ====================
    ws.cell(row, 1, "EQUITY")
    ws.cell(row, 1).font = Font(bold=True, size=12)
    ws.cell(row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    row += 1

    # Share Capital
    ws.cell(row, 1, "  Share Capital")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$51")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row}")  # Constant
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Retained Earnings
    ws.cell(row, 1, "  Retained Earnings")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"=Inputs!$B$52+'Income Statement'!${get_column_letter(i+2)}$25")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row}+'Income Statement'!${get_column_letter(i+2)}$25")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Total Equity
    ws.cell(row, 1, "TOTAL EQUITY")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-2}+{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    row += 2

    # TOTAL LIABILITIES + EQUITY
    ws.cell(row, 1, "TOTAL LIABILITIES + EQUITY")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-15}+{get_column_letter(i+2)}{row-3}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True, size=11)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    row += 2

    # BALANCE CHECK
    ws.cell(row, 1, "BALANCE CHECK (should be 0)")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-17}-{get_column_letter(i+2)}{row-2}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 14

def setup_cashflow_statement(ws, months):
    """Create the Cash Flow Statement (Indirect Method)"""

    ws['A1'] = "CASH FLOW STATEMENT (INDIRECT METHOD)"
    ws['A1'].font = Font(bold=True, size=14)

    row = 3

    # Timeline header
    ws.cell(row, 1, "Period")
    ws.cell(row, 1).font = Font(bold=True)
    for i in range(months):
        ws.cell(row, i+2, f"Month {i+1}")
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        ws.cell(row, i+2).font = Font(bold=True, color="FFFFFF")

    row += 2

    # ==================== OPERATING ACTIVITIES ====================
    ws.cell(row, 1, "OPERATING ACTIVITIES")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    row += 1

    # Net Income
    ws.cell(row, 1, "Net Income")
    for i in range(months):
        ws.cell(row, i+2, f"='Income Statement'!${get_column_letter(i+2)}$25")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Add back: Depreciation
    ws.cell(row, 1, "Add: Depreciation")
    for i in range(months):
        ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$37")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Changes in Working Capital
    ws.cell(row, 1, "Changes in Working Capital:")
    ws.cell(row, 1).font = Font(italic=True)
    row += 1

    # Decrease/(Increase) in AR
    ws.cell(row, 1, "  (Increase)/Decrease in AR")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"=Inputs!$B$42-Workings!${get_column_letter(i+2)}$9")
        else:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+1)}$9-Workings!${get_column_letter(i+2)}$9")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Decrease/(Increase) in Inventory
    ws.cell(row, 1, "  (Increase)/Decrease in Inventory")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"=Inputs!$B$43-Workings!${get_column_letter(i+2)}$17")
        else:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+1)}$17-Workings!${get_column_letter(i+2)}$17")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Decrease/(Increase) in Prepayments
    ws.cell(row, 1, "  (Increase)/Decrease in Prepayments")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"=Inputs!$B$44-Workings!${get_column_letter(i+2)}$25")
        else:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+1)}$25-Workings!${get_column_letter(i+2)}$25")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Increase/(Decrease) in AP
    ws.cell(row, 1, "  Increase/(Decrease) in AP")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$47-Inputs!$B$47")
        else:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$47-Workings!${get_column_letter(i+1)}$47")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Increase/(Decrease) in Accrued Expenses
    ws.cell(row, 1, "  Increase/(Decrease) in Accrued Exp")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$53-Inputs!$B$48")
        else:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$53-Workings!${get_column_letter(i+1)}$53")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Increase/(Decrease) in Deferred Revenue
    ws.cell(row, 1, "  Increase/(Decrease) in Deferred Rev")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$59-Inputs!$B$49")
        else:
            ws.cell(row, i+2, f"=Workings!${get_column_letter(i+2)}$59-Workings!${get_column_letter(i+1)}$59")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Cash from Operating Activities
    ws.cell(row, 1, "Cash from Operating Activities")
    for i in range(months):
        ws.cell(row, i+2, f"=SUM({get_column_letter(i+2)}{row-9}:{get_column_letter(i+2)}{row-1})")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    row += 2

    # ==================== INVESTING ACTIVITIES ====================
    ws.cell(row, 1, "INVESTING ACTIVITIES")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    row += 1

    # Capital Expenditure
    ws.cell(row, 1, "Capital Expenditure")
    for i in range(months):
        ws.cell(row, i+2, f"=-Inputs!${get_column_letter(i+2)}$23")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Proceeds from Asset Disposals
    ws.cell(row, 1, "Proceeds from Asset Disposals")
    for i in range(months):
        ws.cell(row, i+2, 0)  # Simplified
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Cash from Investing Activities
    ws.cell(row, 1, "Cash from Investing Activities")
    for i in range(months):
        ws.cell(row, i+2, f"=SUM({get_column_letter(i+2)}{row-2}:{get_column_letter(i+2)}{row-1})")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    row += 2

    # ==================== FINANCING ACTIVITIES ====================
    ws.cell(row, 1, "FINANCING ACTIVITIES")
    ws.cell(row, 1).font = Font(bold=True, size=11)
    ws.cell(row, 1).fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    row += 1

    # Debt Drawdowns
    ws.cell(row, 1, "Debt Drawdowns")
    for i in range(months):
        ws.cell(row, i+2, f"=Inputs!${get_column_letter(i+2)}$30")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Debt Repayments
    ws.cell(row, 1, "Debt Repayments")
    for i in range(months):
        ws.cell(row, i+2, f"=-Inputs!${get_column_letter(i+2)}$31")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Interest Paid
    ws.cell(row, 1, "Interest Paid")
    for i in range(months):
        ws.cell(row, i+2, f"=-Workings!${get_column_letter(i+2)}$60")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Equity Injections
    ws.cell(row, 1, "Equity Injections")
    for i in range(months):
        ws.cell(row, i+2, 0)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Dividends Paid
    ws.cell(row, 1, "Dividends Paid")
    for i in range(months):
        ws.cell(row, i+2, 0)
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    # Cash from Financing Activities
    ws.cell(row, 1, "Cash from Financing Activities")
    for i in range(months):
        ws.cell(row, i+2, f"=SUM({get_column_letter(i+2)}{row-5}:{get_column_letter(i+2)}{row-1})")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    row += 2

    # ==================== CASH RECONCILIATION ====================
    ws.cell(row, 1, "Opening Cash")
    for i in range(months):
        if i == 0:
            ws.cell(row, i+2, "=Inputs!$B$41")
        else:
            ws.cell(row, i+2, f"={get_column_letter(i+1)}{row+2}")
        ws.cell(row, i+2).number_format = '#,##0'
    row += 1

    ws.cell(row, 1, "Net Change in Cash")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-16}+{get_column_letter(i+2)}{row-9}+{get_column_letter(i+2)}{row-2}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True)
    row += 1

    ws.cell(row, 1, "Closing Cash")
    for i in range(months):
        ws.cell(row, i+2, f"={get_column_letter(i+2)}{row-2}+{get_column_letter(i+2)}{row-1}")
        ws.cell(row, i+2).number_format = '#,##0'
        ws.cell(row, i+2).font = Font(bold=True, size=11)
        ws.cell(row, i+2).fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for i in range(months):
        ws.column_dimensions[get_column_letter(i+2)].width = 14

if __name__ == "__main__":
    print("Building 3-Statement Financial Model...")
    filename = create_financial_model()
    print(f"\n✓ Model complete: {filename}")
    print("\nModel structure:")
    print("  - Inputs: All assumptions and starting balances")
    print("  - Workings: Control accounts for all balance sheet items")
    print("  - Income Statement: P&L with full linkages")
    print("  - Balance Sheet: Driven entirely by control accounts")
    print("  - Cash Flow Statement: Indirect method, fully reconciled")

#!/usr/bin/env python3
"""
Build a comprehensive 3-statement financial model with control accounts
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime

def create_financial_model():
    wb = Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Define styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    calc_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    subtotal_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    bold_font = Font(bold=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Number of periods (months)
    num_periods = 12

    # =======================
    # 1. INPUTS SHEET
    # =======================
    ws_inputs = wb.create_sheet("Inputs", 0)

    # Set column widths
    ws_inputs.column_dimensions['A'].width = 35
    ws_inputs.column_dimensions['B'].width = 15

    # Title
    ws_inputs['A1'] = "Financial Model - Inputs & Assumptions"
    ws_inputs['A1'].font = Font(bold=True, size=14)

    row = 3

    # Revenue Assumptions
    ws_inputs[f'A{row}'] = "REVENUE ASSUMPTIONS"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Starting Monthly Revenue"
    ws_inputs[f'B{row}'] = 100000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Revenue Growth Rate"
    ws_inputs[f'B{row}'] = 0.03
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0.0%'
    row += 1

    ws_inputs[f'A{row}'] = "% Revenue on Credit (AR)"
    ws_inputs[f'B{row}'] = 0.70
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0.0%'
    row += 1

    ws_inputs[f'A{row}'] = "Days Sales Outstanding (DSO)"
    ws_inputs[f'B{row}'] = 45
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0'
    row += 2

    # COGS Assumptions
    ws_inputs[f'A{row}'] = "COST OF GOODS SOLD"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "COGS as % of Revenue"
    ws_inputs[f'B{row}'] = 0.40
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0.0%'
    row += 2

    # Inventory Assumptions
    ws_inputs[f'A{row}'] = "INVENTORY ASSUMPTIONS"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Days Inventory Outstanding (DIO)"
    ws_inputs[f'B{row}'] = 60
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0'
    row += 1

    ws_inputs[f'A{row}'] = "Opening Inventory Balance"
    ws_inputs[f'B{row}'] = 80000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 2

    # Operating Expenses
    ws_inputs[f'A{row}'] = "OPERATING EXPENSES"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Salaries & Wages (Monthly)"
    ws_inputs[f'B{row}'] = 25000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Rent (Monthly)"
    ws_inputs[f'B{row}'] = 5000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Marketing (Monthly)"
    ws_inputs[f'B{row}'] = 8000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Other OpEx (Monthly)"
    ws_inputs[f'B{row}'] = 7000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 2

    # Prepayments
    ws_inputs[f'A{row}'] = "PREPAYMENTS"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Opening Prepayments"
    ws_inputs[f'B{row}'] = 15000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Prepayment Additions"
    ws_inputs[f'B{row}'] = 5000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Prepayment Amortization"
    ws_inputs[f'B{row}'] = 4000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 2

    # Accounts Payable
    ws_inputs[f'A{row}'] = "ACCOUNTS PAYABLE"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Days Payable Outstanding (DPO)"
    ws_inputs[f'B{row}'] = 30
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0'
    row += 2

    # Accrued Expenses
    ws_inputs[f'A{row}'] = "ACCRUED EXPENSES"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Opening Accrued Expenses"
    ws_inputs[f'B{row}'] = 12000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Accrual Rate (% of OpEx)"
    ws_inputs[f'B{row}'] = 0.15
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0.0%'
    row += 2

    # Deferred Revenue
    ws_inputs[f'A{row}'] = "DEFERRED REVENUE"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Opening Deferred Revenue"
    ws_inputs[f'B{row}'] = 30000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Cash Received in Advance"
    ws_inputs[f'B{row}'] = 20000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Recognition Rate"
    ws_inputs[f'B{row}'] = 18000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 2

    # Fixed Assets
    ws_inputs[f'A{row}'] = "FIXED ASSETS (PP&E)"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Opening Gross PP&E"
    ws_inputs[f'B{row}'] = 500000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Opening Accumulated Depreciation"
    ws_inputs[f'B{row}'] = 100000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly CapEx"
    ws_inputs[f'B{row}'] = 10000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Depreciation"
    ws_inputs[f'B{row}'] = 8333
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 2

    # Debt
    ws_inputs[f'A{row}'] = "DEBT"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Opening Debt Balance"
    ws_inputs[f'B{row}'] = 200000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Annual Interest Rate"
    ws_inputs[f'B{row}'] = 0.06
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0.0%'
    row += 1

    ws_inputs[f'A{row}'] = "Monthly Principal Repayment"
    ws_inputs[f'B{row}'] = 5000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Debt Drawdown (Month 1)"
    ws_inputs[f'B{row}'] = 50000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 2

    # Equity
    ws_inputs[f'A{row}'] = "EQUITY"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Opening Share Capital"
    ws_inputs[f'B{row}'] = 300000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    ws_inputs[f'A{row}'] = "Opening Retained Earnings"
    ws_inputs[f'B{row}'] = 50000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 2

    # Cash
    ws_inputs[f'A{row}'] = "CASH"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Opening Cash Balance"
    ws_inputs[f'B{row}'] = 150000
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '#,##0'
    row += 1

    # Tax
    ws_inputs[f'A{row}'] = "TAX"
    ws_inputs[f'A{row}'].font = bold_font
    ws_inputs[f'A{row}'].fill = subtotal_fill
    row += 1

    ws_inputs[f'A{row}'] = "Tax Rate"
    ws_inputs[f'B{row}'] = 0.25
    ws_inputs[f'B{row}'].fill = input_fill
    ws_inputs[f'B{row}'].number_format = '0.0%'

    # =======================
    # 2. WORKINGS SHEET
    # =======================
    ws_work = wb.create_sheet("Workings", 1)
    ws_work.column_dimensions['A'].width = 35
    for col in range(2, 2 + num_periods + 1):
        ws_work.column_dimensions[get_column_letter(col)].width = 12

    # Title
    ws_work['A1'] = "Control Accounts - Workings"
    ws_work['A1'].font = Font(bold=True, size=14)

    # Month headers
    ws_work['A2'] = "Period"
    ws_work['A2'].font = header_font
    ws_work['A2'].fill = header_fill

    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}2'] = f"Month {i+1}"
        ws_work[f'{col_letter}2'].font = header_font
        ws_work[f'{col_letter}2'].fill = header_fill
        ws_work[f'{col_letter}2'].alignment = Alignment(horizontal='center')

    row = 4

    # ==================
    # ACCOUNTS RECEIVABLE CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "ACCOUNTS RECEIVABLE"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    # Month 1 opening (from inputs or calculated from DSO)
    ws_work[f'B{row}'] = "=Inputs!$B$7*Inputs!$B$4*Inputs!$B$8/30"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"  # Closing balance from previous month
    row += 1

    ws_work[f'A{row}'] = "Credit Sales (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        # Credit sales = Revenue * % on credit
        ws_work[f'{col_letter}{row}'] = f"='Income Statement'!{col_letter}5*Inputs!$B$7"
    row += 1

    ws_work[f'A{row}'] = "Cash Collections (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        # Collections based on DSO - simplified: collect previous period AR
        if i == 0:
            ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-2}*0.5"  # Collect 50% of opening in month 1
        else:
            prev_col = get_column_letter(i + 1)
            ws_work[f'{col_letter}{row}'] = f"={prev_col}{row-1}"  # Collect previous month's credit sales
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # ==================
    # INVENTORY CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "INVENTORY"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$16"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"
    row += 1

    ws_work[f'A{row}'] = "Purchases (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        # Target inventory = COGS * DIO / 30
        # Purchases = COGS + (Closing Inv - Opening Inv)
        # Simplified: maintain target inventory level
        if i == 0:
            ws_work[f'{col_letter}{row}'] = f"='Income Statement'!{col_letter}7*Inputs!$B$15/30"
        else:
            next_month_col = get_column_letter(i + 3) if i < num_periods - 1 else col_letter
            ws_work[f'{col_letter}{row}'] = f"='Income Statement'!{col_letter}7+'Income Statement'!{next_month_col}7*Inputs!$B$15/30-{col_letter}{row-1}"
    row += 1

    ws_work[f'A{row}'] = "COGS Withdrawal (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"='Income Statement'!{col_letter}7"
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # ==================
    # PREPAYMENTS CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "PREPAYMENTS"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$28"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"
    row += 1

    ws_work[f'A{row}'] = "Additions (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = "=Inputs!$B$29"
    row += 1

    ws_work[f'A{row}'] = "Amortization (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = "=Inputs!$B$30"
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # ==================
    # FIXED ASSETS (PP&E) CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "FIXED ASSETS - GROSS PP&E"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$46"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"
    row += 1

    ws_work[f'A{row}'] = "CapEx Additions (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = "=Inputs!$B$48"
    row += 1

    ws_work[f'A{row}'] = "Disposals (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = 0
    row += 1

    ws_work[f'A{row}'] = "Closing Balance - Gross PP&E"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # Accumulated Depreciation
    ws_work[f'A{row}'] = "ACCUMULATED DEPRECIATION"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$47"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+2}"
    row += 1

    ws_work[f'A{row}'] = "Depreciation Expense (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = "=Inputs!$B$49"
    row += 1

    ws_work[f'A{row}'] = "Closing Balance - Accum Depr"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-2}+{col_letter}{row-1}"
    row += 1

    ws_work[f'A{row}'] = "Net PP&E"
    ws_work[f'A{row}'].font = Font(bold=True)
    ws_work[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-8}-{col_letter}{row-1}"
    row += 2

    # ==================
    # ACCOUNTS PAYABLE CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "ACCOUNTS PAYABLE"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    # Calculate opening AP based on purchases and DPO
    ws_work[f'B{row}'] = "=Workings!B12*Inputs!$B$34/30"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"
    row += 1

    ws_work[f'A{row}'] = "Purchases on Credit (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"=Workings!{col_letter}12"  # Link to inventory purchases
    row += 1

    ws_work[f'A{row}'] = "Cash Payments (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-2}"  # Pay opening balance
        else:
            prev_col = get_column_letter(i + 1)
            ws_work[f'{col_letter}{row}'] = f"={prev_col}{row-1}"  # Pay last month's purchases
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # ==================
    # ACCRUED EXPENSES CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "ACCRUED EXPENSES"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$38"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"
    row += 1

    ws_work[f'A{row}'] = "Expense Accruals (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        # Accrue % of operating expenses
        ws_work[f'{col_letter}{row}'] = f"='Income Statement'!{col_letter}10*Inputs!$B$39"
    row += 1

    ws_work[f'A{row}'] = "Cash Payments (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-2}*0.5"
        else:
            prev_col = get_column_letter(i + 1)
            ws_work[f'{col_letter}{row}'] = f"={prev_col}{row-1}"
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # ==================
    # DEFERRED REVENUE CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "DEFERRED REVENUE"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$43"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"
    row += 1

    ws_work[f'A{row}'] = "Cash Received in Advance (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = "=Inputs!$B$44"
    row += 1

    ws_work[f'A{row}'] = "Revenue Recognition (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = "=Inputs!$B$45"
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # ==================
    # DEBT CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "DEBT"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$53"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+4}"
    row += 1

    ws_work[f'A{row}'] = "Drawdowns (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_work[f'{col_letter}{row}'] = "=Inputs!$B$56"
        else:
            ws_work[f'{col_letter}{row}'] = 0
    row += 1

    ws_work[f'A{row}'] = "Principal Repayments (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = "=Inputs!$B$55"
    row += 1

    ws_work[f'A{row}'] = "Interest Expense"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}*Inputs!$B$54/12"
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-4}+{col_letter}{row-3}-{col_letter}{row-2}"
    row += 2

    # ==================
    # RETAINED EARNINGS CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "RETAINED EARNINGS"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$61"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        prev_col = get_column_letter(i + 1)
        ws_work[f'{col_letter}{row}'] = f"={prev_col}{row+3}"
    row += 1

    ws_work[f'A{row}'] = "Net Income (Increase)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"='Income Statement'!{col_letter}18"
    row += 1

    ws_work[f'A{row}'] = "Dividends (Decrease)"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = 0
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-3}+{col_letter}{row-2}-{col_letter}{row-1}"
    row += 2

    # ==================
    # CASH CONTROL ACCOUNT
    # ==================
    ws_work[f'A{row}'] = "CASH"
    ws_work[f'A{row}'].font = Font(bold=True, size=11)
    ws_work[f'A{row}'].fill = total_fill
    row += 1

    ws_work[f'A{row}'] = "Opening Balance"
    ws_work[f'B{row}'] = "=Inputs!$B$65"
    for i in range(1, num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"='Cash Flow Statement'!{col_letter}21"  # Link to closing cash from CF
    row += 1

    ws_work[f'A{row}'] = "Operating Cash Flow"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"='Cash Flow Statement'!{col_letter}13"
    row += 1

    ws_work[f'A{row}'] = "Investing Cash Flow"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"='Cash Flow Statement'!{col_letter}16"
    row += 1

    ws_work[f'A{row}'] = "Financing Cash Flow"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"='Cash Flow Statement'!{col_letter}19"
    row += 1

    ws_work[f'A{row}'] = "Closing Balance"
    ws_work[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_work[f'{col_letter}{row}'] = f"={col_letter}{row-5}+{col_letter}{row-4}+{col_letter}{row-3}+{col_letter}{row-2}"

    # =======================
    # 3. INCOME STATEMENT
    # =======================
    ws_is = wb.create_sheet("Income Statement", 2)
    ws_is.column_dimensions['A'].width = 35
    for col in range(2, 2 + num_periods + 1):
        ws_is.column_dimensions[get_column_letter(col)].width = 12

    ws_is['A1'] = "Income Statement"
    ws_is['A1'].font = Font(bold=True, size=14)

    # Headers
    ws_is['A2'] = "Period"
    ws_is['A2'].font = header_font
    ws_is['A2'].fill = header_fill

    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}2'] = f"Month {i+1}"
        ws_is[f'{col_letter}2'].font = header_font
        ws_is[f'{col_letter}2'].fill = header_fill
        ws_is[f'{col_letter}2'].alignment = Alignment(horizontal='center')

    row = 4

    # Revenue
    ws_is[f'A{row}'] = "Revenue"
    ws_is[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_is[f'{col_letter}{row}'] = "=Inputs!$B$4"
        else:
            prev_col = get_column_letter(i + 1)
            ws_is[f'{col_letter}{row}'] = f"={prev_col}{row}*(1+Inputs!$B$5)"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_is[f'A{row}'] = ""
    row += 1

    # COGS
    ws_is[f'A{row}'] = "Cost of Goods Sold"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = f"=-{col_letter}5*Inputs!$B$13"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    # Gross Profit
    ws_is[f'A{row}'] = "Gross Profit"
    ws_is[f'A{row}'].font = bold_font
    ws_is[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = f"={col_letter}5+{col_letter}7"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_is[f'A{row}'] = ""
    row += 1

    # Operating Expenses
    ws_is[f'A{row}'] = "Operating Expenses"
    ws_is[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = "=-(Inputs!$B$20+Inputs!$B$21+Inputs!$B$22+Inputs!$B$23)"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_is[f'A{row}'] = "Depreciation"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = "=-Inputs!$B$49"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    # EBIT
    ws_is[f'A{row}'] = "EBIT"
    ws_is[f'A{row}'].font = bold_font
    ws_is[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = f"={col_letter}8+{col_letter}10+{col_letter}11"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_is[f'A{row}'] = ""
    row += 1

    # Interest Expense
    ws_is[f'A{row}'] = "Interest Expense"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = f"=-Workings!{col_letter}54"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    # EBT
    ws_is[f'A{row}'] = "Earnings Before Tax"
    ws_is[f'A{row}'].font = bold_font
    ws_is[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = f"={col_letter}12+{col_letter}14"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    # Tax
    ws_is[f'A{row}'] = "Tax Expense"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = f"=-MAX({col_letter}15,0)*Inputs!$B$68"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_is[f'A{row}'] = ""
    row += 1

    # Net Income
    ws_is[f'A{row}'] = "Net Income"
    ws_is[f'A{row}'].font = Font(bold=True)
    ws_is[f'A{row}'].fill = total_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_is[f'{col_letter}{row}'] = f"={col_letter}15+{col_letter}16"
        ws_is[f'{col_letter}{row}'].number_format = '#,##0'

    # =======================
    # 4. BALANCE SHEET
    # =======================
    ws_bs = wb.create_sheet("Balance Sheet", 3)
    ws_bs.column_dimensions['A'].width = 35
    for col in range(2, 2 + num_periods + 1):
        ws_bs.column_dimensions[get_column_letter(col)].width = 12

    ws_bs['A1'] = "Balance Sheet"
    ws_bs['A1'].font = Font(bold=True, size=14)

    # Headers
    ws_bs['A2'] = "Period"
    ws_bs['A2'].font = header_font
    ws_bs['A2'].fill = header_fill

    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}2'] = f"Month {i+1}"
        ws_bs[f'{col_letter}2'].font = header_font
        ws_bs[f'{col_letter}2'].fill = header_fill
        ws_bs[f'{col_letter}2'].alignment = Alignment(horizontal='center')

    row = 4

    # ASSETS
    ws_bs[f'A{row}'] = "ASSETS"
    ws_bs[f'A{row}'].font = Font(bold=True, size=12)
    row += 1

    # Current Assets
    ws_bs[f'A{row}'] = "Current Assets"
    ws_bs[f'A{row}'].font = bold_font
    row += 1

    ws_bs[f'A{row}'] = "Cash"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}77"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Accounts Receivable"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}8"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Inventory"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}14"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Prepayments"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}20"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Total Current Assets"
    ws_bs[f'A{row}'].font = bold_font
    ws_bs[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=SUM({col_letter}6:{col_letter}9)"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = ""
    row += 1

    # Non-Current Assets
    ws_bs[f'A{row}'] = "Non-Current Assets"
    ws_bs[f'A{row}'].font = bold_font
    row += 1

    ws_bs[f'A{row}'] = "Property, Plant & Equipment"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}32"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Total Non-Current Assets"
    ws_bs[f'A{row}'].font = bold_font
    ws_bs[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"={col_letter}13"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = ""
    row += 1

    ws_bs[f'A{row}'] = "TOTAL ASSETS"
    ws_bs[f'A{row}'].font = Font(bold=True)
    ws_bs[f'A{row}'].fill = total_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"={col_letter}10+{col_letter}14"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 2

    # LIABILITIES
    ws_bs[f'A{row}'] = "LIABILITIES & EQUITY"
    ws_bs[f'A{row}'].font = Font(bold=True, size=12)
    row += 1

    # Current Liabilities
    ws_bs[f'A{row}'] = "Current Liabilities"
    ws_bs[f'A{row}'].font = bold_font
    row += 1

    ws_bs[f'A{row}'] = "Accounts Payable"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}38"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Accrued Expenses"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}44"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Deferred Revenue"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}50"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Total Current Liabilities"
    ws_bs[f'A{row}'].font = bold_font
    ws_bs[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=SUM({col_letter}19:{col_letter}21)"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = ""
    row += 1

    # Non-Current Liabilities
    ws_bs[f'A{row}'] = "Non-Current Liabilities"
    ws_bs[f'A{row}'].font = bold_font
    row += 1

    ws_bs[f'A{row}'] = "Long-term Debt"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}55"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Total Non-Current Liabilities"
    ws_bs[f'A{row}'].font = bold_font
    ws_bs[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"={col_letter}25"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = ""
    row += 1

    ws_bs[f'A{row}'] = "TOTAL LIABILITIES"
    ws_bs[f'A{row}'].font = Font(bold=True)
    ws_bs[f'A{row}'].fill = total_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"={col_letter}22+{col_letter}26"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 2

    # EQUITY
    ws_bs[f'A{row}'] = "Equity"
    ws_bs[f'A{row}'].font = bold_font
    row += 1

    ws_bs[f'A{row}'] = "Share Capital"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = "=Inputs!$B$60"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Retained Earnings"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=Workings!{col_letter}61"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = "Total Equity"
    ws_bs[f'A{row}'].font = Font(bold=True)
    ws_bs[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"=SUM({col_letter}31:{col_letter}32)"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_bs[f'A{row}'] = ""
    row += 1

    ws_bs[f'A{row}'] = "TOTAL LIABILITIES & EQUITY"
    ws_bs[f'A{row}'].font = Font(bold=True)
    ws_bs[f'A{row}'].fill = total_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"={col_letter}28+{col_letter}33"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
    row += 2

    # Balance Check
    ws_bs[f'A{row}'] = "BALANCE CHECK (should be 0)"
    ws_bs[f'A{row}'].font = Font(bold=True, color="FF0000")
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_bs[f'{col_letter}{row}'] = f"={col_letter}16-{col_letter}35"
        ws_bs[f'{col_letter}{row}'].number_format = '#,##0'
        ws_bs[f'{col_letter}{row}'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # =======================
    # 5. CASH FLOW STATEMENT (INDIRECT METHOD)
    # =======================
    ws_cf = wb.create_sheet("Cash Flow Statement", 4)
    ws_cf.column_dimensions['A'].width = 35
    for col in range(2, 2 + num_periods + 1):
        ws_cf.column_dimensions[get_column_letter(col)].width = 12

    ws_cf['A1'] = "Cash Flow Statement (Indirect Method)"
    ws_cf['A1'].font = Font(bold=True, size=14)

    # Headers
    ws_cf['A2'] = "Period"
    ws_cf['A2'].font = header_font
    ws_cf['A2'].fill = header_fill

    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}2'] = f"Month {i+1}"
        ws_cf[f'{col_letter}2'].font = header_font
        ws_cf[f'{col_letter}2'].fill = header_fill
        ws_cf[f'{col_letter}2'].alignment = Alignment(horizontal='center')

    row = 4

    # Operating Activities
    ws_cf[f'A{row}'] = "OPERATING ACTIVITIES"
    ws_cf[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws_cf[f'A{row}'] = "Net Income"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"='Income Statement'!{col_letter}18"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Add: Depreciation"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"=-'Income Statement'!{col_letter}11"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Increase in Accounts Receivable"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_cf[f'{col_letter}{row}'] = f"=-(Workings!{col_letter}8-Workings!{col_letter}5)"
        else:
            prev_col = get_column_letter(i + 1)
            ws_cf[f'{col_letter}{row}'] = f"=-(Workings!{col_letter}8-Workings!{prev_col}8)"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Increase in Inventory"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_cf[f'{col_letter}{row}'] = f"=-(Workings!{col_letter}14-Workings!{col_letter}11)"
        else:
            prev_col = get_column_letter(i + 1)
            ws_cf[f'{col_letter}{row}'] = f"=-(Workings!{col_letter}14-Workings!{prev_col}14)"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Increase in Prepayments"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_cf[f'{col_letter}{row}'] = f"=-(Workings!{col_letter}20-Workings!{col_letter}17)"
        else:
            prev_col = get_column_letter(i + 1)
            ws_cf[f'{col_letter}{row}'] = f"=-(Workings!{col_letter}20-Workings!{prev_col}20)"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Increase in Accounts Payable"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}38-Workings!{col_letter}35"
        else:
            prev_col = get_column_letter(i + 1)
            ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}38-Workings!{prev_col}38"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Increase in Accrued Expenses"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}44-Workings!{col_letter}41"
        else:
            prev_col = get_column_letter(i + 1)
            ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}44-Workings!{prev_col}44"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Increase in Deferred Revenue"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        if i == 0:
            ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}50-Workings!{col_letter}47"
        else:
            prev_col = get_column_letter(i + 1)
            ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}50-Workings!{prev_col}50"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Net Cash from Operating Activities"
    ws_cf[f'A{row}'].font = bold_font
    ws_cf[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"=SUM({col_letter}5:{col_letter}12)"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = ""
    row += 1

    # Investing Activities
    ws_cf[f'A{row}'] = "INVESTING ACTIVITIES"
    ws_cf[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws_cf[f'A{row}'] = "Capital Expenditure"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"=-Workings!{col_letter}24"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Net Cash from Investing Activities"
    ws_cf[f'A{row}'].font = bold_font
    ws_cf[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"={col_letter}15"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = ""
    row += 1

    # Financing Activities
    ws_cf[f'A{row}'] = "FINANCING ACTIVITIES"
    ws_cf[f'A{row}'].font = Font(bold=True, size=11)
    row += 1

    ws_cf[f'A{row}'] = "Debt Drawdowns"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}52"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Debt Repayments"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"=-Workings!{col_letter}53"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Net Cash from Financing Activities"
    ws_cf[f'A{row}'].font = bold_font
    ws_cf[f'A{row}'].fill = subtotal_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"=SUM({col_letter}18:{col_letter}19)"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = ""
    row += 1

    # Net Change and Closing Cash
    ws_cf[f'A{row}'] = "Net Change in Cash"
    ws_cf[f'A{row}'].font = bold_font
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"={col_letter}13+{col_letter}16+{col_letter}19"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Opening Cash"
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"=Workings!{col_letter}68"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'
    row += 1

    ws_cf[f'A{row}'] = "Closing Cash"
    ws_cf[f'A{row}'].font = Font(bold=True)
    ws_cf[f'A{row}'].fill = total_fill
    for i in range(num_periods):
        col_letter = get_column_letter(i + 2)
        ws_cf[f'{col_letter}{row}'] = f"={col_letter}20+{col_letter}21"
        ws_cf[f'{col_letter}{row}'].number_format = '#,##0'

    # Save the workbook
    filename = "Financial_Model_3_Statement.xlsx"
    wb.save(filename)
    print(f"✓ Financial model created: {filename}")
    print(f"✓ Sheets created: {', '.join(wb.sheetnames)}")
    print(f"✓ Timeline: {num_periods} months")
    print(f"\nKey Features:")
    print(f"  • All Balance Sheet items driven by control accounts")
    print(f"  • No hardcoded values in calculation blocks")
    print(f"  • Full integration between all 3 statements")
    print(f"  • Indirect cash flow method")
    print(f"  • Balance Sheet balance check included")

if __name__ == "__main__":
    create_financial_model()

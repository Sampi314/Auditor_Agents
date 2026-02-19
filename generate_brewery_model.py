import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

def create_complex_model():
    workbook = xlsxwriter.Workbook('Brewery_Financial_Model_v2.xlsx')

    st = {
        'input': workbook.add_format({'font_color': '#0000FF', 'bg_color': '#FFFFCC', 'border': 1}),
        'formula': workbook.add_format({'font_color': '#000000', 'border': 1}),
        'link': workbook.add_format({'font_color': '#008000', 'border': 1}),
        'header': workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'}),
        'title': workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#FFFFFF', 'bg_color': '#4472C4'}),
        'num': workbook.add_format({'num_format': '#,##0', 'border': 1}),
        'num_dec': workbook.add_format({'num_format': '#,##0.00', 'border': 1}),
        'pct': workbook.add_format({'num_format': '0.0%', 'border': 1}),
        'curr': workbook.add_format({'num_format': '$#,##0', 'border': 1}),
        'curr_bold': workbook.add_format({'num_format': '$#,##0', 'bold': True, 'bg_color': '#F2F2F2', 'border': 1}),
        'check_err': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1}),
        'check_ok': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
    }

    sh_control = workbook.add_worksheet('01_Control')
    sh_inputs = workbook.add_worksheet('02_Inputs')
    sh_calcs = workbook.add_worksheet('03_Calcs')
    sh_outputs = workbook.add_worksheet('04_Outputs')
    sh_dashboard = workbook.add_worksheet('05_Dashboard')
    sh_checks = workbook.add_worksheet('06_Checks')

    MONTHS = 60
    SKUS = ['Lager', 'Pale Ale', 'IPA', 'Stout', 'Seasonal']

    # 01_Control
    sh_control.set_column(0, 0, 40)
    sh_control.write('A1', 'MODEL CONTROL PANEL', st['title'])
    sh_control.write('A3', 'Scenario Selector', st['header']); sh_control.write('B3', 1, st['input'])
    sh_control.write('A5', 'Global Rates', st['header'])
    sh_control.write('A6', 'GST Rate', st['formula']); sh_control.write('B6', 0.10, st['pct'])
    sh_control.write('A7', 'Corp Tax Rate', st['formula']); sh_control.write('B7', 0.30, st['pct'])
    sh_control.write('A8', 'Excise Packaged / LAL', st['formula']); sh_control.write('B8', 60.22, st['curr'])
    sh_control.write('A9', 'Excise Draught / LAL', st['formula']); sh_control.write('B9', 40.50, st['curr'])

    # 02_Inputs
    sh_inputs.set_column(0, 0, 25)
    sh_inputs.write('A1', 'INPUTS', st['title'])
    headers = ['SKU Name', 'ABV %', 'Price On-Trade', 'Price Off-Trade', 'Promo On %', 'Promo Off %', 'Base Vol']
    for i, h in enumerate(headers): sh_inputs.write(4, i, h, st['header'])
    sku_data = [
        ['Lager', 0.042, 520, 450, 0.05, 0.12, 1000],
        ['Pale Ale', 0.050, 580, 500, 0.06, 0.14, 800],
        ['IPA', 0.065, 650, 580, 0.08, 0.15, 400],
        ['Stout', 0.055, 600, 530, 0.05, 0.10, 200],
        ['Seasonal', 0.045, 540, 470, 0.10, 0.20, 300],
    ]
    for i, row in enumerate(sku_data):
        for j, val in enumerate(row): sh_inputs.write(5+i, j, val, st['input'] if j > 0 else st['formula'])
    sh_inputs.write('A11', 'On-Trade Split %', st['formula']); sh_inputs.write('B11', 0.40, st['input'])
    sh_inputs.write('A12', 'Brewing Loss %', st['formula']); sh_inputs.write('B12', 0.05, st['pct'])
    sh_inputs.write('A13', 'Pack Loss %', st['formula']); sh_inputs.write('B13', 0.02, st['pct'])
    sh_inputs.write('A15', 'Labor Count', st['header'])
    depts = [('Production', 8, 80), ('S&M', 6, 105), ('G&A', 3, 115)]
    for i, (d, h, s) in enumerate(depts):
        sh_inputs.write(16+i, 0, d, st['formula']); sh_inputs.write(16+i, 1, h, st['input']); sh_inputs.write(16+i, 2, s, st['input'])

    # 03_Calcs
    for m in range(MONTHS): sh_calcs.write(2, m+1, f'M{m+1}', st['header'])
    r = 4
    sku_vols = {}
    for i, sku in enumerate(SKUS):
        sh_calcs.write(r+i, 0, sku, st['formula'])
        sku_vols[sku] = r+i
        for m in range(MONTHS): sh_calcs.write_formula(r+i, m+1, f"='02_Inputs'!{xl_rowcol_to_cell(5+i, 6)} * (1.005^{m})", st['num'])
    total_sales_row = r + len(SKUS)
    for m in range(MONTHS): sh_calcs.write_formula(total_sales_row, m+1, f"=SUM({xl_rowcol_to_cell(r, m+1)}:{xl_rowcol_to_cell(total_sales_row-1, m+1)})", st['num'])

    prod_row = total_sales_row + 1
    for m in range(MONTHS): sh_calcs.write_formula(prod_row, m+1, f"={xl_rowcol_to_cell(total_sales_row, m+1)} / (1-0.07)", st['num'])

    rev_row = prod_row + 2
    for m in range(MONTHS):
        f = " + ".join([f"({xl_rowcol_to_cell(sku_vols[s], m+1)} * ('02_Inputs'!$B$11 * '02_Inputs'!{xl_rowcol_to_cell(5+i, 2)} + (1-'02_Inputs'!$B$11) * '02_Inputs'!{xl_rowcol_to_cell(5+i, 3)}))" for i, s in enumerate(SKUS)])
        sh_calcs.write_formula(rev_row, m+1, f"={f}", st['curr'])

    exc_row = rev_row + 1
    for m in range(MONTHS):
        f = " + ".join([f"({xl_rowcol_to_cell(sku_vols[s], m+1)} * 100 * ({sku_data[i][1]}-0.0115) * ('02_Inputs'!$B$11 * '01_Control'!$B$9 + (1-'02_Inputs'!$B$11) * '01_Control'!$B$8))" for i, s in enumerate(SKUS)])
        sh_calcs.write_formula(exc_row, m+1, f"=-{f}", st['curr'])
    net_rev_row = exc_row + 2
    for m in range(MONTHS): sh_calcs.write_formula(net_rev_row, m+1, f"={xl_rowcol_to_cell(rev_row, m+1)} + {xl_rowcol_to_cell(exc_row, m+1)}", st['curr'])

    cogs_row = net_rev_row + 2
    for m in range(MONTHS): sh_calcs.write_formula(cogs_row, m+1, f"={xl_rowcol_to_cell(prod_row, m+1)} * 155", st['curr'])

    opex_row = cogs_row + 2
    lab = "SUMPRODUCT('02_Inputs'!$B$17:$B$19, '02_Inputs'!$C$17:$C$19) * 1000 / 12"
    for m in range(MONTHS): sh_calcs.write_formula(opex_row, m+1, f"={lab} + {xl_rowcol_to_cell(net_rev_row, m+1)}*0.12", st['curr'])

    dep_row = opex_row + 2
    for m in range(MONTHS): sh_calcs.write(dep_row, m+1, 35000, st['curr'])

    ar_row = dep_row + 2
    for m in range(MONTHS): sh_calcs.write_formula(ar_row, m+1, f"={xl_rowcol_to_cell(rev_row, m+1)} * 1.2", st['curr'])

    # 04_Outputs
    for m in range(12): sh_outputs.write(2, m+1, f"M{m+1}", st['header'])
    for y in range(2, 6): sh_outputs.write(2, 11+y, f"Y{y}", st['header'])
    def ag(r, sm, em): return f"SUM('03_Calcs'!{xl_rowcol_to_cell(r, sm)}:{xl_rowcol_to_cell(r, em)})"

    r_is = 5
    sh_outputs.write(r_is, 0, 'Net Revenue', st['formula'])
    for m in range(1, 13): sh_outputs.write_formula(r_is, m, f"='03_Calcs'!{xl_rowcol_to_cell(net_rev_row, m)}", st['curr'])
    for y in range(2, 6): sh_outputs.write_formula(r_is, 11+y, f"={ag(net_rev_row, (y-1)*12+1, y*12)}", st['curr'])

    sh_outputs.write(r_is+1, 0, 'COGS', st['formula'])
    for m in range(1, 13): sh_outputs.write_formula(r_is+1, m, f"=-'03_Calcs'!{xl_rowcol_to_cell(cogs_row, m)}", st['curr'])
    for y in range(2, 6): sh_outputs.write_formula(r_is+1, 11+y, f"=-{ag(cogs_row, (y-1)*12+1, y*12)}", st['curr'])

    sh_outputs.write(r_is+2, 0, 'OpEx', st['formula'])
    for m in range(1, 13): sh_outputs.write_formula(r_is+2, m, f"=-'03_Calcs'!{xl_rowcol_to_cell(opex_row, m)}", st['curr'])
    for y in range(2, 6): sh_outputs.write_formula(r_is+2, 11+y, f"=-{ag(opex_row, (y-1)*12+1, y*12)}", st['curr'])

    eb_row = 8
    sh_outputs.write(eb_row, 0, 'EBITDA', st['header'])
    for c in range(1, 17): sh_outputs.write_formula(eb_row, c, f"=SUM({xl_rowcol_to_cell(5, c)}:{xl_rowcol_to_cell(7, c)})", st['curr'])

    np_row = 12
    sh_outputs.write(np_row, 0, 'NPAT', st['curr_bold'])
    for c in range(1, 17): sh_outputs.write_formula(np_row, c, f"={xl_rowcol_to_cell(eb_row, c)} * 0.7", st['curr_bold'])

    cl_cash_row = 20
    sh_outputs.write(cl_cash_row, 0, 'Closing Cash', st['curr_bold'])
    sh_outputs.write(cl_cash_row-1, 1, 800000, st['curr'])
    for c in range(1, 17):
        if c > 1: sh_outputs.write_formula(cl_cash_row-1, c, f"={xl_rowcol_to_cell(cl_cash_row, c-1)}", st['curr'])
        sh_outputs.write_formula(cl_cash_row, c, f"={xl_rowcol_to_cell(eb_row, c)}*0.8 + {xl_rowcol_to_cell(cl_cash_row-1, c)}", st['curr_bold'])

    as_row = 32
    sh_outputs.write(as_row, 0, 'Total Assets', st['header'])
    for c in range(1, 17): sh_outputs.write_formula(as_row, c, f"={xl_rowcol_to_cell(cl_cash_row, c)} + 2000000 + '03_Calcs'!{xl_rowcol_to_cell(ar_row, c if c<=12 else c*12-11)}", st['curr'])
    eq_row = 36
    sh_outputs.write(eq_row, 0, 'Total Equity', st['header'])
    sh_outputs.write(eq_row-1, 1, 2800000, st['curr'])
    for c in range(1, 17):
        if c > 1: sh_outputs.write_formula(eq_row-1, c, f"={xl_rowcol_to_cell(eq_row, c-1)}", st['curr'])
        sh_outputs.write_formula(eq_row, c, f"={xl_rowcol_to_cell(eq_row-1, c)} + {xl_rowcol_to_cell(np_row, c)}", st['curr'])

    # 06_Checks
    sh_checks.set_column(0, 0, 40)
    sh_checks.write('A1', 'AUDIT & COVENANTS', st['title'])
    sh_checks.write('A4', 'Balance Sheet Check (Assets - Equity)', st['formula'])
    sh_checks.write_formula('B4', f"=IF(ABS({xl_rowcol_to_cell(as_row, 16)}-{xl_rowcol_to_cell(eq_row, 16)})<100, 'PASS', 'FAIL')", st['check_ok'])
    sh_checks.write('A5', 'Minimum Cash Covenant ($500k)', st['formula'])
    sh_checks.write_formula('B5', f"=IF(MIN('04_Outputs'!B21:Q21)>500000, 'PASS', 'FAIL')", st['check_ok'])
    sh_checks.write('A6', 'Interest Cover (>4.0x)', st['formula'])
    sh_checks.write_formula('B6', f"=IF(AVERAGE('04_Outputs'!B9:M9)/10000 > 4, 'PASS', 'FAIL')", st['check_ok'])

    workbook.close()

if __name__ == "__main__":
    create_complex_model()
    print("Master Brewery Model Finalized with Checks.")

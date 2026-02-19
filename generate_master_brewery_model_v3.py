import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

def create_master_model():
    # Model configuration
    FILENAME = 'Master_Brewery_Financial_Model_v3.xlsx'
    workbook = xlsxwriter.Workbook(FILENAME)

    # Styles
    st = {
        'in': workbook.add_format({'font_color': '#0000FF', 'bg_color': '#FFFFCC', 'border': 1}),
        'fm': workbook.add_format({'font_color': '#000000', 'border': 1}),
        'lk': workbook.add_format({'font_color': '#008000', 'border': 1}),
        'hd': workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'}),
        'title': workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#FFFFFF', 'bg_color': '#4472C4'}),
        'num': workbook.add_format({'num_format': '#,##0', 'border': 1}),
        'num_dec': workbook.add_format({'num_format': '#,##0.00', 'border': 1}),
        'pct': workbook.add_format({'num_format': '0.0%', 'border': 1}),
        'curr': workbook.add_format({'num_format': '$#,##0', 'border': 1}),
        'curr_bold': workbook.add_format({'num_format': '$#,##0', 'bold': True, 'bg_color': '#F2F2F2', 'border': 1}),
        'ok': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1}),
        'err': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
    }

    # Sheets
    sheets = {
        'control': workbook.add_worksheet('01_Control'),
        'inputs_sales': workbook.add_worksheet('02_Inputs_Sales'),
        'inputs_ops': workbook.add_worksheet('03_Inputs_Ops'),
        'calcs_vol': workbook.add_worksheet('04_Calcs_Volume'),
        'calcs_rev': workbook.add_worksheet('05_Calcs_Revenue'),
        'calcs_cogs': workbook.add_worksheet('06_Calcs_COGS'),
        'calcs_opex': workbook.add_worksheet('07_Calcs_OpEx'),
        'calcs_wc': workbook.add_worksheet('08_Calcs_WC_Inv'),
        'calcs_capex': workbook.add_worksheet('09_Calcs_Capex_Tax'),
        'outputs': workbook.add_worksheet('10_Outputs_3Way'),
        'dash': workbook.add_worksheet('11_Dashboard'),
        'audit': workbook.add_worksheet('12_Audit_Checks')
    }

    MONTHS = 60
    SKUS = [
        {'id': 0, 'name': 'Lager - Premium', 'abv': 0.042, 'cat': 'Core'},
        {'id': 1, 'name': 'Lager - Mid Strength', 'abv': 0.035, 'cat': 'Core'},
        {'id': 2, 'name': 'Pale Ale - Pacific', 'abv': 0.048, 'cat': 'Core'},
        {'id': 3, 'name': 'Pale Ale - West Coast', 'abv': 0.056, 'cat': 'Core'},
        {'id': 4, 'name': 'IPA - Hazy', 'abv': 0.062, 'cat': 'Limited'},
        {'id': 5, 'name': 'IPA - Double', 'abv': 0.075, 'cat': 'Limited'},
        {'id': 6, 'name': 'Stout - Oatmeal', 'abv': 0.052, 'cat': 'Core'},
        {'id': 7, 'name': 'Summer Ale', 'abv': 0.040, 'cat': 'Seasonal'},
        {'id': 8, 'name': 'Winter Porter', 'abv': 0.060, 'cat': 'Seasonal'}
    ]

    # --- 01_Control ---
    sh = sheets['control']
    sh.set_column(0, 0, 40)
    sh.write('A1', 'MODEL CONTROL PANEL', st['title'])
    sh.write('A3', 'Model Scenario Selector (1=Base, 2=High Cost, 3=Aggressive)', st['hd'])
    sh.write('B3', 1, st['in'])
    sh.write('A5', 'Global Rates & Statutory Assumptions', st['hd'])
    sh.write('A6', 'GST Rate', st['fm']); sh.write('B6', 0.10, st['pct'])
    sh.write('A7', 'Corporate Tax Rate', st['fm']); sh.write('B7', 0.30, st['pct'])
    sh.write('A8', 'Payroll Tax Rate', st['fm']); sh.write('B8', 0.0485, st['pct'])
    sh.write('A9', 'Superannuation Rate', st['fm']); sh.write('B9', 0.115, st['pct'])
    sh.write('A11', 'Australian Excise Rates (ATO)', st['hd'])
    sh.write('A12', 'Packaged Beer (>1.15% ABV) / LAL', st['fm']); sh.write('B12', 60.22, st['curr'])
    sh.write('A13', 'Draught Beer (>1.15% ABV) / LAL', st['fm']); sh.write('B13', 40.50, st['curr'])

    # --- 02_Inputs_Sales ---
    sh = sheets['inputs_sales']
    sh.set_column(0, 0, 25)
    sh.write('A1', 'SALES & PRICING INPUTS', st['title'])
    sh.write('A3', 'Monthly Seasonality Indices', st['hd'])
    seasonality = [1.2, 1.15, 0.9, 0.8, 0.75, 0.85, 0.95, 1.05, 1.1, 1.25, 1.35, 1.45]
    for m in range(12):
        sh.write(3, m+1, f"Month {m+1}", st['hd'])
        sh.write(4, m+1, seasonality[m], st['in'])
    sh.write('A6', 'SKU Performance Assumptions', st['hd'])
    headers = ['SKU Name', 'ABV', 'On-Trade Price/HL', 'Off-Trade Price/HL', 'On-Trade %', 'Base Vol HL/m']
    for i, h in enumerate(headers): sh.write(7, i, h, st['hd'])
    sku_inputs = [
        ['Lager - Premium', 0.042, 520, 440, 0.35, 1200],
        ['Lager - Mid Strength', 0.035, 480, 410, 0.25, 1000],
        ['Pale Ale - Pacific', 0.048, 560, 480, 0.40, 800],
        ['Pale Ale - West Coast', 0.056, 590, 510, 0.45, 600],
        ['IPA - Hazy', 0.062, 650, 570, 0.55, 400],
        ['IPA - Double', 0.075, 850, 750, 0.60, 150],
        ['Stout - Oatmeal', 0.052, 610, 530, 0.30, 250],
        ['Summer Ale', 0.040, 530, 460, 0.50, 500],
        ['Winter Porter', 0.060, 620, 540, 0.40, 300]
    ]
    for i, row in enumerate(sku_inputs):
        for j, val in enumerate(row):
            fmt = st['in'] if j > 0 else st['fm']
            sh.write(8+i, j, val, fmt)

    # --- 03_Inputs_Ops ---
    sh = sheets['inputs_ops']
    sh.set_column(0, 0, 25)
    sh.write('A1', 'OPERATIONAL & COGS INPUTS', st['title'])
    sh.write('A3', 'Bill of Materials (per HL)', st['hd'])
    sh.write_row(3, 1, ['Malt (kg)', 'Hops (g)', 'Packaging Units', 'Water/Utils ($)'], st['hd'])
    bom_data = [
        [16, 150, 303, 12], [14, 120, 303, 11], [19, 450, 303, 14], [21, 600, 303, 15],
        [24, 900, 303, 18], [28, 1200, 303, 20], [25, 300, 303, 16], [17, 400, 303, 13], [26, 350, 303, 17]
    ]
    for i, row in enumerate(bom_data):
        sh.write(4+i, 0, SKUS[i]['name'], st['fm'])
        for j, val in enumerate(row): sh.write(4+i, j+1, val, st['in'])
    sh.write('A15', 'Unit Cost Inputs', st['hd'])
    sh.write('A16', 'Malt ($/kg)', st['fm']); sh.write('B16', 1.35, st['in'])
    sh.write('A17', 'Hops ($/kg)', st['fm']); sh.write('B17', 52.00, st['in'])
    sh.write('A18', 'Packaging ($/unit)', st['fm']); sh.write('B18', 0.58, st['in'])
    sh.write('A21', 'Labor & Salaries', st['hd'])
    headers = ['Department', 'Headcount', 'Avg Salary ($k)']
    for i, h in enumerate(headers): sh.write(22, i, h, st['hd'])
    depts = [('Brewing', 5, 95), ('Packaging', 6, 70), ('Quality/Lab', 2, 105), ('Sales/Marketing', 8, 110), ('Admin/Finance', 4, 125)]
    for i, (d, h, s) in enumerate(depts):
        sh.write(23+i, 0, d, st['fm']); sh.write(23+i, 1, h, st['in']); sh.write(23+i, 2, s, st['in'])

    # --- 04_Calcs_Volume ---
    sh = sheets['calcs_vol']
    sh.set_column(0, 0, 30)
    for m in range(MONTHS): sh.write(2, m+1, f"M{m+1}", st['hd'])
    r = 4
    sku_v_rows = {}
    for i, sku in enumerate(SKUS):
        sh.write(r+1+i, 0, sku['name'], st['fm'])
        sku_v_rows[sku['name']] = r+1+i
        for m in range(MONTHS):
            sh.write_formula(r+1+i, m+1, f"='02_Inputs_Sales'!{xl_rowcol_to_cell(8+i, 5)} * '02_Inputs_Sales'!{xl_rowcol_to_cell(4, (m%12)+1)} * (1.005^{m})", st['num'])
    tot_vol_row = r+1+len(SKUS)
    for m in range(MONTHS): sh.write_formula(tot_vol_row, m+1, f"=SUM({xl_rowcol_to_cell(r+1, m+1)}:{xl_rowcol_to_cell(r+len(SKUS), m+1)})", st['num'])

    # --- 05_Calcs_Revenue ---
    sh = sheets['calcs_rev']
    sh.set_column(0, 0, 30)
    for m in range(MONTHS): sh.write(2, m+1, f"M{m+1}", st['hd'])
    r = 4
    gross_rev_row = r+1
    for m in range(MONTHS):
        parts = []
        for i, sku in enumerate(SKUS):
            v_c = f"'04_Calcs_Volume'!{xl_rowcol_to_cell(sku_v_rows[sku['name']], m+1)}"
            p_on = f"'02_Inputs_Sales'!{xl_rowcol_to_cell(8+i, 2)}"
            p_off = f"'02_Inputs_Sales'!{xl_rowcol_to_cell(8+i, 3)}"
            on_pct = f"'02_Inputs_Sales'!{xl_rowcol_to_cell(8+i, 4)}"
            parts.append(f"{v_c} * ({on_pct}*{p_on} + (1-{on_pct})*{p_off})")
        sh.write_formula(gross_rev_row, m+1, f"={' + '.join(parts)}", st['curr'])
    excise_row = r+4
    for m in range(MONTHS):
        parts = []
        for i, sku in enumerate(SKUS):
            v_c = f"'04_Calcs_Volume'!{xl_rowcol_to_cell(sku_v_rows[sku['name']], m+1)}"
            on_pct = f"'02_Inputs_Sales'!{xl_rowcol_to_cell(8+i, 4)}"
            lal = f"({v_c}*100*({sku['abv']}-0.0115))"
            rate = f"({on_pct}*'01_Control'!$B$13 + (1-{on_pct})*'01_Control'!$B$12)"
            parts.append(f"MAX(0, {lal}*{rate})")
        sh.write_formula(excise_row, m+1, f"=-({' + '.join(parts)})", st['curr'])
    net_rev_row = r+7
    for m in range(MONTHS): sh.write_formula(net_rev_row, m+1, f"={xl_rowcol_to_cell(gross_rev_row, m+1)} + {xl_rowcol_to_cell(excise_row, m+1)}", st['curr_bold'])

    # --- 06_Calcs_COGS ---
    sh = sheets['calcs_cogs']
    sh.set_column(0, 0, 30)
    for m in range(MONTHS): sh.write(2, m+1, f"M{m+1}", st['hd'])
    r = 4
    sh.write(r+1, 0, 'Total Malt Cost', st['fm'])
    for m in range(MONTHS):
        parts = [f"'04_Calcs_Volume'!{xl_rowcol_to_cell(sku_v_rows[s['name']], m+1)}*'03_Inputs_Ops'!{xl_rowcol_to_cell(4+i, 1)}" for i, s in enumerate(SKUS)]
        sh.write_formula(r+1, m+1, f"=({' + '.join(parts)}) * '03_Inputs_Ops'!$B$16", st['curr'])
    sh.write(r+2, 0, 'Total Hops Cost', st['fm'])
    for m in range(MONTHS):
        parts = [f"'04_Calcs_Volume'!{xl_rowcol_to_cell(sku_v_rows[s['name']], m+1)}*'03_Inputs_Ops'!{xl_rowcol_to_cell(4+i, 2)}/1000" for i, s in enumerate(SKUS)]
        sh.write_formula(r+2, m+1, f"=({' + '.join(parts)}) * '03_Inputs_Ops'!$B$17", st['curr'])
    tot_cogs_row = r+5
    for m in range(MONTHS): sh_calcs_vol_row = 14 # Total Vol row in calcs_vol sheet
    for m in range(MONTHS): sh.write_formula(tot_cogs_row, m+1, f"=SUM({xl_rowcol_to_cell(r+1, m+1)}:{xl_rowcol_to_cell(r+3, m+1)}) + '04_Calcs_Volume'!{xl_rowcol_to_cell(tot_vol_row, m+1)}*85", st['curr'])

    # --- 07_Calcs_OpEx ---
    sh = sheets['calcs_opex']
    sh.set_column(0, 0, 30)
    for m in range(MONTHS): sh.write(2, m+1, f"M{m+1}", st['hd'])
    lab_base = "SUMPRODUCT('03_Inputs_Ops'!$B$23:$B$27, '03_Inputs_Ops'!$C$23:$C$27) * 1000 / 12 * (1 + '01_Control'!$B$9 + '01_Control'!$B$8)"
    for m in range(MONTHS): sh.write_formula(4, m+1, f"={lab_base} * (1.002^{m})", st['curr'])
    sh.write(6, 0, 'Variable OpEx', st['fm'])
    for m in range(MONTHS): sh.write_formula(6, m+1, f"='05_Calcs_Revenue'!{xl_rowcol_to_cell(net_rev_row, m+1)} * 0.12 + 25000", st['curr'])
    tot_opex_row = 8
    for m in range(MONTHS): sh.write_formula(tot_opex_row, m+1, f"=SUM({xl_rowcol_to_cell(4, m+1)}:{xl_rowcol_to_cell(7, m+1)})", st['curr'])

    # --- 10_Outputs_3Way ---
    sh = sheets['outputs']
    sh.set_column(0, 0, 30)
    for m in range(12): sh.write(2, m+1, f"M{m+1}", st['hd'])
    for y in range(2, 6): sh.write(2, 11+y, f"Y{y}", st['hd'])
    def ag(sheet, r, sm, em): return f"SUM('{sheet}'!{xl_rowcol_to_cell(r, sm)}:{xl_rowcol_to_cell(r, em)})"
    sh.write('A4', 'INCOME STATEMENT', st['hd'])
    rows = [('Net Revenue', '05_Calcs_Revenue', net_rev_row, 1), ('COGS', '06_Calcs_COGS', tot_cogs_row, -1), ('OpEx', '07_Calcs_OpEx', tot_opex_row, -1)]
    for i, (l, s, sr, sn) in enumerate(rows):
        sh.write(5+i, 0, l, st['fm'])
        for m in range(1, 13): sh.write_formula(5+i, m, f"={sn}*'{s}'!{xl_rowcol_to_cell(sr, m)}", st['curr'])
        for y in range(2, 6): sh.write_formula(5+i, 11+y, f"={sn}*{ag(s, sr, (y-1)*12+1, y*12)}", st['curr'])
    eb_row = 9
    for c in range(1, 17): sh.write_formula(eb_row, c, f"=SUM({xl_rowcol_to_cell(5, c)}:{xl_rowcol_to_cell(8, c)})", st['curr'])
    sh.write('A20', 'CASH FLOW SUMMARY', st['hd'])
    for c in range(1, 17): sh.write_formula(20, c, f"={xl_rowcol_to_cell(eb_row, c)} * 0.7", st['curr'])
    sh.write(21, 1, 1000000, st['curr'])
    for c in range(1, 17):
        if c > 1: sh.write_formula(21, c, f"={xl_rowcol_to_cell(21, c-1)} + {xl_rowcol_to_cell(20, c)}", st['curr_bold'])
        else: sh.write_formula(21, c, f"1000000 + {xl_rowcol_to_cell(20, c)}", st['curr_bold'])

    # --- 12_Audit_Checks ---
    sh = sheets['audit']
    sh.set_column(0, 0, 45)
    sh.write('A1', 'MODEL INTEGRITY AUDIT', st['title'])
    sh.write_formula('B4', f"=IF(MIN('10_Outputs_3Way'!B22:Q22)>0, 'PASS', 'WARN')", st['ok'])
    sh.write_formula('B5', f"=IF('05_Calcs_Revenue'!{xl_rowcol_to_cell(net_rev_row, 1)} > 0, 'PASS', 'FAIL')", st['ok'])

    workbook.close()
    return FILENAME

if __name__ == "__main__":
    fname = create_master_model()
    print(f"Master Model '{fname}' generated.")

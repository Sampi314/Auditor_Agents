import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

def create_model():
    workbook = xlsxwriter.Workbook('Brewery_Financial_Model.xlsx')

    # --- Styles ---
    def get_styles(wb):
        s = {
            'input': wb.add_format({'font_color': '#0000FF', 'bg_color': '#FFFFCC', 'border': 1}),
            'formula': wb.add_format({'font_color': '#000000', 'border': 1}),
            'link': wb.add_format({'font_color': '#008000', 'border': 1}),
            'header': wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'}),
            'title': wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#FFFFFF', 'bg_color': '#4472C4'}),
            'num': wb.add_format({'num_format': '#,##0', 'border': 1}),
            'pct': wb.add_format({'num_format': '0.0%', 'border': 1}),
            'curr': wb.add_format({'num_format': '$#,##0', 'border': 1}),
            'curr_bold': wb.add_format({'num_format': '$#,##0', 'bold': True, 'bg_color': '#F2F2F2', 'border': 1}),
            'check': wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
        }
        return s

    st = get_styles(workbook)
    sh_control = workbook.add_worksheet('01_Control')
    sh_inputs = workbook.add_worksheet('02_Inputs')
    sh_calcs = workbook.add_worksheet('03_Calcs')
    sh_outputs = workbook.add_worksheet('04_Outputs')
    sh_dashboard = workbook.add_worksheet('05_Dashboard')

    MONTHS = 60
    SKUS = [
        {'name': 'Crisp Lager', 'abv': 0.045, 'p_on': 420, 'p_off': 380, 'malt': 17, 'hops': 0.15, 'vol': 800},
        {'name': 'Pacific Pale', 'abv': 0.052, 'p_on': 550, 'p_off': 480, 'malt': 20, 'hops': 0.50, 'vol': 600},
        {'name': 'Hops Smash IPA', 'abv': 0.065, 'p_on': 680, 'p_off': 600, 'malt': 23, 'hops': 1.10, 'vol': 300},
        {'name': 'Night Owl Stout', 'abv': 0.058, 'p_on': 620, 'p_off': 540, 'malt': 25, 'hops': 0.40, 'vol': 150},
        {'name': 'Summer Ale', 'abv': 0.042, 'p_on': 480, 'p_off': 420, 'malt': 16, 'hops': 0.35, 'vol': 400},
        {'name': 'Barrel Aged Ltd', 'abv': 0.085, 'p_on': 980, 'p_off': 880, 'malt': 30, 'hops': 0.70, 'vol': 80},
    ]

    # --- 01_Control ---
    sh_control.set_column(0, 0, 40)
    sh_control.write('A1', 'MODEL CONTROL PANEL', st['title'])
    sh_control.write('A3', 'Scenario Selector (1=Base, 2=Spike)', st['header']); sh_control.write('B3', 1, st['input'])
    sh_control.write('A5', 'Scenario Assumptions', st['header'])
    sh_control.write('A6', 'Malt Price Spike (%)', st['formula']); sh_control.write('B6', 0.25, st['pct'])
    sh_control.write('A7', 'Hops Price Spike (%)', st['formula']); sh_control.write('B7', 0.50, st['pct'])

    # --- 02_Inputs ---
    sh_inputs.set_column(0, 0, 25)
    sh_inputs.write('A1', 'OPERATIONAL INPUTS', st['title'])
    sh_inputs.write('A3', 'Base Malt Price ($/kg)', st['header']); sh_inputs.write('B3', 1.25, st['input'])
    sh_inputs.write('A4', 'Base Hops Price ($/kg)', st['header']); sh_inputs.write('B4', 48.00, st['input'])

    sh_inputs.write('A6', 'Live Malt Price', st['formula'])
    sh_inputs.write_formula('B6', "=IF('01_Control'!$B$3=2, $B$3*(1+'01_Control'!$B$6), $B$3)", st['curr'])
    sh_inputs.write('A7', 'Live Hops Price', st['formula'])
    sh_inputs.write_formula('B7', "=IF('01_Control'!$B$3=2, $B$4*(1+'01_Control'!$B$7), $B$4)", st['curr'])

    sh_inputs.write('A9', 'Seasonality', st['header'])
    seasonality = [1.2, 1.1, 0.9, 0.8, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3, 1.4]
    for i, s in enumerate(seasonality): sh_inputs.write(10, i, s, st['input'])

    sh_inputs.write('A12', 'Loss Rate', st['formula']); sh_inputs.write('B12', 0.07, st['pct'])
    sh_inputs.write('A13', 'Channel Split On-Trade %', st['formula']); sh_inputs.write('B13', 0.35, st['pct'])

    sh_inputs.write('A15', 'Labor', st['header'])
    depts = [('Brewing', 4, 85000), ('Packaging', 3, 65000), ('Sales', 3, 95000), ('Admin', 2, 80000)]
    for i, (d, h, s) in enumerate(depts):
        sh_inputs.write(16+i, 0, d, st['formula'])
        sh_inputs.write(16+i, 1, h, st['input'])
        sh_inputs.write(16+i, 2, s, st['input'])

    # --- 03_Calcs ---
    sh_calcs.set_column(0, 0, 35)
    for m in range(MONTHS): sh_calcs.write(2, m+1, f'M{m+1}', st['header'])

    # 1. Volume
    r = 4
    sku_vol_rows = {}
    for i, sku in enumerate(SKUS):
        sh_calcs.write(r+i, 0, sku['name'], st['formula'])
        sku_vol_rows[sku['name']] = r+i
        for m in range(MONTHS):
            sh_calcs.write_formula(r+i, m+1, f"={sku['vol']} * '02_Inputs'!{xl_rowcol_to_cell(10, m%12)} * (1.005^{m})", st['num'])
    total_vol_row = r + len(SKUS)
    sh_calcs.write(total_vol_row, 0, 'Total Volume HL', st['header'])
    for m in range(MONTHS): sh_calcs.write_formula(total_vol_row, m+1, f"=SUM({xl_rowcol_to_cell(r, m+1)}:{xl_rowcol_to_cell(total_vol_row-1, m+1)})", st['num'])

    # 2. Revenue & Excise
    r = total_vol_row + 2
    sh_calcs.write(r, 0, 'Net Revenue', st['header'])
    for m in range(MONTHS):
        parts = []
        for i, sku in enumerate(SKUS):
            v_c = xl_rowcol_to_cell(sku_vol_rows[sku['name']], m+1)
            # Simplified Net Revenue formula for brevity in logic but SKU granular
            parts.append(f"{v_c}*('02_Inputs'!$B$13*{sku['p_on']}+(1-'02_Inputs'!$B$13)*{sku['p_off']})*0.85") # 0.85 approx after excise
        sh_calcs.write_formula(r, m+1, f"={' + '.join(parts)}", st['curr'])
    net_rev_row = r

    # 3. COGS
    r = net_rev_row + 2
    sh_calcs.write(r, 0, 'Total COGS', st['header'])
    for m in range(MONTHS):
        pm = []
        lf = "(1/(1-'02_Inputs'!$B$12))"
        for i, sku in enumerate(SKUS):
            v_c = xl_rowcol_to_cell(sku_vol_rows[sku['name']], m+1)
            pm.append(f"({v_c}*{sku['malt']}*{lf}*'02_Inputs'!$B$6 + {v_c}*{sku['hops']}*{lf}*'02_Inputs'!$B$7)")
        sh_calcs.write_formula(r, m+1, f"={' + '.join(pm)} + {xl_rowcol_to_cell(total_vol_row, m+1)}*75", st['curr'])
    total_cogs_row = r

    # 4. OpEx
    r = total_cogs_row + 2
    sh_calcs.write(r, 0, 'Total OpEx', st['header'])
    for m in range(MONTHS):
        sh_calcs.write_formula(r, m+1, f"=(SUMPRODUCT('02_Inputs'!$B$16:$B$19, '02_Inputs'!$C$16:$C$19)/12) + {xl_rowcol_to_cell(net_rev_row, m+1)}*0.1", st['curr'])
    total_opex_row = r

    # --- 04_Outputs (3-Way) ---
    sh_outputs.set_column(0, 0, 30)
    for m in range(12): sh_outputs.write(2, m+1, f"M{m+1}", st['header'])
    for y in range(2, 6): sh_outputs.write(2, 11+y, f"Y{y}", st['header'])

    # IS
    r_is = 4
    sh_outputs.write(r_is, 0, 'INCOME STATEMENT', st['header'])
    def map_row(label, src_row, sign):
        sh_outputs.write(r_is+1, 0, label, st['formula'])
        for m in range(1, 13): sh_outputs.write_formula(r_is+1, m, f"={sign}*'03_Calcs'!{xl_rowcol_to_cell(src_row, m)}", st['curr'])
        for y in range(2, 6): sh_outputs.write_formula(r_is+1, 11+y, f"={sign}*SUM('03_Calcs'!{xl_rowcol_to_cell(src_row, (y-1)*12+1)}:{xl_rowcol_to_cell(src_row, y*12)})", st['curr'])

    map_row('Net Revenue', net_rev_row, 1); r_is += 1
    map_row('COGS', total_cogs_row, -1); r_is += 1
    map_row('OpEx', total_opex_row, -1); r_is += 1
    ebitda_row_out = r_is + 1
    sh_outputs.write(ebitda_row_out, 0, 'EBITDA', st['header'])
    for c in range(1, 17): sh_outputs.write_formula(ebitda_row_out, c, f"=SUM({xl_rowcol_to_cell(ebitda_row_out-3, c)}:{xl_rowcol_to_cell(ebitda_row_out-1, c)})", st['curr'])

    # BS
    r_bs = ebitda_row_out + 3
    sh_outputs.write(r_bs, 0, 'BALANCE SHEET', st['header'])
    sh_outputs.write(r_bs+1, 0, 'Cash Asset', st['formula'])
    # CF
    r_cf = r_bs + 5
    sh_outputs.write(r_cf, 0, 'CASH FLOW', st['header'])
    sh_outputs.write(r_cf+1, 0, 'Net Cash Flow', st['formula'])
    for c in range(1, 17): sh_outputs.write_formula(r_cf+1, c, f"={xl_rowcol_to_cell(ebitda_row_out, c)}*0.7", st['curr'])
    sh_outputs.write(r_cf+2, 0, 'Opening Cash', st['formula']); sh_outputs.write(r_cf+2, 1, 500000, st['curr'])
    for c in range(2, 17): sh_outputs.write_formula(r_cf+2, c, f"={xl_rowcol_to_cell(r_cf+3, c-1)}", st['curr'])
    sh_outputs.write(r_cf+3, 0, 'Closing Cash', st['curr_bold'])
    for c in range(1, 17): sh_outputs.write_formula(r_cf+3, c, f"=SUM({xl_rowcol_to_cell(r_cf+1, c)}:{xl_rowcol_to_cell(r_cf+2, c)})", st['curr'])

    # Link Cash back to BS
    for c in range(1, 17): sh_outputs.write_formula(r_bs+1, c, f"={xl_rowcol_to_cell(r_cf+3, c)}", st['curr'])

    # --- 05_Dashboard ---
    sh_dashboard.set_column(0, 0, 40)
    sh_dashboard.write('A1', 'BREWERY KPI DASHBOARD', st['title'])
    sh_dashboard.write('A4', 'Year 1 EBITDA', st['header'])
    sh_dashboard.write_formula('B4', f"=SUM('04_Outputs'!B9:M9)", st['curr_bold'])

    workbook.close()

if __name__ == "__main__":
    create_model()
    print("Master Brewery Model Finalized.")

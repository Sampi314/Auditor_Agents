import openpyxl
import re
import json
from collections import Counter

def get_r1c1_formula(formula, cell):
    """
    Very basic R1C1 conversion for simple formulas to improve consistency check.
    Handles relative references like A1, B2.
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula

    # Simple regex for A1 style references
    # This won't handle complex sheet names with exclamation marks or absolute refs perfectly,
    # but it's a step up from direct string comparison.
    def replace_ref(match):
        col_str = match.group(1)
        row_str = match.group(2)

        # Convert column to index
        col_idx = 0
        for char in col_str:
            col_idx = col_idx * 26 + (ord(char.upper()) - ord('A') + 1)

        row_idx = int(row_str)

        row_off = row_idx - cell.row
        col_off = col_idx - cell.column

        return f"R[{row_off}]C[{col_off}]"

    # Regex for A1 style (excluding absolute $ refs for now as they are already consistent)
    r1c1 = re.sub(r'(?<![A-Z$])([A-Z]+)([0-9]+)(?![0-9])', replace_ref, formula)
    return r1c1

def run_audit(file_path):
    wb_val = openpyxl.load_workbook(file_path, data_only=True)
    wb_form = openpyxl.load_workbook(file_path, data_only=False)

    findings = []

    # --- SENTRY AUDIT ---
    error_codes = ["#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!"]
    for sheet_name in wb_val.sheetnames:
        ws_v = wb_val[sheet_name]
        ws_f = wb_form[sheet_name]
        for row_v, row_f in zip(ws_v.iter_rows(), ws_f.iter_rows()):
            for cell_v, cell_f in zip(row_v, row_f):
                if cell_v.value in error_codes:
                    findings.append({
                        "Sheet Name": sheet_name, "Cell Reference": cell_v.coordinate,
                        "Description of the Location": f"Cell value in {sheet_name}",
                        "Short Error Categories": "Calculation Error",
                        "Long Description of error": f"Cell evaluates to {cell_v.value}."
                    })
                if isinstance(cell_f.value, str) and "#REF!" in cell_f.value:
                    findings.append({
                        "Sheet Name": sheet_name, "Cell Reference": cell_f.coordinate,
                        "Description of the Location": f"Formula in {sheet_name}",
                        "Short Error Categories": "Broken Reference",
                        "Long Description of error": f"Formula contains #REF!: {cell_f.value}"
                    })

    for name, defn in wb_form.defined_names.items():
        val = str(defn.value) if hasattr(defn, 'value') else str(defn)
        if "#REF!" in val:
            findings.append({
                "Sheet Name": "Name Manager", "Cell Reference": "N/A",
                "Description of the Location": f"Named Range: {name}",
                "Short Error Categories": "Dead Name",
                "Long Description of error": f"Named range {name} points to #REF!: {val}"
            })

    # --- SYNC AUDIT ---
    calc_sheets = [s for s in wb_form.sheetnames if s.startswith('Calc_') or s in ['IS', 'CF', 'BS', 'KPIs']]
    for sheet_name in calc_sheets:
        ws = wb_form[sheet_name]
        # Scan all 120 columns (B to DQ approx)
        for row in ws.iter_rows(min_row=4, max_col=121):
            row_cells = row[1:] # Skip column A
            r1c1_formulas = []
            valid_cells = []

            for cell in row_cells:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    r1c1_formulas.append(get_r1c1_formula(cell.value, cell))
                    valid_cells.append(cell)
                elif isinstance(cell.value, (int, float)) and cell.value != 0:
                    # Check if this cell SHOULD be a formula based on neighbors
                    pass

            # Check for formula inconsistencies (R1C1 based)
            if len(r1c1_formulas) > 10:
                counts = Counter(r1c1_formulas)
                if len(counts) > 1:
                    most_common = counts.most_common(1)[0][0]
                    for cell, f_r1c1 in zip(valid_cells, r1c1_formulas):
                        if f_r1c1 != most_common:
                            # Only flag if it's not the first period
                            if cell.column > 2:
                                findings.append({
                                    "Sheet Name": sheet_name, "Cell Reference": cell.coordinate,
                                    "Description of the Location": f"Calculation row in {sheet_name}",
                                    "Short Error Categories": "Formula Inconsistency",
                                    "Long Description of error": f"Formula pattern break. R1C1 variant: {f_r1c1}"
                                })

            # Hard-coded check
            formulas_in_row = [c for c in row_cells if isinstance(c.value, str) and c.value.startswith('=')]
            nums_in_row = [c for c in row_cells if isinstance(c.value, (int, float)) and c.value != 0]
            if len(formulas_in_row) > 50 and len(nums_in_row) > 0:
                for cell in row_cells:
                    if isinstance(cell.value, (int, float)) and cell.value != 0:
                        findings.append({
                            "Sheet Name": sheet_name, "Cell Reference": cell.coordinate,
                            "Description of the Location": f"Formula row in {sheet_name}",
                            "Short Error Categories": "Hard-coded Value",
                            "Long Description of error": f"Hard-coded number {cell.value} found in a row that appears to be mostly formulas."
                        })

    # --- LINGO AUDIT ---
    for sheet_name in ['IS', 'CF', 'BS', 'KPIs', 'Annual_Summary']:
        ws = wb_val[sheet_name]
        for row in ws.iter_rows(max_col=1):
            cell = row[0]
            if isinstance(cell.value, str) and cell.value.strip():
                if not any(u in cell.value for u in ['$', 'HL', '%', 'AUD', 'Date', 'Period']):
                    has_data = False
                    for c in range(2, 5):
                        if isinstance(ws.cell(row=cell.row, column=c).value, (int, float)):
                            has_data = True; break
                    if has_data:
                        findings.append({
                            "Sheet Name": sheet_name, "Cell Reference": cell.coordinate,
                            "Description of the Location": f"Line item in {sheet_name}",
                            "Short Error Categories": "Missing Label",
                            "Long Description of error": f"Line item '{cell.value}' missing unit indicator."
                        })

    # --- LOGIC AUDIT ---
    bs_ws = wb_val['BS']
    asset_row = 0; le_row = 0
    for r in range(1, bs_ws.max_row + 1):
        v = bs_ws.cell(row=r, column=1).value
        if v == 'TOTAL ASSETS': asset_row = r
        if v == 'TOTAL LIABILITIES & EQUITY': le_row = r
    if asset_row and le_row:
        for col in range(2, 122):
            a = bs_ws.cell(row=asset_row, column=col).value or 0
            le = bs_ws.cell(row=le_row, column=col).value or 0
            if abs(a - le) > 0.1:
                findings.append({
                    "Sheet Name": "BS", "Cell Reference": bs_ws.cell(row=asset_row, column=col).coordinate,
                    "Description of the Location": "Balance Sheet Balance",
                    "Short Error Categories": "Logical Flaw",
                    "Long Description of error": f"BS out of balance by {a-le} in period {col-1}."
                })

    return findings

if __name__ == "__main__":
    results = run_audit("Brewery_Financial_Model_10Y 1.xlsx")
    print("| Sheet Name | Cell Reference | Description of the Location | Short Error Categories | Long Description of error |")
    print("|------------|----------------|-----------------------------|------------------------|---------------------------|")
    seen = set()
    for r in results:
        key = (r["Sheet Name"], r["Cell Reference"], r["Short Error Categories"])
        if key not in seen:
            print(f"| {r['Sheet Name']} | {r['Cell Reference']} | {r['Description of the Location']} | {r['Short Error Categories']} | {r['Long Description of error']} |")
            seen.add(key)

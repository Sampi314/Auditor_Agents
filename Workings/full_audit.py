import openpyxl
import re
import json
from collections import Counter, defaultdict
from spellchecker import SpellChecker
from openpyxl.styles import PatternFill
from openpyxl.worksheet.formula import ArrayFormula

def get_color_str(color):
    if color is None: return "None"
    if color.type == 'rgb': return f"#{color.rgb[2:]}"
    elif color.type == 'theme': return f"Theme({color.theme})"
    elif color.type == 'indexed': return f"Indexed({color.indexed})"
    return "None"

def a1_to_r1c1(formula, cell_row, cell_col):
    if isinstance(formula, ArrayFormula): formula = formula.t
    formula = str(formula)
    def replace_ref(match):
        col_str, row_str = match.groups()
        col_abs = col_str.startswith('$')
        row_abs = row_str.startswith('$')
        col_name = col_str.replace('$', '')
        row_num = int(row_str.replace('$', ''))
        col_index = 0
        for char in col_name: col_index = col_index * 26 + (ord(char.upper()) - ord('A') + 1)
        if row_abs: r_part = f"R{row_num}"
        else:
            r_part = f"R[{row_num - cell_row}]"
            if r_part == "R[0]": r_part = "R"
        if col_abs: c_part = f"C{col_index}"
        else:
            c_part = f"C[{col_index - cell_col}]"
            if c_part == "C[0]": c_part = "C"
        return r_part + c_part
    pattern = r'(\$?[A-Z]+)(\$?[0-9]+)'
    return re.sub(pattern, replace_ref, formula)

def add_finding(findings, sheetname, ref, desc, cat, long_desc, priority, agent):
    findings.append({
        "Sheet Name": sheetname,
        "Cell Reference": ref,
        "Description": desc,
        "Category": cat,
        "Long Description": long_desc,
        "Priority": priority,
        "Agent": agent
    })

def full_audit(filename):
    wb = openpyxl.load_workbook(filename, data_only=False)
    wb_data = openpyxl.load_workbook(filename, data_only=True)
    findings = []

    # 1. Sentry & Efficiency & Architect
    error_values = ["#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#SPILL!", "#CALC!"]
    volatile_functions = ["OFFSET(", "INDIRECT(", "TODAY()", "NOW(", "RAND(", "RANDBETWEEN("]

    # Identify Timeline Sheet for Architect
    timeline_sheet = next((s for s in wb.sheetnames if 'timeline' in s.lower() or 'timing' in s.lower()), None)

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        ws_data = wb_data[sheetname]
        for row in ws.iter_rows():
            for cell in row:
                val = ws_data[cell.coordinate].value
                if str(val) in error_values:
                    add_finding(findings, sheetname, cell.coordinate, "Calculation Error", "Calculation Error",
                                f"Cell contains error {val}. Check source data and formula logic to resolve the calculation error.", "High", "Sentry")

                if cell.data_type == 'f':
                    formula = str(cell.value).upper()

                    # Efficiency: Volatile functions
                    for vf in volatile_functions:
                        if vf in formula:
                            add_finding(findings, sheetname, cell.coordinate, f"Volatile Function: {vf[:-1]}", "Volatile Complexity",
                                        f"Use of volatile function {vf[:-1]} detected. Volatile functions recalculate every time any cell in the workbook changes, which can slow down large models. Consider replacing with INDEX/MATCH if possible.", "Medium", "Efficiency")

                    # Efficiency: Mega-Formula (>4000)
                    if len(formula) > 4000:
                        add_finding(findings, sheetname, cell.coordinate, "Mega-Formula (Bad Practice)", "Mega-Formula",
                                    "Extremely long formula detected (>4000 chars). This is a critical auditability risk. Formulas this complex are impossible to verify or maintain. Break the calculation into multiple modular steps using helper rows.", "High", "Efficiency")
                    # Efficiency: Long Formula (>500)
                    elif len(formula) > 500:
                        add_finding(findings, sheetname, cell.coordinate, "Long Formula", "Auditability Risk",
                                    "Formula exceeds 500 characters. While not a 'Mega-Formula', it still poses an auditability risk. Consider breaking it down into smaller, more manageable components.", "Medium", "Efficiency")

    # 2. Lingo
    spell = SpellChecker()
    spell.word_frequency.load_words(['EBITDA', 'COGS', 'CFADS', 'IRR', 'CAPEX', 'OPEX', 'GST', 'Corality', 'Excel', 'tutorial', 'modelling', 'Offsheet', 'OneLine', 'DSCR', 'Heng'])
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        for row in ws.iter_rows(max_row=30):
            for cell in row:
                if isinstance(cell.value, str) and not cell.value.startswith('=') and cell.column <= 5:
                    words = re.findall(r'\b[A-Za-z]{4,}\b', cell.value)
                    misspelled = spell.unknown(words)
                    for word in misspelled:
                        if word.lower() not in ['dscr', 'heng']:
                            add_finding(findings, sheetname, cell.coordinate, f"Label: {cell.value}", "Typo",
                                        f"Potential typo '{word}' detected in label. Review and correct the spelling for professional presentation.", "Low", "Lingo")

    # 3. Stylist
    # Key: (sheetname, category, long_desc, priority, agent)
    stylist_findings = defaultdict(list)
    for sheetname in wb.sheetnames:
        if sheetname in ['L', '  ', 'Corality']: continue
        ws = wb[sheetname]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None: continue
                f_color = get_color_str(cell.font.color)
                if not cell.data_type == 'f' and isinstance(cell.value, (int, float)):
                    if cell.row > 5 and cell.column > 3:
                        if f_color != "#800000" and f_color != "None" and f_color != "Theme(0)":
                            key = (sheetname, "Colour Coding Error", f"Hard-coded input is formatted with font color {f_color} instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency.", "Low", "Stylist")
                            stylist_findings[key].append(cell.coordinate)
                if cell.data_type == 'f' and '!' in str(cell.value):
                    if f_color != "Indexed(16)":
                        key = (sheetname, "Colour Coding Error", f"Off-sheet link is formatted with font color {f_color} instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies.", "Low", "Stylist")
                        stylist_findings[key].append(cell.coordinate)
    for key, cells in stylist_findings.items():
        sn, cat, desc, priority, agent = key
        add_finding(findings, sn, ", ".join(cells), "Multiple locations", cat, desc, priority, agent)

    # 4. Logic & Numerical Sense Check
    for sheetname in wb.sheetnames:
        if sheetname in ['L', '  ', 'Corality']: continue
        ws = wb[sheetname]
        ws_data = wb_data[sheetname]

        row_labels = {}
        for row in ws.iter_rows(min_col=1, max_col=5):
            label = ""
            for cell in row:
                if isinstance(cell.value, str): label = cell.value; break
            row_labels[row[0].row] = label

        for row_idx, label in row_labels.items():
            label_lower = label.lower()
            for cell in ws_data[row_idx]:
                val = cell.value
                if isinstance(val, (int, float)):
                    if "dscr" in label_lower and val < 0:
                        add_finding(findings, sheetname, cell.coordinate, f"Label: {label}", "Sanity Check Failure",
                                    "Negative DSCR detected. DSCR (Debt Service Coverage Ratio) should typically be positive. Review the calculation for potential logic errors.", "High", "Logic")
                    if "revenue" in label_lower and val < 0:
                        add_finding(findings, sheetname, cell.coordinate, f"Label: {label}", "Sanity Check Failure",
                                    "Negative Revenue detected. Revenue values should typically be positive. Verify the source data and revenue logic.", "High", "Logic")

                orig_cell = ws[cell.coordinate]
                if orig_cell.data_type == 'f':
                    try:
                        r1c1 = a1_to_r1c1(orig_cell.value, orig_cell.row, orig_cell.column)
                        if len(str(orig_cell.value)) > 4000:
                            if val not in error_values:
                                add_finding(findings, sheetname, cell.coordinate, "Deep Logic Check", "Auditability Risk",
                                            "Mega-formula is technically functioning (no calculation error detected), but its complexity makes it an extreme auditability risk. It should be broken out for transparency.", "Medium", "Efficiency")
                    except: pass

        for row in ws.iter_rows():
            row_formulas = {}
            for cell in row:
                if cell.data_type == 'f':
                    try:
                        r1c1 = a1_to_r1c1(cell.value, cell.row, cell.column)
                        row_formulas[cell.coordinate] = r1c1
                    except: pass
            if len(row_formulas) > 5:
                counts = Counter(row_formulas.values())
                if len(counts) > 1:
                    dominant_r1c1, _ = counts.most_common(1)[0]
                    for coord, r1c1 in row_formulas.items():
                        if r1c1 != dominant_r1c1 and ws[coord].column >= 9:
                            add_finding(findings, sheetname, coord, "Formula row", "Formula Pattern Break",
                                        f"Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row.", "Medium", "Logic")

            for cell in row:
                if cell.data_type == 'f':
                    formula = str(cell.value)
                    literals = re.findall(r'(?<![A-Z$])\b([0-9]+\.[0-9]+|[2-9][0-9]+|[0-9]{2,})\b', formula)
                    literals = [l for l in literals if l not in ['100', '12', '365', '52', '1', '0', '2', '4', '8000', '4000']]
                    if literals:
                        # Avoid mentioning specific numbers to allow grouping
                        add_finding(findings, sheetname, cell.coordinate, "Hard-coded literal", "Hard-Code in Formula",
                                    "Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency.", "Medium", "Logic")

    return findings

if __name__ == "__main__":
    import sys
    import os
    if len(sys.argv) > 1:
        target_file = sys.argv[1]
    else:
        # Try to find any xlsx file in the parent directory if no argument is provided
        xlsx_files = [f for f in os.listdir('..') if f.endswith('.xlsx')]
        if xlsx_files:
            target_file = os.path.join('..', xlsx_files[0])
        else:
            print("No .xlsx file found to audit.")
            sys.exit(1)

    print(f"Auditing: {target_file}")
    results = full_audit(target_file)
    with open('full_audit_results.json', 'w') as f: json.dump(results, f)

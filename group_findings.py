import json
from collections import defaultdict
import re

def col_to_int(col):
    num = 0
    for char in col: num = num * 26 + (ord(char) - ord('A') + 1)
    return num

def int_to_col(n):
    res = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(65 + rem) + res
    return res

def split_coord(coord):
    match = re.match(r'([A-Z]+)([0-9]+)', coord)
    if not match: return None, None
    return match.group(1), int(match.group(2))

def group_cells_into_ranges_2d(cell_list):
    if not cell_list: return ""
    cells = set()
    for c in cell_list:
        col_str, row = split_coord(c)
        if col_str: cells.add((row, col_to_int(col_str)))
    ranges = []
    while cells:
        r, c = min(cells)
        w = 1
        while (r, c + w) in cells: w += 1
        h = 1
        while True:
            can_expand_h = True
            for next_c in range(c, c + w):
                if (r + h, next_c) not in cells:
                    can_expand_h = False
                    break
            if not can_expand_h: break
            h += 1
        top_left = f"{int_to_col(c)}{r}"
        if w == 1 and h == 1: ranges.append(top_left)
        else:
            bottom_right = f"{int_to_col(c + w - 1)}{r + h - 1}"
            ranges.append(f"{top_left}:{bottom_right}")
        for i in range(r, r + h):
            for j in range(c, c + w): cells.discard((i, j))
    return ", ".join(ranges)

def main():
    try:
        with open('full_audit_results.json') as f: results = json.load(f)
    except FileNotFoundError: return

    final_grouped = defaultdict(list)
    descriptions = {}
    for r in results:
        key = (r["Sheet Name"], r["Category"], r["Long Description"])
        final_grouped[key].append(r["Cell Reference"])
        if key not in descriptions: descriptions[key] = r["Description"]

    final_report_data = []
    for key, cells in final_grouped.items():
        sn, cat, desc = key
        flat_cells = []
        for c in cells: flat_cells.extend([x.strip() for x in c.split(",") if x.strip()])
        range_str = group_cells_into_ranges_2d(flat_cells)
        final_report_data.append({
            "Sheet Name": sn,
            "Cell Reference": range_str,
            "Description": descriptions[key],
            "Category": cat,
            "Long Description": desc
        })

    def severity_rank(desc):
        if "🔴 HIGH" in desc: return 0
        if "⚠️ MEDIUM" in desc: return 1
        return 2

    final_report_data.sort(key=lambda x: (x["Sheet Name"], severity_rank(x["Long Description"]), x["Category"]))
    with open('final_audit_data.json', 'w') as f: json.dump(final_report_data, f)

if __name__ == "__main__":
    main()

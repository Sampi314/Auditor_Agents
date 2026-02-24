import json

def generate_markdown_report(json_filepath, output_filepath):
    with open(json_filepath, 'r') as f:
        data = json.load(f)

    headers = ["Sheet Name", "Cell Reference", "Description of the Location", "Short Error Category", "Long Description of Error"]

    lines = []
    lines.append("# Audit Report 📊")
    lines.append("")
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("|" + "|".join(["---"] * len(headers)) + "|")

    for row in data:
        line = "| " + " | ".join([
            str(row.get("Sheet Name", "")),
            str(row.get("Cell Reference", "")),
            str(row.get("Description", "")),
            str(row.get("Category", "")),
            str(row.get("Long Description", ""))
        ]) + " |"
        lines.append(line)

    with open(output_filepath, 'w') as f:
        f.write("\n".join(lines))

if __name__ == "__main__":
    generate_markdown_report('final_audit_data.json', 'Audit Report.md')

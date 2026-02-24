import json
import os
import argparse
from collections import defaultdict

def generate_markdown_report(output_dir=None):
    if output_dir:
        input_path = os.path.join(output_dir, 'final_audit_data.json')
        output_path = os.path.join(output_dir, 'Audit Report.md')
    else:
        input_path = 'final_audit_data.json'
        output_path = 'Audit Report.md'

    try:
        with open(input_path, 'r') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"File not found: {input_path}")
        return

    # Group data by Agent
    agent_reports = defaultdict(list)
    for row in data:
        agent = row.get("Agent", "Unknown")
        agent_reports[agent].append(row)

    headers = ["Sheet Name", "Cell Reference", "Description of the Location", "Priority", "Short Error Category", "Long Description of Error"]

    lines = []
    lines.append("# Audit Report 📊")
    lines.append("")
    lines.append("This report is divided into sections based on the specialized agent persona that identified the findings.")
    lines.append("")

    # Sort agents for consistent report order
    sorted_agents = sorted(agent_reports.keys())

    for agent in sorted_agents:
        lines.append(f"## {agent} Report")
        lines.append("")
        lines.append("| " + " | ".join(headers) + " |")
        lines.append("|" + "|".join(["---"] * len(headers)) + "|")

        for row in agent_reports[agent]:
            # Priority icons
            priority = row.get("Priority", "Medium")
            prio_display = priority
            if priority.upper() == "HIGH": prio_display = "🔴 HIGH"
            elif priority.upper() == "MEDIUM": prio_display = "⚠️ MEDIUM"
            elif priority.upper() == "LOW": prio_display = "🟡 LOW"

            line = "| " + " | ".join([
                str(row.get("Sheet Name", "")),
                str(row.get("Cell Reference", "")),
                str(row.get("Description", "")),
                prio_display,
                str(row.get("Category", "")),
                str(row.get("Long Description", ""))
            ]) + " |"
            lines.append(line)
        lines.append("")

    with open(output_path, 'w') as f:
        f.write("\n".join(lines))
    print(f"Report written to {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Report Script")
    parser.add_argument("--output-dir", required=True, help="Path to output directory")
    args = parser.parse_args()

    generate_markdown_report(args.output_dir)

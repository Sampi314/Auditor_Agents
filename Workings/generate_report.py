import json
from collections import defaultdict

def generate_markdown_report(json_filepath, output_filepath):
    with open(json_filepath, 'r') as f:
        data = json.load(f)

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

    with open(output_filepath, 'w') as f:
        f.write("\n".join(lines))

if __name__ == "__main__":
    generate_markdown_report('final_audit_data.json', 'Audit Report.md')

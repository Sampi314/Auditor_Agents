import argparse
import os
import sys

# Import functions from sibling scripts
from full_audit import full_audit
from hyperlinks_audit import audit_hyperlinks
from group_findings import group_findings
from generate_report import generate_markdown_report
from cartographer import generate_cartographer_reports
from navigator import generate_navigator_reports

def main():
    parser = argparse.ArgumentParser(description="Run Full Financial Model Analysis")
    parser.add_argument("--input", required=True, help="Path to input Excel file")
    parser.add_argument("--output-dir", required=True, help="Path to output directory for results")
    args = parser.parse_args()

    input_file = args.input
    output_dir = args.output_dir

    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' does not exist.")
        sys.exit(1)

    if not os.path.exists(output_dir):
        print(f"Creating output directory: {output_dir}")
        os.makedirs(output_dir)

    print("="*50)
    print(f"Starting Analysis for: {input_file}")
    print(f"Output Directory: {output_dir}")
    print("="*50)

    # 1. Full Audit
    print("\n--- Step 1: Full Audit ---")
    full_audit(input_file, output_dir)

    # 2. Hyperlinks Audit
    print("\n--- Step 2: Hyperlinks Audit ---")
    audit_hyperlinks(input_file, output_dir)

    # 3. Group Findings
    print("\n--- Step 3: Group Findings ---")
    group_findings(output_dir)

    # 4. Generate Audit Report
    print("\n--- Step 4: Generate Audit Report ---")
    generate_markdown_report(output_dir)

    # 5. Cartographer (Maps)
    print("\n--- Step 5: Cartographer Maps ---")
    generate_cartographer_reports(input_file, output_dir)

    # 6. Navigator (Summary)
    print("\n--- Step 6: Navigator Reports ---")
    generate_navigator_reports(input_file, output_dir)

    print("\n" + "="*50)
    print("Analysis Complete!")
    print(f"Results are available in: {output_dir}")
    print("="*50)

if __name__ == "__main__":
    main()

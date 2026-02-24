# Generalized Python Scripts for Financial Model Analysis

This folder contains a suite of Python scripts designed to audit, map, and summarize financial models (Excel files). These scripts have been generalized to work with any `.xlsx` file and output results to a specified directory.

## Prerequisites

- Python 3.x
- `openpyxl`
- `pyspellchecker`
- `argparse`

You can install the dependencies using pip:
```bash
pip install openpyxl pyspellchecker
```

## Usage

The main entry point is `run_analysis.py`. It orchestrates the entire analysis pipeline.

### Command

```bash
python run_analysis.py --input <path_to_excel_file> --output-dir <path_to_output_directory>
```

### Example

```bash
python run_analysis.py --input "../Brewery_Financial_Model_10Y 1.xlsx" --output-dir "Analysis_Results"
```

## Pipeline Steps

1.  **Full Audit (`full_audit.py`)**: Checks for calculation errors, volatile functions, mega-formulas, typos, and styling inconsistencies.
2.  **Hyperlinks Audit (`hyperlinks_audit.py`)**: Verifies internal and external hyperlinks.
3.  **Group Findings (`group_findings.py`)**: Groups individual cell findings into ranges for better readability.
4.  **Generate Report (`generate_report.py`)**: Produces a consolidated Markdown report (`Audit Report.md`).
5.  **Cartographer (`cartographer.py`)**: Generates dependency maps (Mermaid diagrams) and a dependency register.
    -   Outputs to `Maps/` subdirectory.
    -   `Flow_Dependency_Register.md`
    -   `Maps/Flow_L1_Workbook.mermaid`
    -   `Maps/Flow_L2_{SheetName}.mermaid`
    -   `Maps/Flow_L3_Critical_Path.mermaid`
6.  **Navigator (`navigator.py`)**: Generates model summary reports.
    -   `01_Model_Summary.md`
    -   `02_Model_Flowchart.md`
    -   `03_Calculation_Reference.md`

## Output Structure

The output directory will contain:

-   `Audit Report.md`: The main findings report.
-   `Flow_Dependency_Register.md`: A detailed list of all inter-sheet dependencies.
-   `01_Model_Summary.md`: High-level summary of the model structure.
-   `02_Model_Flowchart.md`: High-level flowchart of the model.
-   `03_Calculation_Reference.md`: Key calculations reference.
-   `Maps/`: A folder containing Mermaid diagrams for visualization.
-   `full_audit_results.json`: Intermediate raw finding data.
-   `final_audit_data.json`: Intermediate grouped finding data.

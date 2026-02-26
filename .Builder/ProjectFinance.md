# Builder Blueprint: Project Finance (SPV) 🏗️

> *"Precision over the long haul. Timing is everything."*

## Overview

This blueprint is designed for Special Purpose Vehicle (SPV) infrastructure and energy models. It focuses on the distinction between Construction (Capex) and Operations (Revenue) phases, precise debt sculpting, and complex cash flow waterfalls.

## Structure & Architecture

Based on `SP Case Study Interim Model 14 Indirect Cash Flow.xlsm`, the model structure emphasizes a robust timeline and clear separation of inputs.

### 1. Inputs (Assumptions) 📥

- **`Model Parameters`**: High-level toggles (Inflation, Currency, Project Name).
- **`General Assumptions`**:
    - **Timing**: Financial Close, Construction Start, COD (Commercial Operations Date), End Date.
    - **Macro**: Indexation (CPI/RPI), FX rates.
    - **Fiscal**: Tax rates, VAT.
- **`Timing`**:
    - **Master Timeline**: Monthly during Construction, Semi-annual/Annual during Operations.
    - **Flags**: `Construction_Flag`, `Operations_Flag`, `Debt_Repayment_Flag`.
- **`Lookup`**: Dropdown lists for scenarios, currencies, etc.

### 2. Calculations (Processing) ⚙️

- **`Calculations`**: Often a single, large sheet or a series of linked sheets (`Calc_...`).
    - **Construction Costs (Uses)**: EPC contract, development fees, contingency.
    - **Funding (Sources)**: Equity, Shareholder Loans, Senior Debt.
    - **Revenue**: Availability payments or volume-based (PPA) revenue.
    - **Opex**: O&M contracts, insurance, SPV management.
    - **Tax**: Depreciation (Capital Allowances), Tax losses carried forward.
    - **Working Capital**: Debtors, Creditors, VAT account.
- **Debt Sizing**:
    - **Construction**: Interest During Construction (IDC), Commitment Fees.
    - **Operations**: Sculpting principal repayments based on CFADS (Cash Flow Available for Debt Service).
    - **Ratios**: DSCR (Debt Service Coverage Ratio), LLCR (Loan Life Coverage Ratio).
- **Reserve Accounts**: DSRA (Debt Service Reserve Account), MMRA (Major Maintenance Reserve Account).

### 3. Outputs (Reporting) 📊

- **`Income Statement` (IS)**:
    - Revenue - Opex - Depr - Interest - Tax = Profit.
- **`Balance Sheet` (BS)**:
    - **Assets**: PPE (Construction in Progress -> Asset Base), Cash (DSRA, Operating), Receivables.
    - **Liabilities**: Senior Debt Balance, Shareholder Loans.
    - **Equity**: Share Capital + Retained Earnings.
- **`Cash Flow Statement` (CF)**:
    - **CFADS Calculation**: EBITDA - Tax +/- Working Capital - Capex.
    - **Waterfall**: Senior Debt Service -> DSRA Top-up -> Sub-debt -> Equity Distributions.
- **`Error Checks`**:
    - **Binary**: `Asset = Liability + Equity` (must be TRUE).
    - **Cash**: `Min Cash Balance >= 0`.
    - **Debt**: `Debt Balance at End = 0`.

## Key Logic & Formulas

- **Timeline Driven**: All calculations MUST reference the `Timing` sheet flags (e.g., `=IF(Construction_Flag=1, Capex, 0)`).
- **Circularities**: Common in IDC and Debt Sculpting. Use a "copy-paste" macro or iterative calculation switch (with caution) to resolve.
- **Waterfall**: Strict priority of payments in the Cash Flow.

## Implementation Steps

1.  **Timing Sheet**: Establish the master timeline first. This is the spine of the model.
2.  **Capex & Funding**: Build the Construction phase logic (Uses & Sources).
3.  **Operations**: Build Revenue and Opex logic.
4.  **Tax & WC**: Layer in tax and working capital adjustments.
5.  **CFADS**: Calculate Cash Flow Available for Debt Service.
6.  **Debt**: Size the debt based on CFADS and DSCR constraints.
7.  **Financial Statements**: Link the flows to IS/BS/CF.
8.  **Checks & Sensitivities**: Add the `Error Checks` sheet and Data Tables for sensitivity analysis.

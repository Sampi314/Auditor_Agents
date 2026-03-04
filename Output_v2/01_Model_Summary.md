# Model Summary

## 1. Purpose
This is a project finance model for a renewable energy asset (likely wind or solar). It calculates equity IRR, DSCR, and CFADS over a 10-year operations period following a 2-year (8-quarter) construction phase. The model supports debt sizing and funding structure analysis for a total CapEx of AUD 60M.

## 2. Model Structure
| Sheet Name | Role | Description |
|---|---|---|
| ` ` (Cover) | Cover / TOC | Title page with model name, course name, author details (Kenny Chew, Corality Financial Group), and contact hyperlinks |
| QCF | Financial Statement | Quarterly Cash Flow Statement — consolidates revenue, OpEx, CapEx, funding, debt service, dividends, and net cash balance |
| Inputs | Inputs / Assumptions | Centralised assumptions: timing, CapEx, volume & price, OpEx, funding terms, macroeconomics (base interest rates) |
| Cons | Calculations | Construction phase: spend profile, funding sources & uses, construction debt drawdowns, interest during construction (IDC) |
| Ops | Calculations | Operations phase: revenue (volume x price), fixed & variable OpEx, FCFE, equity IRR (XIRR) |
| Debt | Calculations | Operations debt: refinance from construction, linear repayment, interest, DSCR |
| Timing | Timing | Master date spine: quarterly period start/end dates, construction & operations flags, counters (quarter, year, days) |
| L&N | Data / Lookup | Formatting Legend, Named Ranges & Constants: style guide, entity names, structural constants (Days_Yr=365, Qtrs_Yr=4, Million=1,000,000) |

## 3. Timeline
- **Model start**: 31-Dec-2011
- **Model end**: 31-Dec-2024
- **Frequency**: Quarterly
- **Total periods**: 52 quarters (13 years)
- **Construction phase**: 1-Jan-2012 to 31-Dec-2013 (8 quarters)
- **Operations phase**: 1-Jan-2014 to 31-Dec-2023 (40 quarters / 10 years)

## 4. Key Inputs
| Input | Location | Value | Unit |
|---|---|---|---|
| Model Start Date | Inputs!E6 | 31-Dec-2011 | Date |
| Construction Duration | Inputs!E10 | 8 | Qtrs |
| Operations Duration | Inputs!E15 | 10 | Yrs |
| Total CapEx | Inputs!E20 | 60.00 | AUD M |
| Construction Spend Profile | Inputs!F23:M23 | 12.5% per quarter | % |
| Net Generation | Inputs!E28 | 130,000 | MWh p.a. |
| Market Price (Yr 1-4) | Inputs!F32:I32 | 110, 112, 114.5, 115 | AUD/MWh |
| Labour (Fixed OpEx) | Inputs!E37 | 0.515 (escalating ~3% p.a.) | AUD M |
| Land & Leasing (Fixed OpEx) | Inputs!E38 | 1.03 (escalating ~3% p.a.) | AUD M |
| Variable OpEx | Inputs!E43 | 1.545 (escalating ~3% p.a.) | AUD/MWh |
| Senior Debt Limit | Inputs!E49 | 50.00 | AUD M |
| Construction Margin | Inputs!E52 | 3.50% | % p.a. |
| Operations Margin | Inputs!F52 | 2.50% | % p.a. |
| Debt Tenor | Inputs!E55 | 6 | Yrs |
| Repayment Rate | Inputs!E58 | 16.67% | % p.a. (linear) |
| Base Interest Rate (2012-2015) | Inputs!F62:I62 | 3.6%, 3.0%, 2.9%, 3.0% | % p.a. |

## 5. Key Outputs
| Output | Location | Value | Unit |
|---|---|---|---|
| Total Revenue | Ops!I19 (=QCF!I14) | ~155.3 | AUD M |
| Total Fixed OpEx | Ops!I26 | ~70.8 | AUD M |
| Total CapEx (incl. IDC) | Cons!I25 | ~63.4 | AUD M |
| Total Debt Drawn | Cons!I36 | 50.0 | AUD M |
| Total Equity | Cons!I31 | ~13.4 | AUD M |
| Equity IRR | Ops!I59 | XIRR-based | % p.a. |
| Min DSCR | Debt!E45 | Array formula | x |
| Avg DSCR | Debt!F45 | Weighted average | x |

## 6. Scenarios / Switches
No scenario switches or toggles are present in this model. It runs a single base case.

## 7. Key Assumptions
- **Volume**: 130,000 MWh p.a. for the first ~6 years (from Inputs), then extrapolated at 0.5% quarterly growth (hard-coded in Ops)
- **Price**: Escalating from AUD 110/MWh, specified year-by-year in Inputs for 4 years
- **Costs**: Fixed costs (Labour + Land & Leasing) and variable costs (per MWh) escalate at ~3% p.a.
- **Debt**: AUD 50M senior facility, drawn during construction, refinanced at COD, repaid linearly at 16.67% p.a. over 6-year tenor
- **Interest**: Floating rate = Base Rate + Margin, with base rates specified for 2012-2015

## 8. Circular References
No circular references detected.

## 9. Known Limitations / Notes
- The model is a training/self-auditing exercise from Corality Financial Group (June 2014 vintage)
- Volume and price assumptions in Inputs only cover partial periods; the Ops sheet hard-codes 0.5% growth for later periods
- No tax calculations are present in the model
- No working capital or depreciation calculations
- The "Spare" OpEx line items (Inputs rows 39, 44, 45) are placeholders with no values
- Cover sheet is named ` ` (single space character) — poor practice

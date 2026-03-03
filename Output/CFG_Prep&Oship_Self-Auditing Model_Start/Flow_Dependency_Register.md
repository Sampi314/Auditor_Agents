# Flow Dependency Register

| # | Source Sheet | Source Cell | Source Row Label | Target Sheet | Target Cell | Target Row Label | Reference Type |
|---|---|---|---|---|---|---|---|
| 1 | Timing | A4 | Period Start | Cons | A4 |  | INPUT |
| 2 | Timing | A5 | Period End | Cons | A5 |  | INPUT |
| 3 | Timing | A6 | Construction | Cons | A6 |  | INPUT |
| 4 | Timing | A7 | Operations | Cons | A7 |  | INPUT |
| 5 | Timing | C10 | Construction Quarters | Cons | C10 | Num# | INPUT |
| 6 | Timing | C13 | Calendar Year | Cons | C11 | Year | INPUT |
| 7 | Timing | C14 | Days in Period | Cons | C12 | Days | INPUT |
| 8 | Inputs | C20 | All construction costs | Cons | C16 | % | INPUT |
| 9 | Timing | C9 | Counters | Cons | C9 |  | INPUT |
| 10 | Inputs | E20 | All construction costs | Cons | E19 | AUD M | INPUT |
| 11 | Inputs | E11 | End | Cons | E37 | Refinance | LINK |
| 12 | Inputs | E49 | Debt Limit | Cons | E40 | Undrawn Amount | INPUT |
| 13 | Inputs | E52 | Margin | Cons | E44 | Margin | INPUT |
| 14 | Inputs | F22 | Development Schedule | Cons | J16 | % | INPUT |
| 15 | Inputs | F23 | Construction spend profile | Cons | J16 | % | INPUT |
| 16 | Inputs | F62 | Calendar Year | Cons | J43 | Base Rate | LINK |
| 17 | Inputs | F63 | Base Rate | Cons | J43 | Base Rate | INPUT |
| 18 | Timing | A4 | Period Start | Debt | A4 |  | INPUT |
| 19 | Timing | A5 | Period End | Debt | A5 |  | INPUT |
| 20 | Timing | A6 | Construction | Debt | A6 |  | INPUT |
| 21 | Timing | A7 | Operations | Debt | A7 |  | INPUT |
| 22 | Timing | C12 | Operations Year | Debt | C10 | Num# | INPUT |
| 23 | Timing | C13 | Calendar Year | Debt | C11 | Year | INPUT |
| 24 | Timing | C14 | Days in Period | Debt | C12 | Days | INPUT |
| 25 | Timing | C9 | Counters | Debt | C9 |  | INPUT |
| 26 | Inputs | G55 | Tenor | Debt | E16 | Tenor Flag | LINK |
| 27 | Inputs | E58 | % of opening balance at operations | Debt | E34 | Principal repayment (annualised) | LINK |
| 28 | Cons | J37 | Refinance | Debt | J20 | Refinance from Construction | LINK |
| 29 | Cons | J43 | Base Rate | Debt | J25 | Base Rate | LINK |
| 30 | Inputs | F51 | Period | Debt | J26 | Margin | INPUT |
| 31 | Inputs | F52 | Margin | Debt | J26 | Margin | INPUT |
| 32 | QCF | J33 | Cashflow Available for Debt Service (CFADS) | Debt | J39 | CFADS | LINK |
| 33 | Timing | A4 | Period Start | Ops | A4 |  | INPUT |
| 34 | Timing | A5 | Period End | Ops | A5 |  | INPUT |
| 35 | Timing | A6 | Construction | Ops | A6 |  | INPUT |
| 36 | Timing | A7 | Operations | Ops | A7 |  | INPUT |
| 37 | Timing | C12 | Operations Year | Ops | C10 | Num# | INPUT |
| 38 | Timing | C13 | Calendar Year | Ops | C11 | Year | INPUT |
| 39 | Inputs | C37 | Labour | Ops | C23 | AUD M | INPUT |
| 40 | Inputs | C38 | Land & Leasing | Ops | C24 | AUD M | INPUT |
| 41 | Inputs | C39 | Spare | Ops | C25 | AUD M | INPUT |
| 42 | Inputs | C43 | Other | Ops | C30 | AUD / MWh | INPUT |
| 43 | Inputs | C44 | Spare | Ops | C31 | AUD / MWh | INPUT |
| 44 | Timing | C9 | Counters | Ops | C9 |  | INPUT |
| 45 | Inputs | E27 | Period | Ops | J15 | Net Generation | INPUT |
| 46 | Inputs | E28 | Net Generation | Ops | J15 | Net Generation | INPUT |
| 47 | Inputs | F32 | Price | Ops | J18 | Market Price | INPUT |
| 48 | Inputs | E36 | Nominal | Ops | J23 | AUD M | INPUT |
| 49 | QCF | J40 | Cashflow Available for Equity | Ops | J48 | Cashflow available for Equity | LINK |
| 50 | QCF | J46 | Cash Balance B/f | Ops | J49 | Cash Balance B/f | LINK |
| 51 | Cons | J31 | Equity | Ops | J53 | Equity | LINK |
| 52 | Timing | A4 | Period Start | QCF | A4 |  | INPUT |
| 53 | Timing | A5 | Period End | QCF | A5 |  | INPUT |
| 54 | Timing | A6 | Construction | QCF | A6 |  | INPUT |
| 55 | Timing | A7 | Operations | QCF | A7 |  | INPUT |
| 56 | Cons | C23 | Construction Cost | QCF | C24 |  | INPUT |
| 57 | Cons | C24 | Interest During Construction | QCF | C25 |  | INPUT |
| 58 | Ops | J19 | Revenue | QCF | J14 | Sales | LINK |
| 59 | Ops | J43 | Fixed Costs | QCF | J17 | Fixed Costs | LINK |
| 60 | Ops | J44 | Variable Costs | QCF | J18 | Variable Costs | LINK |
| 61 | Cons | J28 | Debt | QCF | J29 | Debt | LINK |
| 62 | Cons | J31 | Equity | QCF | J30 | Equity | LINK |
| 63 | Debt | J21 | Principal Repayment | QCF | J36 | Principal | LINK |
| 64 | Debt | J30 | Interest During Operations | QCF | J37 | Interest | LINK |
| 65 | Ops | J50 | Free Cash Flow to Equity | QCF | J42 | Dividends | LINK |
| 66 | Inputs | E9 | Start | Timing | E6 | Construction | LINK |
| 67 | Inputs | E14 | Start | Timing | E7 | Operations | LINK |
| 68 | Inputs | E11 | End | Timing | F6 | Construction | LINK |
| 69 | Inputs | E16 | End | Timing | F7 | Operations | LINK |
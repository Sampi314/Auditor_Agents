# Calculation Reference

## Glossary
| Term | Definition |
|---|---|
| **CFADS** | Cash Flow Available for Debt Service — operating cash flow after CapEx and funding |
| **DSCR** | Debt Service Coverage Ratio — CFADS / (Interest + Principal Repayment) |
| **IDC** | Interest During Construction — interest accrued on drawn debt before operations begin |
| **FCFE** | Free Cash Flow to Equity — cash available for equity holders after debt service |
| **IRR** | Internal Rate of Return — discount rate at which NPV = 0 |
| **COD** | Commercial Operations Date — start of the operations period |
| **MWh** | Megawatt-hour — unit of electrical energy |

---

## Timing Sheet

| Sheet | Cell | Name | Readable Formula | Excel Formula | Notes |
|---|---|---|---|---|---|
| **Timing** | | **--- Period Dates ---** | | | |
| Timing | I5 | Period End (initial) | *= {Model_Start}* | `=Model_Start` | First period only |
| Timing | J4 | Period Start | = prior Period End + 1 | `=I5+1` | Same J4:BI4 |
| Timing | J5 | Period End | = end of month ({Months_Qtr} - 1 months after Period Start) | `=EOMONTH(J4,Months_Qtr-1)` | Same J5:BI5 |
| **Timing** | | **--- Phase Flags ---** | | | |
| Timing | E6 | Construction Start | *= Inputs :: Construction Start* | `=Inputs!E9` | |
| Timing | F6 | Construction End | *= Inputs :: Construction End* | `=Inputs!E11` | |
| Timing | J6 | Construction Flag | = if Period Start >= Construction Start AND Period End <= Construction End then 1 else 0 | `=IF(AND(J$4>=$E6,J$5<=$F6),1,0)` | Same J6:BI6 |
| Timing | E7 | Operations Start | *= Inputs :: Operations Start* | `=Inputs!E14` | |
| Timing | F7 | Operations End | *= Inputs :: Operations End* | `=Inputs!E16` | |
| Timing | J7 | Operations Flag | = if Period Start >= Operations Start AND Period End <= Operations End then 1 else 0 | `=IF(AND(J$4>=$E7,J$5<=$F7),1,0)` | Same J7:BI7 |
| **Timing** | | **--- Counters ---** | | | |
| Timing | J10 | Construction Quarters | = cumulative count of construction flag | `=SUM($J6:J6)*J6` | Same J10:BI10 |
| Timing | J11 | Operations Quarters | = cumulative count of operations flag | `=SUM($J7:J7)*J7` | Same J11:BI11 |
| Timing | J12 | Operations Year | = round up (Operations Quarters / {Qtrs_Yr}) | `=ROUNDUP(J11/Qtrs_Yr,0)` | Same J12:BI12 |
| Timing | J13 | Calendar Year | = year of Period End | `=YEAR(J5)` | Same J13:BI13 |
| Timing | J14 | Days in Period | = Period End - Period Start + 1 | `=J5-J4+1` | Same J14:BI14 |

---

## Inputs Sheet

| Sheet | Cell | Name | Readable Formula | Excel Formula | Notes |
|---|---|---|---|---|---|
| **Inputs** | | **--- Timing ---** | | | |
| Inputs | E6 | Model Start Date | *Assumption* | `2011-12-31` | |
| Inputs | E9 | Construction Start | *Assumption* | `2012-01-01` | |
| Inputs | E10 | Construction Duration | *Assumption* | `8` Qtrs | |
| Inputs | E11 | Construction End | = end of month (Start + Duration quarters - 1) | `=EOMONTH(E9,E10*Months_Qtr-1)` | = 2013-12-31 |
| Inputs | E14 | Operations Start | = Construction End + 1 day | `=E11+1` | = 2014-01-01 |
| Inputs | E15 | Operations Duration | *Assumption* | `10` Yrs | |
| Inputs | E16 | Operations End | = end of month (Start + Duration years * 12 - 1) | `=EOMONTH(E14,E15*Months_Yr-1)` | = 2023-12-31 |
| **Inputs** | | **--- CapEx ---** | | | |
| Inputs | E20 | Total Construction Cost | *Assumption* | `60` AUD M | |
| Inputs | F23:M23 | Spend Profile (Qtrs 1-8) | *Assumption* | `12.5%` each | Sums to 100% |
| **Inputs** | | **--- Volume & Price ---** | | | |
| Inputs | E28 | Net Generation | *Assumption* | `130,000` MWh p.a. | Constant across 10 years |
| Inputs | F32:I32 | Market Price (Yrs 1-4) | *Assumption* | `110, 112, 114.5, 115` AUD/MWh | |
| **Inputs** | | **--- OpEx ---** | | | |
| Inputs | E37:N37 | Labour (Fixed, Yrs 1-10) | *Assumption* | `0.515` escalating ~3% p.a. | AUD M |
| Inputs | E38:N38 | Land & Leasing (Fixed, Yrs 1-10) | *Assumption* | `1.03` escalating ~3% p.a. | AUD M |
| Inputs | E43:N43 | Other (Variable, Yrs 1-10) | *Assumption* | `1.545` escalating ~3% p.a. | AUD/MWh |
| **Inputs** | | **--- Funding ---** | | | |
| Inputs | E49 | Senior Debt Limit | *Assumption* | `50` AUD M | |
| Inputs | E52 | Construction Margin | *Assumption* | `3.50%` | |
| Inputs | F52 | Operations Margin | *Assumption* | `2.50%` | |
| Inputs | E55 | Debt Tenor | *Assumption* | `6` Yrs | |
| Inputs | E58 | Repayment Rate | *Assumption* | `16.67%` p.a. (linear) | |
| Inputs | F62:I62 | Base Interest Rate (2012-2015) | *Assumption* | `3.6%, 3.0%, 2.9%, 3.0%` | |

---

## Cons Sheet (Construction)

| Sheet | Cell | Name | Readable Formula | Excel Formula | Notes |
|---|---|---|---|---|---|
| **Cons** | | **--- Construction Costs ---** | | | |
| Cons | J16 | Spend Profile % | = lookup Construction Quarter in Inputs spend profile | `=LOOKUP(J$10,Inputs!$F$22:$Q$22,Inputs!$F$23:$Q$23)` | Same J16:BI16 |
| Cons | J19 | Construction Costs | = Total CapEx x Spend Profile x (1 + escalation) | `=$E19*J16*(1+J$16)` | Same J19:BI19 |
| **Cons** | | **--- Funding ---** | | | |
| Cons | J23 | Uses: Construction Cost | = Construction Costs | `=J19` | Same J23:BI23 |
| Cons | J24 | Uses: IDC | = Interest During Construction | `=J48` | Same J24:BI24 |
| Cons | J25 | Total Uses | = sum(Construction Cost to IDC) | `=SUM(J23:J24)` | Same J25:BI25 |
| Cons | J28 | Debt Drawdowns | = min(Total Uses, Undrawn Amount) | `=MIN(J25,J40)` | Same J28:BI28 |
| Cons | J29 | Shortfall (Equity Required) | = Total Uses - Debt | `=J25-J28` | Same J29:BI29 |
| Cons | J31 | Equity | = Shortfall | `=J29` | Same J31:BI31 |
| Cons | J32 | Check | = Total Uses - Debt - Equity (should = 0) | `=J25-J28-J31` | Same J32:BI32 |
| **Cons** | | **--- Debt Account ---** | | | |
| Cons | J35 | Balance B/f | = prior period Balance C/f | `=I38` | Same J35:BI35 |
| Cons | J36 | Drawdowns | = Debt Drawdowns | `=J28` | Same J36:BI36 |
| Cons | J37 | Refinance | = if Period End = Construction End then -(B/f + Drawdowns) else 0 | `=-IF(J5=$E37,SUM(J35:J36),0)` | Zeroes out construction debt at COD |
| Cons | J38 | Balance C/f | = sum(B/f + Drawdowns + Refinance) | `=SUM(J35:J37)` | Same J38:BI38 |
| Cons | J40 | Undrawn Amount | = (Facility Size - Balance B/f) x Construction Flag | `=($E40-J35)*J$6` | Same J40:BI40 |
| **Cons** | | **--- Interest ---** | | | |
| Cons | J43 | Base Rate | = lookup Calendar Year in Inputs base rates | `=LOOKUP(J$13,Inputs!$F$61:$I$61,Inputs!$F$62:$I$62)` | Same J43:BI43 |
| Cons | J44 | Margin | *= Inputs :: Construction Margin* | `=Inputs!$E$52` | Same J44:BI44 |
| Cons | J45 | All-in Rate (% p.a.) | = (Base + Margin) x Construction Flag | `=(J43+J44)*J$6` | Same J45:BI45 |
| Cons | J46 | Per-period Rate | = (1 + Annual Rate) ^ (Days / 365) - 1 | `=(1+J45)^(J$14/Days_Yr)-1` | Same J46:BI46 |
| Cons | J48 | IDC | = Balance B/f x Per-period Rate | `=J35*J46` | Same J48:BI48 |

---

## Ops Sheet (Operations)

| Sheet | Cell | Name | Readable Formula | Excel Formula | Notes |
|---|---|---|---|---|---|
| **Ops** | | **--- Revenue ---** | | | |
| Ops | J15 | Net Generation (MWh) | = if Ops Year = 0 then 0 else lookup Ops Year in Inputs volume / {Qtrs_Yr} | `=IF(J$10=0,0,LOOKUP(J$10,Inputs!$E$27:$M$27,Inputs!$E$28:$M$28))/Qtrs_Yr` | J15:AG15 |
| Ops | AH15 | Net Generation (MWh) | = prior Net Generation x (1 + 0.5%) | `=AG15*(1+0.5%)` | AH15:BI15 (hard-coded growth) |
| Ops | J18 | Market Price (AUD/MWh) | = lookup Calendar Year in Inputs prices x Operations Flag | `=LOOKUP(J$13,Inputs!$F$31:$I$31,Inputs!$F$32:$I$32)*J$7` | Same J18:BI18 |
| Ops | J19 | Revenue (AUD M) | = Net Generation x Market Price / {Million} | `=J15*J18/Million` | Same J19:BI19 |
| **Ops** | | **--- Fixed OpEx ---** | | | |
| Ops | J23 | Labour (AUD M) | = if Ops Year = 0 then 0 else lookup in Inputs Labour / {Qtrs_Yr} | `=IF(J$10=0,0,LOOKUP(J$10,...))` | Same J23:BI23 |
| Ops | J24 | Land & Leasing (AUD M) | = same pattern as Labour | `=IF(J$10=0,0,LOOKUP(J$10,...))` | Same J24:BI24 |
| Ops | J25 | Spare (AUD M) | = same pattern (currently zero) | `=IF(J$10=0,0,LOOKUP(J$10,...))` | Same J25:BI25 |
| Ops | J26 | Total Fixed Costs | = sum(Labour to Spare) | `=SUM(J23:J25)` | **J26 only** |
| Ops | K26 | Total Fixed Costs | = sum(Labour to Land & Leasing) | `=SUM(K23:K24)` | K26:BI26 (excludes Spare) |
| **Ops** | | **--- Variable OpEx ---** | | | |
| Ops | J30 | Other (AUD/MWh) | = Inputs Other Cost x Operations Flag | `=Inputs!E43*$J7` | Same J30:BI30 |
| Ops | J37 | Other (AUD M) | = Cost Price x Net Generation / {Million} | `=J30*J$35/Million` | Same J37:BI37 |
| Ops | J40 | Total Variable Costs | = sum(Other to Spare) | `=SUM(J37:J38)` | **J40 only** |
| Ops | K40 | Total Variable Costs | = sum(Other to Spare to Spare) | `=SUM(K37:K39)` | K40:BI40 (includes Row 39) |
| **Ops** | | **--- Summary ---** | | | |
| Ops | J43 | Fixed Costs | = Total Fixed Costs | `=J26` | Same J43:BI43 |
| Ops | J44 | Variable Costs | = Total Variable Costs | `=J40` | Same J44:BI44 |
| Ops | J45 | Total Costs | = sum(Fixed + Variable) | `=SUM(J43:J44)` | Same J45:BI45 |
| **Ops** | | **--- Equity Returns ---** | | | |
| Ops | J48 | Cashflow (from QCF) | *= QCF :: CF Available for Equity* | `=QCF!J40` | Same J48:BI48 |
| Ops | J49 | Cash Balance B/f (from QCF) | *= QCF :: Cash Balance B/f* | `=QCF!J46` | Same J49:BI49 |
| Ops | J50 | FCFE | = max(Cashflow + Cash Balance, 0) | `=MAX(J48+J49,0)` | Same J50:BI50 |
| Ops | J53 | Equity Investment | = negative of Cons :: Equity | `=-Cons!J31` | Same J53:BI53 |
| Ops | J54 | Dividends | = FCFE | `=J50` | Same J54:BI54 |
| Ops | J55 | Net Cashflow | = Equity Investment + Dividends | `=SUM(J53:J54)` | Same J55:BI55 |
| Ops | I57 | Evaluation CF (seed) | *Hard-coded seed for XIRR* | `=-0.00001` | Ensures sign change |
| Ops | J57 | Evaluation CF | = Net Cashflow | `=J55` | Same J57:BI57 |
| Ops | I59 | Equity IRR | = XIRR of evaluation cashflows against period end dates | `=XIRR(I57:BI57,I5:BI5)` | |

---

## QCF Sheet (Quarterly Cash Flow)

| Sheet | Cell | Name | Readable Formula | Excel Formula | Notes |
|---|---|---|---|---|---|
| **QCF** | | **--- Revenue ---** | | | |
| QCF | J14 | Sales | *= Ops :: Revenue* | `=Ops!J19` | Same J14:BI14 |
| **QCF** | | **--- Operating Expense ---** | | | |
| QCF | J17 | Fixed Costs | = negative of Ops :: Fixed Costs | `=-Ops!J43` | Same J17:BI17 |
| QCF | J18 | Variable Costs | = negative of Ops :: Variable Costs | `=-Ops!J44` | Same J18:BI18 |
| QCF | J19 | Total OpEx | = sum(Fixed + Variable) | `=SUM(J17:J18)` | Same J19:BI19 |
| **QCF** | | **--- Operating Cashflow ---** | | | |
| QCF | J21 | Operating Cashflow | = Revenue + Total OpEx | `=J14+J19` | Same J21:BI21 |
| **QCF** | | **--- Capital Expenditure ---** | | | |
| QCF | J24 | Construction Cost | = negative of Cons :: Construction Cost | `=-Cons!J23` | Same J24:BI24 |
| QCF | J25 | IDC | = negative of Cons :: IDC | `=-Cons!J24` | Same J25:BI25 |
| QCF | J26 | Total CapEx | = sum(Construction + IDC) | `=SUM(J24:J25)` | Same J26:BI26 |
| **QCF** | | **--- Funding ---** | | | |
| QCF | J29 | Debt | *= Cons :: Debt Drawdowns* | `=Cons!J28` | Same J29:BI29 |
| QCF | J30 | Equity | *= Cons :: Equity* | `=Cons!J31` | Same J30:BI30 |
| QCF | J31 | Total Funding | = sum(Debt + Equity) | `=SUM(J29:J30)` | Same J31:BI31 |
| **QCF** | | **--- CFADS ---** | | | |
| QCF | J33 | CFADS | = Operating CF + CapEx + Funding | `=J21+J26+J31` | Same J33:BI33 |
| **QCF** | | **--- Debt Service ---** | | | |
| QCF | J36 | Principal | *= Debt :: Principal Repayment* | `=Debt!J21` | Same J36:BI36 |
| QCF | J37 | Interest | = negative of Debt :: Interest | `=-Debt!J30` | Same J37:BI37 |
| QCF | J38 | Total Debt Service | = sum(Principal + Interest) | `=SUM(J36:J37)` | Same J38:BI38 |
| **QCF** | | **--- Equity ---** | | | |
| QCF | J40 | CF Available for Equity | = CFADS + Debt Service | `=J33+J38` | Same J40:BI40 |
| QCF | J42 | Dividends | = negative of Ops :: FCFE | `=-Ops!J50` | Same J42:BI42 |
| QCF | J44 | Net Cashflow | = CF for Equity + Dividends | `=J40+J42` | Same J44:BI44 |
| **QCF** | | **--- Cash Balance ---** | | | |
| QCF | J46 | Balance B/f | = prior period Balance C/f | `=I47` | Same J46:BI46 |
| QCF | J47 | Balance C/f | = Net Cashflow + Balance B/f | `=J44+J46` | Same J47:BI47 |

---

## Debt Sheet

| Sheet | Cell | Name | Readable Formula | Excel Formula | Notes |
|---|---|---|---|---|---|
| **Debt** | | **--- Tenor ---** | | | |
| Debt | E16 | Tenor Start | *= Inputs :: COD* | `=Inputs!G55` | = 2014-01-01 |
| Debt | F16 | Tenor End | *= Inputs :: Tenor End* | `=Inputs!I55` | = 2019-12-31 |
| Debt | J16 | Tenor Flag | = if Period Start >= Tenor Start AND Period End <= Tenor End then 1 else 0 | `=IF(AND(J4>=E16,J5<=F16),1,0)` | Same J16:BI16 |
| **Debt** | | **--- Debt Account ---** | | | |
| Debt | J19 | Balance B/f | = prior period Balance C/f | `=I22` | Same J19:BI19 |
| Debt | J20 | Refinance from Construction | = negative of Cons :: Refinance | `=-Cons!J37` | Same J20:BI20 |
| Debt | J21 | Principal Repayment | = negative of Quarterly Principal | `=-J35` | Same J21:BI21 |
| Debt | J22 | Balance C/f | = sum(B/f + Refinance + Repayment) | `=SUM(J19:J21)` | Same J22:BI22 |
| **Debt** | | **--- Interest ---** | | | |
| Debt | J25 | Base Rate | *= Cons :: Base Rate* | `=Cons!J43` | Same J25:BI25 |
| Debt | J26 | Margin | = if Ops Year = 0 then 0 else lookup in Inputs margins | `=IF(J10=0,0,LOOKUP(J10,...))` | Same J26:BI26 |
| Debt | J27 | All-in Rate (% p.a.) | = Base + Margin | `=SUM(J25:J26)` | Same J27:BI27 |
| Debt | J28 | All-in Rate (% p.q.) | = Annual Rate / {Qtrs_Yr} | `=J27/Qtrs_Yr` | Same J28:BI28 |
| Debt | J30 | Interest | = Balance B/f x Quarterly Rate x Tenor Flag | `=J19*J28*J16` | Same J30:BI30 |
| **Debt** | | **--- Repayment ---** | | | |
| Debt | E34 | Repayment Rate | *= Inputs :: Repayment Rate* | `=Inputs!E58` | 16.67% p.a. |
| Debt | J34 | Annual Repayment | = Total Refinanced Debt x Repayment Rate x Tenor Flag | `=$I$20*$E34*J16` | Same J34:BI34 |
| Debt | J35 | Quarterly Repayment | = Annual Repayment / {Qtrs_Yr} | `=J34/Qtrs_Yr` | Same J35:BI35 |
| **Debt** | | **--- DSCR ---** | | | |
| Debt | J39 | CFADS | = QCF :: CFADS x Tenor Flag | `=QCF!J33*J16` | Same J39:BI39 |
| Debt | J41 | Interest (mirror) | = Interest | `=J30` | Same J41:BI41 |
| Debt | J42 | Principal (mirror) | = Quarterly Repayment | `=J35` | Same J42:BI42 |
| Debt | J43 | Total Debt Service | = sum(Interest + Principal) | `=SUM(J41:J42)` | Same J43:BI43 |
| Debt | J45 | DSCR | = if Balance B/f > 0 then CFADS / Total Debt Service else 0 | `=IF(J19>0,J39/J43,0)` | Same J45:BI45 |
| Debt | E45 | Min DSCR | = min of all period DSCRs | Array formula | |
| Debt | F45 | Avg DSCR | = if Total Debt Service = 0 then 0 else Total CFADS / Total Debt Service | `=IF(I43=0,0,I39/I43)` | Weighted average |

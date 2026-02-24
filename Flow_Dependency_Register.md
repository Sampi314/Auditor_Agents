# Flow Dependency Register

| # | Source Sheet | Source Cell | Source Row Label | Target Sheet | Target Cell | Target Row Label | Reference Type |
|---|---|---|---|---|---|---|---|
| 1 | IS | B33 | EBIT | Annual_Summary | B10 | EBIT ($000s) | LINK |
| 2 | IS | B40 | NET INCOME | Annual_Summary | B11 | Net Income ($000s) | LINK |
| 3 | Calc_Capex | B14 | Total Capex | Annual_Summary | B12 | Capex ($000s) | LINK |
| 4 | CF | B10 | Cash from Operations | Annual_Summary | B13 | Cash from Ops ($000s) | LINK |
| 5 | CF | M25 | Closing Cash | Annual_Summary | B14 | Closing Cash ($000s) | LINK |
| 6 | Calc_Debt | M20 | Total Debt | Annual_Summary | B15 | Total Debt ($000s) | LINK |
| 7 | BS | M36 | TOTAL EQUITY | Annual_Summary | B16 | Total Equity ($000s) | LINK |
| 8 | Calc_Volume | B14 | Total | Annual_Summary | B4 | Volume (HL) | LINK |
| 9 | Calc_Revenue | B12 | Total Gross Revenue | Annual_Summary | B5 | Gross Revenue ($000s) | LINK |
| 10 | Calc_Revenue | B25 | Net Revenue | Annual_Summary | B6 | Net Revenue ($000s) | LINK |
| 11 | Calc_COGS | B28 | Total COGS | Annual_Summary | B7 | COGS ($000s) | LINK |
| 12 | IS | B19 | Gross Profit | Annual_Summary | B8 | Gross Profit ($000s) | LINK |
| 13 | IS | B29 | EBITDA | Annual_Summary | B9 | EBITDA ($000s) | LINK |
| 14 | Assum_Volume | C16 | On-Trade % | Assum_WC | C13 | Blended Average | INPUT |
| 15 | Assum_Volume | C17 | Off-Trade % | Assum_WC | C13 | Blended Average | LINK |
| 16 | Calc_WC | B10 | Total Inventory | BS | B10 | Inventory | LINK |
| 17 | Calc_WC | B19 | Prepayments | BS | B11 | Prepayments | LINK |
| 18 | Calc_Capex | B20 | Net PP&E - Closing | BS | B15 | Property, Plant & Equipment | LINK |
| 19 | Calc_WC | B16 | Trade Payables | BS | B22 | Trade Payables | LINK |
| 20 | Calc_WC | B20 | Accruals | BS | B23 | Accruals | LINK |
| 21 | Assum_Finance | C7 | Term Loan - Annual Principal Payment | BS | B24 | Current Debt | LINK |
| 22 | Calc_Debt | B9 | Closing Balance | BS | B24 | Current Debt | LINK |
| 23 | Calc_Debt | B20 | Total Debt | BS | B28 | Long-term Debt | LINK |
| 24 | Calc_Volume | B3 | Period | BS | B3 | Period | INPUT |
| 25 | Assum_Finance | C18 | Opening Share Capital | BS | B34 | Share Capital | INPUT |
| 26 | Assum_Finance | C19 | Opening Retained Earnings | BS | B35 | Retained Earnings | INPUT |
| 27 | CF | B19 | Dividends Paid | BS | B35 | Retained Earnings | LINK |
| 28 | IS | B40 | NET INCOME | BS | B35 | Retained Earnings | LINK |
| 29 | Calc_Volume | B4 | Date | BS | B4 | Date | INPUT |
| 30 | CF | B25 | Closing Cash | BS | B8 | Cash | LINK |
| 31 | Calc_WC | B13 | Trade Receivables | BS | B9 | Trade Receivables | LINK |
| 32 | Calc_Capex | B14 | Total Capex | CF | B13 | Capital Expenditure | LINK |
| 33 | Calc_Debt | B8 | Principal Payment | CF | B17 | Term Loan Principal Payments | LINK |
| 34 | Calc_Debt | B14 | Draw / (Repay) | CF | B18 | Revolver Draw / (Repay) | INPUT |
| 35 | Assum_Finance | C23 | Target Payout Ratio (% of Net Income) | CF | B19 | Dividends Paid | INPUT |
| 36 | IS | B40 | NET INCOME | CF | B19 | Dividends Paid | LINK |
| 37 | Assum_Finance | C28 | Opening Cash Balance | CF | B24 | Opening Cash | INPUT |
| 38 | Calc_Volume | B3 | Period | CF | B3 | Period | INPUT |
| 39 | Calc_Volume | B4 | Date | CF | B4 | Date | INPUT |
| 40 | IS | B40 | NET INCOME | CF | B7 | Net Income | LINK |
| 41 | Calc_Capex | B15 | Depreciation | CF | B8 | Add: Depreciation | LINK |
| 42 | Calc_WC | B24 | Change in NWC | CF | B9 | Change in Working Capital | LINK |
| 43 | Assum_Costs | C8 | Yeast | Calc_COGS | B10 | Yeast | INPUT |
| 44 | Calc_Volume | B14 | Total | Calc_COGS | B10 | Yeast | LINK |
| 45 | Assum_Costs | C9 | Water & Treatment | Calc_COGS | B11 | Water & Treatment | INPUT |
| 46 | Calc_Volume | B14 | Total | Calc_COGS | B11 | Water & Treatment | LINK |
| 47 | Assum_Costs | C13 | Bottles (glass, labels, caps, cartons) | Calc_COGS | B15 | Bottles | INPUT |
| 48 | Calc_Volume | B25 | Bottles | Calc_COGS | B15 | Bottles | LINK |
| 49 | Assum_Costs | C14 | Cans (aluminum, labels, cartons) | Calc_COGS | B16 | Cans | INPUT |
| 50 | Calc_Volume | B26 | Cans | Calc_COGS | B16 | Cans | LINK |
| 51 | Assum_Costs | C15 | Kegs (depreciation, cleaning, fittings) | Calc_COGS | B17 | Kegs | INPUT |
| 52 | Calc_Volume | B27 | Kegs | Calc_COGS | B17 | Kegs | LINK |
| 53 | Assum_Costs | C18 | Direct Labour | Calc_COGS | B21 | Direct Labour | INPUT |
| 54 | Calc_Volume | B14 | Total | Calc_COGS | B21 | Direct Labour | LINK |
| 55 | Assum_Costs | C19 | Utilities (power, gas, water) | Calc_COGS | B22 | Utilities | INPUT |
| 56 | Calc_Volume | B14 | Total | Calc_COGS | B22 | Utilities | LINK |
| 57 | Assum_Costs | C20 | Manufacturing Overhead | Calc_COGS | B23 | Manufacturing Overhead | INPUT |
| 58 | Calc_Volume | B14 | Total | Calc_COGS | B23 | Manufacturing Overhead | LINK |
| 59 | Assum_Costs | C21 | Quality Control | Calc_COGS | B24 | Quality Control | INPUT |
| 60 | Calc_Volume | B14 | Total | Calc_COGS | B24 | Quality Control | LINK |
| 61 | Calc_Volume | B3 | Period | Calc_COGS | B3 | Period | INPUT |
| 62 | Calc_Volume | B14 | Total | Calc_COGS | B31 | COGS per HL | LINK |
| 63 | Calc_Revenue | B25 | Net Revenue | Calc_COGS | B32 | Gross Margin % | LINK |
| 64 | Calc_Volume | B4 | Date | Calc_COGS | B4 | Date | INPUT |
| 65 | Assum_Costs | C6 | Malt & Adjuncts | Calc_COGS | B8 | Malt & Adjuncts | INPUT |
| 66 | Calc_Volume | B14 | Total | Calc_COGS | B8 | Malt & Adjuncts | LINK |
| 67 | Assum_Costs | C7 | Hops | Calc_COGS | B9 | Hops | INPUT |
| 68 | Calc_Volume | B14 | Total | Calc_COGS | B9 | Hops | LINK |
| 69 | Assum_Costs | C47 | Base Scenario | Calc_COGS | N5 | Cost Index (cumulative) | INPUT |
| 70 | Assum_Costs | C49 | Downside Scenario | Calc_COGS | N5 | Cost Index (cumulative) | INPUT |
| 71 | Control | D5 | Active Scenario (1=Base, 2=Upside, 3=Downside) | Calc_COGS | N5 | Cost Index (cumulative) | INPUT |
| 72 | Assum_Capex | C31 | Annual Maintenance Capex (% of Gross PP&E) | Calc_Capex | B12 | Maintenance Capex | INPUT |
| 73 | Assum_Capex | F13 | TOTAL | Calc_Capex | B15 | Depreciation | LINK |
| 74 | Assum_Capex | C17 | Fermentation Capacity (HL/year) | Calc_Capex | B23 | Fermentation Capacity (HL/month) | INPUT |
| 75 | Calc_Volume | B14 | Total | Calc_Capex | B24 | Volume Produced (HL) | LINK |
| 76 | Calc_Volume | B3 | Period | Calc_Capex | B3 | Period | INPUT |
| 77 | Calc_Volume | B4 | Date | Calc_Capex | B4 | Date | INPUT |
| 78 | Assum_Capex | B13 | TOTAL | Calc_Capex | B7 | Gross PP&E - Opening | LINK |
| 79 | Assum_Capex | C13 | TOTAL | Calc_Capex | B8 | Accumulated Depreciation - Opening | LINK |
| 80 | Assum_Finance | C5 | Term Loan - Interest Rate | Calc_Debt | B10 | Interest Expense | INPUT |
| 81 | Assum_Finance | C12 | Current Drawn Amount | Calc_Debt | B13 | Opening Balance | INPUT |
| 82 | Assum_Finance | C13 | Interest Rate | Calc_Debt | B16 | Interest Expense | INPUT |
| 83 | Assum_Finance | C11 | Facility Limit | Calc_Debt | B17 | Commitment Fee | INPUT |
| 84 | Assum_Finance | C14 | Commitment Fee (on undrawn) | Calc_Debt | B17 | Commitment Fee | INPUT |
| 85 | Calc_Volume | B3 | Period | Calc_Debt | B3 | Period | INPUT |
| 86 | Calc_Volume | B4 | Date | Calc_Debt | B4 | Date | INPUT |
| 87 | Assum_Finance | C4 | Term Loan - Principal Outstanding | Calc_Debt | B7 | Opening Balance | INPUT |
| 88 | Assum_Finance | C7 | Term Loan - Annual Principal Payment | Calc_Debt | B8 | Principal Payment | LINK |
| 89 | Assum_Costs | C35 | Cold Chain Premium (Kegs) | Calc_Opex | B10 | Cold Chain (Kegs) | INPUT |
| 90 | Calc_COGS | B5 | Cost Index (cumulative) | Calc_Opex | B10 | Cold Chain (Kegs) | INPUT |
| 91 | Calc_Volume | B27 | Kegs | Calc_Opex | B10 | Cold Chain (Kegs) | LINK |
| 92 | Assum_Costs | C39 | Sales Team | Calc_Opex | B14 | Sales Team | INPUT |
| 93 | Calc_Revenue | B12 | Total Gross Revenue | Calc_Opex | B14 | Sales Team | LINK |
| 94 | Assum_Costs | C40 | Marketing | Calc_Opex | B15 | Marketing | INPUT |
| 95 | Calc_Revenue | B12 | Total Gross Revenue | Calc_Opex | B15 | Marketing | LINK |
| 96 | Assum_Costs | C41 | Trade Marketing | Calc_Opex | B16 | Trade Marketing | INPUT |
| 97 | Calc_Revenue | B12 | Total Gross Revenue | Calc_Opex | B16 | Trade Marketing | LINK |
| 98 | Assum_Costs | C42 | General & Admin | Calc_Opex | B17 | General & Admin | INPUT |
| 99 | Calc_Revenue | B12 | Total Gross Revenue | Calc_Opex | B17 | General & Admin | LINK |
| 100 | Calc_Volume | B3 | Period | Calc_Opex | B3 | Period | INPUT |
| 101 | Calc_Volume | B4 | Date | Calc_Opex | B4 | Date | INPUT |
| 102 | Assum_Costs | C31 | Domestic Freight - Australia | Calc_Opex | B7 | Australia Freight | INPUT |
| 103 | Calc_COGS | B5 | Cost Index (cumulative) | Calc_Opex | B7 | Australia Freight | INPUT |
| 104 | Calc_Volume | B17 | Australia | Calc_Opex | B7 | Australia Freight | LINK |
| 105 | Assum_Costs | C32 | Domestic Freight - New Zealand | Calc_Opex | B8 | New Zealand Freight | INPUT |
| 106 | Assum_Costs | C33 | Export Freight (AU to NZ) | Calc_Opex | B8 | New Zealand Freight | INPUT |
| 107 | Calc_COGS | B5 | Cost Index (cumulative) | Calc_Opex | B8 | New Zealand Freight | INPUT |
| 108 | Calc_Volume | B18 | New Zealand | Calc_Opex | B8 | New Zealand Freight | LINK |
| 109 | Assum_Costs | C34 | Warehousing | Calc_Opex | B9 | Warehousing | INPUT |
| 110 | Calc_COGS | B5 | Cost Index (cumulative) | Calc_Opex | B9 | Warehousing | INPUT |
| 111 | Calc_Volume | B14 | Total | Calc_Opex | B9 | Warehousing | LINK |
| 112 | Assum_Pricing | C10 | AllDark | Calc_Revenue | B10 | AllDark On-Trade | INPUT |
| 113 | Assum_Volume | C16 | On-Trade % | Calc_Revenue | B10 | AllDark On-Trade | INPUT |
| 114 | Assum_Volume | C21 | Kegs % | Calc_Revenue | B10 | AllDark On-Trade | INPUT |
| 115 | Assum_Volume | C22 | Bottles % | Calc_Revenue | B10 | AllDark On-Trade | INPUT |
| 116 | Assum_Volume | C23 | Cans % | Calc_Revenue | B10 | AllDark On-Trade | LINK |
| 117 | Calc_Volume | B13 | AllDark | Calc_Revenue | B10 | AllDark On-Trade | LINK |
| 118 | Assum_Pricing | D10 | AllDark | Calc_Revenue | B11 | AllDark Off-Trade | INPUT |
| 119 | Assum_Volume | C17 | Off-Trade % | Calc_Revenue | B11 | AllDark Off-Trade | LINK |
| 120 | Assum_Volume | C26 | Bottles % | Calc_Revenue | B11 | AllDark Off-Trade | INPUT |
| 121 | Assum_Volume | C27 | Cans % | Calc_Revenue | B11 | AllDark Off-Trade | INPUT |
| 122 | Calc_Volume | B13 | AllDark | Calc_Revenue | B11 | AllDark Off-Trade | LINK |
| 123 | Assum_Pricing | C15 | On-Trade Discount % | Calc_Revenue | B15 | On-Trade Discounts | INPUT |
| 124 | Assum_Pricing | C18 | Blended Off-Trade Discount | Calc_Revenue | B16 | Off-Trade Discounts | LINK |
| 125 | Assum_Pricing | C27 | AU SkyBrew (5% ABV) - Standard | Calc_Revenue | B20 | Australia Excise | LINK |
| 126 | Assum_Pricing | C28 | AU SkyBrew (5% ABV) - Keg | Calc_Revenue | B20 | Australia Excise | LINK |
| 127 | Assum_Pricing | C29 | AU AllDark (6% ABV) - Standard | Calc_Revenue | B20 | Australia Excise | LINK |
| 128 | Assum_Pricing | C30 | AU AllDark (6% ABV) - Keg | Calc_Revenue | B20 | Australia Excise | LINK |
| 129 | Assum_Volume | C16 | On-Trade % | Calc_Revenue | B20 | Australia Excise | INPUT |
| 130 | Assum_Volume | C21 | Kegs % | Calc_Revenue | B20 | Australia Excise | INPUT |
| 131 | Assum_Volume | C8 | SkyBrew % | Calc_Revenue | B20 | Australia Excise | INPUT |
| 132 | Assum_Volume | C9 | AllDark % | Calc_Revenue | B20 | Australia Excise | LINK |
| 133 | Calc_Volume | B17 | Australia | Calc_Revenue | B20 | Australia Excise | LINK |
| 134 | Assum_Pricing | C31 | NZ SkyBrew (5% ABV) | Calc_Revenue | B21 | New Zealand Excise | LINK |
| 135 | Assum_Pricing | C32 | NZ AllDark (6% ABV) | Calc_Revenue | B21 | New Zealand Excise | LINK |
| 136 | Assum_Volume | C8 | SkyBrew % | Calc_Revenue | B21 | New Zealand Excise | INPUT |
| 137 | Assum_Volume | C9 | AllDark % | Calc_Revenue | B21 | New Zealand Excise | LINK |
| 138 | Calc_Volume | B18 | New Zealand | Calc_Revenue | B21 | New Zealand Excise | LINK |
| 139 | Calc_Volume | B14 | Total | Calc_Revenue | B28 | Gross Revenue per HL | LINK |
| 140 | Calc_Volume | B14 | Total | Calc_Revenue | B29 | Net Revenue per HL | LINK |
| 141 | Calc_Volume | B3 | Period | Calc_Revenue | B3 | Period | INPUT |
| 142 | Calc_Volume | B14 | Total | Calc_Revenue | B30 | Excise per HL | LINK |
| 143 | Calc_Volume | B4 | Date | Calc_Revenue | B4 | Date | INPUT |
| 144 | Assum_Pricing | C6 | SkyBrew | Calc_Revenue | B8 | SkyBrew On-Trade | INPUT |
| 145 | Assum_Volume | C16 | On-Trade % | Calc_Revenue | B8 | SkyBrew On-Trade | INPUT |
| 146 | Assum_Volume | C21 | Kegs % | Calc_Revenue | B8 | SkyBrew On-Trade | INPUT |
| 147 | Assum_Volume | C22 | Bottles % | Calc_Revenue | B8 | SkyBrew On-Trade | INPUT |
| 148 | Assum_Volume | C23 | Cans % | Calc_Revenue | B8 | SkyBrew On-Trade | LINK |
| 149 | Calc_Volume | B12 | SkyBrew | Calc_Revenue | B8 | SkyBrew On-Trade | LINK |
| 150 | Assum_Pricing | D6 | SkyBrew | Calc_Revenue | B9 | SkyBrew Off-Trade | INPUT |
| 151 | Assum_Volume | C17 | Off-Trade % | Calc_Revenue | B9 | SkyBrew Off-Trade | LINK |
| 152 | Assum_Volume | C26 | Bottles % | Calc_Revenue | B9 | SkyBrew Off-Trade | INPUT |
| 153 | Assum_Volume | C27 | Cans % | Calc_Revenue | B9 | SkyBrew Off-Trade | INPUT |
| 154 | Calc_Volume | B12 | SkyBrew | Calc_Revenue | B9 | SkyBrew Off-Trade | LINK |
| 155 | Assum_Pricing | C36 | Base Scenario | Calc_Revenue | N5 | Price Index (cumulative) | INPUT |
| 156 | Assum_Pricing | C38 | Downside Scenario | Calc_Revenue | N5 | Price Index (cumulative) | INPUT |
| 157 | Control | D5 | Active Scenario (1=Base, 2=Upside, 3=Downside) | Calc_Revenue | N5 | Price Index (cumulative) | INPUT |
| 158 | Assum_Volume | C5 | Total Base Volume (HL) | Calc_Volume | B12 | SkyBrew | INPUT |
| 159 | Assum_Volume | C8 | SkyBrew % | Calc_Volume | B12 | SkyBrew | INPUT |
| 160 | Assum_Volume | C5 | Total Base Volume (HL) | Calc_Volume | B13 | AllDark | INPUT |
| 161 | Assum_Volume | C9 | AllDark % | Calc_Volume | B13 | AllDark | LINK |
| 162 | Assum_Volume | C12 | Australia % | Calc_Volume | B17 | Australia | INPUT |
| 163 | Assum_Volume | C13 | New Zealand % | Calc_Volume | B18 | New Zealand | LINK |
| 164 | Assum_Volume | C16 | On-Trade % | Calc_Volume | B21 | On-Trade | INPUT |
| 165 | Assum_Volume | C17 | Off-Trade % | Calc_Volume | B22 | Off-Trade | LINK |
| 166 | Assum_Volume | C22 | Bottles % | Calc_Volume | B25 | Bottles | INPUT |
| 167 | Assum_Volume | C23 | Cans % | Calc_Volume | B26 | Cans | LINK |
| 168 | Assum_Volume | C21 | Kegs % | Calc_Volume | B27 | Kegs | INPUT |
| 169 | Assum_Volume | B46 | Jan | Calc_Volume | B7 | Seasonality Index | INPUT |
| 170 | Assum_Volume | B47 | Feb | Calc_Volume | C7 | Seasonality Index | INPUT |
| 171 | Assum_Volume | B48 | Mar | Calc_Volume | D7 | Seasonality Index | INPUT |
| 172 | Assum_Volume | B49 | Apr | Calc_Volume | E7 | Seasonality Index | INPUT |
| 173 | Assum_Volume | B50 | May | Calc_Volume | F7 | Seasonality Index | INPUT |
| 174 | Assum_Volume | B51 | Jun | Calc_Volume | G7 | Seasonality Index | INPUT |
| 175 | Assum_Volume | B52 | Jul | Calc_Volume | H7 | Seasonality Index | INPUT |
| 176 | Assum_Volume | B53 | Aug | Calc_Volume | I7 | Seasonality Index | INPUT |
| 177 | Assum_Volume | B54 | Sep | Calc_Volume | J7 | Seasonality Index | INPUT |
| 178 | Assum_Volume | B55 | Oct | Calc_Volume | K7 | Seasonality Index | INPUT |
| 179 | Assum_Volume | B56 | Nov | Calc_Volume | L7 | Seasonality Index | INPUT |
| 180 | Assum_Volume | B57 | Dec | Calc_Volume | M7 | Seasonality Index | INPUT |
| 181 | Assum_Volume | B32 | Year 1 | Calc_Volume | N9 | CUMULATIVE GROWTH INDEX | INPUT |
| 182 | Assum_Volume | D41 | Year 10 | Calc_Volume | N9 | CUMULATIVE GROWTH INDEX | INPUT |
| 183 | Control | D5 | Active Scenario (1=Base, 2=Upside, 3=Downside) | Calc_Volume | N9 | CUMULATIVE GROWTH INDEX | INPUT |
| 184 | Assum_WC | C13 | Blended Average | Calc_WC | B13 | Trade Receivables | LINK |
| 185 | Calc_Revenue | B25 | Net Revenue | Calc_WC | B13 | Trade Receivables | LINK |
| 186 | Assum_WC | C17 | Trade Payables (days of COGS + Opex) | Calc_WC | B16 | Trade Payables | INPUT |
| 187 | Calc_COGS | B28 | Total COGS | Calc_WC | B16 | Trade Payables | LINK |
| 188 | Calc_Opex | B21 | Total Operating Expenses | Calc_WC | B16 | Trade Payables | LINK |
| 189 | Assum_WC | C21 | Prepayments (days of Opex) | Calc_WC | B19 | Prepayments | INPUT |
| 190 | Calc_Opex | B21 | Total Operating Expenses | Calc_WC | B19 | Prepayments | LINK |
| 191 | Assum_WC | C22 | Accruals (days of Opex) | Calc_WC | B20 | Accruals | INPUT |
| 192 | Calc_Opex | B21 | Total Operating Expenses | Calc_WC | B20 | Accruals | LINK |
| 193 | Assum_WC | C32 | Opening Net Working Capital | Calc_WC | B24 | Change in NWC | LINK |
| 194 | Calc_Volume | B3 | Period | Calc_WC | B3 | Period | INPUT |
| 195 | Calc_Volume | B4 | Date | Calc_WC | B4 | Date | INPUT |
| 196 | Assum_WC | C4 | Raw Materials (days of COGS) | Calc_WC | B7 | Raw Materials | INPUT |
| 197 | Calc_COGS | B28 | Total COGS | Calc_WC | B7 | Raw Materials | LINK |
| 198 | Assum_WC | C5 | Work in Progress (days of COGS) | Calc_WC | B8 | Work in Progress | INPUT |
| 199 | Calc_COGS | B28 | Total COGS | Calc_WC | B8 | Work in Progress | LINK |
| 200 | Assum_WC | C6 | Finished Goods (days of COGS) | Calc_WC | B9 | Finished Goods | INPUT |
| 201 | Calc_COGS | B28 | Total COGS | Calc_WC | B9 | Finished Goods | LINK |
| 202 | Calc_Debt | B20 | Total Debt | Checks | B10 | Debt Never Negative | LINK |
| 203 | Calc_Volume | B14 | Total | Checks | B11 | Volume Never Negative | LINK |
| 204 | IS | B20 | Gross Margin % | Checks | B12 | Gross Margin Always Positive | LINK |
| 205 | BS | B18 | TOTAL ASSETS | Checks | B13 | Total Assets = Total L&E (Period 1) | LINK |
| 206 | BS | B38 | TOTAL LIABILITIES & EQUITY | Checks | B13 | Total Assets = Total L&E (Period 1) | LINK |
| 207 | BS | BI18 | TOTAL ASSETS | Checks | B14 | Total Assets = Total L&E (Period 60) | LINK |
| 208 | BS | BI38 | TOTAL LIABILITIES & EQUITY | Checks | B14 | Total Assets = Total L&E (Period 60) | LINK |
| 209 | BS | DQ18 | TOTAL ASSETS | Checks | B15 | Total Assets = Total L&E (Period 120) | LINK |
| 210 | BS | DQ38 | TOTAL LIABILITIES & EQUITY | Checks | B15 | Total Assets = Total L&E (Period 120) | LINK |
| 211 | Assum_Finance | C19 | Opening Retained Earnings | Checks | B16 | Net Income Flows to RE | INPUT |
| 212 | BS | DQ35 | Retained Earnings | Checks | B16 | Net Income Flows to RE | LINK |
| 213 | CF | B19 | Dividends Paid | Checks | B16 | Net Income Flows to RE | LINK |
| 214 | IS | B40 | NET INCOME | Checks | B16 | Net Income Flows to RE | LINK |
| 215 | BS | B41 | Assets - Liabilities - Equity | Checks | B4 | Balance Sheet Balances (all periods) | LINK |
| 216 | Assum_Volume | B58 | Seasonality Check (should = 12) | Checks | B5 | Seasonality Sums to 12 | LINK |
| 217 | Assum_Volume | C8 | SkyBrew % | Checks | B6 | Brand Mix Sums to 100% | INPUT |
| 218 | Assum_Volume | C9 | AllDark % | Checks | B6 | Brand Mix Sums to 100% | LINK |
| 219 | Assum_Volume | C12 | Australia % | Checks | B7 | Market Mix Sums to 100% | INPUT |
| 220 | Assum_Volume | C13 | New Zealand % | Checks | B7 | Market Mix Sums to 100% | LINK |
| 221 | Assum_Volume | C16 | On-Trade % | Checks | B8 | Channel Mix Sums to 100% | INPUT |
| 222 | Assum_Volume | C17 | Off-Trade % | Checks | B8 | Channel Mix Sums to 100% | LINK |
| 223 | CF | B25 | Closing Cash | Checks | B9 | Cash Never Negative | LINK |
| 224 | Calc_Revenue | B25 | Net Revenue | IS | B10 | Net Revenue | LINK |
| 225 | Calc_COGS | B12 | Total Raw Materials | IS | B13 | Raw Materials | LINK |
| 226 | Calc_COGS | B18 | Total Packaging | IS | B14 | Packaging | LINK |
| 227 | Calc_COGS | B25 | Total Direct Manufacturing | IS | B15 | Direct Manufacturing | LINK |
| 228 | Calc_COGS | B28 | Total COGS | IS | B16 | Total COGS | LINK |
| 229 | Calc_Opex | B11 | Total Logistics | IS | B23 | Logistics & Distribution | LINK |
| 230 | Calc_Opex | B14 | Sales Team | IS | B24 | Sales & Marketing | LINK |
| 231 | Calc_Opex | B15 | Marketing | IS | B24 | Sales & Marketing | LINK |
| 232 | Calc_Opex | B16 | Trade Marketing | IS | B24 | Sales & Marketing | LINK |
| 233 | Calc_Opex | B17 | General & Admin | IS | B25 | General & Admin | LINK |
| 234 | Calc_Opex | B21 | Total Operating Expenses | IS | B26 | Total Opex | LINK |
| 235 | Calc_Volume | B3 | Period | IS | B3 | Period | INPUT |
| 236 | Calc_Capex | B15 | Depreciation | IS | B32 | Depreciation | LINK |
| 237 | Calc_Debt | B21 | Total Interest Expense | IS | B36 | Interest Expense | LINK |
| 238 | Control | D16 | Corporate Tax Rate | IS | B39 | Income Tax | INPUT |
| 239 | Calc_Volume | B4 | Date | IS | B4 | Date | INPUT |
| 240 | Calc_Revenue | B12 | Total Gross Revenue | IS | B7 | Gross Revenue | LINK |
| 241 | Calc_Revenue | B17 | Total Discounts | IS | B8 | Less: Discounts | LINK |
| 242 | Calc_Revenue | B22 | Total Excise | IS | B9 | Less: Excise Duty | LINK |
| 243 | Calc_Revenue | B12 | Total Gross Revenue | KPIs | B12 | Gross Revenue ($000s) | LINK |
| 244 | Calc_Revenue | B25 | Net Revenue | KPIs | B13 | Net Revenue ($000s) | LINK |
| 245 | IS | B20 | Gross Margin % | KPIs | B18 | Gross Margin % | LINK |
| 246 | IS | B29 | EBITDA | KPIs | B19 | EBITDA ($000s) | LINK |
| 247 | IS | B30 | EBITDA Margin % | KPIs | B20 | EBITDA Margin % | LINK |
| 248 | IS | B34 | EBIT Margin % | KPIs | B21 | EBIT Margin % | LINK |
| 249 | IS | B40 | NET INCOME | KPIs | B22 | Net Income ($000s) | LINK |
| 250 | IS | B41 | Net Margin % | KPIs | B23 | Net Margin % | LINK |
| 251 | CF | B25 | Closing Cash | KPIs | B26 | Cash Balance ($000s) | LINK |
| 252 | Calc_Debt | B20 | Total Debt | KPIs | B27 | Total Debt ($000s) | LINK |
| 253 | Calc_Volume | B3 | Period | KPIs | B3 | Period | INPUT |
| 254 | Calc_Debt | B21 | Total Interest Expense | KPIs | B30 | Interest Coverage (EBITDA/Interest) | LINK |
| 255 | IS | B29 | EBITDA | KPIs | B30 | Interest Coverage (EBITDA/Interest) | LINK |
| 256 | BS | B36 | TOTAL EQUITY | KPIs | B33 | ROE % | LINK |
| 257 | BS | B18 | TOTAL ASSETS | KPIs | B34 | ROA % | LINK |
| 258 | BS | B36 | TOTAL EQUITY | KPIs | B35 | ROCE % | LINK |
| 259 | IS | B33 | EBIT | KPIs | B35 | ROCE % | LINK |
| 260 | Calc_WC | B23 | Net Working Capital | KPIs | B38 | Net Working Capital ($000s) | LINK |
| 261 | Calc_Volume | B4 | Date | KPIs | B4 | Date | INPUT |
| 262 | Assum_WC | C13 | Blended Average | KPIs | B40 | Cash Conversion Cycle (days) | LINK |
| 263 | Assum_WC | C17 | Trade Payables (days of COGS + Opex) | KPIs | B40 | Cash Conversion Cycle (days) | INPUT |
| 264 | Assum_WC | C7 | Total Inventory Days | KPIs | B40 | Cash Conversion Cycle (days) | LINK |
| 265 | Calc_Volume | B14 | Total | KPIs | B7 | Total Volume (HL) | LINK |
| 266 | Calc_Capex | B25 | Utilization % | KPIs | B9 | Capacity Utilization % | LINK |
| 267 | IS | B29 | EBITDA | KPIs | M29 | Net Debt / EBITDA (LTM) | LINK |
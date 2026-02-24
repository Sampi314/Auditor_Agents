# Flow Dependency Register

| # | Source Sheet | Source Cell | Source Row Label | Target Sheet | Target Cell | Target Row Label | Reference Type |
|---|---|---|---|---|---|---|---|
| 1 | Opening Balance Sheet | C13 | Current assets | Balance Sheet | C13 |  | INPUT |
| 2 | Opening Balance Sheet | C20 | Non-current assets | Balance Sheet | C20 |  | INPUT |
| 3 | Opening Balance Sheet | C25 | Total assets | Balance Sheet | C25 |  | INPUT |
| 4 | Opening Balance Sheet | C27 | Current liabilities | Balance Sheet | C27 |  | INPUT |
| 5 | Opening Balance Sheet | C35 | Non-current liabilities | Balance Sheet | C35 |  | INPUT |
| 6 | Opening Balance Sheet | C40 | Total liabilities | Balance Sheet | C40 |  | INPUT |
| 7 | Opening Balance Sheet | C42 | Net assets | Balance Sheet | C42 |  | INPUT |
| 8 | Opening Balance Sheet | C45 | Equity | Balance Sheet | C45 |  | INPUT |
| 9 | Opening Balance Sheet | D14 | Cash | Balance Sheet | D14 |  | INPUT |
| 10 | Opening Balance Sheet | D15 | Accounts receivable | Balance Sheet | D15 |  | INPUT |
| 11 | Opening Balance Sheet | D16 | Inventory | Balance Sheet | D16 |  | INPUT |
| 12 | Opening Balance Sheet | D17 | Other current assets | Balance Sheet | D17 |  | INPUT |
| 13 | Opening Balance Sheet | D18 | Total current assets | Balance Sheet | D18 |  | INPUT |
| 14 | Opening Balance Sheet | D21 | PP&E | Balance Sheet | D21 |  | INPUT |
| 15 | Opening Balance Sheet | D22 | Deferred tax assets | Balance Sheet | D22 |  | INPUT |
| 16 | Opening Balance Sheet | D23 | Total non-current assets | Balance Sheet | D23 |  | INPUT |
| 17 | Opening Balance Sheet | D28 | Accounts payable | Balance Sheet | D28 |  | INPUT |
| 18 | Opening Balance Sheet | D29 | Interest payable | Balance Sheet | D29 |  | INPUT |
| 19 | Opening Balance Sheet | D30 | Dividends payable | Balance Sheet | D30 |  | INPUT |
| 20 | Opening Balance Sheet | D31 | Tax payable | Balance Sheet | D31 |  | INPUT |
| 21 | Opening Balance Sheet | D32 | Other current liabilities | Balance Sheet | D32 |  | INPUT |
| 22 | Opening Balance Sheet | D33 | Total current liabilities | Balance Sheet | D33 |  | INPUT |
| 23 | Opening Balance Sheet | D36 | Debt | Balance Sheet | D36 |  | INPUT |
| 24 | Opening Balance Sheet | D37 | Deferred tax liabilities | Balance Sheet | D37 |  | INPUT |
| 25 | Opening Balance Sheet | D38 | Total non-current liabilities | Balance Sheet | D38 |  | INPUT |
| 26 | Opening Balance Sheet | D46 | Ordinary equity | Balance Sheet | D46 |  | INPUT |
| 27 | Opening Balance Sheet | D50 | Retained profits | Balance Sheet | D50 |  | INPUT |
| 28 | Opening Balance Sheet | D51 | Total equity | Balance Sheet | D51 |  | INPUT |
| 29 | Opening Balance Sheet | E47 | Opening retained profits | Balance Sheet | E47 |  | INPUT |
| 30 | Opening Balance Sheet | E48 | NPAT | Balance Sheet | E48 |  | INPUT |
| 31 | Opening Balance Sheet | E49 | Dividends declared | Balance Sheet | E49 |  | INPUT |
| 32 | Cash Flow Statement | J37 | Net increase / (decrease) in cash held | Balance Sheet | J14 |  | LINK |
| 33 | Calculations | J38 | Closing receivables | Balance Sheet | J15 |  | LINK |
| 34 | Calculations | J139 | Closing inventory | Balance Sheet | J16 |  | LINK |
| 35 | Calculations | J412 |  | Balance Sheet | J17 |  | LINK |
| 36 | Calculations | J200 | Closing net book value | Balance Sheet | J21 |  | LINK |
| 37 | Calculations | J353 | Closing DTAs | Balance Sheet | J22 |  | LINK |
| 38 | Calculations | J168 | Closing payables | Balance Sheet | J28 |  | LINK |
| 39 | Calculations | J236 | Closing interest payable | Balance Sheet | J29 |  | LINK |
| 40 | Calculations | J256 | Closing interest receivable | Balance Sheet | J29 |  | LINK |
| 41 | Calculations | J401 | Closing dividend payable | Balance Sheet | J30 |  | LINK |
| 42 | Calculations | J362 | Closing tax payable | Balance Sheet | J31 |  | LINK |
| 43 | Calculations | J217 | Closing debt | Balance Sheet | J36 |  | LINK |
| 44 | Calculations | J315 | Closing DTLs | Balance Sheet | J37 |  | LINK |
| 45 | Calculations | J377 | Closing equity | Balance Sheet | J46 |  | LINK |
| 46 | Income Statement | J28 | NPAT | Balance Sheet | J48 |  | LINK |
| 47 | Calculations | J399 | Dividend declared | Balance Sheet | J49 |  | LINK |
| 48 | Timing | J5 |  | Balance Sheet | J5 |  | LINK |
| 49 | Timing | J6 | Start Date | Balance Sheet | J6 | Start Date | LINK |
| 50 | Timing | J7 | End Date | Balance Sheet | J7 | End Date | LINK |
| 51 | Timing | J8 | Number of Days | Balance Sheet | J8 | Number of Days | LINK |
| 52 | Timing | J9 | Counter | Balance Sheet | J9 | Counter | LINK |
| 53 | General Assumptions | C13 | Revenue and related | Calculations | C13 |  | INPUT |
| 54 | General Assumptions | C52 | All expenses are assumed to be paid as they are incurred | Calculations | C143 |  | INPUT |
| 55 | General Assumptions | C61 | Capex and related | Calculations | C171 |  | INPUT |
| 56 | General Assumptions | C76 | Debt and related | Calculations | C203 |  | INPUT |
| 57 | General Assumptions | C77 | Movements are assumed to occur at the end of each period | Calculations | C204 |  | INPUT |
| 58 | General Assumptions | C78 | Interest is assumed to be paid in the following period | Calculations | C205 |  | INPUT |
| 59 | General Assumptions | D54 | Opex | Calculations | D145 |  | INPUT |
| 60 | General Assumptions | D15 | Sales | Calculations | D15 |  | INPUT |
| 61 | General Assumptions | D63 | Capital expenditure | Calculations | D173 |  | INPUT |
| 62 | General Assumptions | D67 | Accounting depreciation - straight line | Calculations | D178 |  | INPUT |
| 63 | General Assumptions | D80 | Debt | Calculations | D207 |  | INPUT |
| 64 | General Assumptions | D85 | Interest | Calculations | D219 |  | INPUT |
| 65 | General Assumptions | E56 | Growth rates used from | Calculations | E147 |  | INPUT |
| 66 | General Assumptions | E57 | Amounts | Calculations | E150 |  | INPUT |
| 67 | General Assumptions | E58 | Growth rates | Calculations | E151 |  | INPUT |
| 68 | General Assumptions | E17 | Projected sales | Calculations | E17 |  | INPUT |
| 69 | General Assumptions | E18 | Unit price | Calculations | E18 |  | INPUT |
| 70 | General Assumptions | E69 | Remaining life of existing assets | Calculations | E180 |  | INPUT |
| 71 | General Assumptions | E70 | Annual rate | Calculations | E181 |  | INPUT |
| 72 | General Assumptions | E72 | Economic life of new capex | Calculations | E183 |  | INPUT |
| 73 | General Assumptions | E19 | Inflation | Calculations | E19 |  | INPUT |
| 74 | General Assumptions | E82 | Debt drawdowns | Calculations | E209 |  | INPUT |
| 75 | General Assumptions | E83 | Debt repayments | Calculations | E210 |  | INPUT |
| 76 | General Assumptions | E87 | Interest rate | Calculations | E221 |  | INPUT |
| 77 | General Assumptions | E34 | Days payable | Calculations | E53 |  | INPUT |
| 78 | General Assumptions | E41 | Purchases | Calculations | E71 |  | INPUT |
| 79 | General Assumptions | E42 | Price | Calculations | E72 |  | INPUT |
| 80 | Opening Balance Sheet | I16 | Inventory | Calculations | I139 | Closing inventory | INPUT |
| 81 | Opening Balance Sheet | I21 | PP&E | Calculations | I175 | Existing assets | INPUT |
| 82 | Opening Balance Sheet | I21 | PP&E | Calculations | I200 | Closing net book value | INPUT |
| 83 | Opening Balance Sheet | I36 | Debt | Calculations | I217 | Closing debt | INPUT |
| 84 | Opening Balance Sheet | I29 | Interest payable | Calculations | I236 | Closing interest payable | INPUT |
| 85 | General Assumptions | I93 | Proportion into period of movement | Calculations | I243 | Proportion into period of movement | INPUT |
| 86 | General Assumptions | I109 | Declining balance multiplier | Calculations | I283 | Declining balance multiplier | INPUT |
| 87 | General Assumptions | I111 | Remaining life of tax assets | Calculations | I285 | Remaining life of tax assets | INPUT |
| 88 | General Assumptions | I112 | Annual rate | Calculations | I286 | Annual rate | LINK |
| 89 | General Assumptions | I114 | Tax asset life of new capex | Calculations | I288 | Tax asset life of new Capex | INPUT |
| 90 | Opening Balance Sheet | I15 | Accounts receivable | Calculations | I38 | Closing receivables | INPUT |
| 91 | Opening Balance Sheet | I28 | Accounts payable | Calculations | I64 | Closing payables | INPUT |
| 92 | General Assumptions | J43 | Amount used per sale | Calculations | J100 | Amount used per sale | INPUT |
| 93 | General Assumptions | J44 | Wastage | Calculations | J101 | Wastage | INPUT |
| 94 | General Assumptions | J91 | Interest receivable rate | Calculations | J241 | Interest receivable rate | INPUT |
| 95 | Balance Sheet | I14 |  | Calculations | J245 | Opening cash balance | INPUT |
| 96 | Opening Balance Sheet | I14 | Cash | Calculations | J245 | Opening cash balance | INPUT |
| 97 | Cash Flow Statement | J22 | Net Operating cash flow | Calculations | J246 | Non-interest cash movement | LINK |
| 98 | Cash Flow Statement | J26 | Purchases of Non-Current Assets | Calculations | J246 | Non-interest cash movement | LINK |
| 99 | Cash Flow Statement | J35 | Net Financing cash flows | Calculations | J246 | Non-interest cash movement | LINK |
| 100 | General Assumptions | I100 | Tax rate | Calculations | J247 | Tax rate | INPUT |
| 101 | General Assumptions | J104 | Non-assessable revenue | Calculations | J263 | Non-assessable revenue | INPUT |
| 102 | General Assumptions | J105 | Disallowable expenses | Calculations | J264 | Disallowable expenses | INPUT |
| 103 | Income Statement | J25 | NPBT | Calculations | J268 | NPBT | LINK |
| 104 | General Assumptions | J23 | Days receivable | Calculations | J27 | Days receivable | INPUT |
| 105 | General Assumptions | J30 | Gross margin (non-inventory) | Calculations | J46 | Gross margin (non-inventory) | INPUT |
| 106 | Timing | J5 |  | Calculations | J5 |  | LINK |
| 107 | Timing | J6 | Start Date | Calculations | J6 | Start Date | LINK |
| 108 | Timing | J7 | End Date | Calculations | J7 | End Date | LINK |
| 109 | Timing | J8 | Number of Days | Calculations | J8 | Number of Days | LINK |
| 110 | General Assumptions | J48 | Days payable | Calculations | J80 | Days payable | INPUT |
| 111 | Timing | J9 | Counter | Calculations | J9 | Counter | LINK |
| 112 | Calculations | J37 | Cash receipts | Cash Flow Statement | J16 | Cash receipts | LINK |
| 113 | Calculations | J63 | Cash payments | Cash Flow Statement | J17 |  | LINK |
| 114 | Calculations | J235 | Interest paid | Cash Flow Statement | J20 | Interest paid | LINK |
| 115 | Calculations | J359 | Tax paid | Cash Flow Statement | J21 | Tax paid | LINK |
| 116 | Calculations | J255 | Interest Received | Cash Flow Statement | J25 | Interest received | LINK |
| 117 | Calculations | J198 | Capital expenditure | Cash Flow Statement | J26 | Purchases of Non-Current Assets | LINK |
| 118 | Calculations | J215 | Debt drawdowns | Cash Flow Statement | J30 | Debt drawdowns | LINK |
| 119 | Calculations | J216 | Debt repayments | Cash Flow Statement | J31 | Debt repayments | LINK |
| 120 | Calculations | J375 | Equity issuances | Cash Flow Statement | J32 | Ordinary equity issuances | LINK |
| 121 | Calculations | J376 | Equity buybacks | Cash Flow Statement | J33 | Ordinary equity buybacks | LINK |
| 122 | Calculations | J400 | Dividends paid | Cash Flow Statement | J34 | Dividends paid | LINK |
| 123 | Timing | J5 |  | Cash Flow Statement | J5 |  | LINK |
| 124 | Timing | J6 | Start Date | Cash Flow Statement | J6 | Start Date | LINK |
| 125 | Timing | J7 | End Date | Cash Flow Statement | J7 | End Date | LINK |
| 126 | Timing | J8 | Number of Days | Cash Flow Statement | J8 | Number of Days | LINK |
| 127 | Timing | J9 | Counter | Cash Flow Statement | J9 | Counter | LINK |
| 128 | Opening Balance Sheet | I22 | Deferred tax assets | General Assumptions | I120 | DTA | INPUT |
| 129 | Opening Balance Sheet | I37 | Deferred tax liabilities | General Assumptions | I125 | DTL | INPUT |
| 130 | Timing | J5 |  | General Assumptions | J5 |  | LINK |
| 131 | Timing | J6 | Start Date | General Assumptions | J6 | Start Date | LINK |
| 132 | Timing | J7 | End Date | General Assumptions | J7 | End Date | LINK |
| 133 | Timing | J8 | Number of Days | General Assumptions | J8 | Number of Days | LINK |
| 134 | Timing | J9 | Counter | General Assumptions | J9 | Counter | LINK |
| 135 | Calculations | J36 | Revenue | Income Statement | J13 | Revenue | LINK |
| 136 | Calculations | J137 | COGS | Income Statement | J14 | COGS | LINK |
| 137 | Calculations | J138 | Wastage | Income Statement | J17 | Wastage | LINK |
| 138 | Calculations | J166 |  | Income Statement | J18 | Operating expenditure | LINK |
| 139 | Calculations | J199 | Depreciation | Income Statement | J21 | Depreciation | LINK |
| 140 | Calculations | J234 | Interest expense | Income Statement | J24 | Interest Expense | LINK |
| 141 | Calculations | J254 | Interest Income | Income Statement | J24 | Interest Expense | LINK |
| 142 | Calculations | J358 | Tax expense / (credit) | Income Statement | J27 | Tax Expense | LINK |
| 143 | Timing | J5 |  | Income Statement | J5 |  | LINK |
| 144 | Timing | J6 | Start Date | Income Statement | J6 | Start Date | LINK |
| 145 | Timing | J7 | End Date | Income Statement | J7 | End Date | LINK |
| 146 | Timing | J8 | Number of Days | Income Statement | J8 | Number of Days | LINK |
| 147 | Timing | J9 | Counter | Income Statement | J9 | Counter | LINK |
| 148 | Timing | J7 | End Date | Lookup | D11 |  | LINK |
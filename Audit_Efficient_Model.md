# Audit Report: 20130401 Efficient Modelling (Tutorial).xlsx

## Style Convention Summary

*Detection method: Style Guide Sheet (L)*
*Confidence: High*

| Cell Type | Font Colour | Fill Colour | Font Style | Source | Confidence |
|---|---|---|---|---|---|
| Input | Dark Red (#800000) | No Fill | Normal | Style Guide | ✅ Definitive |
| Formula | Black (None/Theme 0) | No Fill | Normal | Inferred | ✅ High |
| Link | Indexed 16 | No Fill | Normal | Style Guide | ✅ Definitive |
| Header 1 | Red (#EF4136) | No Fill | Bold | Style Guide | ✅ Definitive |
| Table Heading | Theme 0 | No Fill | Normal | Style Guide | ✅ Definitive |

## Audit Findings

| Sheet Name | Cell Reference | Description of the Location | Short Error Category | Long Description of Error |
|---|---|---|---|---|
| Dynamic Array | J55 | Total sum calculation | Calculation Error | 🔴 HIGH: Cell contains error #VALUE! |
| Multiple | Various | Sheet naming vs Labeling | Inconsistent Naming | 🟡 LOW: Sheet is named 'Inputs' but main heading in A1 is 'Assumptions'. Consistent terminology is preferred. |
| L | C18 | Offsheet Reference | Typo | 🟡 LOW: Potential typo 'Offsheet' instead of 'Off-sheet' or 'Off sheet'. (Note: This may be intentional Corality jargon). |
| One Line | E11, J11:M11, ... (105 cells) | Off-sheet links | Colour Coding Error | 🟡 LOW: Off-sheet links are missing the model's Link style (Indexed 16). |
| Dynamic Array | K7:O7, ... (103 cells) | Hard-coded inputs | Colour Coding Error | 🟡 LOW: Hard-coded inputs are formatted with font color Indexed(23) instead of the model's Input style (#800000). |
| One Line | I5 | Period end date | Formula Pattern Break | ⚠️ MEDIUM: Formula differs from dominant pattern in row. Expected R1C1: =EOMONTH(R[-1]C,Months_Qtr-1) vs Actual R1C1: =Model_Start |
| Dynamic Array | E102 | Tax calculation | Hard-Code in Formula | ⚠️ MEDIUM: Formula contains hard-coded literal(s) ['0.3', '0.3'] (likely 30% tax rate) which should be a cell reference. |

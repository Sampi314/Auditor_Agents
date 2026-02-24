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
| Dynamic Array | J55 | Calculation Error | Calculation Error | 🔴 HIGH: Cell contains error #VALUE! |
| Dynamic Array | E102 | Hard-coded literal | Hard-Code in Formula | ⚠️ MEDIUM: Formula contains hard-coded literal(s) ['0.3', '0.3'] which should likely be cell references. |
| Dynamic Array | K7:BI7, K8:BI8, I74 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(23) instead of the model's Input style (#800000) |
| Dynamic Array | E29, E30, E31, E32, E33, E34 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(16) instead of the model's Input style (#800000) |
| N | F11, F12, F13, F14, F16, F17, F18, F19, F21 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(23) instead of the model's Input style (#800000) |
| One Line | J11:BI11, J14:BI14 | Mega-Formula | Mega-Formula | 🔴 HIGH: Extremely long formulas detected (>4000 chars). Mega-formulas are impossible to audit or maintain. |
| One Line | E11, J11:BI11, J14:BI14 | Multiple locations | Colour Coding Error | 🟡 LOW: Off-sheet link is formatted with font color None instead of the model's Link style (Indexed 16) |
| One Line | I11 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(23) instead of the model's Input style (#800000) |

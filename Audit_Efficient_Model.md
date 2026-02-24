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
| Dynamic Array | K7:BI8, I74 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(23) instead of the model's Input style (#800000) |
| Dynamic Array | E29:E34 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(16) instead of the model's Input style (#800000) |
| N | F11:F14, F16:F19, F21 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(23) instead of the model's Input style (#800000) |
| One Line | J11:Y11 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (4826 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | Z11 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (4832 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | AA11 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (4917 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | AB11 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (4955 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | AC11:BI11 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (4957 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | J14:Y14 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (5193 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | Z14 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (5202 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | AA14 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (5295 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | AB14 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (5333 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | AC14:BI14 | Mega-Formula (Bad Practice) | Mega-Formula | 🔴 HIGH: Extremely long formula (5335 chars) detected. Flagged as CRITICAL BAD PRACTICE due to total lack of auditability. Formulas should be broken into modular steps. |
| One Line | J11:M11, BC11:BI11, J14:M14, AL14:BI14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 0), but its complexity makes it an extreme auditability risk. |
| One Line | N11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.2273457270300705), but its complexity makes it an extreme auditability risk. |
| One Line | O11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 4.165945297812476), but its complexity makes it an extreme auditability risk. |
| One Line | P11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 4.1682945789555035), but its complexity makes it an extreme auditability risk. |
| One Line | Q11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 4.122806486940265), but its complexity makes it an extreme auditability risk. |
| One Line | R11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 6.823785472401559), but its complexity makes it an extreme auditability risk. |
| One Line | S11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 6.596364111891576), but its complexity makes it an extreme auditability risk. |
| One Line | T11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 6.551745577859867), but its complexity makes it an extreme auditability risk. |
| One Line | U11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 6.48262876099958), but its complexity makes it an extreme auditability risk. |
| One Line | V11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 6.1335828698798736), but its complexity makes it an extreme auditability risk. |
| One Line | W11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.006172250940809), but its complexity makes it an extreme auditability risk. |
| One Line | X11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 6.97733700939281), but its complexity makes it an extreme auditability risk. |
| One Line | Y11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 6.917497114037831), but its complexity makes it an extreme auditability risk. |
| One Line | Z11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 5.308518435963977), but its complexity makes it an extreme auditability risk. |
| One Line | AA11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.9315639574796606), but its complexity makes it an extreme auditability risk. |
| One Line | AB11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.917762710209825), but its complexity makes it an extreme auditability risk. |
| One Line | AC11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.873474087859572), but its complexity makes it an extreme auditability risk. |
| One Line | AD11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.816339665562518), but its complexity makes it an extreme auditability risk. |
| One Line | AE11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.8289395488768623), but its complexity makes it an extreme auditability risk. |
| One Line | AF11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.797437581365209), but its complexity makes it an extreme auditability risk. |
| One Line | AG11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.7537102615895783), but its complexity makes it an extreme auditability risk. |
| One Line | AH11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.6993502615798963), but its complexity makes it an extreme auditability risk. |
| One Line | AI11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.7038773802009555), but its complexity makes it an extreme auditability risk. |
| One Line | AJ11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.670829880661218), but its complexity makes it an extreme auditability risk. |
| One Line | AK11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 3.6276751739880897), but its complexity makes it an extreme auditability risk. |
| One Line | AL11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 8.549062168930485), but its complexity makes it an extreme auditability risk. |
| One Line | AM11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 8.497489529691588), but its complexity makes it an extreme auditability risk. |
| One Line | AN11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 8.415001031963008), but its complexity makes it an extreme auditability risk. |
| One Line | AO11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 8.324566153916388), but its complexity makes it an extreme auditability risk. |
| One Line | AP11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 8.235399022827274), but its complexity makes it an extreme auditability risk. |
| One Line | AQ11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 8.16074599503908), but its complexity makes it an extreme auditability risk. |
| One Line | AR11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 8.083988553976777), but its complexity makes it an extreme auditability risk. |
| One Line | AS11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.994149598419372), but its complexity makes it an extreme auditability risk. |
| One Line | AT11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.900588507178982), but its complexity makes it an extreme auditability risk. |
| One Line | AU11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.832043809704565), but its complexity makes it an extreme auditability risk. |
| One Line | AV11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.746306008845999), but its complexity makes it an extreme auditability risk. |
| One Line | AW11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.657074983914301), but its complexity makes it an extreme auditability risk. |
| One Line | AX11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.566518789414821), but its complexity makes it an extreme auditability risk. |
| One Line | AY11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.4892307133900635), but its complexity makes it an extreme auditability risk. |
| One Line | AZ11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.401818991378513), but its complexity makes it an extreme auditability risk. |
| One Line | BA11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 7.313208147178971), but its complexity makes it an extreme auditability risk. |
| One Line | BB11 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 0.07433702984285695), but its complexity makes it an extreme auditability risk. |
| One Line | N14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.1864093740940775), but its complexity makes it an extreme auditability risk. |
| One Line | O14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.639242080201367), but its complexity makes it an extreme auditability risk. |
| One Line | P14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.646259954701511), but its complexity makes it an extreme auditability risk. |
| One Line | Q14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.6432993177094737), but its complexity makes it an extreme auditability risk. |
| One Line | R14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.0762272856248303), but its complexity makes it an extreme auditability risk. |
| One Line | S14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.0517009333121163), but its complexity makes it an extreme auditability risk. |
| One Line | T14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.056100597465643), but its complexity makes it an extreme auditability risk. |
| One Line | U14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.056605347618682), but its complexity makes it an extreme auditability risk. |
| One Line | V14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.0109815385051335), but its complexity makes it an extreme auditability risk. |
| One Line | W14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.16797165317372), but its complexity makes it an extreme auditability risk. |
| One Line | X14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.1765763941505476), but its complexity makes it an extreme auditability risk. |
| One Line | Y14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 2.1800926331656525), but its complexity makes it an extreme auditability risk. |
| One Line | Z14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.9162969120154445), but its complexity makes it an extreme auditability risk. |
| One Line | AA14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.6867276282870387), but its complexity makes it an extreme auditability risk. |
| One Line | AB14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.6925889229298687), but its complexity makes it an extreme auditability risk. |
| One Line | AC14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.6931380954111728), but its complexity makes it an extreme auditability risk. |
| One Line | AD14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.6913737094962613), but its complexity makes it an extreme auditability risk. |
| One Line | AE14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.7023566872673173), but its complexity makes it an extreme auditability risk. |
| One Line | AF14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.7054261541994893), but its complexity makes it an extreme auditability risk. |
| One Line | AG14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.7062743566885923), but its complexity makes it an extreme auditability risk. |
| One Line | AH14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.705118017046661), but its complexity makes it an extreme auditability risk. |
| One Line | AI14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.7153035897210163), but its complexity makes it an extreme auditability risk. |
| One Line | AJ14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.718408142530917), but its complexity makes it an extreme auditability risk. |
| One Line | AK14 | Deep Logic Check | Auditability Risk | ⚠️ MEDIUM: Mega-formula is technically functioning (results in 1.7195920291692137), but its complexity makes it an extreme auditability risk. |
| One Line | I5 | Formula row | Formula Pattern Break | ⚠️ MEDIUM: Formula differs from dominant pattern in row. Expected R1C1: =EOMONTH(R[-1]C,Months_Qtr-1) vs Actual R1C1: =Model_Start |
| One Line | J11:BI11, J14:BI14 | Modularity Failure | Modularity Failure | ⚠️ MEDIUM: Calculation is condensed into a 'black box' Mega-Formula. Structural integrity is compromised as modifications risk breaking complex nested logic. |
| One Line | E11, J11:BI11, J14:BI14 | Multiple locations | Colour Coding Error | 🟡 LOW: Off-sheet link is formatted with font color None instead of the model's Link style (Indexed 16) |
| One Line | I11 | Multiple locations | Colour Coding Error | 🟡 LOW: Hard-coded input is formatted with font color Indexed(23) instead of the model's Input style (#800000) |

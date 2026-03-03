# Audit Report 📊

This report is divided into sections based on the specialized agent persona that identified the findings.

## Logic Report

| Sheet Name | Cell Reference | Description of the Location | Priority | Short Error Category | Long Description of Error |
|---|---|---|---|---|---|
| Cons | I16, I19, I23:I25, I28:I29, I31:I32, I36:I37, I48 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Debt | I20:I21, I30, I34:I35, I39, I41:I43 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Ops | I15:AG15, I19, I23:I26, J26, I37:I40, J40, I43:I45, I48, I50, I53:I55, I57 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Ops | AH15:BI15, I57 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| QCF | I14, I17:I19, I21, I24:I26, I29:I31, I33, I36:I38, I40, I42, I44 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Timing | I5 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |

## Stylist Report

| Sheet Name | Cell Reference | Description of the Location | Priority | Short Error Category | Long Description of Error |
|---|---|---|---|---|---|
| Cons | A4:A7, J4:BI5, C10:C12, J10:BI12, C16, J16:BI16, J43:BI43 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color None instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Cons | J6:BI7 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Indexed(23) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Cons | C9 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color #E58832 instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Debt | A4:A7, J4:BI5, C10:C12, J10:BI12, J20:BI20, J25:BI26, J39:BI39 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color None instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Debt | J6:BI7 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Indexed(23) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Debt | C9 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color #E58832 instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Inputs | E10, E15, E20, F23:M23, E28, F32:I32, E37:N38, E43:N43, E49, E52:I52, E55, F63:I63 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Theme(3) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| L&N | E14 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Theme(3) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| L&N | E16 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Indexed(16) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| L&N | E20, E30:E33, E35 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Indexed(23) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Ops | A4:A7, J4:BI5, C10:C11, J10:BI11, J15:AG15, J18:BI18, C23:C25, J23:BI25, C30:C32, J30:BI32, J48:BI49, J53:BI53 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color None instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Ops | J6:BI7 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Indexed(23) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Ops | C9 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color #E58832 instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| QCF | A4:A7, J4:BI5, J14:BI14, J17:BI18, C24:C25, J24:BI25, J29:BI30, J36:BI37, J42:BI42 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color None instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| QCF | J6:BI7 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Indexed(23) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |

# Audit Report 📊

This report is divided into sections based on the specialized agent persona that identified the findings.

## Lingo Report

| Sheet Name | Cell Reference | Description of the Location | Priority | Short Error Category | Long Description of Error |
|---|---|---|---|---|---|
| Annual_Summary | A10 | Label: EBIT ($000s) | 🟡 LOW | Typo | Potential typo 'ebit' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Capex | C5 | Label: Accum Depr | 🟡 LOW | Typo | Potential typo 'depr' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Capex | C5 | Label: Accum Depr | 🟡 LOW | Typo | Potential typo 'accum' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Costs | A18 | Label: Direct Labour | 🟡 LOW | Typo | Potential typo 'labour' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Finance | A14 | Label: Commitment Fee (on undrawn) | 🟡 LOW | Typo | Potential typo 'undrawn' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Pricing | A6:A8, A27:A28 | Label: SkyBrew | 🟡 LOW | Typo | Potential typo 'skybrew' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Pricing | A9:A11, A29:A30 | Label: AllDark | 🟡 LOW | Typo | Potential typo 'alldark' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Pricing | A22 | Label: Australia - Rate per Litre of Alcohol (LAL) | 🟡 LOW | Typo | Potential typo 'litre' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Volume | A8 | Label: SkyBrew % | 🟡 LOW | Typo | Potential typo 'skybrew' detected in label. Review and correct the spelling for professional presentation. |
| Assum_Volume | A9 | Label: AllDark % | 🟡 LOW | Typo | Potential typo 'alldark' detected in label. Review and correct the spelling for professional presentation. |
| Assum_WC | A17, A29 | Label: Trade Payables (days of COGS + Opex) | 🟡 LOW | Typo | Potential typo 'payables' detected in label. Review and correct the spelling for professional presentation. |
| BS | A22 | Label: Trade Payables | 🟡 LOW | Typo | Potential typo 'payables' detected in label. Review and correct the spelling for professional presentation. |
| Calc_COGS | A21 | Label: Direct Labour | 🟡 LOW | Typo | Potential typo 'labour' detected in label. Review and correct the spelling for professional presentation. |
| Calc_Opex | A17 | Label: General & Admin | 🟡 LOW | Typo | Potential typo 'admin' detected in label. Review and correct the spelling for professional presentation. |
| Calc_Revenue | A8:A9 | Label: SkyBrew On-Trade | 🟡 LOW | Typo | Potential typo 'skybrew' detected in label. Review and correct the spelling for professional presentation. |
| Calc_Revenue | A10:A11 | Label: AllDark On-Trade | 🟡 LOW | Typo | Potential typo 'alldark' detected in label. Review and correct the spelling for professional presentation. |
| Calc_Volume | A12 | Label: SkyBrew | 🟡 LOW | Typo | Potential typo 'skybrew' detected in label. Review and correct the spelling for professional presentation. |
| Calc_Volume | A13 | Label: AllDark | 🟡 LOW | Typo | Potential typo 'alldark' detected in label. Review and correct the spelling for professional presentation. |
| Calc_WC | A15:A16 | Label: PAYABLES ($000s) | 🟡 LOW | Typo | Potential typo 'payables' detected in label. Review and correct the spelling for professional presentation. |
| Cover | B5 | Label: SkyBrew & AllDark Brewing Company | 🟡 LOW | Typo | Potential typo 'alldark' detected in label. Review and correct the spelling for professional presentation. |
| Cover | B5 | Label: SkyBrew & AllDark Brewing Company | 🟡 LOW | Typo | Potential typo 'skybrew' detected in label. Review and correct the spelling for professional presentation. |
| Cover | B15 | Label: 3. SKU_Matrix - Product portfolio (24 SKUs) | 🟡 LOW | Typo | Potential typo 'skus' detected in label. Review and correct the spelling for professional presentation. |
| Cover | B19 | Label: 7. KPIs - Key performance indicators | 🟡 LOW | Typo | Potential typo 'kpis' detected in label. Review and correct the spelling for professional presentation. |
| IS | A25 | Label: General & Admin | 🟡 LOW | Typo | Potential typo 'admin' detected in label. Review and correct the spelling for professional presentation. |
| KPIs | A21 | Label: EBIT Margin % | 🟡 LOW | Typo | Potential typo 'ebit' detected in label. Review and correct the spelling for professional presentation. |
| SKU_Matrix | B4:B15 | Label: SkyBrew | 🟡 LOW | Typo | Potential typo 'skybrew' detected in label. Review and correct the spelling for professional presentation. |
| SKU_Matrix | B16:B27 | Label: AllDark | 🟡 LOW | Typo | Potential typo 'alldark' detected in label. Review and correct the spelling for professional presentation. |

## Logic Report

| Sheet Name | Cell Reference | Description of the Location | Priority | Short Error Category | Long Description of Error |
|---|---|---|---|---|---|
| Annual_Summary | I4:K16 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Assum_Pricing | C27:C32 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| BS | B24:DQ24, B34:DQ34, B35 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| CF | B24 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| Calc_COGS | N5, Z5, AL5, AX5, BJ5, BV5, CH5, CT5, DF5 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Calc_COGS | B8:DQ11, B15:DQ17, B21:DQ24, B31:DQ31 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| Calc_Capex | B7:B8, B15:DQ15 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| Calc_Debt | B7:B8, C8:DQ8, B13, B17:DQ17 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| Calc_Opex | B7:DQ10 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| Calc_Revenue | N5, Z5, AL5, AX5, BJ5, BV5, CH5, CT5, DF5 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Calc_Revenue | B8:DQ11, B20:DQ21, B28:DQ30 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| Calc_Volume | I7:M7, O7:Y7, AA7:AK7, AM7:AW7, AY7:BI7, BK7:BU7, BW7:CG7, CI7:CS7, CU7:DE7, DG7:DQ7, N9, Z9, AL9, AX9, BJ9, BV9, CH9, CT9, DF9 | Formula row | ⚠️ MEDIUM | Formula Pattern Break | Formula differs from the dominant pattern in this row. Inconsistent formulas within a row can lead to calculation errors and make the model harder to audit. Standardize the formula across the row. |
| Calc_Volume | DF9 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| Checks | C5:C8, B16 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| KPIs | B14:DQ14 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |
| SKU_Matrix | G4:G27 | Hard-coded literal | ⚠️ MEDIUM | Hard-Code in Formula | Formula contains one or more hard-coded literals. Literals should be moved to an assumptions sheet and referenced by cell to ensure model flexibility and transparency. |

## Stylist Report

| Sheet Name | Cell Reference | Description of the Location | Priority | Short Error Category | Long Description of Error |
|---|---|---|---|---|---|
| Annual_Summary | B4:K16 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Assum_Capex | F6 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Theme(1) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Assum_Capex | E7:E12 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color #0000FF instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Assum_Pricing | D6:D11 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color #0000FF instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Assum_Volume | D32:D41 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color #0000FF instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Assum_WC | C13 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| BS | B3:DQ4, B8:DQ11, B15:DQ15, B22:DQ24, B28:DQ28, B34:DQ35 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| CF | B3:DQ4, B7:DQ9, B13:DQ13, B17:DQ19, B24 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_COGS | B3:DQ4, N5, Z5, AL5, AX5, BJ5, BV5, CH5, CT5, DF5, B8:DQ11, B15:DQ17, B21:DQ24, B31:DQ32 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_Capex | B3:DQ4, B7:B8, B12:DQ12, B23:DQ24 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_Capex | D13:DQ13 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color #0000FF instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Calc_Capex | B15:DQ15 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color None instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_Debt | B3:DQ4, B7:B8, C8:DQ8, B10:DQ10, B13, B16:DQ17 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_Debt | D14:DQ14 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color #0000FF instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Calc_Opex | B3:DQ4, B7:DQ10, B14:DQ17 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_Revenue | B3:DQ4, N5, Z5, AL5, AX5, BJ5, BV5, CH5, CT5, DF5, B8:DQ11, B15:DQ16, B20:DQ21, B28:DQ30 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_Volume | D6:DQ6 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Theme(1) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Calc_Volume | B7:DQ7 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color #008000 instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_Volume | N9, Z9, AL9, AX9, BJ9, BV9, CH9, CT9, DF9, B12:DQ13, B17:DQ18, B21:DQ22, B25:DQ27 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_WC | B3:DQ4, B7:DQ9, B19:DQ20 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Calc_WC | B13:DQ13, B16:DQ16, B24 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color None instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Checks | B4:B16 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| Control | D9:D10, D15:D16, D19:D20 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color #0000FF instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| IS | B3:DQ4, B7:DQ9, B13:DQ15, B23:DQ25, B32:DQ32, B36:DQ36, B39:DQ39 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| IS | B10:DQ10, B16:DQ16, B26:DQ26 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color None instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| KPIs | B3:DQ4, B7:DQ7, B9:DQ9, B12:DQ13, B18:DQ23, B26:DQ27, M29:DQ30, B30:L30, B33:DQ35, B38:DQ38, B40:DQ40 | Multiple locations | 🟡 LOW | Colour Coding Error | Off-sheet link is formatted with font color Theme(1) instead of the model's standard Link style (Indexed 16). Apply the Link style to help users identify cross-sheet dependencies. |
| KPIs | D8:M8, D15:M15, D29:L29 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Theme(1) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| SKU_Matrix | E6:F27, H6:K27 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color #0000FF instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |
| Timeline | D6:D123 | Multiple locations | 🟡 LOW | Colour Coding Error | Hard-coded input is formatted with font color Theme(1) instead of the model's standard Input style (#800000). Apply the standard Input style to maintain consistency. |

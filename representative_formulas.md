# Representative Formulas (Master Model v3)

This document provides the actual Excel formula logic used in the core calculation engine.

## 1. Australian Excise Duty (Sheet 05, Row 8)
Calculates the volumetric tax liability per SKU, accounting for the 1.15% ABV exemption.
```excel
=MAX(0, '04_Calcs_Volume'!B5 * 100 * ('02_Inputs_Sales'!$B$9 - 0.0115) * ('02_Inputs_Sales'!$E$9 * '01_Control'!$B$13 + (1-'02_Inputs_Sales'!$E$9) * '01_Control'!$B$12))
```
*Logic: HL * 100 * (ABV - 0.0115) * Blended_Rate*

## 2. Direct Labor (Sheet 07, Row 5)
Calculates fully-burdened monthly labor cost by department.
```excel
=('03_Inputs_Ops'!$B$24 * '03_Inputs_Ops'!$C$24 * 1000 / 12) * (1 + '01_Control'!$B$9 + '01_Control'!$B$8)
```
*Logic: Headcount * Salary * 1000 / 12 * (1 + Super + PayrollTax)*

## 3. Materials Usage (Sheet 06, Row 6)
Dynamic BOM-driven cost calculation for Malt.
```excel
=SUMPRODUCT('04_Calcs_Volume'!B5:B13, '03_Inputs_Ops'!$B$5:$B$13) * '03_Inputs_Ops'!$B$16
```
*Logic: SumProduct(SKU_Vols, SKU_Malt_BOM) * Malt_Price*

## 4. Closing Cash (Sheet 10, Row 22)
Integrates the Cash Flow Statement into the Balance Sheet.
```excel
=B21 + B20
```
*Logic: Opening Cash + Net Cash Flow (Monthly)*

## 5. Balance Sheet Check (Sheet 12, Row 4)
Auditable integrity check for the 3-way model.
```excel
=IF(ABS(Total_Assets - Total_Equity) < 0.01, "PASS", "FAIL")
```

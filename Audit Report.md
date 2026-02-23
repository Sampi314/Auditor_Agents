# Audit Report: Brewery_Financial_Model_10Y 1.xlsx

**Orchestrated by:** Manager 👔
**Domains Audited:** Lingo ✍️, Logic 🧠, Sentry 🛡️, Stylist 🎨, Senior 🎓

---

### 🎨 Style Convention Summary
*Detection method: Statistical Inference*
*Confidence: High*

| Cell Type | Font Colour | Fill Colour | Font Style | Source | Confidence |
|---|---|---|---|---|---|
| **Input** | Blue (#0000FF) | Yellow (#FFFF00) | Normal | Inferred | ✅ High |
| **Formula** | Black (#000000) | No Fill | Normal / Bold | Inferred | ✅ High |
| **Link** | Black (#000000) | No Fill | Normal | Inferred | ✅ High |

---

### 📋 Consolidated Audit Findings

| Sheet Name | Cell Reference | Description of the Location | Short Error Category | Long Description of Error |
|:---|:---|:---|:---|:---|
| **Assum_Pricing** | C27:C30 | AU SkyBrew / AllDark Excise per HL | **Excise Error** | 🔴 **HIGH:** Australian Excise calculation misses the 1.15% ABV non-taxable threshold for beer. Current formula: `Rate * ABV * 100`. Expected: `Rate * (ABV - 1.15%) * 100`. |
| **Global** | Workbook | Working Capital / Cash Flow | **Logical Flaw** | 🔴 **HIGH:** GST (10%) is completely missing from the model. Australian brewery models must account for GST collected on sales and GST paid on inputs (Net GST Payable/Receivable) in Working Capital. |
| **IS** | B39:DU39 | Income Tax — P&L | **Logical Flaw** | 🟡 **LOW:** Tax calculation `=-MAX(0,EBT)*Rate` does not account for Tax Loss Carry Forwards. If the business incurs losses in early periods, these should offset future taxable income. |
| **Calc_Revenue** | B12:DU12 (multiple) | Gross Revenue — Time Series | **Formula Pattern Break** | ⚠️ **MEDIUM:** Inconsistent formula pattern found every 12th column (e.g., N, Z, AL). While this represents annual escalation, the first month of each year uses a different calculation logic than the subsequent 11 months. |
| **Checks** | B10:B20 | Various Integrity Checks | **Sentry Check** | ✅ **PASS:** All model integrity checks (Balance Sheet balance, Cash flow tie-in, etc.) currently resolve to TRUE or zero variance. |
| **All Sheets** | N/A | Entire Workbook | **Broken Reference** | ✅ **PASS:** No `#REF!`, `#DIV/0!`, or `#VALUE!` errors detected in the active calculation ranges. |
| **Global** | Multiple | Assumptions & Control Sheets | **Data Validation Missing** | 🔴 **HIGH:** No Data Validation (dropdowns) applied to scenario selections or key assumptions, increasing the risk of user entry errors. |
| **Calc_Volume** | B3:DQ3 | Date headers | **Hardcoded Date/Header** | ⚠️ **MEDIUM:** Period numbers are hardcoded rather than being linked to the Timeline sheet date spine. |
| **Assum_Finance** | C14, C24, C29 | Commitment Fee, Min Cash, Min Op Cash | **Unused Assumption** | 🟡 **LOW:** These assumptions appear to be unused in the current calculation logic (no downstream dependents). |
| **Annual_Summary**| B14:K14 | Closing Cash row | **Formula Inefficiency** | 🟡 **LOW:** Formulas in the closing cash row use direct cell references instead of consistent SUM or rollup logic used in other rows. |

---

### 👔 Manager's Summary & Recommendations

1.  **Immediate Fix Required (Excise):** The Excise Duty is being over-calculated by approx. 23% (for a 5% ABV beer) because the 1.15% ABV threshold is ignored. This significantly impacts Net Revenue and EBITDA.
2.  **Missing Component (GST):** The model is currently "GST exclusive" in its cash flows. In a real Australian business context, the timing of GST payments/refunds can have a material impact on monthly liquidity.
3.  **Consistency:** Technically, the model is very robust with no broken links or circular references. The formatting is consistent with financial modeling best practices.

---

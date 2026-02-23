# Audit Report: 20130401 Efficient Modelling (Tutorial).xlsx

**Orchestrated by:** Manager 👔
**Domains Audited:** Lingo ✍️, Logic 🧠, Sentry 🛡️, Stylist 🎨

---

### 🎨 Style Convention Summary
*Detection method: Statistical Inference*
*Confidence: High*

| Cell Type | Font Colour | Fill Colour | Font Style | Source | Confidence |
|---|---|---|---|---|---|
| **Input** | Maroon (#800000) | Pale Yellow (#FFFF99) | Normal | Inferred | ✅ High |
| **Formula** | Black (Default) | No Fill | Normal | Inferred | ✅ High |
| **Link** | Green / Default | No Fill | Normal | Inferred | 🟡 Moderate |

---

### 📋 Consolidated Audit Findings

| Sheet Name | Cell Reference | Description of the Location | Short Error Category | Long Description of Error |
|:---|:---|:---|:---|:---|
| **Dynamic Array** | J55 | Calculation Area | **Calculation Error** | 🔴 **HIGH:** Cell contains a `#VALUE!` error, indicating a breakdown in the calculation chain in this period. |
| **One Line** | AD11 | Equity NPV Calculation | **Formula Pattern Break** | 🔴 **HIGH:** This cell breaks the row pattern by replacing a standard `XNPV` formula with a massive manual calculation. (Expected R1C1: `=XNPV(Inputs!R81C6,R11C9:R11C61,R5C9:R5C61)` vs Actual R1C1: Complex multi-line manual NPV). |
| **Inputs** | G21:S21 | Timeline / Control Row | **Colour Coding Error** | ⚠️ **MEDIUM:** Cells are formatted with a bright red fill (#EF4136) which deviates from the established model style and may indicate an accidental override or "work-in-progress" marker. |
| **Dynamic Array** | C27, C67, C116 | Row Labels | **Dominant Term Mismatch** | 🟡 **LOW:** Inconsistent casing in row labels: "Development costs" vs "Development Costs", "Interest rate" vs "Interest Rate", "Base rate" vs "Base Rate". The model predominantly uses Title Case for labels. |
| **L** | C18, G18 | Reference Labels | **Typo** | 🟡 **LOW:** Possible typo or non-standard term: "Offsheet" (Expected: "Off-sheet" or "Link"). |
| **Inputs** | D101 | Taxation section | **Number Format Mismatch** | 🟡 **LOW:** Tax Rate label uses a "%" sign in a separate cell (D101) but the input value (F101) is correctly formatted as a percentage. This is a minor presentation inconsistency. |

---

### 👔 Manager's Summary & Recommendations

1.  **Immediate Fix (Dynamic Array):** The `#VALUE!` error in cell J55 should be investigated as it likely invalidates any summary results or NPVs derived from the Dynamic Array sheet.
2.  **Logic Integrity (One Line):** The pattern break in AD11 is highly suspicious. It appears the modeler manually "unrolled" an NPV calculation for a specific column, which compromises the scalability and auditability of the sheet.
3.  **Style Standardisation:** The bright red fill in the Inputs sheet should be reverted to the standard style unless it serves a specific, documented purpose.
4.  **Overall Quality:** Aside from the specific errors noted, the model follows a structured approach (Corality/SMART standards) with clear separation of Inputs and Calculations. Tax logic matches the Australian 30% corporate rate.

---

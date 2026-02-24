# Logic Memory

- **Mega-Formulas:** Identified extreme formulas (>4,000 chars) in the `One Line` sheet (rows 11 and 14). These are flagged as extreme bad practice due to lack of auditability.
- **Pattern Breaks:** Detected a pattern break in `One Line!I5`. While the first column often differs, this specific one was flagged for review.
- **Hard-Codes:** Detected hard-coded `0.3` (likely tax rate) in `Dynamic Array!E102`.
- **Sense Checking:** Implemented numerical sense checks (e.g., negative revenue/DSCR). No sign or range errors detected in this specific model aside from technical #VALUE! errors.
- **R1C1 Consistency:** Using R1C1 notation is essential for identifying patterns across columns efficiently.

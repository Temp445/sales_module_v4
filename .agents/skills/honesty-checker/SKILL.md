---
name: honesty-checker
description: Rigorous validation and equivalence analysis between Next.js and VB6 forms, databases, and UI fields.
---

# Honesty Checker Skill

This skill guarantees 100% equivalence, absolute accuracy, and transparency when validating Next.js migrated forms against legacy VB6 implementations.

It prevents faking or assuming matches by enforcing strict, step-by-step documentation of all fields, types, and logic discrepancies.

## core Guidelines

1. **Verify both labels and bindings:** 
   - A field may have a label in VB6 (e.g., `Customer Formula Name`) but save to a database column with a different meaning or type (e.g., `CustomerFormulaCode`).
   - You must trace how the field is loaded, how it is displayed, how it is updated in VB6 (e.g., Tag properties vs Text properties), and verify that the Next.js implementation provides equivalent behavior (e.g., showing the name but storing the code).

2. **Strict Mappings Check:**
   - Write down every field in a clear table:
     | VB6 Control Name | VB6 Label / UI Display | VB6 Database Column / Binding | Next.js State / Binding | Next.js Label / UI Display | Match Status | Discrepancy details |
   - Never declare "100% match" until this table has been fully populated and verified.

3. **Validation Equivalence:**
   - Match field lengths (e.g. `maxLength` in Next.js matching the size limit in database and VB6 text boxes).
   - Match types (e.g. `parseInt` for integer codes, string trimming, boolean checkboxes).
   - Match behavior of custom lookup selectors (e.g., if VB6 uses a Tag for an ID/Code and Text for the Name, the Next.js form must implement a search popup, dropdown, or search field that maps to the database ID but shows the name).

4. **No Dummy Data (Strict Interoperability):**
   - You are forbidden from guessing dropdown options or hardcoding dummy select values.
   - You must search the VB6 `.frm` file for `cmbBox.AddItem` lines or extract strings from binary `.frx` files.
   - If a Next.js dropdown option deviates by even one space, it will corrupt the database for legacy users. Flag this immediately.

5. **Bidirectional Data Flow Tracing:**
   - Do not just map UI fields to database columns.
   - **Load Path:** Prove how the UI receives data (e.g., trace `FetchValue_New` and document the exact mapping line: `txtTinNo.Text = snap("PanNo")`).
   - **Save Path:** Prove how the UI saves data (e.g., trace string manipulations and SQL concatenations like `Sql = Sql & "@UdyamNo="`).

6. **Honest Output Enforcement:**
   - If a discrepancy is found, report it immediately to the user.
   - Do not make up matching names or assume that a field can be simplified without user consent.

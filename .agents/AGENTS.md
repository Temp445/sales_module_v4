# ERP Migration Strict Agent Workflow Rules

These rules apply globally to ALL agent actions when migrating legacy ERP forms to Next.js. They are designed to prevent "lazy execution", missed fields, and untestable code.

## 1. Zero-Assumption UI Generation
When building a React form (`.tsx`), you MUST map **every single input** from the `[FormName]_extracted.json` `controls` array.
- **Rule:** Before calling the `write_to_file` or `replace_file_content` tool to create/update the UI component, you must explicitly read the `[FormName]_extracted.json` artifact using `view_file`.
- **Rule:** Do NOT skip fields because they "seem unnecessary" or you "assume" a modern equivalent is better. If the JSON has `txtEnqType`, the React form MUST have `enqType` state.

## 2. Test-Driven API Development
You are forbidden from writing a Next.js `/api/` route without first proving the SQL query works.
- **Rule:** Before writing any API route, you MUST create a `scratch/test_db_[feature].js` script using Prisma `$queryRawUnsafe` to execute the exact SQL query against the real database.
- **Rule:** You MUST execute this script and verify the console output matches expectations (e.g., ensuring `Compid` is correct and rows are returned). Only after the script succeeds may you write the `route.ts`.
- **Rule:** Every fetch API that filters by `session.compId` or `session.branchId` MUST include a debug `console.log` printing the exact session values used as filters. You MUST also run a scratch `.mjs` script WITHOUT session filters (`findMany({take:3})`) to confirm the actual stored DB values match the session. Never declare a dropdown "working" until the fetched count is `> 0` with real credentials.
- **Rule:** If a scratch test script fails mid-run (e.g., crashes at step 2 of 4), all remaining steps are also considered FAILED. Fix and re-run the entire script until every step passes.

## 3. Mandatory Check-Ins & Logging
You cannot silently ignore the `Activity_task.md` file.
- **Rule:** Upon completing ANY step (UI creation, DB script test, API creation, API integration), you MUST immediately update `[FormName]_Activity_task.md` with a timestamped log entry.
- **Rule:** You must run a final `scratch/test_[form]_save.js` script to prove end-to-end data submission works before declaring the migration "complete" to the user.

## 4. No Hardcoding allowed
You are strictly forbidden from hardcoding global session fallback values anywhere in the codebase.
- **Rule:** Never hardcode fallback values for `branchId` (e.g., "00") or `compId` (e.g., "07" or "02"). Always read them from the actual session or context variables. Let the application fail gracefully if missing.
- **Rule:** If a dropdown or selection list exists in VB6, it must be dynamically fetched from the database via an API in Next.js. Do not hardcode `<option>` values based on guesses.

## 5. Proactive API Generation
Do not leave dropdowns empty or wait for failures.
- **Rule:** When migrating a form, you MUST proactively parse the original `.frm` file (or use the JSON) to identify exactly where every dropdown gets its data.
- **Rule:** You MUST automatically write and test the new `/api/` route to fetch that required data *before* hooking it up to the React UI, preventing "empty dropdown" bugs entirely.
- **Rule:** For APIs that filter by a discriminator like `ModuleCode`, you MUST first run a scratch `.mjs` script (`findMany({take:5})`) WITHOUT any filter to observe what values actually exist in the DB. Never assume a default (e.g., do NOT use `20` if the data only contains `4`). Derive all defaults from observed DB values.
- **Rule:** When an `isActive` or similar boolean flag filters out ALL rows (common in legacy/test databases where old records have `isActive: false`), you MUST remove or make the filter optional rather than assuming the data is missing.

## 6. Stored Procedure Contract Verification (Prevents Formula/Terms type bugs)
Before calling ANY stored procedure via `$executeRawUnsafe`, you MUST verify the SP parameter contract.
- **Rule:** Look up the corresponding model in `prisma/schema.prisma` and verify EVERY column's data type (`Int`, `String`, `DateTime`, `Float`, `Boolean`) before writing any raw SQL.
- **Rule:** For Boolean columns in Prisma `where` clauses, you MUST pass `true` or `false` — NEVER `1` or `0`. Passing an integer to a Boolean field causes a runtime crash (`Expected BoolFilter or Boolean, provided Int`). Always check `schema.prisma` — do not assume type from field name.
- **Rule:** For EVERY parameter passed as a bare number (unquoted in SQL), verify that the source value is sanitized with `parseInt()` or `parseFloat()` before interpolation. NEVER interpolate a React state variable directly into an unquoted SQL parameter.
- **Rule:** For EVERY dropdown that feeds a numeric SP parameter (e.g., `@FormulaCode [int]`), the dropdown's `value` in React MUST be the numeric ID, not the display label/name. The API that fetches dropdown options MUST return the numeric ID as `value` and the human-readable name as `label`.
- **Rule:** For EVERY auto-fill API (like `customer-defaults`), verify it returns the same ID type the SP expects. NEVER return a display string (name) when the SP expects an integer code.

## 7. No Placeholder/Mocked Save Logic (NEW â€” Prevents fake success messages)
You are absolutely forbidden from shipping a save API with commented-out or placeholder SQL.
- **Rule:** A save API route MUST NOT be created until you have verified the target table model exists in `prisma/schema.prisma`, all column data types are mapped, and a scratch test script successfully executes the raw query.
- **Rule:** If the SP test script fails, you MUST fix the SQL issue before creating the API route. You MUST NOT create the route with a TODO comment or fake return value.
- **Rule:** A save API is only "complete" when `scratch/test_[form]_save.js` returns HTTP 200 and the data is physically verifiable in the database via a follow-up SELECT query in `scratch/test_[form]_list.js`.

## 8. Value vs. Label Separation Rule (NEW â€” Prevents string-in-SQL bugs)
Every `SearchableSelect` dropdown that feeds a database save MUST separate `value` (DB key) from `label` (display text).
- **Rule:** The `value` prop of any dropdown option must ALWAYS be the primary key or integer code used by the database (e.g., `FrmID`, `PartyCode`, `Code`).
- **Rule:** The `label` prop is purely for display. It MUST NEVER be sent to the backend as a parameter value.
- **Rule:** When auto-filling related fields from another form (e.g., customer-defaults), the API response MUST return the same type of key that was stored â€” not the human-readable name.

## 9. Continuous Testing Until 100% Success (NEW Prevents premature "done" declarations)
You are forbidden from declaring any feature "complete" or "working" unless ALL tests pass without errors.

- **Rule:** After every save attempt, you MUST immediately run `scratch/test_[form]_save.js` again. If it returns anything other than HTTP 200 with a valid record ID, you MUST fix the failure and re-run. Do NOT ask the user to retry fix it yourself first.
- **Rule:** After the save test passes, you MUST run `scratch/test_[form]_list.js` to verify the saved record physically EXISTS in the database. A HTTP 200 save response alone is NOT enough â€” you must confirm the row is visible via SELECT.
- **Rule:** If ANY test fails at ANY point in the loop, you MUST:
  1. Read the exact error message.
  2. Identify the root cause (wrong data type, missing field, wrong SP parameter, string in a numeric column, etc.).
  3. Fix the root cause in the source file.
  4. Re-run the failing test.
  5. Repeat until the test passes cleanly with zero errors.
- **Rule:** The loop is: **Write / Test / Fail / Fix / Re-Test / Repeat** until the output is 100% success. You may NOT break out of this loop by adding placeholder code, fake success responses, or TODO comments.
- **Rule:** You MUST report the final successful test output to the user (e.g., paste the `Status: 200`, `quotationNo: "SQ-5"`, and `SELECT` confirmation result) as proof of completion. A verbal claim of "it works now" without test output is not acceptable.
- **Rule:** For UI-level bugs (empty dropdowns, wrong values displayed, missing auto-fill), you MUST open the browser test page, screenshot or verify the data flow end-to-end, and confirm the UI correctly reflects the saved database values before closing the task.

## 10. No Summarized Documentation (NEW — Prevents incomplete analysis)
You are strictly forbidden from "summarizing" or "grouping" controls when creating the `[FormName]_Analysis.md` document to save tokens or time.
- **Rule:** Every single control extracted in the JSON MUST be explicitly listed in the Markdown file with its full properties (Caption, Data Type, Purpose, etc.).
- **Rule:** If a form has a large number of controls (e.g., >20), you MUST write and execute a Node.js/Python script to dynamically generate the Markdown file from the JSON. Do NOT attempt to manually type out a large analysis document, as you will inevitably take shortcuts.

## 11. Programmatic UI Mapping Verification (Anti-Missing-Field Rule)
You are strictly forbidden from declaring a UI migration complete based on visual assumptions or by copying existing React templates. Before finalizing any React form, you MUST write and execute a validation script `scratch/verify_[form]_ui.mjs` that performs ALL of the following steps automatically:
- **Step 1 – VB6 Census**: Read `[FormName]_extracted.json` and enumerate every control of type `TextBox`, `ComboBox`, `CheckBox`, `DTPicker`, `OptionButton`. Build a checklist of every control name.
- **Step 2 – Grid Enum Census**: For any `AceSpread`, `MSHFlexGrid`, or similar grid in the `.frm`, locate its `Private Enum` (e.g., `Private Enum eItem`). Count the entries — your React `<table>` MUST render exactly that many `<th>` columns. Each enum entry's computed value (e.g., `mAssGoodsValue = Qty x AssRate`) MUST also be implemented in `handleItemChange`.
- **Step 3 – React Census**: Parse the `.tsx` file and extract: all `useState` variable names, and every JSX element of type `<input>`, `<select>`, `<SearchableSelect>`, `<textarea>`, `<input type="checkbox">`. Map each to a VB6 control by name similarity.
- **Step 4 – Gap Report**: Output a table with columns: `VB6 Control | VB6 Type | React State | React JSX Element | Status`. Status must be one of: `PRESENT`, `MISSING`, `WRONG TYPE`.
- **Step 5 – Fail Hard**: If ANY row shows `MISSING` or `WRONG TYPE`, the script MUST `process.exit(1)`. You are forbidden from declaring the form complete until the script exits with code 0.

## 12. Strict Input Type Enforcement
Every VB6 control type MUST map to a specific, non-negotiable React JSX element type:
- `TextBox` (Locked=False) → `<input type="text">` or `<input type="number">`
- `TextBox` (Locked=True, DblClick Lookup) → `<SearchableSelect>` (NOT a text input)
- `DTPicker` → `<input type="date">`
- `ComboBox` (DB driven) → `<select>` with dynamic API options
- `ComboBox` (Static) → `<select>` with exact `.frx` options
Violation of this mapping is a critical defect. Your verify script (Rule 11) must assert these types and report `WRONG TYPE` for mismatches.

## 13. Mandatory Helper & Lookup API Testing
- **Helper/Cascading APIs** (e.g., Country → State → City): Verify each endpoint returns data and cascading steps filter correctly. Check that API response keys perfectly match what React expects (e.g. `StateName` vs `label`).
- **Terms/TypeFlg APIs**: Test each `TypeFlg` value individually (`'Form'`, `'Payment Terms'`, etc.). You MUST ensure the API Prisma query strictly filters by the discriminator (e.g. `TypeFlg`). If you fetch all records unfiltered from a shared master table, the React UI will silently auto-fill wrong values because identical ID codes are reused across different categories (e.g. Code `3` is both a Scrap Reason and a Form).
- **Dropdown APIs**: Assert `count > 0`, PK exists, display name exists, and ALL string values are trimmed (no trailing spaces).
Print a PASS/FAIL table. You are forbidden from wiring the UI until assertions pass.

## 14. SP vs Schema vs UI Checklist
Upon completing a form migration, you MUST create a `[FormName]_SP_Checklist.md`. You must map and explicitly compare three layers for every saved field:
- **Stored Procedure Parameter** (e.g., `@PartyCode varchar(10)`)
- **Prisma Schema** column name and type
- **React UI Input** state variable and component
Mark each with `[x]` to verify they align perfectly in intent and data type. Fix any mismatches before marking the task complete.

## 15. API Reuse and Centralization (No API Duplication)
Do not create duplicate API endpoints for shared master data (e.g., Currency, City, Terms, Formula, Country). Before creating a new dropdown or lookup API for a form, you MUST first search the existing codebase to check if an endpoint that serves this exact master data already exists (e.g., checking if `/api/sales/masters/terms/fetch` already exists). If a suitable endpoint exists, you MUST reuse it. Only create a new API if the required data is strictly form-specific.

## 16. SQL Server Trailing Spaces (Anti-Empty-Result Rule)
Legacy SQL Server databases often use fixed-length `char` columns which pad values with trailing spaces (e.g., `"Q" ` becomes `"Q    "`).
- **Rule:** When writing Prisma `where` clauses for exact string matching on legacy discriminator columns (like `PertainingTo`, `TypeFlg`), you MUST use `{ startsWith: value }` instead of direct equality to prevent empty results.
- **Rule:** When filtering or mapping data on the frontend client, you MUST use `.trim()` on the string before doing an `===` comparison to prevent silent failures.

## 17. Reactive Auto-Fill vs. Legacy OK Buttons (Anti-Clunk Rule)
In VB6, auto-filling related fields (like Terms & Conditions based on a Customer) often required the user to click an "OK" button to trigger a subroutine (e.g., `cmdOk_Click` calling `TermsDetails`). 
- **Rule:** You MUST NOT migrate this two-step "Header OK" workflow to React.
- **Rule:** Instead, make the UI fully reactive. When the user selects a value in a primary dropdown (e.g., Customer), use a `useEffect` hook to instantly fetch and auto-fill all dependent default fields (e.g., terms, formula) without requiring a manual confirmation button.

## 18. Mandatory Event-Loop Mapping (Anti-Silent-Failure Rule)
The JSON and Schema are static; VB6 behavior is dynamic. You are forbidden from implementing React `onChange` or `useEffect` logic based purely on guesses from the JSON structure.
- **Rule:** Before writing any state or validation logic for a form control (especially grids), you MUST open the original `.frm` file and read the exact `_LeaveCell`, `_KeyDown`, `_Click`, and `_ComboSelChange` event subroutines for that control line-by-line.
- **Rule:** Every `Spread_Lock`, `Spread_Clear_Duplicate`, or `If` condition found in those VB6 event handlers MUST be explicitly translated into the Next.js form logic.
- **Rule:** Do not declare a feature complete until you have rigorously compared the VB6 dynamic behaviors against your Next.js implementation using the `honesty-checker` or manual verification.

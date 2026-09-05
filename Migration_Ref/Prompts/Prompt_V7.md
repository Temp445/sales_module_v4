# New Analysis – Frontend Inputs, Controls, Database Usage & Business Logic

**CRITICAL THREE-STEP PIPELINE:**
You must perform this analysis in three distinct steps to ensure maximum accuracy. 

**PRIORITY RULE:**
1. **HIGHEST PRIORITY:** Database Stored Procedures, Functions, and Queries. You must analyze the internal SQL logic to determine exactly what tables are being modified (e.g., Does it INSERT into a specific table? Does it UPDATE a rates table? Does it write to an audit log?).
2. **SECOND PRIORITY:** The VB6 Frontend Code (`.frm`, `.bas`, `.cls`).

Treat the original SQL scripts and VB6 source code as the ONLY sources of truth. Beware of "lazy reading" (e.g., assuming a DatePicker drives database logic just because it is named `dtpMonth` when the code actually uses `ServerDate`).

1. **JSON Extraction:** First, extract all facts from the source code and SQL scripts into a strictly structured `.json` artifact. JSON = extracted facts, not narrative. If a detail is not explicitly in the code, use **"UNKNOWN"** or **"UNVERIFIED"** instead of guessing.
2. **Markdown Generation:** Second, use the JSON artifact to generate the final detailed Markdown (`.md`) file.
3. **Independent Verification:** Third, perform an independent verification step at the end of the document. Do not claim a 100% match unless the verification shows no critical missing/conflicting items.

### Step 1: JSON Schema Requirements
Your intermediate JSON must strictly decouple facts from narrative and use the exact structural skeleton below. Do not deviate from this schema:

```json
{
  "formMetadata": {},
  "controls": [],
  "dataBindings": [],
  "events": [],
  "procedures": [],
  "databaseOperations": [],
  "transactionHandling": [],
  "businessRules": [],
  "validations": [],
  "calculations": [],
  "conditions": [],
  "databaseObjects": {
    "directTables": [],
    "storedProcedures": [],
    "sqlFunctions": [],
    "indirectTables": []
  },
  "formDependencies": [],
  "globalHelpers": [],
  "externalDependencies": [],
  "messages": [],
  "variables": [],
  "globalVariables": [],
  "constants": [],
  "errorHandling": [],
  "sourceReferences": [],
  "uiStateTransitions": {},
  "unknowns": []
}
```

*Note: For `controls`, explicitly EXCLUDE purely visual/formatting properties (e.g., `ClientHeight`, `ClientLeft`, `ClientTop`, `ClientWidth`, `Font`, `Name`, `Size`, `Charset`, `Weight`, `Underline`, `Italic`, `Strikethrough`, `BackColor`, `ForeColor`). ONLY capture logical and functional properties: `name`, `type`, `maxLength`, `visible`, `enabled`, `required`, `toolTip`, and explicit `usage`. For `dataBindings`, identify if grids rely on `.frx` files instead of explicit SQL. For `databaseOperations`, use explicit keys (triggerEvent, type, tables, columnsSelected, whereConditions, ifCondition) and capture ALL implicit parameters like `gFinyear`. For `transactionHandling`, document `DBStartTransaction`, `DBExecuteSql`, and `DBCommitTransaction`. For `globalHelpers`, document custom application functions like `CheckNull`, `RSOpen`, `sShowMessage`. Analyze referenced `.bas` or `.cls` files when necessary if the module relies on them for business logic.*

### Step 2: Markdown Generation
Using the extracted JSON data, create a **new detailed analysis Markdown (`.md`) file** for the selected VB6 form/module.

The analysis must cover the following:

## 1. Total Frontend Inputs & Controls

Identify and document **all Frontend UI Inputs and Controls**.
For each input/control, clearly describe:

* Control name & Control type
* Caption / label & Data type
* Key Properties (e.g., MaxCols, MaxRows, Bindings)
* Whether it is required or optional
* Default value & Validation rules
* Exact purpose & User interaction
* Data Source / Bindings (Is it populated by SQL or an `.frx` binary file?)
* **Actual Usage vs Named Usage:** Does the control *actually* affect logic, or is it purely cosmetic? (Verify in the source code).

At the end, provide the **total number of UI inputs and controls**.

## 2. Database Tables Used

Identify **all database tables used by the form/module**. Do **not duplicate table names**.
For every table, document: Table name, Purpose, Columns used, and Operations (SELECT/INSERT/UPDATE/DELETE).

## 3. Data Fetching & Database Operations

For every database operation, identify:
* Where data is fetched from (Table accessed)
* Columns selected & Conditions used
* Parameter values (Do not miss implicit parameters like global variables passed to SQL functions).
* What UI control receives/displays the result

Clearly distinguish between: **Fetch/SELECT, INSERT, UPDATE, DELETE, Validation queries, Lookup queries.**

**For Stored Procedures, you MUST explicitly list out their internal operations in the analysis and JSON.** For example:
* Does it INSERT into a table called `Bank_Info`?
* Does it UPDATE a table called `Item_Rates`?
* Does it write to an audit log?

Do not assume a database operation if it is not present in the existing VB6 code or SQL scripts. Use UNKNOWN if verification is impossible.

## 4. Detailed Business Rules, Conditions & Procedures

Identify and document **all business rules, functions, and event handlers implemented in the form/module**.
For each business rule/procedure, you MUST explicitly document:
* **Scope & Type:** Is it `Private` or `Public`? Is it a `Sub` or `Function`?
* **Conditional Logic:** Document the exact `If/Else` and `ElseIf` branches.
* **Database Filters:** Document the exact `WHERE` clauses used in any inline SQL queries.
* **Trigger condition:** When does this rule fire?
* **Calculations & Branches:** Source line, Inputs/Database values involved, Calculations, and True/False branches.

**CRITICAL:** Do not just summarize the UI controls. You must deeply extract the actual programmatic procedures, their scopes, and their conditional branching logic.

**Transaction Handling & Global Helpers:**
* Explicitly document how transactions are handled (e.g., `DBStartTransaction`, error rollbacks, `DBCommitTransaction`).
* Document all global helper functions used (e.g., `CheckNull`, `RSOpen`, `Spread_Clear`).

## 5. Form Events & Workflow

Document all important VB6 events and their behavior (`Form_Load`, `_Click`, `_Change`, etc).
For each event, explain the complete workflow and any database or business-logic operations it performs.

## 6. Input → Business Logic → Database Mapping

Create a clear mapping showing:
**UI Control → Validation → Condition → Business Logic → Database Table/Column → Operation → UI Result**

## 7. Avoid Duplication
Do not duplicate tables, rules, or controls.

## 8. VB6 Source References
For every important functionality, identify the original source location wherever possible (Form name, Control name, Event, Function, Relevant `.frm` / `.bas` / `.cls`, Line number). Use **UNKNOWN** if it cannot be determined.

## 9. Migration Requirements
The analysis must provide enough detail to implement the same functionality in **Next.js + TypeScript + Prisma** without changing the existing business behavior.

### Step 3: Independent Verification & Final Summary

After generating the Markdown, independently compare the original VB6 source against the generated JSON and Markdown.

The verification must identify:
- Controls/Properties present in VB6 but missing from JSON/Markdown
- Data bindings (like `.frx`) missed by the analysis
- Events/Procedures missing from analysis
- Database tables/operations/parameters missing from analysis
- Conditions/Business rules missing from analysis
- Global Helper Functions / Transactions missing from analysis
- Information present in Markdown but not supported by source (Hallucinations)

Classify every difference as:
* **MATCH**
* **MISSING**
* **EXTRA**
* **CONFLICT**
* **UNVERIFIED**

The original VB6 source code must always be treated as the source of truth.
Do not declare the analysis complete or accurate if any critical item is MISSING, EXTRA, or CONFLICT.

Finally, provide a summary containing:
* Total UI Inputs / Controls
* Unique Database Tables Used
* Database Operations (SELECT/INSERT/UPDATE/DELETE)
* Validation / Business Rules
* Conditional Logic Locations
* Form/External Dependencies & Global Helpers
* Potential Migration Risks
* **Verification Status:** (Based strictly on the classifications above).

## 5. Next.js Implementation Rules

1. **Binary Data Extraction & JSON Reference**: If the VB6 form relies on an `.frx` file, you MUST use the provided extraction script `@[Migration_Ref/Scripts/extract_frx.js]` to generate a `[FormName]_extracted.json` file. You MUST explicitly read and reference this JSON file during implementation to accurately recreate grid columns, dropdown options, and form controls.
2. **Strict Data Accuracy**: Use the generated `[FormName]_extracted.json` file as the absolute source of truth for the UI data model. Do NOT add dummy data, fake mappings, or placeholder fallback data in the Next.js implementation.
3. **Mandatory Analysis Document**: Before generating an implementation plan or writing any Next.js code, you MUST create a `[FormName]_Analysis.md` file in the `Migration_Ref` folder. This document must break down the original VB6 form's UI components, database interactions, business logic, and potential migration risks. Never skip this step.
4. **Rigorous Continuous Testing**: Once the Next.js UI is built, you MUST run continuous end-to-end testing of the data fetching and saving operations. **All testing and verification scripts MUST be created and saved inside the `scratch/` folder.** Do NOT mark a task as complete until the form data saves 100% correctly without errors. Ensure all global variables (like `branchId`, `compId`, `userId`) are properly passed to the API and SQL queries.
5. **Mandatory Activity Logging (`Activity_task.md`)**: For every form conversion, you must create and maintain an `Activity_task.md` file. You must log every single activity (conversion steps, testing attempts, passes, fails, form validations) with an explicit timestamp. This serves as the undeniable proof of work and testing.
5. **Skill Utilization**: If required during the migration, actively call the specialized agent skills to assist with business logic extraction or equivalence testing.
6. **Standard SP Parameters**: Legacy ERP stored procedures almost always expect standard audit parameters. When calling an `EXECUTE` query, you MUST always include `@Userid`, `@Branchid`, and `@Compid` (and `@ModuleName` if applicable), even if the extracted JSON does not explicitly show them. Extract these values from the session or request globals.
7. **Stored Procedure Transactions**: Legacy SQL Server stored procedures often include an internal `COMMIT TRANSACTION` statement but expect the client application (e.g., ADO) to open the transaction. When calling them via Prisma (`prisma.$queryRawUnsafe()` or inside `prisma.$transaction`), you MUST prepend `BEGIN TRAN;` to the raw SQL string before the `EXECUTE` statement to prevent "no corresponding BEGIN TRANSACTION" errors from SQL Server.
8. **Dynamic Dropdowns (No Hardcoding/Disabling)**: Never hardcode options or disable dropdowns (ComboBoxes) simply because the data source isn't immediately obvious. You must trace the global helpers (e.g., `LoadDocType`, `FetchValue`) to identify the underlying database tables and dynamically fetch the options via a new Next.js API route. If a field was a dropdown in VB6, it MUST be a fully functioning dynamic dropdown in Next.js.
9. **No Hardcoded Fallbacks**: Never hardcode fallback values for session variables like `branchId = "00"`, `compId = "07"`, or `userId = "USER1"` in component props or server code. Strictly enforce reading from the actual Next.js session context and allow the application to fail gracefully if the session is missing.
10. **Searchable Dropdowns**: For dropdowns containing master data (e.g., Customers, Items, Vendors) that could have dozens or hundreds of records, NEVER use a plain native HTML `<select>`. You MUST implement a searchable custom dropdown (combobox) with an integrated search bar so the user can easily filter and find records by name or code.
11. **Grid Column Cross-Checking**: When building UI data grids (`<table>`), you MUST strictly cross-check your React table headers (`<th>`) and data cells (`<td>`) against the extracted JSON `dataBindings` array. Additionally, for any `AceSpread`, `MSHFlexGrid`, or similar VB6 grid control, you MUST locate its `Private Enum` in the `.frm` source (e.g., `Private Enum eItem`). The enum is the authoritative column list — your React `<table>` MUST render exactly as many `<th>` columns as there are enum entries. Any computed column value defined by the enum (e.g., `mAssGoodsValue = Qty × AssRate`) MUST also be implemented in the React `handleItemChange` function. Do not omit any column from the UI markup.
12. **Professional UI & Aesthetics**: Ensure a modern, enterprise-grade look for all converted forms. You MUST use professional icons for standard UI actions (Save, Clear, Trash, Plus, etc.) from the project's standard icon library. Use thin, custom-styled WebKit scrollbars for overflowing grids (e.g., `[&::-webkit-scrollbar]:h-2`). Keep vertical spacing compact but breathable (`gap-y-5`), and do NOT use horizontal line dividers (`border-t`) inside a single cohesive form section to avoid clutter. Ensure all inputs have subtle hover/focus states, and use professional empty-state graphics when a grid is empty or locked.
13. **Proactive API Generation**: Do not leave dropdowns empty or wait for failures. When migrating a form, you MUST proactively parse the original `.frm` file (or use the JSON) to identify exactly where every dropdown gets its data. You MUST automatically write and test the new `/api/` route to fetch that required data *before* hooking it up to the React UI, preventing "empty dropdown" bugs entirely. For APIs that filter by a discriminator like `ModuleCode`, you MUST run a scratch `.mjs` script WITHOUT any filter first to observe what values actually exist in the DB — never assume a default value. For every fetch API that filters by `session.compId` or `session.branchId`, add a debug `console.log` of the exact session values and verify they match the actual DB stored values by running an unfiltered scratch query.
14. **Stored Procedure Contract Verification**: Before calling ANY stored procedure via `$executeRawUnsafe`, you MUST look up the corresponding model in `prisma/schema.prisma` and verify EVERY column's data type (`Int`, `String`, `DateTime`, `Float`, `Boolean`). For Boolean columns in Prisma `where` clauses, you MUST pass `true` or `false` — NEVER `1` or `0`. Passing an integer to a Boolean field causes a runtime crash (`Expected BoolFilter or Boolean, provided Int`). Do not assume the type from the field name — always verify in `schema.prisma`. For EVERY parameter passed as a bare number (unquoted in SQL), you MUST sanitize the source value with `parseInt()` or `parseFloat()` before interpolation. For EVERY dropdown that feeds a numeric SP parameter (e.g., `@FormulaCode Int`), the dropdown `value` in React MUST be the numeric ID. For EVERY auto-fill API (like `customer-defaults`), verify it returns the same ID type the SP expects — NEVER return a display name string when the SP expects an integer code.
15. **No Placeholder or Mocked Save Logic**: You are absolutely forbidden from creating a save API route with commented-out SQL, TODO comments, or fake return values. A save API route MUST NOT be created until: (a) the target table/model is verified in `prisma/schema.prisma`, (b) all column data types are mapped, and (c) a scratch test script successfully executes the raw query. The save API is only "complete" when `scratch/test_[form]_save.js` returns HTTP 200 AND the data is physically verifiable in the database via a follow-up SELECT.
16. **Value vs. Label Separation**: Every `SearchableSelect` dropdown that feeds a database save MUST separate `value` (DB primary key / integer code) from `label` (human-readable display text). The `label` MUST NEVER be sent to the backend as a parameter value. The `value` prop must always be the key the stored procedure expects (e.g., `FrmID`, `PartyCode`, `Code`). Auto-fill APIs must return the same type of key as the SP expects — not the display name.
17. **Continuous Testing Until 100% Success**: You are forbidden from declaring any feature "complete" unless ALL tests pass. The mandatory loop is: **Write → Test → Fail → Fix → Re-Test → Repeat**. After every save attempt, immediately run `scratch/test_[form]_save.js`. If it returns anything other than HTTP 200 with a valid record ID, fix the root cause and re-run — do NOT ask the user to retry. After the save test passes, run `scratch/test_[form]_list.js` to confirm the row physically EXISTS in the database. A HTTP 200 response alone is NOT proof. You MUST paste the actual test output (`Status: 200`, record ID, and SELECT confirmation) to the user as proof of completion. A verbal claim of "it works now" without test output is not acceptable.
18. **No Summarized Documentation (Scripted Analysis Mandate)**: You are strictly forbidden from "summarizing" or "grouping" controls in the `[FormName]_Analysis.md` document to save output tokens. Every single control from the extracted JSON MUST be explicitly listed in the Markdown file with its full properties. If the form has a large number of controls (e.g., >20), you MUST write and execute a Node.js/Python script to dynamically generate the Markdown file from the JSON. Do NOT attempt to manually type out a large analysis document, as this guarantees missing fields.

19. **Programmatic UI Mapping Verification**: You are strictly forbidden from declaring a UI migration complete based on visual assumptions or by copying existing React templates. Before finalizing any React form, you MUST write and execute a validation script `scratch/verify_[form]_ui.mjs` that performs ALL of the following steps automatically:
   - **Step 1 – VB6 Census**: Read `[FormName]_extracted.json` and enumerate every control of type `TextBox`, `ComboBox`, `CheckBox`, `DTPicker`, `OptionButton`. Build a checklist of every control name.
   - **Step 2 – React Census**: Parse the `.tsx` file and extract: all `useState` variable names, and every JSX element of type `<input>`, `<select>`, `<SearchableSelect>`, `<textarea>`, `<input type="checkbox">`. Map each to a VB6 control by name similarity.
   - **Step 3 – Gap Report**: Output a table with columns: `VB6 Control | VB6 Type | React State | React JSX Element | Status`. Status must be one of: `PRESENT`, `MISSING`, `WRONG TYPE`.
   - **Step 4 – Fail Hard**: If ANY row shows `MISSING` or `WRONG TYPE`, the script MUST `process.exit(1)`. You are forbidden from declaring the form complete until the script exits with code 0.
   - Re-run the script after every `.tsx` edit until it exits cleanly with zero failures.

20. **Deep Analysis & Phased Migration (Anti-Superficial Rule)**: When analyzing and migrating a legacy VB6 form, you must strictly adhere to Deep Analysis and Phased Implementation guidelines. Failure to extract all data points or blindly copying VB6 code is unacceptable.
   - **Exhaustive Data Extraction (No Missing Fields)**: Do not blindly parse VB6 .frm properties. Perform a deep semantic analysis to find every single input field. Cross-reference the .frm controls with the Prisma schema definitions. If a table has 50 fields, your UI and API must account for all 50 fields. Generate a highly detailed `_extracted.json` that maps every UI control to its exact Database Column, Table, Data Type, and related Stored Procedure.
   - **Dependency & Sub-Master Discovery**: Identify if the master form relies on other forms/lookups. If APIs for these sub-masters do not exist, you must analyze them, document them, and include their creation in your implementation plan.
   - **Phased Implementation**: Do not attempt to write the entire migration in one massive output. Break the migration down into logical, manageable phases in your `implementation_plan.md` (e.g., Phase 1: Deep Analysis, Phase 2: Database & API Layer, Phase 3: UI Foundation, Phase 4: Integration & Testing).

21. **Strict UI Control Type Matching**: When validating the UI (Rule 19), you must NOT just check if a state variable name exists in the React component. Your verification script MUST enforce the following non-negotiable VB6-to-JSX mapping and report `WRONG TYPE` for any violation:

   | VB6 Control Type | Condition | Required React Element | Forbidden Elements |
   |---|---|---|---|
   | `TextBox` | Free text (`Locked=False`) | `<input type="text">`, `<input type="email">`, or `<input type="number">` | Any other type |
   | `TextBox` | Lookup (`Locked=True`, popup on DblClick/F2) | `<SearchableSelect>` | `<input type="text">` is a critical failure |
   | `DTPicker` | Any | `<input type="date">` | `<input type="text">` is a critical failure |
   | `ComboBox` | Data from DB (`LoadComboBox` / SQL) | `<select>` with dynamic API-fetched options | Hardcoded `<option>` values |
   | `ComboBox` | Static data (`.frx` / `AddItem`) | `<select>` with exact options from `.frx` extraction JSON | Guessed or approximate options |
   | `CheckBox` | Any | `<input type="checkbox">` | Any other type |
   | `OptionButton` | Any | `<input type="radio">` | Any other type |

   The verification script (Rule 19) MUST implement this table as assertions and output `WRONG TYPE` for every mismatch.
22. **Highly Specific API Naming**: Never group distinct dropdown APIs under a generic route (e.g., `/api/formulas` instead of `/api/customer-formula` and `/api/vendor-formula`). If the database stores different contexts in the same table, you MUST create explicit, separate API endpoints for each specific UI dropdown to prevent data bleed and UI filtering bugs.
23. **Mandatory Continuous Testing for ALL APIs**: Do not limit continuous testing to just "Save" APIs. As soon as you create ANY data-fetching API (for a dropdown, defaults, lookup), you MUST immediately create a scratch test script and run it. You are forbidden from wiring any API to the UI until its test script passes all assertions. The test script MUST cover all of the following categories:

   - **Helper / Cascading APIs** (e.g., Country → State → City chains): Write `scratch/test_[form]_helpers.mjs` and assert:
     1. Each individual endpoint (e.g., `/api/helpers?action=country`) returns a non-empty array.
     2. Each cascading step returns correctly filtered child records when given a valid parent code.
     3. The exact field names in the API response match the field names the React component reads. A mismatch (e.g., API returns `StateName` but React reads `label`) is a defect that must be fixed before UI wiring.

   - **Terms / TypeFlg-filtered APIs** (e.g., TERMSMASTER): Test EACH `TypeFlg` value **individually** in the script. For Party Master this means asserting separately for: `'Form'`, `'Payment Terms'`, `'Freight'`, `'Mode'`, `'Insurance'`, `'Packing & Forwarding'`, `'Delivery Instruction'`. Each must return a non-empty array. An empty array for any TypeFlg is a bug to fix immediately.

   - **Sub-Master Dropdown APIs** (Currencies, Formulas, Document Types, etc.): Test each in isolation and assert all four conditions: (a) `count > 0`, (b) the primary key field exists in each record, (c) the display name field exists in each record, (d) all string values are trimmed — no trailing spaces (trailing spaces silently break `===` comparisons in React dropdowns).

   - Print a clear per-API PASS / FAIL summary table at the end of the test script. A verbal claim of "it works" without script output is not acceptable.


24. **Mandatory Stored Procedure (SP) vs Schema vs UI Checklist**: Upon completing a form migration, you MUST create a `[FormName]_SP_Checklist.md` artifact. This checklist must explicitly map and compare three layers for every field being saved:
   - **Stored Procedure Parameter**: The exact parameter name and type expected by the legacy SQL Stored Procedure (e.g., `@PartyCode varchar(10)`).
   - **Prisma Schema**: The corresponding column name and type in `prisma/schema.prisma`.
   - **React UI Input**: The corresponding state variable and input field implemented in the Next.js `page.tsx`.
   
   You must mark each mapped field with a `[x]` to verify they align perfectly in intent and data type. If any field exists in the SP but is missing from the UI, or if data types mismatch (e.g., UI sends a String but SP expects an Int), it must be explicitly flagged and fixed before the migration can be marked complete.

25. **API Reuse and Centralization (No API Duplication)**: Do not create duplicate API endpoints for shared master data (e.g., Currency, City, Terms, Formula, Country). Before creating a new dropdown or lookup API for a form, you MUST first search the existing codebase to check if an endpoint that serves this exact master data already exists (e.g., checking if `/api/sales/masters/terms/fetch` already exists). If a suitable endpoint exists, you MUST reuse it. Only create a new API if the required data is strictly form-specific or if an existing shared endpoint genuinely cannot satisfy the requirement.

26. **Data Hydration and Join Checks**: API test scripts for helper APIs and lookup APIs MUST explicitly assert that display fields (like `description`, `name`) are hydrated with actual text (`length > 0`) when fetched from a database join. Do not accept empty strings just because they are valid JSON. If a legacy table only stores an ID or Code, your test script must verify that the corresponding description is successfully joined and returned.

27. **Legacy Data Simulation in Tests**: When writing and running test scripts for "Load" or "Fetch" APIs, you MUST test them against older, historical database rows, not just freshly created dummy data. This is critical for catching bugs where older records use values (like inactive formulas or deprecated codes) that are filtered out by default logic designed only for new records.

28. **Functional State Updates for Grids**: When updating individual properties within a React state array (like a data grid or list of items), you MUST use functional state updates (`setItems(prev => ...)`) to prevent React state batching bugs. Do not use closure-based array cloning (`const newItems = [...items]`) when making sequential property updates, as this will result in earlier updates being silently overwritten within the same render cycle.

29. **SQL Server Trailing Spaces (Anti-Empty-Result Rule)**: Legacy SQL Server databases often use fixed-length `char` columns which pad values with trailing spaces (e.g., `"Q"` becomes `"Q    "`). When writing Prisma `where` clauses for exact string matching on legacy discriminator columns (like `PertainingTo`, `TypeFlg`), you MUST use `{ startsWith: value }` instead of direct equality to prevent empty results. Furthermore, you MUST ensure that APIs querying shared master tables strictly enforce these discriminator filters. If you fetch all records unfiltered, the React UI will silently auto-fill the wrong values because identical ID codes are reused across different categories (e.g. Code `3` is both a Scrap Reason and a Form). Additionally, when filtering or mapping data on the frontend client, you MUST use `.trim()` on the string before doing an `===` comparison to prevent silent failures.

30. **Reactive Auto-Fill vs. Legacy OK Buttons (Anti-Clunk Rule)**: In VB6, auto-filling related fields (like Terms & Conditions based on a Customer) often required the user to click an "OK" button to lock the header and trigger a fetch subroutine (e.g., `cmdOk_Click` calling `TermsDetails`). When migrating to Next.js, you MUST NOT implement this clunky two-step workflow. Instead, make the UI fully reactive: attach a `useEffect` hook to the primary dropdown (e.g., Customer) that instantly fetches and auto-fills all dependent defaults (like FormCode, PackingCode, etc.) without requiring a manual "OK" confirmation.

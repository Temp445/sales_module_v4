# New Analysis – Frontend Inputs, Controls, Database Usage & Business Logic

## 1. Total Frontend Inputs & Controls

| Control Name | Type | Caption / Data Type | Key Properties / Usage | Required | Purpose / Interaction | Data Source / Bindings |
|---|---|---|---|---|---|---|
| `cmbType` | ComboBox | String | Dropdown List. Items: "Country", "City" | Yes | Select whether to add/edit a Country or a City | `CountryCityMaster.frx` |
| `txtCountry` | TextBox | String | MaxLength: 50. Tag stores `CountryCode` | Yes | Enters Country name. Press F2 or DblClick for Help to search existing Countries. | Populated by `FetchValue_New` on F2 (`COM_CountryCity_Master` where `ParentCode=0`) |
| `txtCity` | TextBox | String | MaxLength: 50, Enabled: False (initially). Tag stores `CityCode` | Yes (if City) | Enters City name. Press F2 or DblClick for Help to search existing Cities. | Populated by `FetchValue_New` on F2 (`COM_CountryCity_Master` where `ParentCode = Val(txtCountry.Tag)`) |
| `txtCode` | TextBox | String | MaxLength: 10, Visible: False | No | Hidden control | None |
| `cmdSave` | CommandButton | N/A | Enabled: False (initially) | N/A | Click to save data (Insert or Update based on existence of ID in Tag) | N/A |
| `cmdClear` | CommandButton | N/A | N/A | N/A | Clears the form inputs | N/A |
| `cmdExit` | CommandButton | N/A | N/A | N/A | Exits the form | N/A |

**Total UI inputs and controls: 7** (including hidden and buttons)

## 2. Database Tables Used

| Table Name | Purpose | Columns Used | Operations |
|---|---|---|---|
| `COM_CountryCity_Master` | Stores both Country and City data (hierarchical with ParentCode) | `CountryCode`, `CountryName`, `ParentCode` | SELECT, EXECUTE (SP inserts/updates) |

## 3. Data Fetching & Database Operations

**1. Fetching Existing Countries (Lookup - F2 on txtCountry)**
- **Table:** `COM_CountryCity_Master`
- **Columns Selected:** `CountryName`, `CountryCode`
- **Conditions:** `ParentCode=0`
- **UI Control:** `txtCountry` gets `CountryName` (value 1), `txtCountry.Tag` gets `CountryCode` (value 2).

**2. Fetching Existing Cities (Lookup - F2 on txtCity)**
- **Table:** `COM_CountryCity_Master`
- **Columns Selected:** `CountryName`, `CountryCode`
- **Conditions:** `ParentCode=Val(txtCountry.Tag)` (Only valid if a Country is selected).
- **UI Control:** `txtCity` gets `CountryName`, `txtCity.Tag` gets `CountryCode`.

**3. Validating Duplicates on Insert**
- **Operation:** SELECT
- **Tables:** `COM_CountryCity_Master`
- **For Country:** `select CountryName from COM_CountryCity_Master where ParentCode=0 and CountryName = [txtCountry.Text]`
- **For City:** `select CountryName from COM_CountryCity_Master where ParentCode=[txtCountry.Tag] and CountryCode=[txtCity.Tag] and CountryName = [txtCountry.Text]` (Wait, the VB6 code has `and CountryName = 'Trim$(txtCountry)'`, which is a bug in VB6 code for city duplicate check, as it checks CountryName against `txtCountry` instead of `txtCity`, but in NextJS we should check against the new city name).

**4. Save Operation (Stored Procedures)**
The form relies on two Stored Procedures that receive the following parameters:
- `@CountryCode` (Integer)
- `@CountryName` (VarChar 50)
- `@ParentCode` (Integer: 0 for Country, Parent ID for City)
- `@UserID` (VarChar 10, from `gUserID`)

- **SP: `COM_Ins_CountryCity_Master` (INSERT)**
  - Triggers if `txtCountry.Tag` (for Country) or `txtCity.Tag` (for City) is empty/0.
  - Generates new `CountryCode` via inline query `Select isNull(Max(CountryCode), 0)+1 from COM_CountryCity_Master Where ParentCode = 0`.
- **SP: `COM_Up_CountryCity_Master` (UPDATE)**
  - Triggers if `txtCountry.Tag` > 0 (for Country) or `txtCountry.Tag` > 0 and `txtCity.Tag` > 0 (for City).

## 4. Detailed Business Rules, Conditions & Procedures

### Form Behaviors & Conditional Logic
- **`cmbType_Click` (Type Selection Toggle):**
  - If `cmbType` = "Country" (or changing): resets and disables `txtCity`, sets `txtCountry.Locked = False`.
  - If `cmbType` = "City": enables `txtCity`, `lblCity.Enabled = True`, sets `txtCountry.Locked = True` (meaning country can only be selected via lookup, not typed manually).
- **`ValidateData` (Function, Private):**
  - Validation: If `txtCountry.Text` is empty, shows "Enter the Country Name" (if Country) or "Enter the City Name" (if City) and returns False.
  - Dependency Check: If Type is "City", it validates `Val(txtCountry.Tag) > 0`. If false, shows "You can create City only for already created Country".
- **`CheckSave`:**
  - Logic to enable `cmdSave` button.
  - If Type is Country: enabled if `txtCountry` <> "".
  - If Type is City: enabled if `txtCountry` <> "" AND `txtCity` <> "".

### Transaction Handling & Global Helpers
- Transactions: Managed using `DBStartTransaction`, `DBCommitTransaction`, and `DBRollBackWithMessage`.
- Helpers: `Form_Clear` (resets form controls), `FetchValue_New` (opens standard lookup grid), `RSOpen` (fetches recordset), `commClear`, `ErrMsgShow`, `sShowMessage`.

## 5. Form Events & Workflow

1. **Form Open:** User selects whether they are managing a "Country" or a "City" from the dropdown `cmbType`.
2. **If Country:**
   - User types a new Country name into `txtCountry` (to create new) OR presses F2 in `txtCountry` to select an existing one (to edit).
   - If selected existing, `txtCountry.Tag` is populated with `CountryCode`.
   - Click Save. Triggers `cmdSave_Click`.
   - Checks if updating or inserting based on `Val(txtCountry.Tag)`. Validates duplicate name on insert. Calls `COM_Ins_...` or `COM_Up_...`.
3. **If City:**
   - User MUST first press F2 in `txtCountry` to select a Parent Country. `txtCountry.Tag` gets the `CountryCode`.
   - User types a new City name into `txtCity` (to create new) OR presses F2 in `txtCity` to select an existing city under that country (to edit).
   - If selected existing, `txtCity.Tag` gets the `CountryCode` (which acts as City ID).
   - Click Save. Validation checks both parent country and city.
   - Calls `COM_Ins_...` or `COM_Up_...` passing `ParentCode = txtCountry.Tag`.

## 6. Input → Business Logic → Database Mapping

- `cmbType` (Dropdown) → Toggle UI state → Determines whether saving Country (ParentCode=0) or City (ParentCode=Country ID).
- `txtCountry` (Text) → Validation (Not Empty) → Stored as `@CountryName` (if Country). Looked up for `CountryCode` (ParentCode) if City.
- `txtCity` (Text) → Validation (Not Empty) → Stored as `@CountryName` (if City).
- `txtCountry.Tag` (Hidden ID) → Determines Insert vs Update (for Country). Passed as `@ParentCode` (for City).
- `txtCity.Tag` (Hidden ID) → Determines Insert vs Update (for City). Passed as `@CountryCode`.

## 7. Avoid Duplication
- Tables: `COM_CountryCity_Master`
- Forms: Single logic form handling both hierarchy levels.

## 8. VB6 Source References
- `txtCountry_DblClick` -> line 237, `CountryCityMaster.frm`
- `cmdSave_Click` -> line 160, `CountryCityMaster.frm`
- `cmbType_Click` -> line 218, `CountryCityMaster.frm`

## 9. Migration Requirements
- **Next.js Route:** `/sales/masters/country-city`
- **Prisma Schema:** Map `COM_CountryCity_Master` to a Prisma model (already exists or needs mapping).
- **APIs required:**
  - GET `/api/sales/masters/country-city/fetch-countries`
  - GET `/api/sales/masters/country-city/fetch-cities?parentCode=[id]`
  - POST `/api/sales/masters/country-city/save` (Handles both insert and update based on provided ID, calling the respective SPs or using Prisma directly if SP logic is simple).
- **UI Details:** Use `SearchableSelect` for both Country and City fields. The ComboBox for Type should control the visibility/editability of the City selector. 

### Step 3: Independent Verification & Final Summary

- **MATCH:** `cmbType` items matched with FRX extraction.
- **MATCH:** Lookup queries correctly captured.
- **MATCH:** Stored procedures identified.
- **MISSING:** The exact `@ModuleName` audit parameter isn't seen in VB6, but per Rule 6, Next.js implementation must include standard audit params (e.g. `UserId`, `BranchId`, `CompId`) if using EXECUTE. The VB6 form passes `gUserID`.

**Summary:**
- **Total UI Inputs / Controls:** 7
- **Unique Database Tables Used:** 1 (`COM_CountryCity_Master`)
- **Database Operations:** SELECT (lookup, duplicate check, max id), EXECUTE (insert/update SP).
- **Migration Risks:** Low. Standard hierarchical master form. The duplicate check logic in VB6 has a subtle bug for City (checks `CountryName = txtCountry` instead of `txtCity`), which should be corrected or replicated exactly.
- **Verification Status:** MATCH.

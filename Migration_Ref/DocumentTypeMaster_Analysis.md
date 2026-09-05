# DocumentTypeMaster Analysis

## 1. UI Components (from Extracted JSON & VB6 Form)
- **`txttype`**: TextBox, max length 3. Mapped to `DocumentType`. Converted to uppercase on input.
- **`txtdesc`**: TextBox, max length 30. Mapped to `Description`. Converted to uppercase on input.
- **`txtstart`**: TextBox, max length 5. Mapped to `StartingNo`. Only allows numeric input.
- **`cmbPertaining`**: ComboBox (Dropdown). Mapped to `PertainingTo`. Contains 20 hardcoded string values like "Sales Enquiry", "Sales Order", etc. 
- **`sprdoctype`**: Data Grid (AceSpread). Displays existing document types for the current Branch and Company.
- **`cmdSave`**: Button to trigger save/validation.
- **`cmdClear`**: Button to clear inputs.
- **`cmdExit`**: Button to close form.

## 2. Business Rules & Logic
- **Validation**: 
  - `txttype`, `txtdesc`, and `txtstart` cannot be empty.
  - Must check if the `DocumentType` already exists for the current `Branchid` and `Compid` before saving. If it does, throw "Document Type already exists" and prevent save.
- **PertainingTo Mapping**:
  - The UI drop-down displays full text (e.g., "Sales Enquiry").
  - The database stores a 1-2 character code (e.g., "E").
  - `PerToName_Code` translates the dropdown string to the DB code before executing the save stored procedure.
  - `PerToCode_Name` translates the DB code back to the dropdown string when loading the data grid.

## 3. Database Interactions
- **Load Grid (`LoadSpr`)**:
  - Queries `COM_DOCUMENTTYPE_MASTER`.
  - Filter: `Branchid=gBranchID AND Compid=gCompID AND ModuleName='SAL'`.
- **Duplicate Check (`ValidateData`)**:
  - Queries `COM_DOCUMENTTYPE_MASTER`.
  - Filter: `DocumentType = [txttype] AND Branchid=gBranchID AND Compid=gCompID`.
- **Save Operation (`cmdSave_Click`)**:
  - Calls `COM_INS_DOCUMENTTYPE_MASTER` Stored Procedure.
  - Parameters: `@DocumentType`, `@Description`, `@StartingNo`, `@PertainingTo` (as code), `@Module` (hardcoded 'SAL'), `@Status` (hardcoded True).
  - Also requires common parameters (`@Userid`, `@Branchid`, `@Compid`) per ERP standards.

## 4. Potential Migration Risks
- **Dropdown State Matching**: The hardcoded list of 20 "PertainingTo" strings must exactly match the Next.js dropdown values, or the translation functions (`PerToName_Code`) will fail to resolve the code, inserting empty values into the database.
- **Global Variables**: We must dynamically inject `Branchid` and `Compid` from the Next.js session using `getSession()`, never hardcoding fallback values like "00" or "02".
- **Stored Procedure Error**: As seen with Country/City, the legacy SP `COM_INS_DOCUMENTTYPE_MASTER` might throw a "Too many arguments" error in Next.js Prisma if the exact signature is not matched. We must test the SP directly in a scratch script before writing the API route. If the SP fails, we will fallback to native Prisma ORM `create()` / `update()` as per established patterns.

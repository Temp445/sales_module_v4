# TermsMaster Analysis

## 1. UI Components (from Extracted JSON & VB6 Form)
- **`cmbTerms`**: ComboBox (Dropdown). Mapped to `TypeFlg`. Contains 8 hardcoded string values: "Form", "Packing & Forwarding", "Delivery Instruction", "Insurance", "Freight", "Mode", "Payment Terms", "Contract Review". 
- **`cmbStatus`**: ComboBox. Values: "Active", "Inactive". Mapped to `Status`.
- **`tbDesc`**: TextBox, max length 100. Mapped to `Description`. Converted to uppercase on input.
- **Hidden `Code` Field**: The VB6 form uses `tbDesc.Tag` to secretly store the primary key `Code`.
- **`cbsave`, `cbclear`, `cbexit`**: Standard command buttons.
- **Lookup Grid**: Double clicking `tbDesc` (or F2) opens a lookup grid to select existing records.

## 2. Business Rules & Logic
- **Validation**: 
  - `tbDesc` (Description) cannot be empty.
  - Must check if the `Description` already exists for the selected `cmbTerms` (TypeFlg), `Branchid`, and `Compid`. If it does, throw "Duplicate Entry not allowed".
- **Auto Code Generation (`GenerateCode`)**:
  - The `Code` field is NOT entered by the user. When a user selects a term type from `cmbTerms`, the system queries the DB: `isnull(max(convert(int,Code)),0) + 1` for that specific `TypeFlg`.
  - Wait! Since it uses `convert(int,Code)`, the `Code` field is a `VarChar` but holds numeric sequential strings ("1", "2", "3").

## 3. Database Interactions
- **Table Name**: `TERMSMASTER`. Wait, is it `TERMSMASTER` or `COM_TermsMaster`? The VB6 code queries `TERMSMASTER` directly. In our Prisma schema, we saw `COM_TermsMaster` and `TermsMaster`. Wait, let me double check which one to use. I should use `TermsMaster` if the VB6 query explicitly says `select Code from TERMSMASTER`.
- **Duplicate Description Check**:
  - `select Code from TERMSMASTER where typeflg = cmbTerms and Description = tbDesc and Branchid=gBranchID and Compid=gCompID`
- **Save/Update**:
  - The script checks if the specific `Code` exists: `select Code from TERMSMASTER where typeflg = cmbTerms and Code = tbDesc.Tag and Branchid=gBranchID and Compid=gCompID`.
  - If NOT EOF, it calls `upTERMSMASTER` SP.
  - If EOF, it calls `insTERMSMASTER` SP.
  - Parameters: `@TypeFlg`, `@Code`, `@Description`, `@Status` (Boolean), + `@Userid`, `@Branchid`, `@Compid`.

## 4. Potential Migration Risks
- **Table Ambiguity**: The Prisma schema contains both `COM_TermsMaster` and `TermsMaster`. The VB6 explicitly uses `TERMSMASTER`. We must test the DB to ensure we query the correct one.
- **Max Code Logic**: Generating the max code securely in a web app requires handling concurrency (e.g., two users adding a term at the exact same time). Since Prisma doesn't natively do `isnull(max(convert(int,Code)),0)`, we will use Prisma's `aggregate` or raw query to fetch the max code.
- **Stored Procedure Errors**: As seen with previous forms, `insTERMSMASTER` might fail with transaction errors (`3902 The COMMIT TRANSACTION request has no corresponding BEGIN TRANSACTION`). We must test the SP in a scratch script. If it fails, we will bypass it and use native Prisma ORM `create()` / `update()`.
- **Dropdown List**: The dropdown options must perfectly match the VB6 `.frx` output.

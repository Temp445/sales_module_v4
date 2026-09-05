# Customer Item Master - Deep Analysis

## 1. Form Overview
- **Name:** `frmCustomerItemMaster`
- **Purpose:** Manages the item mapping and pricing specific to individual customers, allowing sales to link an internal item code to a customer's specific part number, formula, rate, and validity dates.

## 2. Controls & Bindings (Zero-Assumption Mapping)

### Header Controls
| VB6 Control | Type | Data Type | Purpose / Mapping |
| :--- | :--- | :--- | :--- |
| `txtcustcode` | `TextBox (Locked)` | `SearchableSelect` | Stores Customer Name. Tag stores `PartyCode` (CustomerCode). Triggers `FAS_PARTY_MASTER` lookup. |
| `txtDefaultFormula` | `TextBox (Locked)` | `SearchableSelect` | Stores Default Formula Name. Tag stores `FormulaID`. Triggers `COM_FormulaMast_Hdr` lookup. |

### Grid Controls (`spItem`)
| Col | VB6 Column Name | Data Type | Required | Business Rule |
| :--- | :--- | :--- | :--- | :--- |
| 1 | `Item Code` | `SearchableSelect` | Yes | Triggers F2 lookup to `SAL_Customer_ItemMaster` or `BOM_Material_Master`. |
| 2 | `Description` | `Text` | No | Auto-filled from Item lookup. |
| 3 | `Model` | `Text` | No | Auto-filled from Item lookup. |
| 4 | `Drawing Issued No` | `Text` | No | Auto-filled from Item lookup. |
| 5 | `Customer Part No` | `Text` | Yes | Manually entered or updated. |
| 6 | `Rate` | `Number (Double)` | Yes | Must be > 0. |
| 7 | `Assessable Rate` | `Number (Double)` | Yes | Must be > 0. |
| 8 | `Valid From` | `Date` | Yes | Format: `yyyy-MM-dd`. |
| 9 | `Valid To` | `Date` | Yes | Format: `yyyy-MM-dd`. Must be > `Valid From`. |
| 10 | `Formula Code` | `Hidden (Int)` | Yes | Backend ID for formula. |
| 11 | `Formula Name` | `SearchableSelect` | Yes | F2 lookup to `COM_FormulaMast_Hdr`. |
| 12 | `QC Check Required` | `ComboBox (Bool)`| Yes | `Yes` -> 1, `No` -> 0. |
| 13 | `Status` | `ComboBox (Bool)`| Yes | `Active` -> 1, `Inactive` -> 0. |

## 3. Database Interactions
- **Fetch Lookups:**
  - `FAS_PARTY_MASTER` (Customers)
  - `COM_FormulaMast_Hdr` (Formulas)
  - `BOM_Material_Master` (Items)
- **Save Operation:**
  - Uses `SAL_INS_CUSTOMER_ITEMMASTER` (Stored Procedure).
  - Target table: `SAL_CUSTOMER_ITEMMASTER`.

## 4. Migration Risks & Enforcement
- The SP `SAL_INS_CUSTOMER_ITEMMASTER` is encrypted. However, because the `.frm` explicitly maps the ADO parameter array (with exact SQL types like `adInteger`, `adDouble`, `adDBTimeStamp`), we can safely infer the Prisma contract or just use `prisma.sAL_CUSTOMER_ITEMMASTER.create()`.
- Rule 8 (Value vs Label) must be strictly enforced for `CustomerCode`, `ItemCode`, and `FormulaCode`.
- Rule 13 (Helper API Tests) must assert the lookups return correct primary keys.

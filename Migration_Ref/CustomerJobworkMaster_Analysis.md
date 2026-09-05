# Customer Jobwork Master - Deep Analysis

## 1. Form Overview
- **Name:** `frmCustomerJobworkMaster`
- **Purpose:** Manages the mapping of jobwork items, defining the relationship between received materials and sent finished goods/components for a specific customer. It manages rates, conversion factors, and validities.

## 2. Controls & Bindings (Zero-Assumption Mapping)

### Header Controls
| VB6 Control | Type | Data Type | Purpose / Mapping |
| :--- | :--- | :--- | :--- |
| `txtVendCode` | `TextBox (Locked)` | `SearchableSelect` | Stores Customer Name. Tag stores `CustomerCode`. Triggers `FAS_PARTY_MASTER` lookup. |
| `txtDefaultFormula` | `TextBox (Locked)` | `SearchableSelect` | Stores Default Formula Name. Tag stores `FormulaID`. Triggers `COM_FormulaMast_Hdr` lookup. |

### Grid Controls (`sprItem` - 20 Columns)
| Col | VB6 Column Name | Data Type | Required | Business Rule |
| :--- | :--- | :--- | :--- | :--- |
| 1 | `mFirstCol` | `Hidden` | Yes | Internal row tracking identifier. |
| 2 | `mrecdItem` | `SearchableSelect` | Yes | Received Item Code. |
| 3 | `mRecdDesc` | `Text` | No | Received Item Description. |
| 4 | `mRecdUOM` | `Text` | No | Received UOM. |
| 5 | `msentItem` | `SearchableSelect` | Yes | Sent Item Code (Parent Item). Triggers structure lookup. |
| 6 | `mSentDesc` | `Text` | No | Sent Item Description. |
| 7 | `mPartNo` | `Text` | Yes | Customer Part No. |
| 8 | `mSentUOM` | `Text` | No | Sent UOM. |
| 9 | `mRatePer` | `Number` | Yes | Rate Per Quantity. |
| 10 | `mRate` | `Number` | Yes | Jobwork Rate. |
| 11 | `mAssRate` | `Number` | Yes | Assessable Rate. |
| 12 | `mValidFrom` | `Date` | Yes | Format: `yyyy-MM-dd`. |
| 13 | `mValidTo` | `Date` | Yes | Format: `yyyy-MM-dd`. Must be > `Valid From`. |
| 14 | `mQCCheck` | `ComboBox (Bool)`| Yes | `Yes` -> 1, `No` -> 0. |
| 15 | `mFrmCode` | `Hidden (Int)` | Yes | Backend ID for formula. |
| 16 | `mFrmName` | `SearchableSelect` | Yes | F2 lookup to `COM_FormulaMast_Hdr`. |
| 17 | `mStatus` | `ComboBox (Bool)`| Yes | `Active` -> 1, `Inactive` -> 0. |
| 18 | `mRelation` | `Hidden` | Yes | Tracks parent/child relation. |
| 19 | `mItemGroupKey` | `Hidden` | Yes | Used for grouping hierarchy via `fn_ItemGroupKey`. |
| 20 | `mRecdItemCheck`| `Hidden` | Yes | Internal state for validation. |

## 3. Database Interactions
- **Fetch Lookups:**
  - `FAS_PARTY_MASTER` (Customers)
  - `COM_FormulaMast_Hdr` (Formulas)
  - `BOM_MPS_Master` and `BOM_STRUCTURE_HDR` (Process and Item structures)
- **Save Operation:**
  - Target Prisma Table: `SAL_Jobwork_ItemMaster` (Maps directly to DB schema).
  - Grouping function logic requires verifying `ItemGroupKey` generation.

## 4. Migration Risks & Enforcement
- The VB6 logic heavily relies on nested lookups between `BOM_MPS_Master` and `BOM_Structure_Dtl`. We must accurately recreate this logic in an API `/api/sales/jobwork-master/fetch-structure` instead of doing it client-side.
- The stored procedures for insertion might be encrypted, so we must rely on Prisma ORM mapping based on `SAL_Jobwork_ItemMaster` schema and explicitly documented ADO parameters.
- Rule 8 (Value vs Label) must be strictly enforced. All hidden IDs (`CustomerCode`, `FormulaCode`, `ItemCode`) must be tracked as `value` in `SearchableSelect`.

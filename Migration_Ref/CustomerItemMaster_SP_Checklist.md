# Customer Item Master SP Checklist

| Field | SP Parameter (Inferred from ADO) | Prisma Schema Column | UI State Variable | Status |
| :--- | :--- | :--- | :--- | :--- |
| **CustomerCode** | `@CustomerCode [int]` | `CustomerCode Int` | `customerCode` (from txtcustcode) | [x] |
| **ItemCode** | `@ItemCode [varchar(20)]` | `ItemCode String` | `itemCode` (from spItem col 1) | [x] |
| **Model** | `@Model [varchar(20)]` | `Model String` | `model` (from spItem col 3) | [x] |
| **IssuedNo** | `@IssuedNo [varchar(10)]` | `IssuedNo String` | `issuedNo` (from spItem col 4) | [x] |
| **PartNo** | `@PartNo [varchar(20)]` | `PartNo String?` | `partNo` (from spItem col 5) | [x] |
| **Rate** | `@Rate [float]` | `Rate Float` | `rate` (from spItem col 6) | [x] |
| **AssessableRate**| `@AssessableRate [float]` | `AssessableRate Float` | `assessableRate` (from spItem col 7) | [x] |
| **ValidFrom** | `@ValidFrom [datetime]` | `ValidFrom DateTime?`| `validFrom` (from spItem col 8) | [x] |
| **ValidTo** | `@ValidTo [datetime]` | `ValidTo DateTime?` | `validTo` (from spItem col 9) | [x] |
| **FormulaCode** | `@FormulaCode [int]` | `FormulaCode Int` | `formulaCode` (from spItem col 10) | [x] |
| **QCCheck** | `@QCCheck [bit]` | `QCCheck Boolean` | `qcCheck` (from spItem col 12) | [x] |
| **Status** | `@Status [bit]` | `Status Boolean` | `status` (from spItem col 13) | [x] |
| **Userid** | `@Userid [varchar(10)]` | `Userid String?` | (Derived from context) | [x] |
| **Branchid** | `@Branchid [varchar(3)]` | `Branchid String` | (Derived from context) | [x] |
| **Compid** | `@Compid [varchar(3)]` | `Compid String` | (Derived from context) | [x] |

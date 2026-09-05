# SalesQuotation SP to UI Checklist

## Header (SAL_INS_QUOTATION_HDR)

| SP Parameter | Data Type | Prisma/DB equivalent | React State / Input | Status |
| :--- | :--- | :--- | :--- | :--- |
| `@QuotationType` | `Char(3)` | `QuotationType` | `quoType` (cmbQuoType) | [x] |
| `@QuotationNo` | `Int` | `QuotationNo` | `quoNo` (txtQuoNo) | [x] |
| `@QuotationDate` | `DBTimeStamp` | `QuotationDate` | `quoDate` (dpQuoDate) | [x] |
| `@RevisionNo` | `Int` | `RevisionNo` | `0` (hardcoded initially) | [x] |
| `@RevisionDate` | `DBTimeStamp` | `RevisionDate` | `new Date()` (auto) | [x] |
| `@OrderSource` | `Char(3)` | `OrderSource` | `orderSource` (optEnq/optdirect) | [x] |
| `@RefType` | `Char(3)` | `RefType` | `enqType` (txtEnqType) | [x] |
| `@RefNo` | `Int` | `RefNo` | `enqNo` (txtEnqNo) | [x] |
| `@RefDate` | `DBTimeStamp` | `RefDate` | `enqDate` (dtpEnqDate) | [x] |
| `@CustomerCode` | `Int` | `CustomerCode` | `customerCode` (txtcustcd.Tag) | [x] |
| `@CustomerName` | `VarChar(50)` | `CustomerName` | `customerName` (txtcustcd.Text) | [x] |
| `@ValidFrom` | `DBTimeStamp` | `ValidFrom` | `validFrom` (dpvalidfrom) | [x] |
| `@ValidTo` | `DBTimeStamp` | `ValidTo` | `validTo` (dpvalidto) | [x] |
| `@FormulaCode` | `Int` | `FormulaCode` | `formulaCode` (txtFormula.Tag) | [x] |
| `@Remarks` | `VarChar(150)` | `Remarks` | `remarks` (txtremarks) | [x] |
| `@FormCode` | `Int` | `FormCode` | `formCode` (txtForm.Tag) | [x] |
| `@PackingCode` | `Int` | `PackingCode` | `packingCode` (txtPack.Tag) | [x] |
| `@DeliveryCode` | `Int` | `DeliveryCode` | `DeliveryCode` (txtDelIns.Tag)| [x] |
| `@PaymentCode` | `Int` | `PaymentCode` | `paymentCode` (txtTerms.Tag) | [x] |
| `@InsuranceCode` | `Int` | `InsuranceCode` | `insuranceCode` (txtInsurance.Tag)| [x] |
| `@ModeCode` | `Int` | `ModeCode` | `modeCode` (txtMode.Tag) | [x] |
| `@FreightCode` | `Int` | `FreightCode` | `freightCode` (txtFreight.Tag) | [x] |
| `@TotalAmount` | `Double` | `TotalAmount` | `grandTotal` (txtGrandTotal) | [x] |

## Details (SAL_INS_QUOTATION_DTL)

| SP Parameter | Data Type | Prisma/DB equivalent | React State / Input in Grid | Status |
| :--- | :--- | :--- | :--- | :--- |
| `@QuotationType` | `Char(3)` | `QuotationType` | Header | [x] |
| `@QuotationNo` | `Int` | `QuotationNo` | Header | [x] |
| `@QuotationDate` | `DBTimeStamp` | `QuotationDate` | Header | [x] |
| `@RevisionNo` | `Int` | `RevisionNo` | `0` | [x] |
| `@ItemCode` | `VarChar(20)` | `ItemCode` | `itemCode` (col 1) | [x] |
| `@ItemDescription`| `VarChar(50)` | `ItemDescription`| `itemDescription` (col 2) | [x] |
| `@UOM` | `VarChar(3)` | `UOM` | `uom` (col 3) | [x] |
| `@Qty` | `Double` | `Qty` | `qty` (col 4) | [x] |
| `@Rate` | `Double` | `Rate` | `rate` (col 5) | [x] |
| `@AssessableRate` | `Double` | `AssessableRate` | `assessableRate` (col 6) | [x] |

*Note: All lookup textboxes (txtFormula, txtForm, etc.) will be converted to SearchableSelect components, separating value (ID) from label (Description).*

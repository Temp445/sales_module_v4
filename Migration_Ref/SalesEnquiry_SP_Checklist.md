# Sales Enquiry SP Checklist

## SAL_INS_ENQUIRY_HDR

| SP Parameter | SP Type | Prisma Schema Column | Prisma Type | React UI State | React Component | Matched |
| --- | --- | --- | --- | --- | --- | --- |
| `@EnquiryType` | `varchar(3)` | `EnquiryType` | `String` | `formData.enquiryType` | `<SearchableSelect>` | `[x]` |
| `@EnquiryNo` | `int` | `EnquiryNo` | `Int` | `formData.enquiryNo` | `<input type="number">` | `[x]` |
| `@EnquiryDate` | `datetime` | `EnquiryDate` | `DateTime` | `formData.enquiryDate` | `<input type="date">` | `[x]` |
| `@RevisionNo` | `int` | `RevisionNo` | `Int?` | `0` (Hardcoded per VB) | (None) | `[x]` |
| `@RevisionDate` | `datetime` | `RevisionDate` | `DateTime?` | `currentDate` (Auto) | (None) | `[x]` |
| `@ReferenceNo` | `varchar(15)` | `ReferenceNo` | `String?` | `formData.referenceNo` | `<input type="text">` | `[x]` |
| `@ReferenceDate` | `datetime` | `ReferenceDate` | `DateTime?` | `formData.referenceDate` | `<input type="date">` | `[x]` |
| `@CustomerType` | `char(1)` | `CustomerType` | `String` | `formData.customerType` | `<SearchableSelect>` | `[x]` |
| `@CustomerCode` | `int` | `CustomerCode` | `Int` | `formData.customerCode` | `<SearchableSelect>` | `[x]` |
| `@CustomerName` | `varchar(50)` | `CustomerName` | `String` | `formData.customerName` | (Derived) | `[x]` |
| `@Remarks` | `varchar(150)` | `Remarks` | `String?` | `formData.remarks` | `<input type="text">` | `[x]` |
| `@Userid` | `varchar(10)` | `Userid` | `String` | `session.userId` | (Session) | `[x]` |
| `@Branchid` | `varchar(3)` | `Branchid` | `String` | `session.branchId` | (Session) | `[x]` |
| `@Compid` | `varchar(3)` | `Compid` | `String` | `session.compId` | (Session) | `[x]` |
| `@Finyear` | `varchar(20)` | `Finyear` | `String?` | `session.finYear` | (Session) | `[x]` |

## SAL_INS_ENQUIRY_DTL

| SP Parameter | SP Type | Prisma Schema Column | Prisma Type | React UI State (Items) | React Component | Matched |
| --- | --- | --- | --- | --- | --- | --- |
| `@EnquiryType` | `varchar(3)` | `EnquiryType` | `String` | `formData.enquiryType` | (Header) | `[x]` |
| `@EnquiryNo` | `int` | `EnquiryNo` | `Int` | `formData.enquiryNo` | (Header) | `[x]` |
| `@EnquiryDate` | `datetime` | `EnquiryDate` | `DateTime` | `formData.enquiryDate` | (Header) | `[x]` |
| `@RevisionNo` | `int` | `RevisionNo` | `Int?` | `0` (Hardcoded) | (Header) | `[x]` |
| `@ItemType` | `char(1)` | `ItemType` | `String?` | `item.itemType` | `<td><select>` | `[x]` |
| `@ItemCode` | `varchar(20)` | `ItemCode` | `String` | `item.itemCode` | `<td><SearchableSelect>` | `[x]` |
| `@ItemDescription`| `varchar(50)` | `ItemDescription` | `String?` | `item.itemDesc` | `<td><input>` | `[x]` |
| `@UOM` | `varchar(3)` | `UOM` | `String?` | `item.uom` | `<td><SearchableSelect>` | `[x]` |
| `@Qty` | `float` | `Qty` | `Float?` | `item.qty` | `<td><input type="number">`| `[x]` |
| `@Volume` | `float` | `Volume` | `Float?` | `item.tarVolme` | `<td><input type="number">`| `[x]` |
| `@VolumePer` | `varchar(15)` | `VolumePer` | `String?` | `item.volmeper` | `<td><input>` | `[x]` |
| `@Rate` | `float` | `Rate` | `Float?` | `item.rate` | `<td><input type="number">`| `[x]` |
| `@AssessableRate` | `float` | `AssessableRate` | `Float?` | `item.assRate` | `<td><input type="number">`| `[x]` |
| `@TargetDate` | `datetime` | `TargetDate` | `DateTime?` | `item.targtDate` | `<td><input type="date">`| `[x]` |
| `@Specification` | `varchar(30)` | `Specification` | `String?` | `item.speci` | `<td><input>` | `[x]` |
| `@Userid` | `varchar(10)` | `Userid` | `String` | `session.userId` | (Session) | `[x]` |
| `@Branchid` | `varchar(3)` | `Branchid` | `String` | `session.branchId` | (Session) | `[x]` |
| `@Compid` | `varchar(3)` | `Compid` | `String` | `session.compId` | (Session) | `[x]` |
| `@Finyear` | `varchar(20)` | `Finyear` | `String?` | `session.finYear` | (Session) | `[x]` |

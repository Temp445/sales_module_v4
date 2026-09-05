const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function testSave() {
  try {
    const res = await prisma.$transaction(async (tx) => {
      // 1. Save HDR
      await tx.$executeRawUnsafe(`
        EXEC SAL_INS_ENQUIRY_HDR 
          @EnquiryType = 'SE',
          @EnquiryNo = 9999,
          @EnquiryDate = '2026-09-04',
          @RevisionNo = 0,
          @RevisionDate = '2026-09-04',
          @ReferenceNo = 'TEST001',
          @ReferenceDate = '2026-09-04',
          @CustomerType = 'E',
          @CustomerCode = 11,
          @CustomerName = 'TEST CUST',
          @Remarks = 'Test remarks',
          @Userid = 'ADMIN',
          @Branchid = '00',
          @Compid = '02',
          @Finyear = '04-2026 03-2027'
      `);
      
      // 2. Save DTL
      await tx.$executeRawUnsafe(`
        EXEC SAL_INS_ENQUIRY_DTL
          @EnquiryType = 'SE',
          @EnquiryNo = 9999,
          @EnquiryDate = '2026-09-04',
          @RevisionNo = 0,
          @ItemType = 'E',
          @ItemCode = '2U070770Z0',
          @ItemDescription = 'JET HOLDER',
          @UOM = 'NOS',
          @Qty = 10,
          @Volume = 0,
          @VolumePer = '',
          @Rate = 5,
          @AssessableRate = 5,
          @TargetDate = '2026-09-05',
          @Specification = 'TEST SPEC',
          @Userid = 'ADMIN',
          @Branchid = '00',
          @Compid = '02',
          @Finyear = '04-2026 03-2027'
      `);
      return "SUCCESS";
    });
    console.log("Transaction Result:", res);
  } catch (error) {
    console.error("Save Test Failed:", error.message);
  } finally {
    // Cleanup for test repeatability
    await prisma.$executeRawUnsafe(`DELETE FROM SAL_ENQUIRY_DTL WHERE EnquiryNo = 9999`);
    await prisma.$executeRawUnsafe(`DELETE FROM SAL_ENQUIRY_HDR WHERE EnquiryNo = 9999`);
    await prisma.$disconnect();
  }
}

testSave();

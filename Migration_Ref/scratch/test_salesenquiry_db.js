const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function testDB() {
  try {
    const resHdr = await prisma.$queryRawUnsafe(`SELECT TOP 1 * FROM SAL_ENQUIRY_HDR`);
    console.log("SAL_ENQUIRY_HDR Row:", resHdr[0]);
    
    if (resHdr[0]) {
      const resDtl = await prisma.$queryRawUnsafe(`SELECT TOP 1 * FROM SAL_ENQUIRY_DTL WHERE EnquiryNo = ${resHdr[0].EnquiryNo}`);
      console.log("SAL_ENQUIRY_DTL Row:", resDtl[0]);
    }
  } catch(e) {
    console.error("DB Test Failed:", e.message);
  } finally {
    await prisma.$disconnect();
  }
}
testDB();

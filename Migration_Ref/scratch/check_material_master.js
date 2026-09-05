const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function checkView() {
  try {
    const res = await prisma.$queryRawUnsafe(`SELECT TOP 1 * FROM MATERIAL_MASTER`);
    console.log("MATERIAL_MASTER exists. Columns:", Object.keys(res[0] || {}));
  } catch(e) {
    console.log("Error querying MATERIAL_MASTER:", e.message);
  } finally {
    await prisma.$disconnect();
  }
}
checkView();

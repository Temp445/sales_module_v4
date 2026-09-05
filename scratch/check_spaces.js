const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();
async function run() {
    try {
        const itemMaster = await prisma.iNV_MATERIAL_MASTER.findFirst();
        console.log("INV_MATERIAL_MASTER ItemCode:", `'${itemMaster.ItemCode}'`);

        const custItem = await prisma.sAL_CUSTOMER_ITEMMASTER.findFirst();
        console.log("SAL_CUSTOMER_ITEMMASTER ItemCode:", `'${custItem.ItemCode}'`);
    } catch(e) {
        console.error(e);
    } finally {
        await prisma.$disconnect();
    }
}
run();

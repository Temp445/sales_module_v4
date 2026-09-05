const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();
async function run() {
    try {
        const item = await prisma.iNV_MATERIAL_MASTER.findFirst();
        console.log("INV_MATERIAL_MASTER:", item);
    } catch(e) {
        console.error(e);
    } finally {
        await prisma.$disconnect();
    }
}
run();

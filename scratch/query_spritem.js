const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();
async function run() {
    try {
        const item = await prisma.sprItem.findFirst();
        console.log("sprItem:", item);
    } catch(e) {
        console.error(e);
    } finally {
        await prisma.$disconnect();
    }
}
run();

const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function main() {
    try {
        const data = await prisma.sAL_CUSTOMER_ITEMMASTER.findMany();
        console.log("Total records:", data.length);
        if (data.length > 0) {
            console.log("Sample record:", JSON.stringify(data[0], null, 2));
        }
    } catch (e) {
        console.error(e);
    } finally {
        await prisma.$disconnect();
    }
}
main();

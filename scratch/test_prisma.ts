import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function main() {
    try {
        console.log("Connecting to Prisma...");
        const exists = await prisma.sAL_COUNTRYCITY_MASTER.findFirst({
            where: {
                ParentCode: 0
            }
        });
        console.log("Exists:", exists);
    } catch (e) {
        console.error("Prisma Error:", e);
    } finally {
        await prisma.$disconnect();
    }
}

main();

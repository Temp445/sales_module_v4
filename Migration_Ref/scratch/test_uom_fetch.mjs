import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function checkUomData() {
    console.log("Checking UOM Fetch...");
    try {
        const uoms = await prisma.$queryRawUnsafe(`
            SELECT TOP 5 Code, SDesc, CompId, Branchid
            FROM INV_UOM_MASTER 
        `);
        console.log(uoms);
    } catch (e) {
        console.error("Test failed:", e);
    } finally {
        await prisma.$disconnect();
    }
}

checkUomData();

import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function testFetch() {
    try {
        const code = 914; // demo3 party code from screenshot
        const branchId = '00';
        const compId = '02';
        
        console.log(`Fetching PartyCode: ${code}, Branch: ${branchId}, Comp: ${compId}`);

        const countries = await prisma.sAL_COUNTRYCITY_MASTER.findMany({
            where: { ParentCode: 0 }
        });
        const p2 = await prisma.$queryRaw`SELECT CountryCode, CountryName FROM COM_CountryCity_Master_Old WHERE ParentCode = 2`;
        console.log("ParentCode=2:", p2);
        
        const p3 = await prisma.$queryRaw`SELECT CountryCode, CountryName FROM COM_CountryCity_Master_Old WHERE ParentCode = 3`;
        console.log("ParentCode=3:", p3);
        
        console.log("Fetch logic works.");
        
    } catch (error) {
        console.error("Fetch Test Failed:", error);
    } finally {
        await prisma.$disconnect();
    }
}

testFetch();

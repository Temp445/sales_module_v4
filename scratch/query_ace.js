const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function runTest() {
    try {
        const ace = await prisma.fAS_Party_Master.findFirst({
            where: {
                PartyName: { contains: "ACE SOFTWARE" }
            }
        });
        
        if (!ace) {
            console.log("No ACE party found in FAS_Party_Master.");
            return;
        }
        console.log(`Found ACE party. PartyCode: ${ace.PartyCode}, PartyName: ${ace.PartyName}`);

        const items = await prisma.sAL_CUSTOMER_ITEMMASTER.findMany({
            where: {
                CustomerCode: ace.PartyCode
            }
        });

        console.log(`Found ${items.length} items for ACE in SAL_CUSTOMER_ITEMMASTER.`);
        if (items.length > 0) {
            console.log(items);
        }

    } catch (e) {
        console.error("Error:", e);
    } finally {
        await prisma.$disconnect();
    }
}
runTest();

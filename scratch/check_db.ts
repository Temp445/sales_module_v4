import prisma from '../src/lib/prisma';

async function check() {
    try {
        const records = await prisma.sAL_COUNTRYCITY_MASTER.findMany({
            select: {
                Branchid: true,
                Compid: true
            },
            distinct: ['Branchid', 'Compid']
        });
        console.log("Distinct Branchid/Compid combinations:", records);
        
        // Also check the specific records we inserted vs existing
        const india = await prisma.sAL_COUNTRYCITY_MASTER.findFirst({
            where: { CountryName: 'INDIA' }
        });
        console.log("INDIA record:", india);

        const delhi = await prisma.sAL_COUNTRYCITY_MASTER.findMany({
            where: { CountryName: 'DELHI' }
        });
        console.log("DELHI records:", delhi);
    } catch (e) {
        console.error(e);
    }
}
check();

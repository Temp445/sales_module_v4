import prisma from '../src/lib/prisma';

async function fix() {
    try {
        await prisma.sAL_COUNTRYCITY_MASTER.updateMany({
            where: {
                CountryName: 'DELHI',
                Compid: '07'
            },
            data: {
                Compid: '02'
            }
        });
        console.log("Fixed DELHI record Compid to 02");
    } catch (e) {
        console.error(e);
    }
}
fix();

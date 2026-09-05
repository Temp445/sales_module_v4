import prisma from '../src/lib/prisma';

async function check() {
    try {
        const maxRes: any = await prisma.$queryRaw`SELECT ISNULL(MAX(CountryCode),0) + 1 AS NextCode FROM SAL_COUNTRYCITY_MASTER`;
        const nextCode = maxRes[0].NextCode;

        console.log("Next Code:", nextCode);

        const newRecord = await prisma.sAL_COUNTRYCITY_MASTER.create({
            data: {
                CountryCode: nextCode,
                CountryName: 'TestCountryAPI_' + Date.now(),
                ParentCode: 0,
                userid: 'Admin',
                Branchid: '00',
                Compid: '07'
            }
        });
        
        console.log("Inserted:", newRecord);

        // cleanup
        await prisma.sAL_COUNTRYCITY_MASTER.delete({
            where: {
                CountryCode_ParentCode_Branchid_Compid: {
                    CountryCode: nextCode,
                    ParentCode: 0,
                    Branchid: '00',
                    Compid: '07'
                }
            }
        });
        console.log("Cleaned up");
    } catch (e) {
        console.error(e);
    }
}
check();

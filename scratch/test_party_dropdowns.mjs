import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function testDropdowns() {
    console.log("Testing Party Master Dropdowns...\n");

    try {
        console.log("1. Currencies");
        const currencies = await prisma.cOM_Currency_Master1.findMany({ take: 2 });
        console.log("Currencies found:", currencies.length);

        console.log("\n2. Customer Formulas (SAL)");
        const custForms = await prisma.cOM_FormulaMast_Hdr.findMany({ where: { ModuleCode: 4 }, take: 2 });
        console.log("Customer Formulas found:", custForms.length);

        console.log("\n3. Vendor Formulas (PUR)");
        const venForms = await prisma.cOM_FormulaMast_Hdr.findMany({ where: { ModuleCode: 9 }, take: 2 });
        console.log("Vendor Formulas found:", venForms.length);

        console.log("\n4. TermsMaster (TypeFlg='Payment Terms')");
        const terms = await prisma.termsMaster.findMany({ where: { TypeFlg: 'Payment Terms' }, take: 2 });
        console.log("Payment Terms found:", terms.length);

        console.log("\n5. Helpers (City, State, Country)");
        try {
            const countries = await prisma.$queryRaw`SELECT CountryCode, CountryName FROM COM_CountryCity_Master_Old WHERE ParentCode = 0 ORDER BY CountryName`;
            console.log("Countries found:", countries.length);
            
            if (countries.length > 0) {
                const cCode = countries[0].CountryCode;
                const states = await prisma.$queryRaw`SELECT CountryCode AS StateCode, CountryName AS StateName FROM COM_CountryCity_Master_Old WHERE ParentCode = ${cCode} ORDER BY CountryName`;
                console.log("States found:", states.length);
                
                if (states.length > 0) {
                    const sCode = states[0].StateCode;
                    const cities = await prisma.$queryRaw`SELECT CountryCode AS CityCode, CountryName AS CityName FROM COM_CountryCity_Master_Old WHERE ParentCode = ${sCode} ORDER BY CountryName`;
                    console.log("Cities found:", cities.length);
                }
            }
        } catch(e) {
            console.log("Error:", e.message);
        }

    } catch (e) {
        console.error("Test Failed:", e);
    } finally {
        await prisma.$disconnect();
    }
}

testDropdowns();

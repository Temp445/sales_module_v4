import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function dumpSps() {
    try {
        const country = await prisma.$queryRaw`SELECT OBJECT_DEFINITION(OBJECT_ID('FAS_PartyMaster_Hlp_Country')) AS definition`;
        if (country.length) console.log("--- FAS_PartyMaster_Hlp_Country ---\n", country[0].definition);

        const state = await prisma.$queryRaw`SELECT OBJECT_DEFINITION(OBJECT_ID('FAS_PartyMaster_Hlp_State')) AS definition`;
        if (state.length) console.log("--- FAS_PartyMaster_Hlp_State ---\n", state[0].definition);

        const city = await prisma.$queryRaw`SELECT OBJECT_DEFINITION(OBJECT_ID('FAS_PartyMaster_Hlp_City')) AS definition`;
        if (city.length) console.log("--- FAS_PartyMaster_Hlp_City ---\n", city[0].definition);

    } catch (e) {
        console.error("Error:", e);
    } finally {
        await prisma.$disconnect();
    }
}

dumpSps();

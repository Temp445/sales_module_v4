import prisma from '../src/lib/prisma';

async function testDB() {
    console.log("Testing TermsMaster Database Queries...");

    const branchId = '00';
    const compId = '02'; // Using Unicorn's compId for test

    try {
        console.log("\n--- 1. Testing Table Selection ---");
        // Count rows in TermsMaster
        const count1: any[] = await prisma.$queryRawUnsafe(`SELECT COUNT(*) as count FROM TERMSMASTER`);
        console.log(`TERMSMASTER count: ${count1[0].count}`);

        // Count rows in COM_TermsMaster
        const count2: any[] = await prisma.$queryRawUnsafe(`SELECT COUNT(*) as count FROM COM_TermsMaster`);
        console.log(`COM_TermsMaster count: ${count2[0].count}`);

        console.log("\n--- 2. Testing Max Code Fetch ---");
        const typeFlg = 'Form';
        const maxSql = `
            SELECT isnull(max(convert(int,Code)),0) as maxCode 
            FROM TERMSMASTER 
            WHERE TypeFlg='${typeFlg}'
        `;
        const maxCode: any[] = await prisma.$queryRawUnsafe(maxSql);
        console.log(`Max code for '${typeFlg}': ${maxCode[0].maxCode}`);

        console.log("\n--- 3. Testing SP insTERMSMASTER ---");
        const newCode = (maxCode[0].maxCode + 1).toString();
        const desc = 'TEST TERM DESCRIPTION';
        const userId = 'ADMIN';

        try {
            await prisma.$executeRawUnsafe(
                `EXEC insTERMSMASTER '${typeFlg}', '${newCode}', '${desc}', 1, '${userId}', '${branchId}', '${compId}'`
            );
            console.log("SP Executed Successfully! Record saved.");
        } catch (spError) {
            console.error("SP Execution Failed (Likely parameter mismatch or transaction error):", spError);
        }

    } catch (e) {
        console.error("Database Test Failed:", e);
    } finally {
        await prisma.$disconnect();
    }
}

testDB();

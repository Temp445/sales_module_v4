import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function testSave() {
    console.log("Testing TermsMaster Prisma End-to-End Save Logic...");

    const branchId = '00';
    const compId = '02'; // Using Unicorn's compId for test
    const userId = 'ADMIN';

    const typeFlgStr = 'Form';
    const descriptionStr = 'TEST DOC FINAL';
    const statusBool = true;

    try {
        console.log("\n--- Testing Insert via SP ---");
        
        const maxSql = `
            SELECT isnull(max(convert(int,Code)),0) as maxCode 
            FROM TERMSMASTER 
            WHERE TypeFlg='${typeFlgStr}'
        `;
        const maxResult = await prisma.$queryRawUnsafe(maxSql);
        const newCode = (Number(maxResult[0].maxCode) + 1).toString();

        await prisma.$executeRawUnsafe(
            `EXEC insTERMSMASTER '${typeFlgStr}', '${newCode}', '${descriptionStr}', ${statusBool ? 1 : 0}, '${userId}', '${branchId}', '${compId}'`
        );
        console.log("Insert Success! Record saved via insTERMSMASTER.");

        console.log("\n--- Testing Update via SP ---");
        
        await prisma.$executeRawUnsafe(
            `EXEC upTERMSMASTER '${typeFlgStr}', '${newCode}', 'UPDATED TEST DOC FINAL', ${statusBool ? 1 : 0}, '${userId}', '${branchId}', '${compId}'`
        );

        console.log("Update Success! Record saved via upTERMSMASTER.");

        // Final cleanup
        await prisma.termsMaster.deleteMany({
            where: {
                Code: newCode,
                Branchid: branchId,
                Compid: compId
            }
        });
        console.log("\nCleanup successful. End-to-end Test Passed.");

    } catch (e) {
        console.error("Save Test Failed:", e);
    } finally {
        await prisma.$disconnect();
    }
}

testSave();

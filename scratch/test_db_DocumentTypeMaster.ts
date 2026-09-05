import prisma from '../src/lib/prisma';

async function testDB() {
    console.log("Testing DocumentTypeMaster Database Queries...");

    const branchId = '00';
    const compId = '02'; // Using Unicorn's compId for test

    try {
        console.log("\n--- 1. Testing Grid Fetch ---");
        const fetchSql = `
            SELECT DocumentType, Description, StartingNo, PertainingTo 
            FROM COM_DOCUMENTTYPE_MASTER 
            WHERE Branchid='${branchId}' AND Compid='${compId}' AND ModuleName='SAL'
        `;
        const docs: any[] = await prisma.$queryRawUnsafe(fetchSql);
        console.log(`Found ${docs.length} documents.`);
        if (docs.length > 0) {
            console.log(docs.slice(0, 3));
        }

        console.log("\n--- 2. Testing SP COM_INS_DOCUMENTTYPE_MASTER ---");
        // Test an insert. Let's use a dummy test DocumentType 'TST'
        const docType = 'TST';
        const desc = 'Test Document';
        const startNo = 1;
        const pertainingTo = 'E'; // E = Sales Enquiry
        const userId = 'ADMIN';

        // Check if it exists first
        const exists: any[] = await prisma.$queryRawUnsafe(`
            SELECT DocumentType FROM COM_DOCUMENTTYPE_MASTER 
            WHERE DocumentType='${docType}' AND Branchid='${branchId}' AND Compid='${compId}'
        `);

        if (exists.length === 0) {
            try {
                await prisma.$executeRawUnsafe(
                    `EXEC COM_INS_DOCUMENTTYPE_MASTER '${docType}', '${desc}', ${startNo}, '${pertainingTo}', 'SAL', 1, '${userId}', '${branchId}', '${compId}'`
                );
                console.log("SP Executed Successfully! Record saved.");
            } catch (spError) {
                console.error("SP Execution Failed (Likely parameter mismatch):", spError);
            }
        } else {
            console.log("Test record already exists, skipping SP execution test.");
        }

    } catch (e) {
        console.error("Database Test Failed:", e);
    } finally {
        await prisma.$disconnect();
    }
}

testDB();

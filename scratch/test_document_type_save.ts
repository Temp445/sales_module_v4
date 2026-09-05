import prisma from '../src/lib/prisma';

async function testSave() {
    console.log("Testing DocumentTypeMaster Prisma Save Logic...");

    const branchId = '00';
    const compId = '02'; // Using Unicorn's compId for test
    const userId = 'ADMIN';

    const documentTypeStr = 'TS2';
    const descriptionStr = 'TEST DOC 2';
    const startingNoInt = 100;
    const pertainingToCode = 'E'; // Sales Enquiry

    try {
        console.log("\n--- Testing Insert via Prisma ORM ---");
        
        // Cleanup previous if exists
        await prisma.cOM_DocumentType_Master.deleteMany({
            where: {
                DocumentType: documentTypeStr,
                Branchid: branchId,
                Compid: compId
            }
        });

        // Test Insert
        const created = await prisma.cOM_DocumentType_Master.create({
            data: {
                DocumentType: documentTypeStr,
                Description: descriptionStr,
                StartingNo: startingNoInt,
                PertainingTo: pertainingToCode,
                ModuleName: "SAL",
                Status: true,
                Userid: userId,
                Branchid: branchId,
                Compid: compId
            }
        });

        console.log("Insert Success! Record:", created);

        console.log("\n--- Testing Update via Prisma ORM ---");
        
        const updated = await prisma.cOM_DocumentType_Master.update({
            where: {
                PK_COM_DocumentType_Master: {
                    DocumentType: documentTypeStr,
                    PertainingTo: pertainingToCode,
                    ModuleName: "SAL",
                    Branchid: branchId,
                    Compid: compId
                }
            },
            data: {
                Description: "UPDATED DOC 2",
                StartingNo: 200,
                Userid: userId,
                Status: true
            }
        });

        console.log("Update Success! Record:", updated);

        // Final cleanup
        await prisma.cOM_DocumentType_Master.deleteMany({
            where: {
                DocumentType: documentTypeStr,
                Branchid: branchId,
                Compid: compId
            }
        });
        console.log("\nCleanup successful. Test Passed.");

    } catch (e) {
        console.error("Save Test Failed:", e);
    } finally {
        await prisma.$disconnect();
    }
}

testSave();

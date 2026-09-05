import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function main() {
    try {
        console.log('Connecting to database...');
        
        const compId = "07";
        const branchId = "00";

        // Try to fetch one record to verify table access
        const record = await prisma.sAL_CUSTOMER_ITEMMASTER.findFirst({
            where: {
                Compid: compId,
                Branchid: branchId
            }
        });

        if (record) {
            console.log('Successfully fetched a record:');
            console.log(record);
        } else {
            console.log('No records found, but table connection works.');
        }

    } catch (e) {
        console.error('Error:', e.message);
        process.exit(1);
    } finally {
        await prisma.$disconnect();
    }
}

main();

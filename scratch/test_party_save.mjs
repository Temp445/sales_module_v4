import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function testSave() {
    try {
        console.log("Testing Save Logic (with intentional rollback)...");
        
        // 1. Fetch an existing party to update
        const party = await prisma.fAS_Party_Master.findFirst();
        if (!party) {
            console.log("No parties found to test update.");
            return;
        }

        const PartyCode = party.PartyCode;
        const compCode = party.Compid;
        const branchId = party.Branchid;
        const compId = party.Compid;
        const userId = party.Userid;

        const curDate = new Date().toISOString().slice(0, 19).replace('T', ' ');

        await prisma.$transaction(async (tx) => {
            console.log("Updating LedgerMaster...");
            await tx.$queryRawUnsafe(`
                Update FAS_LedgerMaster 
                Set isActive='0',
                    Userid='${userId}',
                    RegisteredName='Test Registered Name',
                    TranStamp='${curDate}'
                Where LedgerCode=${PartyCode} and CompCode='${compCode}'
            `);

            console.log("Executing FAS_UP_PARTY_MASTER...");
            await tx.$queryRawUnsafe(`
                Execute FAS_UP_PARTY_MASTER
                @PartyCode=${PartyCode},
                @VendorCode='TEST01',
                @SuplierCode='',
                @PartyName='Test Party Name',
                @AcType='P',
                @FormCode=0,
                @PackingCode=0,
                @DeliveryCode=0,
                @PaymentCode=0,
                @InsuranceCode=0,
                @ModeCode=0,
                @Freightcode=0,
                @IsSubcontractor='0',
                @AliasName='',
                @EffStatus='0',
                @EffDate='${curDate}',
                @CreditDays=0,
                @CreditDaysPeriod='',
                @CurrencyCode='INR',
                @CustomerFormulaCode=0,
                @VendorFormulaCode=0,
                @TinNo='',
                @PANNo='',
                @Userid='${userId}',
                @Branchid='${branchId}',
                @Compid='${compId}'
            `);

            console.log("Executing FAS_UP_PARTY_ADDRESS...");
            await tx.$queryRawUnsafe(`
                Execute FAS_UP_PARTY_ADDRESS
                @PartyCode=${PartyCode},
                @add1='Test Address 1',
                @add2='',
                @add3='',
                @Country='INDIA',
                @State='KERALA',
                @City='COCHIN',
                @pincode='682020',
                @Phone='9999999999',
                @fax='',
                @Cellular='',
                @email='test@example.com',
                @cst='',
                @cstdate=Null,
                @lst='',
                @lstdate=Null,
                @GSTRegNo='',
                @GSTRegDate=Null,
                @ProvisionalID='',
                @UdyamNo='',
                @UdyamDate=Null,
                @areaCode='',
                @eccno='',
                @Range='', 
                @EDRegNo='',
                @edregdate=Null,
                @contactperson='Test Person',
                @Userid='${userId}',
                @Branchid='${branchId}',
                @Compid='${compId}'
            `);

            console.log("All updates executed successfully.");
            throw new Error("INTENTIONAL_ROLLBACK"); // Rollback the transaction so we don't modify actual data
        });
        
    } catch (e) {
        if (e.message === "INTENTIONAL_ROLLBACK") {
            console.log("Test Passed: Save logic completed successfully and rolled back.");
        } else {
            console.error("Save Test Failed:", e);
        }
    } finally {
        await prisma.$disconnect();
    }
}

testSave();

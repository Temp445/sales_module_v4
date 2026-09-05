import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function POST(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const body = await req.json();
        const {
            isUpdate,
            PartyCode, // Important: Generate this on frontend or pass for update
            PartyName,
            RegisteredName,
            AcType,
            VendorCode,
            SuplierCode,
            Status, // 0 for Active, 1 for Inactive
            CreditDays,
            CreditDaysPeriod,
            AliasName,
            IsSubcontractor,
            FormCode,
            PackingCode,
            DeliveryCode,
            PaymentCode,
            InsuranceCode,
            ModeCode,
            FreightCode,
            EffDate,
            CurrencyCode,
            CustomerFormulaCode,
            VendorFormulaCode,
            TinNo,
            PanNo,
            
            // Address Details
            Add1, Add2, Add3,
            Country, State, City,
            PinCode, Phone, Fax, Cellular, Email,
            Cst, CstDate, Lst, LstDate,
            GSTRegNo, GSTRegDate, ProvisionalID,
            UdyamType, UdyamNo, UdyamDate,
            AreaCode, EccNo, Range, EDRegNo, EDRegDate,
            ContactPerson
        } = body;

        const userId = session.username;
        const branchId = session.branchId;
        const compId = session.companyId;
        const compCode = compId; // Often first 4 chars of CompID in legacy
        
        // Construct Udyam No correctly based on VB6 logic
        const finalUdyamNo = UdyamType ? `${UdyamNo || ''} ${UdyamType}`.trim() : UdyamNo;

        // Current timestamp for TranStamp
        const curDate = new Date().toISOString().slice(0, 19).replace('T', ' ');

        return await prisma.$transaction(async (tx) => {
            // 1. LedgerMaster Insert or Update
            if (isUpdate) {
                await tx.$queryRawUnsafe(`
                    Update FAS_LedgerMaster 
                    Set isActive='${Status}',
                        Userid='${userId}',
                        RegisteredName='${RegisteredName || ''}',
                        TranStamp='${curDate}'
                    Where LedgerCode=${PartyCode} and CompCode='${compCode}'
                `);
            } else {
                await tx.$queryRawUnsafe(`
                    Execute FAS_Ins_LedgerMaster
                    @LedgerCode=${PartyCode},
                    @LedgerName='${PartyName.replace(/'/g, "''")}',
                    @Alias='',
                    @RegisteredName='${(RegisteredName || '').replace(/'/g, "''")}',
                    @LrRoute='${AcType}',
                    @MaxReached='0',
                    @isBillwise='1',
                    @CreditDays=${CreditDays || 0},
                    @OPbalance=0,
                    @isCostCentreApplicable=0,
                    @isActive='${Status}',
                    @CompCode='${compCode}',
                    @Userid='${userId}',
                    @TranStamp='${curDate}'
                `);
            }

            // 2. PartyMaster Insert or Update
            const checkParty = await tx.$queryRawUnsafe(`
                Select PartyCode from FAS_PARTY_MASTER 
                where PartyCode = ${PartyCode} and Branchid ='${branchId}' and Compid='${compId}'
            `);

            const partySp = (checkParty as any[]).length > 0 ? 'FAS_UP_PARTY_MASTER' : 'FAS_INS_PARTY_MASTER';

            await tx.$queryRawUnsafe(`
                Execute ${partySp}
                @PartyCode=${PartyCode},
                @VendorCode='${VendorCode || ''}',
                @SuplierCode='${SuplierCode || ''}',
                @PartyName='${PartyName.replace(/'/g, "''")}',
                @AcType='${AcType}',
                @FormCode=${FormCode || 0},
                @PackingCode=${PackingCode || 0},
                @DeliveryCode=${DeliveryCode || 0},
                @PaymentCode=${PaymentCode || 0},
                @InsuranceCode=${InsuranceCode || 0},
                @ModeCode=${ModeCode || 0},
                @Freightcode=${FreightCode || 0},
                @IsSubcontractor='${IsSubcontractor ? 1 : 0}',
                @AliasName='${(AliasName || '').replace(/'/g, "''")}',
                @EffStatus='${Status}',
                @EffDate='${EffDate}',
                @CreditDays=${CreditDays || 0},
                @CreditDaysPeriod='${CreditDaysPeriod || ''}',
                @CurrencyCode='${CurrencyCode || ''}',
                @CustomerFormulaCode=${CustomerFormulaCode || 0},
                @VendorFormulaCode=${VendorFormulaCode || 0},
                @TinNo='${TinNo || ''}',
                @PANNo='${PanNo || ''}',
                @Userid='${userId}',
                @Branchid='${branchId}',
                @Compid='${compId}'
            `);

            // 3. PartyAddress Insert or Update
            const checkAddress = await tx.$queryRawUnsafe(`
                Select PartyCode from FAS_PARTY_ADDRESS 
                where PartyCode = ${PartyCode} and Branchid ='${branchId}' and Compid='${compId}'
            `);

            const addressSp = (checkAddress as any[]).length > 0 ? 'FAS_UP_PARTY_ADDRESS' : 'FAS_INS_PARTY_ADDRESS';

            await tx.$queryRawUnsafe(`
                Execute ${addressSp}
                @PartyCode=${PartyCode},
                @add1='${(Add1 || '').replace(/'/g, "''")}',
                @add2='${(Add2 || '').replace(/'/g, "''")}',
                @add3='${(Add3 || '').replace(/'/g, "''")}',
                @Country='${Country || ''}',
                @State='${State || ''}',
                @City='${City || ''}',
                @pincode='${PinCode || ''}',
                @Phone='${Phone || ''}',
                @fax='${Fax || ''}',
                @Cellular='${Cellular || ''}',
                @email='${Email || ''}',
                @cst='${Cst || ''}',
                @cstdate=${CstDate ? `'${CstDate}'` : 'Null'},
                @lst='${Lst || ''}',
                @lstdate=${LstDate ? `'${LstDate}'` : 'Null'},
                @GSTRegNo='${GSTRegNo || ''}',
                @GSTRegDate=${GSTRegDate ? `'${GSTRegDate}'` : 'Null'},
                @ProvisionalID='${ProvisionalID || ''}',
                @UdyamNo='${finalUdyamNo || ''}',
                @UdyamDate=${UdyamDate ? `'${UdyamDate}'` : 'Null'},
                @areaCode='${AreaCode || ''}',
                @eccno='${EccNo || ''}',
                @Range='${EccNo || ''}', 
                @EDRegNo='${EDRegNo || ''}',
                @edregdate=${EDRegDate ? `'${EDRegDate}'` : 'Null'},
                @contactperson='${(ContactPerson || '').replace(/'/g, "''")}',
                @Userid='${userId}',
                @Branchid='${branchId}',
                @Compid='${compId}'
            `);

            return NextResponse.json({ success: true, PartyCode });
        });

    } catch (error: any) {
        console.error("Save Party Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

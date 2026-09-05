import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const { searchParams } = new URL(req.url);
        // By default, for Sales module, only show Debtors (D) and Parties (Y)
        // If they want to search specifically, we can allow passing acType
        const acTypeQuery = searchParams.get("acType"); 

        let actypeFilter = { in: ['D', 'Y'] };
        
        if (acTypeQuery === 'Creditor') actypeFilter = { in: ['P', 'Y'] };
        else if (acTypeQuery === 'Debtor') actypeFilter = { in: ['D', 'Y'] };
        else if (acTypeQuery === 'Party') actypeFilter = { in: ['Y'] };
        else if (acTypeQuery === 'All') actypeFilter = { in: ['D', 'P', 'Y'] };

        const parties = await prisma.fAS_Party_Master.findMany({
            where: {
                Branchid: session.branchId,
                Compid: session.compId,
                AcType: actypeFilter,
                EffStatus: true
            },
            select: {
                PartyName: true,
                PartyCode: true,
                VendorCode: true,
                AcType: true,
                EffStatus: true,
                EffDate: true,
                SuplierCode: true,
                CustomerFormulaCode: true
            },
            orderBy: {
                PartyName: 'asc'
            }
        });

        // Map AcType codes to UI strings
        const formattedParties = parties.map(p => {
            let acTypeStr = "Party";
            if (p.AcType === 'D') acTypeStr = "Debtor";
            else if (p.AcType === 'P') acTypeStr = "Creditor";
            else if (p.AcType === 'Y') acTypeStr = "Party";

            return {
                ...p,
                PartyName: p.PartyName.trim(),
                VendorCode: p.VendorCode?.trim() || "",
                SuplierCode: p.SuplierCode?.trim() || "",
                AcTypeStr: acTypeStr,
                EffStatus: p.EffStatus ? 1 : 0,
                customerFormulaCode: p.CustomerFormulaCode
            };
        });
        
        return NextResponse.json(formattedParties);
    } catch (error: any) {
        console.error("Fetch Party Search Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

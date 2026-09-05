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
        const partyCodeStr = searchParams.get("partyCode");
        
        if (!partyCodeStr) {
            return NextResponse.json({ error: "partyCode is required" }, { status: 400 });
        }

        const partyCode = parseInt(partyCodeStr, 10);
        if (isNaN(partyCode)) {
            return NextResponse.json({ error: "Invalid partyCode format" }, { status: 400 });
        }

        console.log(`[party/defaults] Fetching defaults for partyCode: ${partyCode}, compId: ${session.compId}, branchId: ${session.branchId}`);

        const defaults = await prisma.fAS_Party_Master.findFirst({
            where: {
                PartyCode: partyCode,
                Compid: session.compId,
                Branchid: session.branchId
            },
            select: {
                FormCode: true,
                PackingCode: true,
                DeliveryCode: true,
                PaymentCode: true,
                InsuranceCode: true,
                ModeCode: true,
                Freightcode: true,
                CustomerFormulaCode: true
            }
        });

        if (!defaults) {
             return NextResponse.json({ error: "Party not found" }, { status: 404 });
        }

        return NextResponse.json(defaults);

    } catch (error: any) {
        console.error("Fetch Party Defaults Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

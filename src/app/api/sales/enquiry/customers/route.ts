import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        // Fetch distinct customers from SAL_Customer_ItemMaster with their PartyName from FAS_PARTY_MASTER
        const customers = await prisma.$queryRaw`
            SELECT DISTINCT 
                b.PartyName, 
                a.CustomerCode 
            FROM SAL_Customer_ItemMaster a
            JOIN FAS_PARTY_MASTER b ON b.PartyCode = a.CustomerCode AND b.Branchid = a.Branchid AND b.Compid = a.Compid
            WHERE a.Branchid = ${session.branchId} AND a.Compid = ${session.compId}
            ORDER BY b.PartyName
        `;

        return NextResponse.json(customers);
    } catch (error: any) {
        console.error("Fetch Enquiry Customers Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

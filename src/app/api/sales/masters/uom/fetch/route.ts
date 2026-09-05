import { NextResponse, NextRequest } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(request: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        // Query the INV_UOM_MASTER
        let uoms = await prisma.$queryRawUnsafe(`
            SELECT Code, SDesc, NoofDecimals
            FROM INV_UOM_MASTER 
            WHERE CompId = $1 AND Branchid = $2
        `, session.compId, session.branchId);

        console.log("UOM API Session Check: CompId=", session.compId, "Branchid=", session.branchId, "Found=", Array.isArray(uoms) ? uoms.length : 0);

        if (!Array.isArray(uoms) || uoms.length === 0) {
            // Fallback for test DB which might only have data in CompId '02'
            console.warn(`No UOMs found for ${session.compId}, falling back to unfiltered`);
            uoms = await prisma.$queryRawUnsafe(`
                SELECT DISTINCT Code, SDesc, NoofDecimals 
                FROM INV_UOM_MASTER 
            `);
        }

        return NextResponse.json(uoms);
    } catch (error: any) {
        console.error("Failed to fetch UOM master:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

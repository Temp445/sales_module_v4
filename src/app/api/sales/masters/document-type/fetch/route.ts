import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        // Read PertainingTo from query param (e.g. ?PertainingTo=Q for Quotation)
        const { searchParams } = new URL(req.url);
        const pertainingTo = searchParams.get("PertainingTo");

        console.log("[document-type/fetch] PertainingTo:", pertainingTo, "compId:", session.compId, "branchId:", session.branchId);

        const documents = await prisma.cOM_DocumentType_Master.findMany({
            where: {
                Branchid: session.branchId,
                Compid: session.compId,
                ModuleName: "SAL",
                // Use 'startsWith' to handle SQL Server char columns with trailing spaces
                // e.g., PertainingTo stored as "Q   " won't match exact "Q" but matches startsWith
                ...(pertainingTo ? { PertainingTo: { startsWith: pertainingTo } } : {})
            },
            select: {
                DocumentType: true,
                Description: true,
                StartingNo: true,
                PertainingTo: true,
            },
            orderBy: {
                DocumentType: 'asc'
            }
        });

        console.log("[document-type/fetch] found:", documents.length, "types");

        return NextResponse.json(documents);
    } catch (error: any) {
        console.error("Fetch Document Type Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

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
        const typeFlg = searchParams.get("TypeFlg");

        const terms = await prisma.termsMaster.findMany({
            where: {
                Branchid: session.branchId,
                Compid: session.compId,
                ...(typeFlg && { TypeFlg: { startsWith: typeFlg } })
            },
            select: {
                TypeFlg: true,
                Code: true,
                Description: true,
                Status: true,
            },
            orderBy: {
                Code: 'asc'
            }
        });

        // Parse boolean status back into "Active" or "Inactive"
        const formattedTerms = terms.map(t => ({
            ...t,
            TypeFlg: t.TypeFlg.trim(),
            Code: t.Code.trim(),
            StatusStr: t.Status ? "Active" : "Inactive"
        }));
        
        return NextResponse.json(formattedTerms);
    } catch (error: any) {
        console.error("Fetch Terms Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

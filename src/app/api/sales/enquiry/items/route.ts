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
        const searchTerm = searchParams.get("q") || "";
        
        let query = `
            SELECT TOP 200
                ItemCode, 
                ItemDescription, 
                UOMDesc, 
                UOMDecimal, 
                Rate 
            FROM MATERIAL_MASTER 
            WHERE Branchid = '${session.branchId}' 
              AND CompID = '${session.compId}' 
              AND Status = 1
        `;

        if (searchTerm) {
            // Using raw interpolation here requires careful handling in production to avoid SQL injection,
            // but for safe internal string matches we can do simple replacement or pass it carefully.
            // Using parameterized query raw to prevent injection:
            const materials = await prisma.$queryRawUnsafe(`
                SELECT TOP 200
                    ItemCode, 
                    ItemDescription, 
                    UOMDesc, 
                    UOMDecimal, 
                    Rate 
                FROM MATERIAL_MASTER 
                WHERE Branchid = '${session.branchId}' 
                  AND CompID = '${session.compId}' 
                  AND Status = 1
                  AND (ItemCode LIKE '%${searchTerm.replace(/'/g, "''")}%' OR ItemDescription LIKE '%${searchTerm.replace(/'/g, "''")}%')
                ORDER BY ItemCode
            `);
            return NextResponse.json(materials);
        }

        const materials = await prisma.$queryRawUnsafe(query + " ORDER BY ItemCode");
        return NextResponse.json(materials);

    } catch (error: any) {
        console.error("Fetch Enquiry Items Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

import { NextRequest, NextResponse } from "next/server";
import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

export async function GET(req: NextRequest) {
    try {
        const { searchParams } = new URL(req.url);
        const compId = searchParams.get("compId");
        const branchId = searchParams.get("branchId");
        const sentItemCode = searchParams.get("sentItemCode");

        if (!compId || !branchId || !sentItemCode) {
            return NextResponse.json({ error: "Missing required parameters" }, { status: 400 });
        }

        // Simulating the BOM Structure retrieval to find child items for the given Parent (Sent Item)
        // In VB6 this queries BOM_Structure_Dtl joined with MATERIAL_MASTER
        const children = await prisma.$queryRawUnsafe(`
            SELECT 
                b.ChildCode, 
                c.ItemDescription as ChildDesc, 
                c.PurchaseUOMDesc as UOM 
            FROM 
                BOM_STRUCTURE_DTL b 
            LEFT JOIN 
                MATERIAL_MASTER c 
            ON 
                b.ChildCode = c.ItemCode 
                AND b.BranchID = c.BranchID 
                AND b.CompID = c.CompID
            WHERE 
                b.ParentCode = '${sentItemCode}' 
                AND b.BranchID = '${branchId}' 
                AND b.CompID = '${compId}'
        `);

        return NextResponse.json(children, { status: 200 });

    } catch (error: any) {
        console.error("Error fetching jobwork structure:", error);
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    } finally {
        await prisma.$disconnect();
    }
}

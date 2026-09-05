import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";

export async function GET(req: NextRequest) {
    try {
        // ModuleCode for Purchases (Vendors) is 9 (eMdlCode.mPUR)
        const formulas = await prisma.cOM_FormulaMast_Hdr.findMany({
            where: {
                ModuleCode: 9,
                isActive: true
            },
            select: {
                FrmID: true,
                FrmLName: true
            },
            orderBy: {
                FrmLName: 'asc'
            }
        });
        
        return NextResponse.json(formulas);
    } catch (error: any) {
        console.error("Fetch Vendor Formulas Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

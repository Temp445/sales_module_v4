import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";

export async function GET(req: NextRequest) {
    try {
        // ModuleCode for Sales is 4 (eMdlCode.mSAL)
        const formulas = await prisma.cOM_FormulaMast_Hdr.findMany({
            where: {
                ModuleCode: 4
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
        console.error("Fetch Customer Formulas Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

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
        const moduleCodeStr = searchParams.get("moduleCode");
        const moduleCode = moduleCodeStr ? parseInt(moduleCodeStr) : 4; // Default to 4 for Sales

        const formulas = await prisma.cOM_FormulaMast_Hdr.findMany({
            where: {
                ModuleCode: moduleCode,
                BranchID: session.branchId,
                CompID: session.compId
            },
            select: {
                FrmID: true,
                FrmLName: true
            },
            orderBy: {
                FrmLName: 'asc'
            }
        });

        const formatted = formulas.map(f => ({
            value: f.FrmID.toString(),
            label: f.FrmLName.trim()
        }));
        
        return NextResponse.json(formatted);
    } catch (error: any) {
        console.error("Fetch Formulas Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

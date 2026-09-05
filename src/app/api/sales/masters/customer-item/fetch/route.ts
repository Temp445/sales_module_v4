import { NextRequest, NextResponse } from "next/server";
import { PrismaClient } from "@prisma/client";
import { getSession } from "@/lib/session";

const prisma = new PrismaClient();

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        const compId = session.compId;
        const branchId = session.branchId;

        const { searchParams } = new URL(req.url);
        const customerCode = searchParams.get("customerCode");

        if (!compId || !branchId || !customerCode) {
            return NextResponse.json({ error: "Missing required parameters: compId, branchId, customerCode" }, { status: 400 });
        }

        const items = await prisma.sAL_CUSTOMER_ITEMMASTER.findMany({
            where: {
                Compid: compId,
                Branchid: branchId,
                CustomerCode: parseInt(customerCode)
            }
        });

        // Fetch item descriptions from material master
        const itemCodes = items.map(i => i.ItemCode);
        const materials = await prisma.iNV_MATERIAL_MASTER.findMany({
            where: {
                Compid: compId,
                BranchId: branchId,
                ItemCode: { in: itemCodes }
            },
            select: { ItemCode: true, ItemDescription: true }
        });
        const materialMap = new Map(materials.map(m => [m.ItemCode, m.ItemDescription]));

        // Format for frontend mapping
        const formattedItems = items.map(item => ({
            itemCode: { value: item.ItemCode, label: item.ItemCode }, 
            description: materialMap.get(item.ItemCode) || "", 
            model: item.Model,
            drawingIssuedNo: item.IssuedNo,
            customerPartNo: item.PartNo || "",
            rate: item.Rate,
            assessableRate: item.AssessableRate,
            validFrom: item.ValidFrom ? item.ValidFrom.toISOString().split('T')[0] : "",
            validTo: item.ValidTo ? item.ValidTo.toISOString().split('T')[0] : "",
            formulaCode: { value: item.FormulaCode, label: item.FormulaCode.toString() }, 
            formulaName: "", 
            qCCheckRequired: item.QCCheck ? "Yes" : "No",
            status: item.Status ? "Active" : "Inactive"
        }));

        return NextResponse.json(formattedItems, { status: 200 });

    } catch (error: any) {
        console.error("Error fetching customer items:", error);
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    } finally {
        await prisma.$disconnect();
    }
}

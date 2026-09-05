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
            return NextResponse.json({ error: "Missing required parameters" }, { status: 400 });
        }

        const items = await prisma.sAL_Jobwork_ItemMaster.findMany({
            where: {
                CompID: compId,
                BranchID: branchId,
                CustomerCode: parseInt(customerCode)
            },
            orderBy: {
                ItemGroupKey: 'asc'
            }
        });

        // Fetch item descriptions from material master
        const itemCodes = Array.from(new Set([
            ...items.map(i => i.ReceivedItemCode),
            ...items.map(i => i.SentItemCode)
        ]));
        
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
            mFirstCol: item.ItemGroupKey.toString(),
            mrecdItem: { value: item.ReceivedItemCode, label: item.ReceivedItemCode },
            mRecdDesc: materialMap.get(item.ReceivedItemCode) || "", 
            mRecdUOM: "",
            msentItem: { value: item.SentItemCode, label: item.SentItemCode },
            mSentDesc: materialMap.get(item.SentItemCode) || "",
            mPartNo: item.PartNo || "",
            mSentUOM: "",
            mRatePer: item.RatePer,
            mRate: item.Rate,
            mAssRate: item.AssessableRate,
            mValidFrom: item.ValidFrom ? item.ValidFrom.toISOString().split('T')[0] : "",
            mValidTo: item.ValidTo ? item.ValidTo.toISOString().split('T')[0] : "",
            mQCCheck: item.QcCheck ? "Yes" : "No",
            mFrmCode: { value: item.FormulaCode, label: item.FormulaCode.toString() },
            mFrmName: "", 
            mStatus: item.isActive ? "Active" : "Inactive",
            mRelation: item.RelationType,
            mItemGroupKey: item.ItemGroupKey.toString(),
            mRecdItemCheck: ""
        }));

        return NextResponse.json(formattedItems, { status: 200 });

    } catch (error: any) {
        console.error("Error fetching jobwork items:", error);
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    } finally {
        await prisma.$disconnect();
    }
}

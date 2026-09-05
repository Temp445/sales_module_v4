import { NextRequest, NextResponse } from "next/server";
import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

export async function POST(req: NextRequest) {
    try {
        const body = await req.json();
        const { compId, branchId, userId, customerCode, items } = body;

        if (!compId || !branchId || !customerCode || !Array.isArray(items)) {
            return NextResponse.json({ error: "Missing or invalid parameters" }, { status: 400 });
        }

        const custCodeInt = parseInt(customerCode);
        if (isNaN(custCodeInt)) {
            return NextResponse.json({ error: "Customer Code must be a number" }, { status: 400 });
        }

        await prisma.$transaction(async (tx) => {
            await tx.sAL_Jobwork_ItemMaster.deleteMany({
                where: {
                    CompID: compId,
                    BranchID: branchId,
                    CustomerCode: custCodeInt
                }
            });

            const dataToInsert = items.map((item, idx) => ({
                CompID: compId,
                BranchID: branchId,
                UserID: userId || "",
                CustomerCode: custCodeInt,
                ItemGroupKey: parseInt(item.mItemGroupKey) || (idx + 1), // Simplistic fallback if empty
                ReceivedItemCode: item.mrecdItem?.value || "",
                SentItemCode: item.msentItem?.value || "",
                RelationType: item.mRelation || "CHI", // Example default
                PartNo: item.mPartNo || "",
                RatePer: parseFloat(item.mRatePer) || 1,
                Rate: parseFloat(item.mRate) || 0,
                AssessableRate: parseFloat(item.mAssRate) || 0,
                ValidFrom: item.mValidFrom ? new Date(item.mValidFrom) : null,
                ValidTo: item.mValidTo ? new Date(item.mValidTo) : null,
                QcCheck: item.mQCCheck === "Yes",
                FormulaCode: parseInt(item.mFrmCode?.value) || 0,
                isActive: item.mStatus === "Active"
            }));

            if (dataToInsert.length > 0) {
                await tx.sAL_Jobwork_ItemMaster.createMany({
                    data: dataToInsert
                });
            }
        });

        return NextResponse.json({ status: 200, message: "Saved successfully", customerCode: custCodeInt }, { status: 200 });

    } catch (error: any) {
        console.error("Error saving jobwork items:", error);
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    } finally {
        await prisma.$disconnect();
    }
}

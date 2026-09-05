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

        // Validate types before database interaction (Rule 6)
        const custCodeInt = parseInt(customerCode);
        if (isNaN(custCodeInt)) {
            return NextResponse.json({ error: "Customer Code must be a number" }, { status: 400 });
        }

        // Run transaction
        await prisma.$transaction(async (tx) => {
            // First delete all existing items for this customer
            await tx.sAL_CUSTOMER_ITEMMASTER.deleteMany({
                where: {
                    Compid: compId,
                    Branchid: branchId,
                    CustomerCode: custCodeInt
                }
            });

            // Prepare items for insertion
            const dataToInsert = items.map(item => ({
                Compid: compId,
                Branchid: branchId,
                Userid: userId || "",
                CustomerCode: custCodeInt,
                ItemCode: item.itemCode?.value || "",
                Model: item.model || "",
                IssuedNo: item.issuedNo || "",
                PartNo: item.partNo || null,
                Rate: parseFloat(item.rate) || 0,
                AssessableRate: parseFloat(item.assessableRate) || 0,
                ValidFrom: item.validFrom ? new Date(item.validFrom) : null,
                ValidTo: item.validTo ? new Date(item.validTo) : null,
                FormulaCode: parseInt(item.formulaCode?.value) || 0,
                QCCheck: item.qcCheck === "Yes",
                Status: item.status === "Active"
            }));

            if (dataToInsert.length > 0) {
                await tx.sAL_CUSTOMER_ITEMMASTER.createMany({
                    data: dataToInsert
                });
            }
        });

        return NextResponse.json({ status: 200, message: "Saved successfully", customerCode: custCodeInt }, { status: 200 });

    } catch (error: any) {
        console.error("Error saving customer items:", error);
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    } finally {
        await prisma.$disconnect();
    }
}

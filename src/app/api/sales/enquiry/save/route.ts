import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";
import { Prisma } from "@prisma/client";

export async function POST(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const body = await req.json();
        const { header, items } = body;

        // Validation
        if (!header.enquiryType || !header.enquiryNo || !header.customerCode) {
            return NextResponse.json({ error: "Missing required header fields" }, { status: 400 });
        }
        if (!items || items.length === 0) {
            return NextResponse.json({ error: "Enquiry must have at least one item" }, { status: 400 });
        }

        // Format dates
        const curDate = new Date();
        const enquiryDate = header.enquiryDate ? new Date(header.enquiryDate).toISOString().split('T')[0] : curDate.toISOString().split('T')[0];
        const referenceDate = header.referenceDate ? new Date(header.referenceDate).toISOString().split('T')[0] : curDate.toISOString().split('T')[0];
        
        // Execute inside transaction (Do NOT use BEGIN TRAN inside executeRaw as Prisma handles it)
        await prisma.$transaction(async (tx) => {
            // Save HDR
            await tx.$executeRawUnsafe(`
                EXEC SAL_INS_ENQUIRY_HDR 
                    @EnquiryType = '${header.enquiryType}',
                    @EnquiryNo = ${parseInt(header.enquiryNo)},
                    @EnquiryDate = '${enquiryDate}',
                    @RevisionNo = 0,
                    @RevisionDate = '${curDate.toISOString().split('T')[0]}',
                    @ReferenceNo = '${header.referenceNo || ''}',
                    @ReferenceDate = '${referenceDate}',
                    @CustomerType = '${header.customerType === 'Existing Customer' ? 'E' : 'N'}',
                    @CustomerCode = ${parseInt(header.customerCode) || 0},
                    @CustomerName = '${header.customerName || ''}',
                    @Remarks = '${header.remarks || ''}',
                    @Userid = '${session.userId}',
                    @Branchid = '${session.branchId}',
                    @Compid = '${session.compId}',
                    @Finyear = '${session.finYear}'
            `);

            // Save DTL (Items)
            for (const item of items) {
                if (!item.itemCode) continue;

                const targetDate = item.targtDate ? new Date(item.targtDate).toISOString().split('T')[0] : curDate.toISOString().split('T')[0];

                await tx.$executeRawUnsafe(`
                    EXEC SAL_INS_ENQUIRY_DTL
                        @EnquiryType = '${header.enquiryType}',
                        @EnquiryNo = ${parseInt(header.enquiryNo)},
                        @EnquiryDate = '${enquiryDate}',
                        @RevisionNo = 0,
                        @ItemType = '${item.itemType === 'New Item' ? 'N' : 'E'}',
                        @ItemCode = '${item.itemCode}',
                        @ItemDescription = '${item.itemDesc || ''}',
                        @UOM = '${item.uom || ''}',
                        @Qty = ${parseFloat(item.qty) || 0},
                        @Volume = ${parseFloat(item.tarVolme) || 0},
                        @VolumePer = '${item.volmeper || ''}',
                        @Rate = ${parseFloat(item.rate) || 0},
                        @AssessableRate = ${parseFloat(item.assRate) || 0},
                        @TargetDate = '${targetDate}',
                        @Specification = '${item.speci || ''}',
                        @Userid = '${session.userId}',
                        @Branchid = '${session.branchId}',
                        @Compid = '${session.compId}',
                        @Finyear = '${session.finYear}'
                `);
            }
        });

        return NextResponse.json({ message: "Sales Enquiry saved successfully", enquiryNo: header.enquiryNo });

    } catch (error: any) {
        console.error("Save Sales Enquiry Error:", error);
        return NextResponse.json({ error: error.message || "An unexpected error occurred" }, { status: 500 });
    }
}

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
        const {
            quotationType,
            quotationNo,
            quotationDate,
            orderSource, // 'D' or 'E'
            refType, // Enquiry Type
            refNo, // Enquiry No
            refDate, // Enquiry Date
            customerCode,
            customerName,
            validFrom,
            validTo,
            formulaCode,
            remarks,
            formCode,
            packingCode,
            deliveryCode,
            paymentCode,
            insuranceCode,
            modeCode,
            freightCode,
            totalAmount,
            items // Array of items
        } = body;

        // Basic Validation
        if (!quotationType || !quotationNo || !customerCode || !items || !items.length) {
            return NextResponse.json({ error: "Missing required fields" }, { status: 400 });
        }

        const finyear = session.finYear || '2026-2027'; // Fallback for tests if needed, but session should have it
        const userId = session.userId || 'SYSTEM';

        // Perform Save inside a Transaction
        await prisma.$transaction(async (tx) => {
            // 1. Insert Header
            const headerSql = Prisma.sql`
                EXEC SAL_INS_QUOTATION_HDR
                    @QuotationType = ${quotationType},
                    @QuotationNo = ${parseInt(quotationNo)},
                    @QuotationDate = ${new Date(quotationDate)},
                    @RevisionNo = 0,
                    @RevisionDate = ${new Date()},
                    @OrderSource = ${orderSource || 'D'},
                    @RefType = ${refType || ''},
                    @RefNo = ${refNo ? parseInt(refNo) : 0},
                    @RefDate = ${refDate ? new Date(refDate) : null},
                    @CustomerCode = ${parseInt(customerCode)},
                    @CustomerName = ${customerName || ''},
                    @ValidFrom = ${validFrom ? new Date(validFrom) : new Date()},
                    @ValidTo = ${validTo ? new Date(validTo) : new Date()},
                    @FormulaCode = ${formulaCode ? parseInt(formulaCode) : 0},
                    @Remarks = ${remarks || ''},
                    @FormCode = ${formCode ? parseInt(formCode) : 0},
                    @PackingCode = ${packingCode ? parseInt(packingCode) : 0},
                    @DeliveryCode = ${deliveryCode ? parseInt(deliveryCode) : 0},
                    @PaymentCode = ${paymentCode ? parseInt(paymentCode) : 0},
                    @InsuranceCode = ${insuranceCode ? parseInt(insuranceCode) : 0},
                    @ModeCode = ${modeCode ? parseInt(modeCode) : 0},
                    @FreightCode = ${freightCode ? parseInt(freightCode) : 0},
                    @TotalAmount = ${totalAmount ? parseFloat(totalAmount) : 0.00},
                    @UserId = ${userId},
                    @BranchId = ${session.branchId},
                    @CompId = ${session.compId},
                    @Finyear = ${finyear}
            `;
            await tx.$executeRaw(headerSql);

            // 2. Insert Details
            for (const item of items) {
                const detailSql = Prisma.sql`
                    EXEC SAL_INS_QUOTATION_DTL
                        @QuotationType = ${quotationType},
                        @QuotationNo = ${parseInt(quotationNo)},
                        @QuotationDate = ${new Date(quotationDate)},
                        @RevisionNo = 0,
                        @ItemCode = ${item.itemCode},
                        @ItemDescription = ${item.itemDescription || ''},
                        @UOM = ${item.uom || ''},
                        @Qty = ${parseFloat(item.qty || 0)},
                        @Rate = ${parseFloat(item.rate || 0)},
                        @AssessableRate = ${parseFloat(item.assessableRate || 0)},
                        @UserId = ${userId},
                        @BranchId = ${session.branchId},
                        @CompId = ${session.compId},
                        @Finyear = ${finyear}
                `;
                await tx.$executeRaw(detailSql);

                // Update Enquiry Details if source is Enquiry
                if (orderSource === 'E' && refNo && refType) {
                    const enqSql = Prisma.sql`
                        EXEC SAL_UP_ENQUIRY_DTL
                            @EnquiryType = ${refType},
                            @EnquiryNo = ${parseInt(refNo)},
                            @EnquiryDate = ${new Date(refDate)},
                            @ItemCode = ${item.itemCode},
                            @QuotationType = ${quotationType},
                            @QuotationNo = ${parseInt(quotationNo)},
                            @QuotationDate = ${new Date(quotationDate)},
                            @UserId = ${userId},
                            @BranchId = ${session.branchId},
                            @CompId = ${session.compId}
                    `;
                    await tx.$executeRaw(enqSql);
                }
            }
        });

        return NextResponse.json({ message: "Sales Quotation Saved Successfully", quotationNo });
    } catch (error: any) {
        console.error("Save Sales Quotation Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

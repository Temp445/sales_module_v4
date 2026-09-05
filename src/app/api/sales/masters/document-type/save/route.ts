import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

const PertainingMapping: Record<string, string> = {
    "Sales Enquiry": "E",
    "Jobwork Enquiry": "N",
    "Sales Quotation": "Q",
    "Jobwork Quotation": "U",
    "Customer PO Entry": "C",
    "Sales Order": "O",
    "Jobwork Order": "J",
    "Schedule Entry": "S",
    "Inventory Schedule Entry": "V",
    "Weekly Plan": "W",
    "Sales Order Amendment": "A",
    "Jobwork Order Amendment": "M",
    "Delivery Challan": "D",
    "Sales Invoice": "I",
    "Sales Plan Yearly": "YP",
    "Sales Plan Monthly": "MP",
    "Sales Plan Weekly": "WP",
    "Sales Plan Daily": "DP",
    "Contract Review": "CR",
    "Downstream Document": "DD"
};

export async function POST(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const body = await req.json();
        const { type, desc, start, pertainingTo, isUpdate } = body;

        // Validation
        if (!type || type.trim() === "") {
            return NextResponse.json({ error: "Enter the Document Type" }, { status: 400 });
        }
        if (!desc || desc.trim() === "") {
            return NextResponse.json({ error: "Enter the Description" }, { status: 400 });
        }
        if (start === undefined || start === null || isNaN(parseInt(start, 10))) {
            return NextResponse.json({ error: "Enter a valid Starting Number" }, { status: 400 });
        }

        const documentTypeStr = type.trim().toUpperCase();
        const descriptionStr = desc.trim().toUpperCase();
        const startingNoInt = parseInt(start, 10);
        const pertainingToCode = PertainingMapping[pertainingTo] || "";

        const branchId = session.branchId;
        const compId = session.compId;
        const userId = session.userId || "ADMIN";

        // Check for duplicate Document Type before insert
        if (!isUpdate) {
            const exists = await prisma.cOM_DocumentType_Master.findFirst({
                where: {
                    DocumentType: documentTypeStr,
                    Branchid: branchId,
                    Compid: compId
                }
            });
            
            if (exists) {
                return NextResponse.json({ error: "Document Type already exists" }, { status: 400 });
            }

            // Create new record
            await prisma.cOM_DocumentType_Master.create({
                data: {
                    DocumentType: documentTypeStr,
                    Description: descriptionStr,
                    StartingNo: startingNoInt,
                    PertainingTo: pertainingToCode,
                    ModuleName: "SAL",
                    Status: true,
                    Userid: userId,
                    Branchid: branchId,
                    Compid: compId
                }
            });
            
            return NextResponse.json({ success: true, message: "Record Saved" });
        } else {
            // Update existing record
            await prisma.cOM_DocumentType_Master.update({
                where: {
                    PK_COM_DocumentType_Master: {
                        DocumentType: documentTypeStr,
                        PertainingTo: pertainingToCode,
                        ModuleName: "SAL",
                        Branchid: branchId,
                        Compid: compId
                    }
                },
                data: {
                    Description: descriptionStr,
                    StartingNo: startingNoInt,
                    Userid: userId,
                    Status: true
                }
            });

            return NextResponse.json({ success: true, message: "Record Updated" });
        }

    } catch (error: any) {
        console.error("Save Document Type Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function POST(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const body = await req.json();
        const { typeFlg, desc, status, isUpdate, code } = body;

        // Validation
        if (!typeFlg || typeFlg.trim() === "") {
            return NextResponse.json({ error: "Select the Term" }, { status: 400 });
        }
        if (!desc || desc.trim() === "") {
            return NextResponse.json({ error: "Enter the Description" }, { status: 400 });
        }

        const typeFlgStr = typeFlg.trim();
        const descriptionStr = desc.trim().toUpperCase();
        const statusBool = status === "Active" ? true : false;
        
        const branchId = session.branchId;
        const compId = session.compId;
        const userId = session.userId || "ADMIN";

        // Check for duplicate Description
        const duplicateCheck = await prisma.termsMaster.findFirst({
            where: {
                TypeFlg: typeFlgStr,
                Description: descriptionStr,
                Branchid: branchId,
                Compid: compId,
                ...(isUpdate ? { Code: { not: code } } : {}) // If updating, ignore its own code
            }
        });
        
        if (duplicateCheck) {
            return NextResponse.json({ error: "Duplicate Entry not allowed" }, { status: 400 });
        }

        if (!isUpdate) {
            // Generate Code manually as the VB6 did
            const maxSql = `
                SELECT isnull(max(convert(int,Code)),0) as maxCode 
                FROM TERMSMASTER 
                WHERE TypeFlg='${typeFlgStr}'
            `;
            const maxResult: any[] = await prisma.$queryRawUnsafe(maxSql);
            const newCode = (Number(maxResult[0].maxCode) + 1).toString();

            // Execute original Stored Procedure
            await prisma.$executeRawUnsafe(
                `EXEC insTERMSMASTER '${typeFlgStr}', '${newCode}', '${descriptionStr}', ${statusBool ? 1 : 0}, '${userId}', '${branchId}', '${compId}'`
            );
            
            return NextResponse.json({ success: true, message: "Record Saved" });
        } else {
            // Execute original Update Stored Procedure
            await prisma.$executeRawUnsafe(
                `EXEC upTERMSMASTER '${typeFlgStr}', '${code}', '${descriptionStr}', ${statusBool ? 1 : 0}, '${userId}', '${branchId}', '${compId}'`
            );

            return NextResponse.json({ success: true, message: "Record Updated" });
        }

    } catch (error: any) {
        console.error("Save Terms Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

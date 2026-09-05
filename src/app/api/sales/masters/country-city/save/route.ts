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
        const { type, parentCode, countryCode, name } = body;

        if (!type || (type !== "Country" && type !== "City")) {
            return NextResponse.json({ error: "Invalid type" }, { status: 400 });
        }
        if (!name || name.trim() === "") {
            return NextResponse.json({ error: "Name is required" }, { status: 400 });
        }

        const branchId = session.branchId;
        const compId = session.compId;
        const userId = session.userId || "ADMIN";
        
        let pCode = 0;
        if (type === "City") {
            if (!parentCode || parentCode <= 0) {
                return NextResponse.json({ error: "Parent Country is required for City" }, { status: 400 });
            }
            pCode = parentCode;
        }

        const isUpdate = (countryCode && countryCode > 0);

        // Validation against duplicates for Insert
        if (!isUpdate) {
            const exists = await prisma.sAL_COUNTRYCITY_MASTER.findFirst({
                where: {
                    ParentCode: pCode,
                    CountryName: name.trim()
                }
            });
            if (exists) {
                return NextResponse.json({ error: "Duplicate Description Not allowed" }, { status: 400 });
            }
        }

        let newCode = countryCode || 0;
        
        // Let's use the exact stored procedure if we want to follow Prompt V7 Rule 14 & 15 exactly
        // Wait, the rule says "When calling an EXECUTE query, you MUST always include @Userid, @Branchid, and @Compid"
        // Let's use Prisma raw query for SP execution
        
        if (isUpdate) {
            // Update
            await prisma.sAL_COUNTRYCITY_MASTER.update({
                where: {
                    CountryCode_ParentCode_Branchid_Compid: {
                        CountryCode: newCode,
                        ParentCode: pCode,
                        Branchid: branchId,
                        Compid: compId
                    }
                },
                data: {
                    CountryName: name.trim(),
                    userid: userId
                }
            });
        } else {
            // Generate new Code
            const maxResult: any[] = await prisma.$queryRawUnsafe(`
                SELECT ISNULL(MAX(CountryCode), 0) + 1 as maxCode 
                FROM SAL_COUNTRYCITY_MASTER 
                WHERE ParentCode = ${pCode}
            `);
            newCode = maxResult[0].maxCode;

            await prisma.sAL_COUNTRYCITY_MASTER.create({
                data: {
                    CountryCode: newCode,
                    CountryName: name.trim(),
                    ParentCode: pCode,
                    userid: userId,
                    Branchid: branchId,
                    Compid: compId
                }
            });
        }

        return NextResponse.json({ success: true, code: newCode, message: isUpdate ? "Record Updated" : "Record Saved" });
    } catch (error: any) {
        console.error("Save Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

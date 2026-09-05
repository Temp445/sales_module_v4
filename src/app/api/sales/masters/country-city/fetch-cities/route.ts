import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const searchParams = req.nextUrl.searchParams;
        const parentCodeStr = searchParams.get("parentCode");

        if (!parentCodeStr) {
            return NextResponse.json({ error: "Missing parentCode" }, { status: 400 });
        }

        const parentCode = parseInt(parentCodeStr, 10);
        if (isNaN(parentCode)) {
            return NextResponse.json({ error: "Invalid parentCode" }, { status: 400 });
        }

        const cities = await prisma.sAL_COUNTRYCITY_MASTER.findMany({
            where: {
                ParentCode: parentCode,
                Branchid: session.branchId,
                Compid: session.compId
            },
            select: {
                CountryCode: true,
                CountryName: true,
            },
            orderBy: {
                CountryName: 'asc'
            }
        });

        const options = cities.map(c => ({
            value: c.CountryCode,
            label: c.CountryName
        }));

        return NextResponse.json(options);
    } catch (error: any) {
        console.error("Fetch Cities Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const countries = await prisma.sAL_COUNTRYCITY_MASTER.findMany({
            where: {
                ParentCode: 0,
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

        // Map to format SearchableSelect expects, though front-end handles mapping often
        const options = countries.map(c => ({
            value: c.CountryCode,
            label: c.CountryName
        }));

        return NextResponse.json(options);
    } catch (error: any) {
        console.error("Fetch Countries Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

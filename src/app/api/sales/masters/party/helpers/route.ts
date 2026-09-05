import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const { searchParams } = new URL(req.url);
        const action = searchParams.get("action"); // 'country', 'state', 'city'

        if (action === 'country') {
            const countries = await prisma.$queryRaw`SELECT CountryCode, CountryName FROM COM_CountryCity_Master_Old WHERE ParentCode = 0 ORDER BY CountryName`;
            return NextResponse.json(countries);
        }

        if (action === 'state') {
            const countryCode = searchParams.get("countryCode");
            if (!countryCode) return NextResponse.json({ error: "countryCode required" }, { status: 400 });
            
            const states = await prisma.$queryRaw`SELECT DISTINCT CountryName AS StateName FROM COM_CountryCity_Master_Old WHERE ParentCode = ${parseInt(countryCode, 10)} ORDER BY CountryName`;
            return NextResponse.json(states);
        }

        if (action === 'city') {
            const stateCode = searchParams.get("stateCode");
            if (!stateCode) return NextResponse.json({ error: "stateCode required" }, { status: 400 });
            
            const cities = await prisma.$queryRaw`SELECT CountryCode AS CityCode, CountryName AS CityName FROM COM_CountryCity_Master_Old WHERE ParentCode = ${parseInt(stateCode, 10)} ORDER BY CountryName`;
            return NextResponse.json(cities);
        }

        return NextResponse.json({ error: "Invalid action" }, { status: 400 });
    } catch (error: any) {
        console.error("Fetch Helper Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

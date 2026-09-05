import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";

export async function GET(req: NextRequest) {
    try {
        // Prisma ignores COM_Currency_Master because it lacks a primary key.
        // We must use $queryRaw to fetch the currencies.
        const currencies = await prisma.$queryRaw<any[]>`
            SELECT CurrencyCode, CurrencyName 
            FROM COM_Currency_Master 
            ORDER BY CurrencyName ASC
        `;
        
        const formattedCurrencies = currencies.map(c => ({
            CurrencyCode: c.CurrencyCode?.trim() || '',
            CurrencyName: c.CurrencyName?.trim() || ''
        }));

        return NextResponse.json(formattedCurrencies);
    } catch (error: any) {
        console.error("Fetch Currencies Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

import { NextRequest, NextResponse } from "next/server";
import { PrismaClient } from "@prisma/client";
import { getSession } from "@/lib/session";

const prisma = new PrismaClient();

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        // Debug: log session values to identify mismatch
        console.log("[item/search] session compId:", session.compId, "branchId:", session.branchId);

        const { searchParams } = new URL(req.url);
        const searchTerm = searchParams.get("q") || "";

        const items = await prisma.iNV_MATERIAL_MASTER.findMany({
            where: {
                Compid: session.compId,
                BranchId: session.branchId,
                Status: true,
                ...(searchTerm ? {
                    OR: [
                        { ItemCode: { contains: searchTerm } },
                        { ItemDescription: { contains: searchTerm } }
                    ]
                } : {})
            },
            select: {
                ItemCode: true,
                ItemDescription: true,
                Rate: true,
                DrawingNo: true,
                ProductCategoryCode: true
            },
            orderBy: {
                ItemCode: 'asc'
            },
            take: 200
        });

        console.log("[item/search] found:", items.length, "items");

        // Format for frontend mapping
        const formattedItems = items.map(item => ({
            value: item.ItemCode,
            label: `${item.ItemCode} - ${item.ItemDescription}`,
            description: item.ItemDescription || "",
            model: "",
            rate: item.Rate || 0,
            drawingNo: item.DrawingNo || ""
        }));

        return NextResponse.json(formattedItems, { status: 200 });

    } catch (error: any) {
        console.error("Error fetching items:", error);
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    } finally {
        await prisma.$disconnect();
    }
}

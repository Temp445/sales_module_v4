import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

async function run() {
    try {
        const items = await prisma.iNV_MATERIAL_MASTER.findMany({
            where: {
                Compid: "07",
                BranchId: "00",
                Status: true
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
            }
        });

        // Format for frontend mapping
        const formattedItems = items.map((item: any) => ({
            value: item.ItemCode,
            label: `${item.ItemCode} - ${item.ItemDescription}`,
            description: item.ItemDescription || "",
            model: "",
            rate: item.Rate || 0,
            drawingNo: item.DrawingNo || ""
        }));

        console.log("Success! Items:", formattedItems.length);
        if (formattedItems.length > 0) {
            console.log(formattedItems[0]);
        }
    } catch (error: any) {
        console.error("Error fetching items:", error);
    } finally {
        await prisma.$disconnect();
    }
}
run();

const fs = require('fs');

const schemaPath = 'e:\\Nivin_Sync\\ERP\\next_sales_app_v4\\prisma\\schema.prisma';
const schemaContent = fs.readFileSync(schemaPath, 'utf8');

const regex = /model SAL_Jobwork_ItemMaster \{[\s\S]*?\}/g;
const match = schemaContent.match(regex);

if (match) {
    console.log(match[0]);
} else {
    // maybe case insensitive or different spelling
    const allModels = Array.from(schemaContent.matchAll(/model\s+(\w+)\s*\{/g)).map(m => m[1]);
    const matches = allModels.filter(m => m.toLowerCase().includes('customer_itemmaster'));
    console.log("Not exactly matched. Found:", matches);
    if (matches.length > 0) {
        const regex2 = new RegExp(`model ${matches[0]} \\{[\\s\\S]*?\\}`, 'g');
        console.log(schemaContent.match(regex2)[0]);
    }
}

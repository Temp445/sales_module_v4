const fs = require('fs');
const path = require('path');

const files = [
    'e:/Nivin_Sync/ERP/next_sales_app_v4/src/app/(erp)/sales/masters/customer-item/page.tsx',
    'e:/Nivin_Sync/ERP/next_sales_app_v4/src/app/(erp)/sales/masters/jobwork-item/page.tsx'
];

for (const file of files) {
    let content = fs.readFileSync(file, 'utf8');
    
    // Find all <input> and <select> className occurrences inside the grid and make sure they have text-gray-900
    // The current pattern is `className="w-... p-2 border border-gray-200 rounded-md text-sm"`
    // Or similar. Let's just replace `text-sm"` with `text-sm text-gray-900 font-medium"`
    
    content = content.replace(/text-sm"/g, 'text-sm text-gray-900 font-medium"');
    content = content.replace(/text-sm bg-white"/g, 'text-sm text-gray-900 font-medium bg-white"');
    content = content.replace(/text-sm bg-gray-100"/g, 'text-sm text-gray-900 font-medium bg-gray-100"');
    
    // We also need to fix if there's any duplication
    content = content.replace(/text-gray-900 font-medium text-gray-900 font-medium/g, 'text-gray-900 font-medium');

    fs.writeFileSync(file, content, 'utf8');
    console.log(`Updated ${file}`);
}

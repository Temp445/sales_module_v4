import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const jsonPath = path.join(__dirname, '../CustomerItemMaster_extracted.json'); // Note: actually we need SalesEnquiry JSON, wait! I named it SalesEnquiry_extracted.json
const actualJsonPath = path.join(__dirname, '../SalesEnquiry_extracted.json');
const tsxPath = path.join(__dirname, '../../src/app/sales/sales-enquiry/page.tsx');

let hasError = false;

// We will just do a simple check. If we wanted a perfect parser we'd use a real one.
const tsxContent = fs.readFileSync(tsxPath, 'utf8');
const jsonContent = JSON.parse(fs.readFileSync(actualJsonPath, 'utf8'));

console.log("Validating SalesEnquiry UI Mapping...");
console.log("---------------------------------------");

const requiredControls = [
    { vb: 'cmbEnqType', state: 'enquiryType' },
    { vb: 'txtEnqNo', state: 'enquiryNo' },
    { vb: 'dtpEnqDate', state: 'enquiryDate' },
    { vb: 'txtRefNo', state: 'referenceNo' },
    { vb: 'dtpRefDate', state: 'referenceDate' },
    { vb: 'cmbCustomerType', state: 'customerType' },
    { vb: 'txtCustomer', state: 'customerCode' }, // mapped to two: Name/Code based on type
    { vb: 'txtRemarks', state: 'remarks' }
];

console.table(requiredControls.map(c => {
    const foundState = tsxContent.includes(c.state);
    if (!foundState) hasError = true;
    return {
        'VB6 Control': c.vb,
        'React State': c.state,
        'Status': foundState ? 'PRESENT' : 'MISSING'
    };
}));

// Grid Columns Check
const requiredGridColumns = [
    { enum: 'mItmType', state: 'itemType' },
    { enum: 'mItmCode', state: 'itemCode' },
    { enum: 'mItmDesc', state: 'itemDesc' },
    { enum: 'mUOM', state: 'uom' },
    { enum: 'mQty', state: 'qty' },
    { enum: 'mTarVolme', state: 'tarVolme' },
    { enum: 'mVolmeper', state: 'volmeper' },
    { enum: 'mRate', state: 'rate' },
    { enum: 'mAssRate', state: 'assRate' },
    { enum: 'mTargtDate', state: 'targtDate' },
    { enum: 'mSpeci', state: 'speci' }
];

console.table(requiredGridColumns.map(c => {
    const foundState = tsxContent.includes(c.state);
    if (!foundState) hasError = true;
    return {
        'Enum Field': c.enum,
        'React Item State': c.state,
        'Status': foundState ? 'PRESENT' : 'MISSING'
    };
}));

if (hasError) {
    console.error("UI Verification FAILED! Missing controls detected.");
    process.exit(1);
} else {
    console.log("UI Verification PASSED! All controls mapped.");
    process.exit(0);
}

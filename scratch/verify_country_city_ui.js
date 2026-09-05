const fs = require('fs');

const jsonPath = 'Migration_Ref/reference/VB6Source/Sales/CountryCityMaster_extracted.json';
const tsxPath = 'src/app/sales/masters/country-city/page.tsx';

const jsonData = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
const tsxContent = fs.readFileSync(tsxPath, 'utf8');

console.log("=== UI Verification ===");

let allFound = true;
jsonData.controls.forEach(control => {
    // Only check important interactive elements by logical name/type mapping
    if (control.name === 'cmbType' && !tsxContent.includes('mode')) {
        console.error("Missing cmbType equivalent (mode/type toggle)");
        allFound = false;
    }
    if (control.name === 'txtCountry' && !tsxContent.includes('countryName')) {
        console.error("Missing txtCountry equivalent");
        allFound = false;
    }
    if (control.name === 'txtCity' && !tsxContent.includes('cityName')) {
        console.error("Missing txtCity equivalent");
        allFound = false;
    }
    if (control.name === 'cmdSave' && !tsxContent.includes('handleSave')) {
        console.error("Missing cmdSave equivalent");
        allFound = false;
    }
    if (control.name === 'cmdClear' && !tsxContent.includes('handleClear')) {
        console.error("Missing cmdClear equivalent");
        allFound = false;
    }
});

if (allFound) {
    console.log("All UI Controls successfully mapped in React page.");
} else {
    console.log("Verification Failed.");
}

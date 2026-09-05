import fs from "fs";

function runVerification() {
    const jsonStr = fs.readFileSync("e:\\Nivin_Sync\\ERP\\next_sales_app_v4\\Migration_Ref\\CustomerJobworkMaster_extracted.json", "utf-8");
    let tsxStr;
    try {
        tsxStr = fs.readFileSync("e:\\Nivin_Sync\\ERP\\next_sales_app_v4\\src\\app\\(erp)\\sales\\masters\\jobwork-item\\page.tsx", "utf-8");
    } catch (e) {
        console.error("MISSING: page.tsx not found yet. Run this after UI creation.");
        process.exit(1);
    }

    const json = JSON.parse(jsonStr);

    let allPassed = true;
    const requiredState = [
        "customerCode", "defaultFormula", "items"
    ];

    console.log("=== React Census vs VB6 Census ===");
    for (const state of requiredState) {
        if (!tsxStr.includes(`[${state}, set${state.charAt(0).toUpperCase() + state.slice(1)}]`)) {
            console.error(`MISSING: React state '${state}' not found.`);
            allPassed = false;
        } else {
            console.log(`PRESENT: React state '${state}'`);
        }
    }

    // specific field assertions inside the table based on JSON columns
    const columns = json.controls.find(c => c.name === "sprItem").columns;
    for (const col of columns) {
        let expectedField = col.name;
        
        if (!tsxStr.includes(expectedField)) {
            console.error(`MISSING: Column mapping '${expectedField}' not found in TSX.`);
            allPassed = false;
        } else {
            console.log(`PRESENT: Column '${expectedField}'`);
        }
    }

    // Enforce Rule 12 Input types
    if (!tsxStr.includes("SearchableSelect")) {
         console.error("WRONG TYPE: Missing SearchableSelect for F2 Lookups (Rule 12).");
         allPassed = false;
    }

    if (!allPassed) {
        console.error("Verification FAILED. Please fix the UI.");
        process.exit(1);
    } else {
        console.log("Verification PASSED!");
    }
}

runVerification();

// Using built-in fetch

async function runTest() {
    console.log("Starting Customer Item Master Save API Test...");

    const payload = {
        compId: "07",
        branchId: "00",
        userId: "admin",
        customerCode: "5", // Using test code 5
        items: [
            {
                itemCode: { value: "ITM001", label: "Test Item 1" },
                model: "MOD-A",
                issuedNo: "IS-01",
                partNo: "CUST-PART-1",
                rate: "100.50",
                assessableRate: "90.00",
                validFrom: "2024-01-01",
                validTo: "2025-12-31",
                formulaCode: { value: "1", label: "Default Formula" },
                qcCheck: "Yes",
                status: "Active"
            }
        ]
    };

    try {
        const response = await fetch("http://localhost:3000/api/sales/masters/customer-item/save", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        const data = await response.json();
        
        console.log(`Status: ${response.status}`);
        console.log(data);

        if (response.status !== 200) {
            console.error("Test Failed!");
            process.exit(1);
        } else {
            console.log("Test Passed!");
        }

    } catch (e) {
        console.error("Error executing fetch:", e);
        process.exit(1);
    }
}

runTest();

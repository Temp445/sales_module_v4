async function runTest() {
    console.log("Starting Customer Jobwork Master Save API Test...");

    const payload = {
        compId: "07",
        branchId: "00",
        userId: "admin",
        customerCode: "5", // Using test code 5
        items: [
            {
                mItemGroupKey: "1",
                mrecdItem: { value: "RECD001", label: "Recd Item 1" },
                msentItem: { value: "SENT001", label: "Sent Item 1" },
                mRelation: "CHI",
                mPartNo: "PART-JW-01",
                mRatePer: "10",
                mRate: "500",
                mAssRate: "450",
                mValidFrom: "2024-01-01",
                mValidTo: "2025-12-31",
                mQCCheck: "Yes",
                mFrmCode: { value: "1", label: "Default Formula" },
                mStatus: "Active"
            }
        ]
    };

    try {
        const response = await fetch("http://localhost:3000/api/sales/masters/jobwork-item/save", {
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

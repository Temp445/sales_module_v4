// Native fetch

async function runFetchTest() {
    try {
        const response = await fetch("http://localhost:3000/api/sales/masters/customer-item/fetch?compId=07&branchId=00&customerCode=5");
        const data = await response.json();
        
        console.log(`Status: ${response.status}`);
        console.log("Data:");
        console.log(JSON.stringify(data, null, 2));

    } catch (e) {
        console.error("Error executing fetch:", e);
        process.exit(1);
    }
}

runFetchTest();

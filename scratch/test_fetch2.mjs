async function runTest() {
    try {
        const partyRes = await fetch("http://localhost:3000/api/sales/masters/party/search?partyType=1&limit=50");
        const text = await partyRes.text();
        console.log("Raw response:", text.substring(0, 500));
    } catch (e) {
        console.error("Error:", e);
    }
}
runTest();

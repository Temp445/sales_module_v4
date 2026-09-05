async function debug() {
    try {
        const res = await fetch('http://localhost:3000/api/sales/masters/country-city/fetch-countries');
        const text = await res.text();
        console.log("Status:", res.status);
        console.log("Response:", text.substring(0, 500));
    } catch (e) {
        console.error("Fetch failed:", e);
    }
}
debug();

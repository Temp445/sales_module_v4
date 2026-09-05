const fs = require('fs');

const filePath = process.argv[2];
if (!filePath) {
    console.error("Please provide the path to an .frx file.");
    process.exit(1);
}
const buffer = fs.readFileSync(filePath);

// Extract printable strings of length >= 3
const strings = [];
let currentString = '';

for (let i = 0; i < buffer.length; i++) {
    const charCode = buffer[i];
    // Printable ASCII range (32 to 126)
    if (charCode >= 32 && charCode <= 126) {
        currentString += String.fromCharCode(charCode);
    } else {
        if (currentString.length >= 3) {
            strings.push(currentString);
        }
        currentString = '';
    }
}
if (currentString.length >= 3) {
    strings.push(currentString);
}

// Group them and look for patterns that look like 3-letter codes followed by descriptions
console.log("Found Strings:");
const uniqueStrings = [...new Set(strings)];
uniqueStrings.forEach(s => {
    if (s.trim().length > 0) {
        console.log(s);
    }
});

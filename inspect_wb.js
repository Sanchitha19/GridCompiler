const XLSX = require('xlsx');
const path = require('path');
const wb = XLSX.readFile(path.join(__dirname, 'demo', 'inventory.xlsx'));
const ws = wb.Sheets['Summary'];

console.log('--- RAW EXTRACTION (Summary Sheet) ---');
Object.keys(ws).filter(k => !k.startsWith('!')).forEach(addr => {
    console.log(`[RAW EXTRACTION] Summary!${addr} =>`, JSON.stringify(ws[addr]));
});

console.log('\n--- Specific Cells ---');
console.log('Summary B2:', ws['B2']);
console.log('Summary B3:', ws['B3']);

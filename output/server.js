
const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const multer = require('multer');
const fs = require('fs');

const app = express();
const PORT = 3000;
const upload = multer({ storage: multer.memoryStorage() });

app.use(bodyParser.json());
app.use(express.static(__dirname));

// In-memory store
let data = {}; 

// Load initial data from Excel if it exists
function loadInitialData() {
    const excelPath = path.join(__dirname, '../demo/inventory.xlsx');
    if (fs.existsSync(excelPath)) {
        console.log(`Loading initial data from ${excelPath}`);
        const workbook = XLSX.readFile(excelPath);
        const sheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        rows.forEach((row, index) => {
            const rowNo = index + 2;
            data[`Products!A${rowNo}`] = row.ID || (101 + index);
            data[`Products!B${rowNo}`] = row.Name || '';
            data[`Products!C${rowNo}`] = row.Category || autoDetectCategory(row.Name || '');
            data[`Products!D${rowNo}`] = row.Price || 0;
            data[`Products!E${rowNo}`] = row.Stock || 0;
        });
        evaluateFormulas();
    } else {
        console.log("No demo/inventory.xlsx found. Initializing empty.");
    }
}

// Category rules engine
const categoryRules = {
  Electronics: ['laptop','computer','pc','monitor','phone','tablet','keyboard','mouse','camera','tv','router','gpu','cpu'],
  Furniture: ['chair','desk','table','sofa','couch','bed','shelf','cabinet','stool'],
  Clothing: ['shirt','pant','dress','jacket','shoe','sock','hat','belt','watch','bag'],
  Stationery: ['pen','pencil','notebook','paper','folder','stapler','tape','glue'],
  Food: ['rice','wheat','sugar','salt','oil','milk','bread','flour','spice','tea','coffee'],
  Sports: ['ball','bat','racket','glove','helmet','jersey','shoe','gym','cycle']
};

function autoDetectCategory(productName) {
  if (!productName) return 'Other';
  const name = productName.toLowerCase();
  for (const [category, keywords] of Object.entries(categoryRules)) {
    if (keywords.some(k => name.includes(k))) return category;
  }
  return 'Other';
}

function evaluateFormulas() {
    let totalStock = 0, totalValue = 0;
    const productRows = new Set();
    Object.keys(data).forEach(k => { if (k.startsWith('Products!')) { const row = k.match(/\d+$/)?.[0]; if (row) productRows.add(row); } });
    productRows.forEach(row => {
        const stock = Number(data[`Products!E${row}`]) || 0;
        const price = Number(data[`Products!D${row}`]) || 0;
        totalStock += stock; totalValue += (stock * price);
    });
    data["Summary!B2"] = totalStock; data["Summary!B3"] = totalValue; data["Summary!B6"] = totalStock;
    let maxVal = -1, maxProd = "---";
    productRows.forEach(row => {
        const stock = Number(data[`Products!E${row}`]) || 0, price = Number(data[`Products!D${row}`]) || 0, val = stock * price;
        if (val > maxVal) { maxVal = val; maxProd = data[`Products!B${row}`] || "Unknown"; }
    });
    data["Summary!B4"] = maxProd;
}

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.get('/api/detect-category', (req, res) => {
    res.json({ category: autoDetectCategory(req.query.name || '') });
});

app.get('/api/:sheet', (req, res) => {
    evaluateFormulas();
    const sheet = req.params.sheet;
    const sheetData = Object.keys(data).filter(k => k.startsWith(sheet + '!')).reduce((obj, key) => {
        obj[key.split('!')[1]] = data[key]; return obj;
    }, {});
    res.json(sheetData);
});

app.post('/api/cell', (req, res) => {
    const { address, value, type } = req.body;
    data[address] = value === null ? null : (type === 'NUMBER' ? Number(value) : value);
    evaluateFormulas();
    res.json({ success: true });
});

app.post('/api/import', upload.single('file'), (req, res) => {
    try {
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        data = {}; 
        rows.forEach((row, index) => {
            const rowNo = index + 2;
            data[`Products!A${rowNo}`] = row.ID || (101 + index);
            data[`Products!B${rowNo}`] = row.Name || '';
            data[`Products!C${rowNo}`] = row.Category || autoDetectCategory(row.Name || '');
            data[`Products!D${rowNo}`] = row.Price || 0;
            data[`Products!E${rowNo}`] = row.Stock || 0;
        });
        evaluateFormulas();
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: 'Import failed' });
    }
});

app.get('/api/download', (req, res) => {
    evaluateFormulas();
    const wb = XLSX.utils.book_new();
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);

    const productRows = new Set();
    Object.keys(data).forEach(k => { if (k.startsWith('Products!')) productRows.add(k.match(/\d+$/)[0]); });
    const pAOA = [["ID", "Name", "Category", "Price", "Stock"]];
    Array.from(productRows).sort((a,b)=>a-b).forEach(r => pAOA.push([data[`Products!A${r}`], data[`Products!B${r}`], data[`Products!C${r}`], data[`Products!D${r}`], data[`Products!E${r}`]]));
    const ws1 = XLSX.utils.aoa_to_sheet(pAOA);
    ws1['!cols'] = [{wch: 8}, {wch: 25}, {wch: 15}, {wch: 12}, {wch: 10}];
    XLSX.utils.book_append_sheet(wb, ws1, "Products");

    const sAOA = [["METRIC", "VALUE"], ["Total Units", data["Summary!B2"]], ["Total Value", data["Summary!B3"]], ["Top Product", data["Summary!B4"]]];
    const ws2 = XLSX.utils.aoa_to_sheet(sAOA);
    ws2['!cols'] = [{wch: 25}, {wch: 20}];
    XLSX.utils.book_append_sheet(wb, ws2, "Summary");

    const lowStock = Array.from(productRows).filter(r => data[`Products!E${r}`] < 10).map(r => data[`Products!B${r}`]).join(', ');
    const aAOA = [
        ["ANALYTICS REPORT", ""], ["Export Timestamp", new Date().toLocaleString()],
        ["Total Products", productRows.size], ["Total Categories", new Set(Array.from(productRows).map(r => data[`Products!C${r}`])).size],
        ["Low Stock Items", lowStock || "None"], ["Highest Value Product", data["Summary!B4"]]
    ];
    const ws3 = XLSX.utils.aoa_to_sheet(aAOA);
    ws3['!cols'] = [{wch: 25}, {wch: 60}];
    XLSX.utils.book_append_sheet(wb, ws3, "Analytics");

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="excel2app_export_${ts}.xlsx"`);
    res.send(buf);
});

loadInitialData();
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));

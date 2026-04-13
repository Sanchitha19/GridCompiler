
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

app.get('/api/download', (req, res) => {
    try {
        const XLSX = require('xlsx');
        
        const wb = XLSX.utils.book_new();
        
        const productRows = [['ID','Name','Category','Price','Stock']];
        const allKeys = Object.keys(data).filter(k => k.startsWith('Products!A'));
        const rowNums = allKeys.map(k => parseInt(k.replace('Products!A',''))).filter(n => n >= 2).sort((a,b) => a-b);
        
        rowNums.forEach(i => {
            productRows.push([
                data['Products!A'+i] || '',
                data['Products!B'+i] || '',
                data['Products!C'+i] || '',
                Number(data['Products!D'+i]) || 0,
                Number(data['Products!E'+i]) || 0
            ]);
        });
        
        const wsProducts = XLSX.utils.aoa_to_sheet(productRows);
        XLSX.utils.book_append_sheet(wb, wsProducts, 'Products');
        
        const totalStock = rowNums.reduce((sum,i) => sum + (Number(data['Products!E'+i])||0), 0);
        const totalValue = rowNums.reduce((sum,i) => sum + ((Number(data['Products!D'+i])||0) * (Number(data['Products!E'+i])||0)), 0);
        
        const summaryRows = [
            ['Metric', 'Value'],
            ['Total Stock', totalStock],
            ['Total Value', totalValue]
        ];
        const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
        XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');
        
        const buffer = XLSX.write(wb, { 
            bookType: 'xlsx', 
            type: 'buffer'
        });
        
        console.log('Buffer length:', buffer.length);
        console.log('First 4 bytes:', buffer[0], buffer[1], buffer[2], buffer[3]);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="export.xlsx"');
        res.setHeader('Content-Length', buffer.length);
        res.end(buffer);
        
    } catch(err) {
        console.error('Download error:', err);
        res.status(500).json({ error: err.message });
    }
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

loadInitialData();
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));

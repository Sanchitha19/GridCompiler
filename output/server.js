
const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

app.use(bodyParser.json());
app.use(express.static(__dirname));

// In-memory store
let data = {
    "Products!A2": 101, "Products!B2": "Laptop", "Products!C2": "Electronics", "Products!D2": 1200, "Products!E2": 15,
    "Products!A3": 102, "Products!B3": "Chair",  "Products!C3": "Furniture",   "Products!D3": 150,  "Products!E3": 40,
    "Products!A4": 103, "Products!B4": "Desk",   "Products!C4": "Furniture",   "Products!D4": 300,  "Products!E4": 10,
    "Products!A5": 104, "Products!B5": "Monitor","Products!C5": "Electronics", "Products!D5": 250,  "Products!E5": 25
}; 

// Dynamic formula evaluator
function evaluateFormulas() {
    let totalStock = 0;
    let totalValue = 0;

    const productRows = new Set();
    Object.keys(data).forEach(k => {
        if (k.startsWith('Products!')) {
            const row = k.match(/\d+$/)?.[0];
            if (row) productRows.add(row);
        }
    });

    productRows.forEach(row => {
        const stock = Number(data[`Products!E${row}`]) || 0;
        const price = Number(data[`Products!D${row}`]) || 0;
        totalStock += stock;
        totalValue += (stock * price);
    });

    data["Summary!B2"] = totalStock;
    data["Summary!B3"] = totalValue;
    data["Summary!B6"] = totalStock;

    let maxVal = -1;
    let maxProd = "---";
    productRows.forEach(row => {
        const stock = Number(data[`Products!E${row}`]) || 0;
        const price = Number(data[`Products!D${row}`]) || 0;
        const val = stock * price;
        if (val > maxVal) {
            maxVal = val;
            maxProd = data[`Products!B${row}`] || "Unknown";
        }
    });
    data["Summary!B4"] = maxProd;
}

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/api/:sheet', (req, res) => {
    const sheet = req.params.sheet;
    evaluateFormulas();
    const sheetData = Object.keys(data)
        .filter(k => k.startsWith(sheet + '!'))
        .reduce((obj, key) => {
            obj[key.split('!')[1]] = data[key];
            return obj;
        }, {});
    res.json(sheetData);
});

app.post('/api/cell', (req, res) => {
    const { address, value, type } = req.body;
    data[address] = value === null ? null : (type === 'NUMBER' ? Number(value) : value);
    evaluateFormulas();
    res.json({ success: true, newValue: data[address] });
});

// Excel Download Route
app.get('/api/download', (req, res) => {
    evaluateFormulas();
    const wb = XLSX.utils.book_new();

    const productRows = new Set();
    Object.keys(data).forEach(k => {
        if (k.startsWith('Products!')) {
            const row = k.match(/\d+$/)?.[0];
            if (row) productRows.add(row);
        }
    });

    const productsAOA = [["ID", "Name", "Category", "Price", "Stock"]];
    Array.from(productRows).sort((a, b) => a - b).forEach(row => {
        productsAOA.push([
            data[`Products!A${row}`],
            data[`Products!B${row}`],
            data[`Products!C${row}`],
            data[`Products!D${row}`],
            data[`Products!E${row}`]
        ]);
    });
    const productsWS = XLSX.utils.aoa_to_sheet(productsAOA);
    XLSX.utils.book_append_sheet(wb, productsWS, "Products");

    const summaryAOA = [
        ["Metric", "Value"],
        ["Total Stock Units", data["Summary!B2"]],
        ["Total Inventory Value", data["Summary!B3"]],
        ["Most Valuable Product", data["Summary!B4"]]
    ];
    const summaryWS = XLSX.utils.aoa_to_sheet(summaryAOA);
    XLSX.utils.book_append_sheet(wb, summaryWS, "Summary");

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="excel2app_export.xlsx"');
    res.send(buf);
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});

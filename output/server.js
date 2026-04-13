
const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');

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

    // Scan all data keys to find Products sheet cells
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

    // Dynamic XLOOKUP for Summary!B4 (Most valuable product)
    let maxVal = -1;
    let maxProd = "---";
    productRows.forEach(row => {
        const stock = Number(data[`Products!E${row}`]) || 0;
        const price = Number(data[`Products!D${row}`]) || 0;
        if (stock * price > maxVal) {
            maxVal = stock * price;
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
    evaluateFormulas(); // Compute before returning
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
    console.log(`Updating ${address} to ${value}`);
    data[address] = type === 'NUMBER' ? Number(value) : value;
    evaluateFormulas(); // Recompute
    res.json({ success: true, newValue: data[address] });
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});

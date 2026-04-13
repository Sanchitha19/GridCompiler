const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const demoDir = path.join(__dirname, 'demo');
if (!fs.existsSync(demoDir)) {
    fs.mkdirSync(demoDir);
}

const wb = XLSX.utils.book_new();

// Sheet 1: Products
const productsData = [
    ['ID', 'Name', 'Category', 'Price', 'Stock'],
    [101, 'Laptop', 'Electronics', 1200, 15],
    [102, 'Chair', 'Furniture', 150, 40],
    [103, 'Desk', 'Furniture', 300, 10],
    [104, 'Monitor', 'Electronics', 250, 25],
];
const wsProducts = XLSX.utils.aoa_to_sheet(productsData);
XLSX.utils.book_append_sheet(wb, wsProducts, 'Products');

// Sheet 2: Categories
const categoriesData = [
    ['Category Name', 'Description'],
    ['Electronics', 'Gadgets and hardware'],
    ['Furniture', 'Office and home furniture'],
];
const wsCategories = XLSX.utils.aoa_to_sheet(categoriesData);
XLSX.utils.book_append_sheet(wb, wsCategories, 'Categories');

// Sheet 3: Summary
const wsSummary = {
    '!ref': 'A1:B6',
    'A1': { v: 'Metric', t: 's' },
    'B1': { v: 'Value', t: 's' },
    'A2': { v: 'Total Stock', t: 's' },
    'B2': { f: 'SUM(Products!E2:E5)', t: 'n', v: 0 },
    'A3': { v: 'Total Value', t: 's' },
    'B3': { f: 'Products!D2*Products!E2 + Products!D3*Products!E3 + Products!D4*Products!E4 + Products!D5*Products!E5', t: 'n', v: 0 },
    'A4': { v: 'Unknown Function Test', t: 's' },
    'B4': { f: 'XLOOKUP(101, Products!A2:A5, Products!B2:B5)', t: 'n', v: 0 },
    'A5': { v: 'Undefined Ref Test', t: 's' },
    'B5': { f: 'OtherSheet!A1 + 10', t: 'n', v: 0 },
    // CSE Opportunity: same formula as B2
    'A6': { v: 'Total Stock (Again)', t: 's' },
    'B6': { f: 'SUM(Products!E2:E5)', t: 'n', v: 0 }
};

// Add a sheet with a cycle
const wsCycles = {
    '!ref': 'A1:B1',
    'A1': { f: 'B1 + 1', t: 'n', v: 0 },
    'B1': { f: 'A1 + 2', t: 'n', v: 0 }
};

XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');
XLSX.utils.book_append_sheet(wb, wsCycles, 'Cycles');

const filePath = path.join(demoDir, 'inventory.xlsx');
XLSX.writeFile(wb, filePath);
console.log(`Created ${filePath} with CSE opportunity.`);

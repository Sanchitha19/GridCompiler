const XLSX = require('xlsx');
const path = require('path');
const { tokenize } = require('./src/lexer');
const { SemanticAnalyzer } = require('./src/analyzer');
const { IRGenerator } = require('./src/ir');

function main() {
    const filePath = path.join(__dirname, 'demo', 'inventory.xlsx');
    
    console.log(`Loading ${filePath}...`);
    let workbook;
    try {
        workbook = XLSX.readFile(filePath);
    } catch (e) {
        console.error(`Could not read file: ${e.message}`);
        return;
    }

    console.log('--- Phase 1: Lexing ---');
    const tokens = tokenize(workbook);

    console.log('--- Phase 2 & 3: Parsing & Semantic Analysis ---');
    const analyzer = new SemanticAnalyzer(tokens);
    const results = analyzer.analyze();

    console.log('--- Phase 4: IR Generation & Optimization ---');
    const irGen = new IRGenerator(results.symbolTable, results.dependencies);
    const irResults = irGen.generate();

    irGen.printIR();
}

main();

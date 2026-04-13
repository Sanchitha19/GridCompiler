const XLSX = require('xlsx');
const path = require('path');
const { tokenize } = require('./src/lexer');
const { SemanticAnalyzer } = require('./src/analyzer');

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

    console.log('Tokenizing workbook (Phase 1)...');
    const tokens = tokenize(workbook);
    console.log(`Produced ${tokens.length} tokens.`);

    console.log('Analyzing semantics (Phase 2 & 3)...');
    const analyzer = new SemanticAnalyzer(tokens);
    const results = analyzer.analyze();

    // 1. Full Symbol Table
    console.log('\n=== FULL SYMBOL TABLE ===');
    console.log('CELL'.padEnd(15) + ' | ' + 'TYPE'.padEnd(10) + ' | ' + 'VALUE/AST');
    console.log('-'.repeat(60));
    for (const [addr, entry] of results.symbolTable) {
        const typeStr = entry.inferredType.padEnd(10);
        const displayVal = entry.tokenType === 'FORMULA' ? entry.value : entry.value;
        console.log(`${addr.padEnd(15)} | ${typeStr} | ${displayVal}`);
    }

    // 2. Dependency Graph Summary
    let totalEdges = 0;
    for (const [source, targets] of results.dependencies) {
        totalEdges += targets.size;
    }
    console.log('\n=== DEPENDENCY GRAPH SUMMARY ===');
    console.log(`Total Nodes: ${results.symbolTable.size}`);
    console.log(`Active Dependencies (Nodes with deps): ${results.dependencies.size}`);
    console.log(`Total Edges (A -> B): ${totalEdges}`);

    // 3. Any Cycles Detected
    const cycles = results.errors.filter(e => e.type === 'CyclicReference');
    console.log('\n=== CYCLES DETECTED ===');
    if (cycles.length === 0) {
        console.log('None.');
    } else {
        cycles.forEach(c => console.log(`[!] ${c.details}`));
    }

    // 4. Semantic Error Report
    analyzer.printSemanticReport();
}

main();

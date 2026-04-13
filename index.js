const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const { tokenize } = require('./src/lexer');
const { SemanticAnalyzer } = require('./src/analyzer');
const { IRGenerator } = require('./src/ir');
const { CodeGenerator } = require('./src/codegen');

function main() {
    const args = process.argv.slice(2);
    
    if (args.length < 2) {
        printUsage();
        return;
    }

    const command = args[0];
    const filePath = args[1];

    if (command !== 'compile') {
        process.error(`Unknown command: ${command}`);
        printUsage();
        return;
    }

    if (!fs.existsSync(filePath)) {
        process.error(`File not found: ${filePath}`);
        return;
    }

    console.log(`\n🚀 Excel2App Compiler: Compiling ${filePath}...\n`);
    
    try {
        // Phase 1: Lexing
        console.log('--- Phase 1: Lexing ---');
        const workbook = XLSX.readFile(filePath);
        const tokens = tokenize(workbook);
        console.log(`✓ Processed ${tokens.length} tokens.`);

        // Phase 2 & 3: Parsing & Semantic Analysis
        console.log('\n--- Phase 2 & 3: Semantic Analysis ---');
        const analyzer = new SemanticAnalyzer(tokens);
        const semanticResults = analyzer.analyze();
        analyzer.printSemanticReport();

        // Phase 4: IR & Optimization
        console.log('\n--- Phase 4: IR & Optimization ---');
        const irGen = new IRGenerator(semanticResults.symbolTable, semanticResults.dependencies);
        const irResults = irGen.generate();
        irGen.printIR();

        // Phase 5: Code Generation
        console.log('\n--- Phase 5: Code Generation ---');
        const codegen = new CodeGenerator(irResults, semanticResults.symbolTable, workbook);
        codegen.generate();
        console.log('✓ Code generation complete. Output in ./output folder.');

        console.log('\n✨ Build Successful! Run "node output/server.js" to start the app.');
        
    } catch (error) {
        console.error(`\n❌ Compilation Failed: ${error.message}`);
        console.error(error.stack);
    }
}

function printUsage() {
    console.log('Usage: node index.js compile <file.xlsx>');
}

main();

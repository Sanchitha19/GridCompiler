const { FormulaParser, printAST } = require('./src/parser');

const formulas = [
    '=SUM(A1:B5)',
    '=IF(C3>100, "High", "Low")',
    '=A1*B1+C1/D1',
    '=VLOOKUP(A2, Sheet2!B1:C10, 2, FALSE)',
    '=A1+*B2'
];

function runTests() {
    console.log('=== Excel2App Formula Parser Phase 2 Tests ===\n');

    formulas.forEach(formula => {
        console.log(`Parsing Formula: ${formula}`);
        const parser = new FormulaParser(formula);
        const ast = parser.parse();
        printAST(ast);
        console.log('\n' + '-'.repeat(40) + '\n');
    });
}

runTests();

const XLSX = require('xlsx');
const { FormulaParser } = require('../parser');

/**
 * Types supported by the semantic analyzer
 */
const Types = {
    NUMBER: 'NUMBER',
    STRING: 'STRING',
    BOOLEAN: 'BOOLEAN',
    MIXED: 'MIXED',
    UNKNOWN: 'UNKNOWN',
    ERROR: 'ERROR'
};

class SemanticAnalyzer {
    constructor(tokens) {
        this.tokens = tokens; // Cell-level tokens from Phase 1
        this.symbolTable = new Map();
        this.dependencies = new Map(); // Source -> Set of Targets (Source depends on Targets)
        this.dependents = new Map();   // Target -> Set of Sources (Target supports Sources)
        this.errors = [];
        this.supportedFunctions = ['SUM', 'AVERAGE', 'COUNT', 'IF', 'VLOOKUP', 'PRODUCT', 'MIN', 'MAX'];
    }

    /**
     * Helper to format cell addresses consistently
     */
    formatFullAddr(sheet, addr) {
        // Strip any $ for consistency in graph
        const cleanAddr = addr.replace(/\$/g, '');
        return `${sheet}!${cleanAddr}`;
    }

    /**
     * Expands a range string (e.g. A1:B2) into an array of addresses
     */
    expandRange(start, end) {
        const range = XLSX.utils.decode_range(`${start}:${end}`);
        const cells = [];
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                cells.push(XLSX.utils.encode_cell({ r: R, c: C }));
            }
        }
        return cells;
    }

    analyze() {
        this.initializeSymbolTable();
        this.buildDependencyGraph();
        this.detectCycles();
        this.inferTypes();
        this.validateReferences();
        return {
            symbolTable: this.symbolTable,
            dependencies: this.dependencies,
            errors: this.errors
        };
    }

    initializeSymbolTable() {
        for (const token of this.tokens) {
            const addr = this.formatFullAddr(token.sheet, token.cellAddress);
            const entry = {
                ast: null,
                inferredType: Types.UNKNOWN,
                sheet: token.sheet,
                row: token.row,
                col: token.col,
                tokenType: token.type,
                value: token.value
            };

            if (token.type === 'FORMULA') {
                const parser = new FormulaParser(token.value);
                entry.ast = parser.parse();
                if (entry.ast.type === 'ErrorNode') {
                    this.errors.push({
                        cell: addr,
                        type: 'SyntaxError',
                        details: entry.ast.message
                    });
                    entry.inferredType = Types.ERROR;
                }
            } else {
                switch (token.type) {
                    case 'NUMBER': entry.inferredType = Types.NUMBER; break;
                    case 'STRING': entry.inferredType = Types.STRING; break;
                    case 'BOOLEAN': entry.inferredType = Types.BOOLEAN; break;
                    case 'EMPTY': entry.inferredType = Types.UNKNOWN; break;
                }
            }
            this.symbolTable.set(addr, entry);
        }
    }

    addDependency(source, target) {
        if (!this.dependencies.has(source)) this.dependencies.set(source, new Set());
        this.dependencies.get(source).add(target);

        if (!this.dependents.has(target)) this.dependents.set(target, new Set());
        this.dependents.get(target).add(source);
    }

    buildDependencyGraph() {
        for (const [addr, entry] of this.symbolTable) {
            if (entry.ast && entry.ast.type !== 'ErrorNode') {
                this.walkAST(entry.ast, addr, entry.sheet);
            }
        }
    }

    walkAST(node, sourceAddr, sourceSheet) {
        if (!node) return;

        switch (node.type) {
            case 'CellRef': {
                const depSheet = node.sheet || sourceSheet;
                const fullDepAddr = this.formatFullAddr(depSheet, node.col + node.row);
                this.addDependency(sourceAddr, fullDepAddr);
                break;
            }
            case 'RangeRef': {
                const depSheet = node.sheet || sourceSheet;
                const cells = this.expandRange(node.startCell, node.endCell);
                for (const c of cells) {
                    const fullDepAddr = this.formatFullAddr(depSheet, c);
                    this.addDependency(sourceAddr, fullDepAddr);
                }
                break;
            }
            case 'FunctionCall': {
                if (!this.supportedFunctions.includes(node.name)) {
                    this.errors.push({
                        cell: sourceAddr,
                        type: 'UnknownFunction',
                        details: `${node.name} not supported`
                    });
                }
                if (node.args) {
                    node.args.forEach(arg => this.walkAST(arg, sourceAddr, sourceSheet));
                }
                break;
            }
            case 'BinaryOp':
                this.walkAST(node.left, sourceAddr, sourceSheet);
                this.walkAST(node.right, sourceAddr, sourceSheet);
                break;
            case 'UnaryOp':
                this.walkAST(node.operand, sourceAddr, sourceSheet);
                break;
        }
    }

    detectCycles() {
        const visited = new Set();
        const visiting = new Set();
        const path = [];
        const detectedCycles = [];

        const check = (u) => {
            if (visiting.has(u)) {
                const startIndex = path.indexOf(u);
                const cycle = path.slice(startIndex).concat(u);
                detectedCycles.push(cycle);
                return;
            }
            if (visited.has(u)) return;

            visiting.add(u);
            path.push(u);

            const deps = this.dependencies.get(u);
            if (deps) {
                for (const v of deps) {
                    check(v);
                }
            }

            path.pop();
            visiting.delete(u);
            visited.add(u);
        };

        for (const addr of this.symbolTable.keys()) {
            if (!visited.has(addr)) {
                check(addr);
            }
        }

        for (const cycle of detectedCycles) {
            const cycleStr = cycle.join(' \u2192 ');
            for (const cell of cycle) {
                const entry = this.symbolTable.get(cell);
                if (entry) {
                    // Only add error once per cell if involved in multiple cycles
                    if (entry.inferredType !== Types.ERROR) {
                        entry.inferredType = Types.ERROR;
                        this.errors.push({
                            cell: cell,
                            type: 'CyclicReference',
                            details: cycleStr
                        });
                    }
                }
            }
        }
    }

    inferTypes() {
        const memo = new Map();
        const visiting = new Set();

        const getInferred = (addr) => {
            if (memo.has(addr)) return memo.get(addr);
            
            const entry = this.symbolTable.get(addr);
            if (!entry) return Types.UNKNOWN;
            if (entry.inferredType === Types.ERROR) return Types.ERROR;
            if (entry.inferredType !== Types.UNKNOWN) return entry.inferredType;

            if (visiting.has(addr)) return Types.ERROR; // Cycle protection (though already handled)
            visiting.add(addr);

            let type = Types.UNKNOWN;
            if (entry.ast) {
                type = this.inferNode(entry.ast, entry.sheet, getInferred);
            }

            visiting.delete(addr);
            entry.inferredType = type;
            memo.set(addr, type);
            return type;
        };

        for (const addr of this.symbolTable.keys()) {
            getInferred(addr);
        }
    }

    inferNode(node, sheet, getInferred) {
        if (!node) return Types.UNKNOWN;

        switch (node.type) {
            case 'NumberLiteral': return Types.NUMBER;
            case 'StringLiteral': return Types.STRING;
            case 'BooleanLiteral': return Types.BOOLEAN;
            case 'CellRef': {
                const depSheet = node.sheet || sheet;
                const fullAddr = this.formatFullAddr(depSheet, node.col + node.row);
                return getInferred(fullAddr);
            }
            case 'RangeRef': return Types.MIXED;
            case 'BinaryOp': {
                const left = this.inferNode(node.left, sheet, getInferred);
                const right = this.inferNode(node.right, sheet, getInferred);
                if (node.operator === '+') {
                    if (left === Types.NUMBER && right === Types.NUMBER) return Types.NUMBER;
                    if (left === Types.STRING || right === Types.STRING) return Types.STRING;
                }
                if (['-', '*', '/', '^'].includes(node.operator)) return Types.NUMBER;
                if (['=', '<>', '<', '>', '<=', '>='].includes(node.operator)) return Types.BOOLEAN;
                if (node.operator === '&') return Types.STRING;
                return Types.MIXED;
            }
            case 'UnaryOp': {
                if (node.operator === '-') return Types.NUMBER;
                if (node.operator === 'NOT') return Types.BOOLEAN;
                return Types.UNKNOWN;
            }
            case 'FunctionCall': {
                if (['SUM', 'AVERAGE', 'COUNT', 'PRODUCT', 'MIN', 'MAX'].includes(node.name)) return Types.NUMBER;
                if (node.name === 'IF') {
                    const t1 = this.inferNode(node.args[1], sheet, getInferred);
                    const t2 = this.inferNode(node.args[2], sheet, getInferred);
                    return t1 === t2 ? t1 : Types.MIXED;
                }
                return Types.UNKNOWN;
            }
        }
        return Types.UNKNOWN;
    }

    validateReferences() {
        for (const [addr, entry] of this.symbolTable) {
            if (entry.ast) {
                this.checkRefs(entry.ast, addr, entry.sheet);
            }
        }
    }

    checkRefs(node, sourceAddr, sourceSheet) {
        if (!node) return;
        if (node.type === 'CellRef') {
            const depSheet = node.sheet || sourceSheet;
            const fullAddr = this.formatFullAddr(depSheet, node.col + node.row);
            if (!this.symbolTable.has(fullAddr)) {
                this.errors.push({
                    cell: sourceAddr,
                    type: 'UndefinedReference',
                    details: `Referenced address ${fullAddr} is not in workbook`
                });
            }
        } else if (node.type === 'RangeRef') {
            // Check if start/end are at least valid formats? Or just assume for now.
        } else if (node.args) {
            node.args.forEach(a => this.checkRefs(a, sourceAddr, sourceSheet));
        } else if (node.left) {
            this.checkRefs(node.left, sourceAddr, sourceSheet);
            this.checkRefs(node.right, sourceAddr, sourceSheet);
        } else if (node.operand) {
            this.checkRefs(node.operand, sourceAddr, sourceSheet);
        }
    }

    getDependencies(addr) {
        return this.dependencies.get(addr) || new Set();
    }

    getDependents(addr) {
        return this.dependents.get(addr) || new Set();
    }

    printSemanticReport() {
        console.log('\n=== SEMANTIC ERROR REPORT ===');
        if (this.errors.length === 0) {
            console.log('No errors found.');
            return;
        }

        console.log('CELL'.padEnd(15) + ' | ' + 'ERROR TYPE'.padEnd(20) + ' | ' + 'DETAILS');
        console.log('-'.repeat(80));
        for (const err of this.errors) {
            console.log(err.cell.padEnd(15) + ' | ' + err.type.padEnd(20) + ' | ' + err.details);
        }
    }
}

module.exports = {
    SemanticAnalyzer,
    Types
};

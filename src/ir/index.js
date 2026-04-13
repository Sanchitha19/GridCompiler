const XLSX = require('xlsx');

class IRNode {
    constructor(id, cellAddress, operation, dataType) {
        this.id = id;
        this.cellAddress = cellAddress;
        this.operation = operation; // CONST | LOAD | BINOP | CALL | PHI
        this.operands = [];
        this.value = null;
        this.dataType = dataType;
        this.isLive = true;
        this.optimizedBy = null; // null | 'fold' | 'cse' | 'dead'
    }
}

class IRGenerator {
    constructor(symbolTable, dependencies) {
        this.symbolTable = symbolTable;
        this.dependencies = dependencies; // Map<addr, Set<dep_addr>>
        this.irNodes = new Map(); // address -> IRNode
        this.order = [];
        this.stats = { folded: 0, eliminated: 0, cse: 0 };
        this.fingerprints = new Map(); // hash -> IRNode (for CSE)
    }

    generate() {
        this.topologicalSort();
        this.createNodes();
        
        // Optimizations
        this.constantFolding();
        this.commonSubexpressionElimination();
        this.deadCellElimination();
        
        return {
            nodes: this.irNodes,
            order: this.order,
            stats: this.stats
        };
    }

    /**
     * Topological sort using Kahn's algorithm
     */
    topologicalSort() {
        const inDegree = new Map();
        const adj = new Map(); // B -> A (if A depends on B)
        
        for (const addr of this.symbolTable.keys()) {
            inDegree.set(addr, 0);
        }

        for (const [source, targets] of this.dependencies) {
            inDegree.set(source, targets.size);
            for (const target of targets) {
                if (!adj.has(target)) adj.set(target, new Set());
                adj.get(target).add(source);
            }
        }

        const queue = [];
        for (const [addr, degree] of inDegree) {
            if (degree === 0) queue.push(addr);
        }

        while (queue.length > 0) {
            const u = queue.shift();
            this.order.push(u);

            const neighbors = adj.get(u);
            if (neighbors) {
                for (const v of neighbors) {
                    const d = inDegree.get(v) - 1;
                    inDegree.set(v, d);
                    if (d === 0) queue.push(v);
                }
            }
        }

        // Check for cycles (nodes not in order)
        if (this.order.length < this.symbolTable.size) {
            const diff = [...this.symbolTable.keys()].filter(a => !this.order.includes(a));
            console.warn(`[WARNING] Cycle detected or unresolved dependencies. Skipping ${diff.length} nodes: ${diff.join(', ')}`);
        }
    }

    createNodes() {
        // We must create literal nodes for cells that have no AST
        for (const [addr, entry] of this.symbolTable) {
            const id = `ir_${addr.replace(/!/g, '_').replace(/[\$\:]/g, '')}`;
            const node = new IRNode(id, addr, 'CONST', entry.inferredType);
            
            if (entry.ast && entry.ast.type !== 'ErrorNode') {
                // If it's a formula, we'll refine it below
                this.irNodes.set(addr, node);
            } else {
                node.operation = 'CONST';
                node.value = entry.value;
                this.irNodes.set(addr, node);
            }
        }

        // Now translate ASTs for formula cells
        for (const addr of this.order) {
            const entry = this.symbolTable.get(addr);
            if (entry && entry.ast && entry.ast.type !== 'ErrorNode') {
                const node = this.translateAST(entry.ast, addr, entry.sheet);
                const existing = this.irNodes.get(addr);
                Object.assign(existing, node);
            }
        }
    }

    translateAST(node, cellAddr, sheet) {
        const id = `ir_${cellAddr.replace(/!/g, '_').replace(/[\$\:]/g, '')}`;
        
        switch (node.type) {
            case 'NumberLiteral': {
                const n = new IRNode(id, cellAddr, 'CONST', 'NUMBER');
                n.value = node.value;
                return n;
            }
            case 'StringLiteral': {
                const n = new IRNode(id, cellAddr, 'CONST', 'STRING');
                n.value = node.value;
                return n;
            }
            case 'BooleanLiteral': {
                const n = new IRNode(id, cellAddr, 'CONST', 'BOOLEAN');
                n.value = node.value;
                return n;
            }
            case 'CellRef': {
                const n = new IRNode(id, cellAddr, 'LOAD', 'UNKNOWN');
                const targetAddr = `${node.sheet || sheet}!${node.col}${node.row}`.replace(/\$/g, '');
                const targetNode = this.irNodes.get(targetAddr);
                if (targetNode) {
                    n.operands.push(targetNode);
                    n.dataType = targetNode.dataType;
                }
                return n;
            }
            case 'RangeRef': {
                // Should not occur as a top-level node for a cell usually, 
                // but if it does, treat as LOAD array (simplified)
                const n = new IRNode(id, cellAddr, 'LOAD', 'MIXED');
                return n;
            }
            case 'BinaryOp': {
                const n = new IRNode(id, cellAddr, 'BINOP', 'UNKNOWN');
                n.operation = `BINOP:${node.operator}`;
                n.operands.push(this.translateAST(node.left, cellAddr + '_L', sheet));
                n.operands.push(this.translateAST(node.right, cellAddr + '_R', sheet));
                return n;
            }
            case 'UnaryOp': {
                const n = new IRNode(id, cellAddr, 'UNOP', 'UNKNOWN');
                n.operation = `UNOP:${node.operator}`;
                n.operands.push(this.translateAST(node.operand, cellAddr + '_U', sheet));
                return n;
            }
            case 'FunctionCall': {
                const n = new IRNode(id, cellAddr, 'CALL', 'UNKNOWN');
                n.operation = `CALL:${node.name}`;
                if (node.args) {
                    node.args.forEach((arg, i) => {
                        // Special handling for RangeRef in function arguments
                        if (arg.type === 'RangeRef') {
                            const cells = this.expandRange(arg.startCell, arg.endCell);
                            const depSheet = arg.sheet || sheet;
                            for (const c of cells) {
                                const fullDepAddr = `${depSheet}!${c}`.replace(/\$/g, '');
                                const targetNode = this.irNodes.get(fullDepAddr);
                                if (targetNode) {
                                    const loadNode = new IRNode(`${id}_arg_${c}`, fullDepAddr, 'LOAD', targetNode.dataType);
                                    loadNode.operands.push(targetNode);
                                    n.operands.push(loadNode);
                                }
                            }
                        } else {
                            n.operands.push(this.translateAST(arg, `${cellAddr}_arg_${i}`, sheet));
                        }
                    });
                }
                return n;
            }
        }
        return new IRNode(id, cellAddr, 'UNKNOWN', 'UNKNOWN');
    }

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

    /**
     * Optimization: Constant Folding
     */
    constantFolding() {
        const fold = (node) => {
            if (node.operation === 'CONST') return false; // Already const, no new folding
            
            // Recurse on operands
            let allConst = node.operands.length > 0;
            for (const op of node.operands) {
                // We don't care about return value here, just visiting
                this.visitAndFold(op); 
                if (op.operation !== 'CONST') allConst = false;
            }

            if (allConst) {
                try {
                    const result = this.evaluateNode(node);
                    if (result !== undefined) {
                        node.operation = 'CONST';
                        node.value = result;
                        node.operands = [];
                        node.optimizedBy = 'fold';
                        this.stats.folded++;
                        return true;
                    }
                } catch (e) {
                }
            }
            return false;
        };

        for (const addr of this.order) {
            const node = this.irNodes.get(addr);
            if (node) fold(node);
        }
    }

    visitAndFold(node) {
        if (node.operation === 'CONST') return;
        let allConst = node.operands.length > 0;
        for (const op of node.operands) {
            this.visitAndFold(op);
            if (op.operation !== 'CONST') allConst = false;
        }
        if (allConst) {
            const result = this.evaluateNode(node);
            if (result !== undefined) {
                node.operation = 'CONST';
                node.value = result;
                node.operands = [];
                node.optimizedBy = 'fold';
                this.stats.folded++;
            }
        }
    }

    evaluateNode(node) {
        const ops = node.operands.map(o => o.value);
        if (node.operation.startsWith('BINOP:')) {
            const op = node.operation.split(':')[1];
            switch (op) {
                case '+': return ops[0] + ops[1];
                case '-': return ops[0] - ops[1];
                case '*': return ops[0] * ops[1];
                case '/': return ops[0] / ops[1];
                case '^': return Math.pow(ops[0], ops[1]);
                case '&': return String(ops[0]) + String(ops[1]);
                case '=': return ops[0] === ops[1];
                case '<>': return ops[0] !== ops[1];
            }
        } else if (node.operation.startsWith('UNOP:')) {
            const op = node.operation.split(':')[1];
            switch (op) {
                case '-': return -ops[0];
                case 'NOT': return !ops[0];
            }
        } else if (node.operation.startsWith('CALL:')) {
            const name = node.operation.split(':')[1];
            switch (name) {
                case 'SUM': return ops.reduce((a, b) => a + (Number(b) || 0), 0);
                case 'COUNT': return ops.filter(o => typeof o === 'number').length;
                case 'PRODUCT': return ops.reduce((a, b) => a * (Number(b) || 1), 1);
            }
        } else if (node.operation === 'LOAD') {
            return ops[0];
        }
        return undefined;
    }

    /**
     * Optimization: Common Subexpression Elimination
     */
    commonSubexpressionElimination() {
        for (const addr of this.order) {
            const node = this.irNodes.get(addr);
            if (!node || node.operation === 'CONST' || node.operation === 'LOAD') continue;

            const fingerprint = `${node.operation}(${node.operands.map(o => o.id).join(',')})`;
            if (this.fingerprints.has(fingerprint)) {
                const original = this.fingerprints.get(fingerprint);
                // Replace THIS node with a LOAD to the original node's result
                node.operation = 'LOAD';
                node.operands = [original];
                node.optimizedBy = 'cse';
                this.stats.cse++;
            } else {
                this.fingerprints.set(fingerprint, node);
            }
        }
    }

    /**
     * Optimization: Dead Cell Elimination
     */
    deadCellElimination() {
        const liveSet = new Set();

        const markLive = (node) => {
            if (liveSet.has(node.id)) return;
            liveSet.add(node.id);
            node.operands.forEach(op => markLive(op));
        };

        // Start from formula cells or any cell referenced by others
        // Actually, start from cells that are explicitly formulas since they are "outputs" 
        // Or in Excel, anything in the Summary sheet is likely an output.
        // For this task, we mark anything reached from a formula as live.
        for (const [addr, entry] of this.symbolTable) {
            if (entry.tokenType === 'FORMULA') {
                const node = this.irNodes.get(addr);
                if (node) markLive(node);
            }
        }

        // Any cell that is not in liveSet and is a constant with no side effects
        for (const [addr, node] of this.irNodes) {
            if (!liveSet.has(node.id)) {
                node.isLive = false;
                node.optimizedBy = 'dead';
                this.stats.eliminated++;
            }
        }
    }

    printIR() {
        console.log('\n=== INTERMEDIATE REPRESENTATION (IR) ===');
        console.log('ID'.padEnd(20) + ' | ' + 'CELL'.padEnd(15) + ' | ' + 'OP'.padEnd(10) + ' | ' + 'STATUS');
        console.log('-'.repeat(80));
        
        for (const [addr, node] of this.irNodes) {
            let status = node.isLive ? 'LIVE' : 'DEAD';
            if (node.optimizedBy) status += ` (${node.optimizedBy.toUpperCase()})`;
            
            console.log(node.id.padEnd(20) + ' | ' + addr.padEnd(15) + ' | ' + node.operation.padEnd(10) + ' | ' + status);
        }

        console.log(`\nOptimization Summary:`);
        console.log(`${this.stats.folded} nodes folded, ${this.stats.eliminated} nodes eliminated, ${this.stats.cse} CSE replacements`);
    }
}

module.exports = {
    IRGenerator,
    IRNode
};

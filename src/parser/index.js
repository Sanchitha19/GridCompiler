/**
 * AST Node Definitions
 */
const NodeTypes = {
    NumberLiteral: (value) => ({ type: 'NumberLiteral', value }),
    StringLiteral: (value) => ({ type: 'StringLiteral', value }),
    BooleanLiteral: (value) => ({ type: 'BooleanLiteral', value }),
    CellRef: (sheet, col, row) => ({ type: 'CellRef', sheet, col, row }),
    RangeRef: (sheet, startCell, endCell) => ({ type: 'RangeRef', sheet, startCell, endCell }),
    BinaryOp: (operator, left, right) => ({ type: 'BinaryOp', operator, left, right }),
    UnaryOp: (operator, operand) => ({ type: 'UnaryOp', operator, operand }),
    FunctionCall: (name, args) => ({ type: 'FunctionCall', name, args }),
    ErrorNode: (message, formulaString) => ({ type: 'ErrorNode', message, formulaString })
};

/**
 * FormulaLexer
 * Simple lexer to tokenize Excel formula strings.
 */
class FormulaLexer {
    constructor(formula) {
        this.formula = formula.startsWith('=') ? formula.substring(1) : formula;
        this.pos = 0;
        this.tokens = [];
        this.tokenize();
    }

    tokenize() {
        while (this.pos < this.formula.length) {
            const char = this.formula[this.pos];

            if (/\s/.test(char)) {
                this.pos++;
                continue;
            }

            // String literals: "Text" or "Text ""with"" quotes"
            if (char === '"') {
                let val = '';
                this.pos++;
                while (this.pos < this.formula.length) {
                    if (this.formula[this.pos] === '"') {
                        if (this.formula[this.pos + 1] === '"') {
                            val += '"';
                            this.pos += 2;
                        } else {
                            this.pos++;
                            break;
                        }
                    } else {
                        val += this.formula[this.pos];
                        this.pos++;
                    }
                }
                this.tokens.push({ type: 'STRING', value: val });
                continue;
            }

            // Numbers: 123, 123.45
            if (/\d/.test(char) || (char === '.' && /\d/.test(this.formula[this.pos + 1]))) {
                let val = '';
                while (this.pos < this.formula.length && /[\d\.]/.test(this.formula[this.pos])) {
                    val += this.formula[this.pos];
                    this.pos++;
                }
                this.tokens.push({ type: 'NUMBER', value: parseFloat(val) });
                continue;
            }

            // Two-character operators
            const sub2 = this.formula.substring(this.pos, this.pos + 2);
            if (['<>', '<=', '>='].includes(sub2)) {
                this.tokens.push({ type: 'OPERATOR', value: sub2 });
                this.pos += 2;
                continue;
            }

            // Single-character operators and punctuation
            if ('+-*/^&=<>(),:!'.includes(char)) {
                this.tokens.push({ type: 'OPERATOR', value: char });
                this.pos++;
                continue;
            }

            // Identifiers (Function names, Sheet names, Cell refs), Booleans
            if (/[a-zA-Z_\$]/.test(char) || char === '\'') {
                let val = '';
                if (char === '\'') {
                    this.pos++;
                    while (this.pos < this.formula.length) {
                        if (this.formula[this.pos] === '\'') {
                            if (this.formula[this.pos + 1] === '\'') {
                                val += '\'';
                                this.pos += 2;
                            } else {
                                this.pos++;
                                break;
                            }
                        } else {
                            val += this.formula[this.pos];
                            this.pos++;
                        }
                    }
                } else {
                    while (this.pos < this.formula.length && /[a-zA-Z0-9_\.\$]/.test(this.formula[this.pos])) {
                        val += this.formula[this.pos];
                        this.pos++;
                    }
                }

                const upper = val.toUpperCase();
                if (upper === 'TRUE') {
                    this.tokens.push({ type: 'BOOLEAN', value: true });
                } else if (upper === 'FALSE') {
                    this.tokens.push({ type: 'BOOLEAN', value: false });
                } else if (upper === 'NOT') {
                    this.tokens.push({ type: 'OPERATOR', value: 'NOT' });
                } else {
                    this.tokens.push({ type: 'IDENTIFIER', value: val });
                }
                continue;
            }

            // Fallback for unknown characters
            this.tokens.push({ type: 'UNKNOWN', value: char });
            this.pos++;
        }
        this.tokens.push({ type: 'EOF', value: null });
    }
}

/**
 * FormulaParser
 * Recursive descent parser for Excel formulas.
 * 
 * Formal BNF to implement:
 *   expr     ::= term (('+' | '-' | '&' | '=' | '<>' | '<' | '>') term)*
 *   term     ::= factor (('*' | '/') factor)*
 *   factor   ::= unary ('^' unary)*
 *   unary    ::= '-' unary | 'NOT' unary | primary
 *   primary  ::= NUMBER | STRING | BOOLEAN | cellRef | rangeRef 
 *              | FUNCTION '(' argList ')' | '(' expr ')'
 *   argList  ::= expr (',' expr)*
 */
class FormulaParser {
    constructor(formulaString) {
        this.formulaString = formulaString;
        const lexer = new FormulaLexer(formulaString);
        this.tokens = lexer.tokens;
        this.pos = 0;
    }

    peek() {
        return this.tokens[this.pos];
    }

    consume() {
        return this.tokens[this.pos++];
    }

    match(type, value = null) {
        const token = this.peek();
        if (token.type === type && (value === null || token.value === value)) {
            return this.consume();
        }
        return null;
    }

    expect(type, value = null) {
        const token = this.match(type, value);
        if (!token) {
            const current = this.peek();
            throw new Error(`Expected ${type}${value ? ':' + value : ''}, but found ${current.type}${current.value !== null ? ':' + current.value : ''}`);
        }
        return token;
    }

    parse() {
        try {
            const node = this.parseExpression();
            if (this.peek().type !== 'EOF') {
               throw new Error(`Unexpected extra content at end of formula: ${this.peek().value || this.peek().type}`);
            }
            return node;
        } catch (e) {
            return NodeTypes.ErrorNode(e.message, this.formulaString);
        }
    }

    parseExpression() {
        let node = this.parseTerm();
        
        while (true) {
            const token = this.peek();
            if (token.type === 'OPERATOR' && ['+', '-', '&', '=', '<>', '<', '>', '<=', '>='].includes(token.value)) {
                this.consume();
                const right = this.parseTerm();
                node = NodeTypes.BinaryOp(token.value, node, right);
            } else {
                break;
            }
        }
        return node;
    }

    parseTerm() {
        let node = this.parseFactor();

        while (true) {
            const token = this.peek();
            if (token.type === 'OPERATOR' && ['*', '/'].includes(token.value)) {
                this.consume();
                const right = this.parseFactor();
                node = NodeTypes.BinaryOp(token.value, node, right);
            } else {
                break;
            }
        }
        return node;
    }

    parseFactor() {
        let node = this.parseUnary();

        while (true) {
            const token = this.peek();
            if (token.type === 'OPERATOR' && token.value === '^') {
                this.consume();
                const right = this.parseUnary();
                node = NodeTypes.BinaryOp('^', node, right);
            } else {
                break;
            }
        }
        return node;
    }

    parseUnary() {
        const token = this.peek();
        if (token.type === 'OPERATOR' && (token.value === '-' || token.value === 'NOT')) {
            this.consume();
            const operand = this.parseUnary();
            return NodeTypes.UnaryOp(token.value, operand);
        }
        return this.parsePrimary();
    }

    parsePrimary() {
        const token = this.peek();

        if (token.type === 'NUMBER') {
            this.consume();
            return NodeTypes.NumberLiteral(token.value);
        }

        if (token.type === 'STRING') {
            this.consume();
            return NodeTypes.StringLiteral(token.value);
        }

        if (token.type === 'BOOLEAN') {
            this.consume();
            return NodeTypes.BooleanLiteral(token.value);
        }

        if (token.type === 'OPERATOR' && token.value === '(') {
            this.consume();
            const node = this.parseExpression();
            this.expect('OPERATOR', ')');
            return node;
        }

        // Identifier could be Function name, Sheet name, or Cell Ref
        if (token.type === 'IDENTIFIER') {
            // Peek for '(' to see if it's a function
            const next = this.tokens[this.pos + 1];
            if (next && next.type === 'OPERATOR' && next.value === '(') {
                return this.parseFunction();
            }
            
            // Otherwise, it's part of a ref (Cell or Range)
            return this.parseRef();
        }

        throw new Error(`Unexpected token at start of expression: ${token.type} (${token.value})`);
    }

    parseFunction() {
        const nameToken = this.expect('IDENTIFIER');
        this.expect('OPERATOR', '(');
        
        const args = [];
        if (this.peek().type !== 'OPERATOR' || this.peek().value !== ')') {
            args.push(this.parseExpression());
            while (this.match('OPERATOR', ',')) {
                args.push(this.parseExpression());
            }
        }
        
        this.expect('OPERATOR', ')');
        return NodeTypes.FunctionCall(nameToken.value.toUpperCase(), args);
    }

    parseRef() {
        let sheet = null;
        
        // Check for cross-sheet reference: Identifier ! something
        if (this.peek().type === 'IDENTIFIER') {
            const lookahead = this.tokens[this.pos + 1];
            if (lookahead && lookahead.type === 'OPERATOR' && lookahead.value === '!') {
                sheet = this.consume().value;
                this.consume(); // consume '!'
            }
        }

        // Now we expect a cell or range starting with an IDENTIFIER (like A1)
        const startToken = this.expect('IDENTIFIER');
        const startCell = startToken.value;
        
        if (this.match('OPERATOR', ':')) {
            const endToken = this.expect('IDENTIFIER');
            const endCell = endToken.value;
            return NodeTypes.RangeRef(sheet, startCell, endCell);
        } else {
            const { col, row } = this.splitCellAddress(startCell);
            return NodeTypes.CellRef(sheet, col, row);
        }
    }

    splitCellAddress(addr) {
        // Simple regex to split A1 or $A$1 into col and row
        const match = addr.match(/^(\$?[A-Z]+)(\$?[0-9]+)$/i);
        if (match) {
            return { col: match[1], row: match[2] };
        }
        return { col: addr, row: '' };
    }
}

/**
 * Debug Utility
 */
function printAST(node, indent = 0) {
    const pad = '  '.repeat(indent);
    if (!node) {
        console.log(`${pad}null`);
        return;
    }

    switch (node.type) {
        case 'NumberLiteral':
            console.log(`${pad}NumberLiteral(${node.value})`);
            break;
        case 'StringLiteral':
            console.log(`${pad}StringLiteral("${node.value}")`);
            break;
        case 'BooleanLiteral':
            console.log(`${pad}BooleanLiteral(${node.value})`);
            break;
        case 'CellRef':
            console.log(`${pad}CellRef(sheet=${node.sheet || 'null'}, col=${node.col}, row=${node.row})`);
            break;
        case 'RangeRef':
            console.log(`${pad}RangeRef(sheet=${node.sheet || 'null'}, start=${node.startCell}, end=${node.endCell})`);
            break;
        case 'BinaryOp':
            console.log(`${pad}BinaryOp(${node.operator})`);
            printAST(node.left, indent + 1);
            printAST(node.right, indent + 1);
            break;
        case 'UnaryOp':
            console.log(`${pad}UnaryOp(${node.operator})`);
            printAST(node.operand, indent + 1);
            break;
        case 'FunctionCall':
            console.log(`${pad}FunctionCall(${node.name})`);
            if (node.args) {
                node.args.forEach(arg => printAST(arg, indent + 1));
            }
            break;
        case 'ErrorNode':
            console.log(`${pad}ErrorNode: ${node.message}`);
            break;
        default:
            console.log(`${pad}UnknownNode(${node.type})`);
    }
}

module.exports = {
    FormulaParser,
    printAST,
    NodeTypes
};

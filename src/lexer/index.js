const XLSX = require('xlsx');
const TokenTypes = require('./tokenTypes');

/**
 * Formats a cell address as "SheetName!Address" (e.g., "Sheet1!B3").
 */
function formatCellAddress(sheetName, row, col) {
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    return `${sheetName}!${address}`;
}

/**
 * Tokenizes an entire workbook.
 * Walks every cell in every sheet and emits a flat array of Tokens.
 */
function tokenize(workbook) {
    const tokens = [];

    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        
        // Get the range of the sheet
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');

        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = sheet[cellAddress];
                
                // RAW LOGGING REQUESTED BY USER
                if (cell) {
                    console.log(`[RAW EXTRACTION] ${sheetName}!${cellAddress} =>`, JSON.stringify(cell));
                }

                let token = {
                    type: TokenTypes.EMPTY,
                    value: null,
                    row: R,
                    col: C,
                    cellAddress: cellAddress,
                    sheet: sheetName
                };

                if (cell) {
                    // Check if it's a formula - ALWAYS prioritize the formula string
                    // to avoid using stale cached values ("v":0) from the Excel file.
                    if (cell.f) {
                        token.type = TokenTypes.FORMULA;
                        token.value = `=${cell.f}`;
                    } else {
                        // Determine type based on cell property 't' (type)
                        // 'b' for Boolean, 'n' for Number, 's' for String, 'z' for Stub
                        switch (cell.t) {
                            case 'n':
                                token.type = TokenTypes.NUMBER;
                                token.value = cell.v;
                                break;
                            case 's':
                                token.type = TokenTypes.STRING;
                                token.value = cell.v;
                                break;
                            case 'b':
                                token.type = TokenTypes.BOOLEAN;
                                token.value = cell.v;
                                break;
                            default:
                                if (cell.v === undefined || cell.v === null) {
                                    token.type = TokenTypes.EMPTY;
                                } else {
                                    token.type = TokenTypes.STRING;
                                    token.value = String(cell.v);
                                }
                        }
                    }
                }
                
                // Add token if it's not empty
                if (token.type !== TokenTypes.EMPTY) {
                   tokens.push(token);
                }
            }
        }
    });

    return tokens;
}

/**
 * Debug utility to print tokens.
 */
function printTokens(tokens) {
    console.log('--- TOKENS ---');
    tokens.forEach(t => {
        const addr = formatCellAddress(t.sheet, t.row, t.col);
        console.log(`[${addr}] ${t.type.padEnd(12)} : ${t.value}`);
    });
    console.log(`Total Tokens: ${tokens.length}`);
}

module.exports = {
    tokenize,
    formatCellAddress,
    printTokens,
    TokenTypes
};

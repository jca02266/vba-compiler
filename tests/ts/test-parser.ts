import * as fs from 'fs';
import { Lexer } from './src/compiler/lexer';
import { Parser } from './src/compiler/parser';

const code = fs.readFileSync('test.vba', 'utf-8');

try {
    const lexer = new Lexer(code);
    const tokens = lexer.tokenize();
    // console.log(`Tokens: ${tokens.length}`);

    const parser = new Parser(tokens);
    const ast = parser.parse();

    console.log("Successfully parsed test.vba!");
    console.log(`Statements parsed: ${ast.body.length}`);
} catch (e: any) {
    console.error(e.message);
    process.exit(1);
}

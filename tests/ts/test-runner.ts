import * as fs from 'fs';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';

export function runVBATest(filePath: string, procedureName: string, args: any[]): any {
    const start = Date.now();

    // Read and parse
    const sourceCode = fs.readFileSync(filePath, 'utf-8');
    const lexer = new Lexer(sourceCode);
    const tokens = lexer.tokenize();

    const parser = new Parser(tokens);
    const ast = parser.parse();

    // Evaluate
    const evaluator = new Evaluator();
    evaluator.evaluate(ast);

    // Call the specific procedure
    const result = evaluator.callProcedure(procedureName, args);
    const duration = Date.now() - start;
    const formatArgs = args.map(a => typeof a === 'object' ? '[Object]' : String(a)).join(', ');
    console.log(`[PASS] ${procedureName}(${formatArgs}) -> ${result} (${duration}ms)`);
    return result;
}

// Minimal assert framework
export const assert = {
    strictEqual: (actual: any, expected: any, message?: string) => {
        if (actual !== expected) {
            console.error(`[FAIL] ${message || 'Assertion failed'} - Expected ${expected} but got ${actual}`);
            throw new Error(`Assertion Failed`);
        }
    }
};

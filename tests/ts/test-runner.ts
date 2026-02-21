import * as fs from 'fs';
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';

export class VBATest {
    private evaluator: Evaluator;

    constructor(filePath: string) {
        const sourceCode = fs.readFileSync(filePath, 'utf-8');
        const lexer = new Lexer(sourceCode);
        const tokens = lexer.tokenize();
        const parser = new Parser(tokens);
        const ast = parser.parse();
        this.evaluator = new Evaluator(console.log);
        this.evaluator.evaluate(ast);
    }

    run(procedureName: string, args: any[]): any {
        const start = Date.now();
        const result = this.evaluator.callProcedure(procedureName, args);
        const duration = Date.now() - start;
        const formatArgs = args.map(a => typeof a === 'object' ? '[Object]' : String(a)).join(', ');
        console.log(`[PASS] ${procedureName}(${formatArgs}) -> ${result} (${duration}ms)`);
        return result;
    }

    eval(exprString: string): any {
        return this.evaluator.evalExpression(exprString);
    }
}

// Keep backward compatibility
export function runVBATest(filePath: string, procedureName: string, args: any[]): any {
    const vbaTest = new VBATest(filePath);
    return vbaTest.run(procedureName, args);
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

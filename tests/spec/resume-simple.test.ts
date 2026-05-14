import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Simple test: Resume Next
{
    const code = `
    Function Test()
        Dim x As Integer
        x = 0
        On Error GoTo Err1
        x = 1
        Err.Raise 11
        x = 999
        Test = x
        Exit Function
    Err1:
        x = x + 100
        Resume Next
    End Function
    `;
    try {
        const result = runFunc(code, 'Test');
        console.log('Result:', result);
        // x = 1, then Err.Raise, error handler: x = 101, Resume Next: execute x = 999
        assert.strictEqual(result, 999, 'Resume Next should skip Err.Raise, handler changes x to 101, then x = 999 overwrites it');
        console.log('[PASS] Simple Resume Next test');
    } catch (e: any) {
        console.log('[ERROR]', e);
    }
}

console.log('\n✅ Resume Simple: Test complete');

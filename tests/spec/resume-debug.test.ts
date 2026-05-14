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

// Test: Check if error handler is triggered
{
    const code = `
    Function Test()
        Dim x As Integer
        x = 0
        On Error GoTo Err1
        x = 1
        Err.Raise 11
        Test = 999  ' Should be skipped by error
        Exit Function
    Err1:
        x = 100  ' Error handler
        Test = x
    End Function
    `;
    try {
        const result = runFunc(code, 'Test');
        console.log('Result:', result);
        if (result === 100) {
            console.log('[PASS] Error handler was triggered');
        } else {
            console.log('[FAIL] Error handler was NOT triggered, got:', result);
        }
    } catch (e: any) {
        console.log('[ERROR]', e);
    }
}

console.log('\n✅ Resume Debug: Test complete');

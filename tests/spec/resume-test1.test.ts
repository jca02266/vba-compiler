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

// Test 1: Resume re-executes the error-causing statement
{
    const code = `
    Function Test1()
        Dim count As Integer
        count = 0
        On Error GoTo ErrorHandler
        count = count + 1
        If count <= 1 Then
            Err.Raise 11
        End If
        count = count + 100
        Test1 = count
        Exit Function
    ErrorHandler:
        Resume
    End Function
    `;
    const result = runFunc(code, 'Test1');
    assert.strictEqual(result, 102, 'Resume should re-execute the error-causing statement');
    console.log('[PASS] Test 1: Resume re-executes error statement');
}

console.log('\n✅ Test 1 complete');

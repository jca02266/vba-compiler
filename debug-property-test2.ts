import { Lexer } from './src/compiler/lexer';
import { Parser } from './src/compiler/parser';
import { Evaluator } from './src/compiler/evaluator';

const code = `
    Class Container
        Private val
        Property Get Value()
            Value = val
        End Property
        Property Let Value(v)
            val = v
        End Property
    End Class

    Function Test()
        Dim x As New Container
        x.Value = 5
        Test = x.Value
    End Function
`;

try {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    const result = ev.callProcedure('Test', []);
    console.log('Final result:', result);
    console.log('Type:', typeof result);

    // Also test the value getter directly
    const testDirect = `
        Class Box
            Private mVal
            Property Get Value()
                Value = mVal
            End Property
            Property Let Value(v)
                mVal = v
            End Property
        End Class

        Function GetVal()
            Dim b As New Box
            b.Value = 42
            GetVal = b.Value
        End Function
    `;

    const tokens2 = new Lexer(testDirect).tokenize();
    const ast2 = new Parser(tokens2).parse();
    const ev2 = new Evaluator(console.log);
    ev2.evaluate(ast2);
    const result2 = ev2.callProcedure('GetVal', []);
    console.log('\nDirect test result:', result2);
    console.log('Type:', typeof result2);
} catch (e: any) {
    console.error('Error:', e.message || e);
    console.error('Stack:', e.stack);
}

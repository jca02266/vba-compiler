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

        ' Debug: Check x.Value after assignment
        Dim check1
        check1 = x.Value

        ' Now manually modify it within the same function scope
        x.Value = x.Value + 10

        ' Debug: Check x.Value after modification
        Dim check2
        check2 = x.Value

        Test = x.Value
    End Function
`;

try {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    const result = ev.callProcedure('Test', []);
    console.log('Test result:', result);
    console.log('Type:', typeof result);
} catch (e: any) {
    console.error('Error:', e.message || e);
}

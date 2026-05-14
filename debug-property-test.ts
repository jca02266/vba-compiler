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

    Sub ModifyValue(c As Container)
        c.Value = c.Value + 10
    End Sub

    Function TestPropertyByRef()
        Dim x As New Container
        x.Value = 5
        ModifyValue x
        TestPropertyByRef = x.Value
    End Function
`;

try {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    const result = ev.callProcedure('TestPropertyByRef', []);
    console.log('Result:', result);
    console.log('Result type:', typeof result);
    console.log('Result is object:', result && typeof result === 'object');
} catch (e: any) {
    console.error('Error:', e.message || e);
}

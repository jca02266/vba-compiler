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

    Function Test1()
        Dim x As New Container
        x.Value = 5
        ModifyValue x
        Test1 = x.Value
    End Function

    Function Test2()
        Dim x As New Container
        x.Value = 5
        Test2 = x.Value
    End Function
`;

try {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);

    const result2 = ev.callProcedure('Test2', []);
    console.log('Test2 (without ModifyValue call) result:', result2);
    console.log('Type:', typeof result2);

    const result1 = ev.callProcedure('Test1', []);
    console.log('\nTest1 (with ModifyValue call) result:', result1);
    console.log('Type:', typeof result1);

    if (typeof result1 === 'object') {
        console.log('Object details:', JSON.stringify(result1, null, 2).substring(0, 300));
    }
} catch (e: any) {
    console.error('Error:', e.message || e);
}

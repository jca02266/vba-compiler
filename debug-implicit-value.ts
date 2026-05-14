import { Lexer } from './src/compiler/lexer';
import { Parser } from './src/compiler/parser';
import { Evaluator } from './src/compiler/evaluator';

const code = `
    Class ValueHolder
        Private mValue
        Property Get Value()
            Value = mValue
        End Property
        Property Let Value(v)
            mValue = v
        End Property
    End Class

    Function Test()
        Dim holder As New ValueHolder
        holder = 42
        Test = holder
    End Function
`;

try {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    const result = ev.callProcedure('Test', []);
    console.log('Result type:', typeof result);
    console.log('Is class instance:', result && result.__vbaClass__);
    if (result && result.__instanceEnv__) {
        const mValue = result.__instanceEnv__.variables.get('mvalue');
        console.log('mValue field:', mValue);
    }
} catch (e: any) {
    console.error('Error:', e.message);
}

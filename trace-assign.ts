import { Lexer } from './src/compiler/lexer';
import { Parser } from './src/compiler/parser';

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
`;

const tokens = new Lexer(code).tokenize();
const ast = new Parser(tokens).parse();

// Find the ModifyValue procedure
for (const stmt of ast.body) {
    if (stmt.type === 'ProcedureDeclaration' && (stmt as any).name.name === 'ModifyValue') {
        console.log('ModifyValue procedure found');
        const proc = stmt as any;
        console.log('Parameters:', proc.parameters.map((p: any) => ({ name: p.name, paramType: p.paramType })));
        console.log('Body:');
        for (const s of proc.body) {
            console.log('  Statement type:', s.type);
            if (s.type === 'AssignmentStatement') {
                const assign = s as any;
                console.log('    Left:', JSON.stringify(assign.left, null, 2).substring(0, 200));
                console.log('    Right type:', assign.right.type);
            }
        }
    }

    if (stmt.type === 'ClassDeclaration' && (stmt as any).name === 'Container') {
        console.log('\nContainer class found');
        const cls = stmt as any;
        for (const proc of cls.procedures) {
            if (proc.name.name === 'Value') {
                console.log(`Property ${proc.propertyType}:`);
                console.log('  Parameters:', proc.parameters.map((p: any) => ({ name: p.name, paramType: p.paramType })));
                console.log('  Body:');
                for (const s of proc.body) {
                    console.log('    Type:', s.type);
                }
            }
        }
    }
}

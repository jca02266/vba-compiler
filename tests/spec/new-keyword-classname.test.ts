import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

function parseStatements(src: string) {
    const tokens = new Lexer(src).tokenize();
    return new Parser(tokens).parse().body;
}

// Bug: Parser rejected `New Collection` because `Collection` is tokenized as
// KeywordCollection, not Identifier. The check `classNameToken.type !== Identifier`
// threw, causing error recovery to abort the containing function declaration.

// Test 1: New Collection in function return — declaration survives
{
    const stmts = parseStatements(`
Function MakeList() As Collection
    Set MakeList = New Collection
End Function
`);
    assert.strictEqual(stmts.length, 1, 'Function declaration not aborted by New Collection');
    assert.strictEqual(stmts[0].type, 'ProcedureDeclaration', 'Parsed as ProcedureDeclaration');
    console.log('[PASS] New Collection: function declaration survives parsing');
}

// Test 2: New Collection with body statements — no statements leak to top level
{
    const stmts = parseStatements(`
Function MakeList() As Collection
    Set MakeList = New Collection
    MakeList.Add 42
End Function
Function OtherFunc()
    OtherFunc = 1
End Function
`);
    assert.strictEqual(stmts.length, 2, 'Both function declarations present (no leaked statements)');
    assert.strictEqual(stmts[0].type, 'ProcedureDeclaration', 'First is ProcedureDeclaration');
    assert.strictEqual(stmts[1].type, 'ProcedureDeclaration', 'Second is ProcedureDeclaration');
    console.log('[PASS] New Collection with body: no statements leak to top level');
}

// Test 3: New Scripting.Dictionary — dotted class name captured correctly
{
    const stmts = parseStatements(`
Function MakeDict()
    Set MakeDict = New Scripting.Dictionary
End Function
`);
    assert.strictEqual(stmts.length, 1, 'New Scripting.Dictionary parses as one declaration');
    // Find NewExpression in the body
    const body = (stmts[0] as any).body;
    function findNew(nodes: any[]): any {
        for (const n of nodes) {
            if (!n) continue;
            if (n.type === 'NewExpression') return n;
            for (const v of Object.values(n)) {
                if (Array.isArray(v)) { const r = findNew(v); if (r) return r; }
                if (v && typeof v === 'object' && (v as any).type) { const r = findNew([v]); if (r) return r; }
            }
        }
        return null;
    }
    const newExpr = findNew(body);
    assert.strictEqual(newExpr?.className, 'Scripting.Dictionary', 'className captured as dotted name');
    console.log('[PASS] New Scripting.Dictionary: className = "Scripting.Dictionary"');
}

// Test 4: New VBA.Collection — keyword after dot
{
    const stmts = parseStatements(`
Function F()
    Set F = New VBA.Collection
End Function
`);
    assert.strictEqual(stmts.length, 1, 'New VBA.Collection parses as one declaration');
    console.log('[PASS] New VBA.Collection: parses without error');
}

console.log('\n✅ new-keyword-classname: 全テスト通過');

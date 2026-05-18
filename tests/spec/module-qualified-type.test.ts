import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { assert } from '../../test-libs/test-runner';

// Bug: `As Module.Type` in function return type declarations was only partially
// consumed — the parser read `Module` as the return type and left `.Type` in the
// token stream. When Type was a keyword (e.g. VBA.Collection), the body parser
// then failed on `.Collection`, aborting the entire function declaration.

function parseStatements(src: string) {
    const tokens = new Lexer(src).tokenize();
    return new Parser(tokens).parse().body;
}

// Test 1: As VBA.Collection return type
{
    const stmts = parseStatements(`
Function F() As VBA.Collection
End Function
`);
    assert.strictEqual(stmts.length, 1, 'As VBA.Collection: function declaration not lost');
    assert.strictEqual(stmts[0].type, 'ProcedureDeclaration', 'Parsed as ProcedureDeclaration');
    assert.strictEqual((stmts[0] as any).returnType, 'VBA.Collection', 'returnType captured as dotted name');
    console.log('[PASS] As VBA.Collection: returnType = "VBA.Collection"');
}

// Test 2: As Scripting.Dictionary return type
{
    const stmts = parseStatements(`
Function F() As Scripting.Dictionary
End Function
`);
    assert.strictEqual(stmts.length, 1, 'Function not aborted');
    assert.strictEqual((stmts[0] as any).returnType, 'Scripting.Dictionary', 'dotted returnType captured');
    console.log('[PASS] As Scripting.Dictionary: returnType = "Scripting.Dictionary"');
}

// Test 3: Dim x As Scripting.Dictionary
{
    const stmts = parseStatements(`
Sub S()
    Dim d As Scripting.Dictionary
End Sub
`);
    assert.strictEqual(stmts.length, 1, 'Sub parsed');
    const body = (stmts[0] as any).body;
    const dimStmt = body.find((s: any) => s.type === 'VariableDeclaration');
    assert.strictEqual(dimStmt?.declarations[0]?.objectType, 'Scripting.Dictionary', 'Dim objectType is dotted');
    console.log('[PASS] Dim As Scripting.Dictionary: objectType = "Scripting.Dictionary"');
}

// Test 4: Dim x As VBA.Collection (keyword after dot)
{
    const stmts = parseStatements(`
Sub S()
    Dim col As VBA.Collection
End Sub
`);
    assert.strictEqual(stmts.length, 1, 'Sub parsed');
    const body = (stmts[0] as any).body;
    const dimStmt = body.find((s: any) => s.type === 'VariableDeclaration');
    assert.strictEqual(dimStmt?.declarations[0]?.objectType, 'VBA.Collection', 'Dim objectType keyword after dot');
    console.log('[PASS] Dim As VBA.Collection: objectType = "VBA.Collection"');
}

// Test 5: non-qualified types still work
{
    const stmts = parseStatements(`
Function F() As String
    F = "hello"
End Function
`);
    assert.strictEqual((stmts[0] as any).returnType, 'String', 'plain return type unaffected');
    console.log('[PASS] As String: plain return type still works');
}

console.log('\n✅ module-qualified-type: 全テスト通過');

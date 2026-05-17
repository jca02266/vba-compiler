import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { findByRefAssignments } from '../../test-libs/vba-analyzer';
import { assert } from '../../test-libs/test-runner';

function parseProc(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    return ast.body.find((s: any) => s.type === 'ProcedureDeclaration');
}

// ByRef パラメータに代入している場合は検出
{
    const proc = parseProc(`
Sub Calc(ByRef result As Long, ByVal x As Long)
    result = x * 2
End Sub
`);
    const r = findByRefAssignments(proc);
    assert.strictEqual(r.length, 1, 'ByRef result を検出');
    assert.strictEqual(r[0].paramName, 'result', 'paramName');
    assert.strictEqual(r[0].paramType, 'Long', 'paramType');
    assert.strictEqual(r[0].lines.length, 1, '代入1件');
    console.log('[PASS] ByRef パラメータへの代入を検出');
}

// ByVal パラメータへの代入は検出しない
{
    const proc = parseProc(`
Sub Calc(ByVal x As Long)
    x = 10
End Sub
`);
    const r = findByRefAssignments(proc);
    assert.strictEqual(r.length, 0, 'ByVal は検出しない');
    console.log('[PASS] ByVal パラメータは検出しない');
}

// VBA デフォルト（修飾子なし）は ByRef として検出
{
    const proc = parseProc(`
Sub GetValues(outA As Long, outB As String)
    outA = 42
    outB = "hello"
End Sub
`);
    const r = findByRefAssignments(proc);
    assert.strictEqual(r.length, 2, '修飾子なしは ByRef として検出');
    console.log('[PASS] 修飾子なしパラメータ（ByRef デフォルト）を検出');
}

// 複数回代入は行をすべて収集
{
    const proc = parseProc(`
Sub Multi(ByRef n As Long)
    n = 1
    n = 2
    n = 3
End Sub
`);
    const r = findByRefAssignments(proc);
    assert.strictEqual(r[0].lines.length, 3, '3回の代入をすべて収集');
    console.log('[PASS] 複数回代入の行を収集');
}

// ByRef パラメータを読むだけ（代入なし）は検出しない
{
    const proc = parseProc(`
Sub ReadOnly(ByRef n As Long)
    Dim x As Long
    x = n + 1
End Sub
`);
    const r = findByRefAssignments(proc);
    assert.strictEqual(r.length, 0, '読み取りのみは検出しない');
    console.log('[PASS] ByRef パラメータの読み取りのみは検出しない');
}

// パラメータなしは空配列
{
    const proc = parseProc(`
Sub NoParams()
    Dim x As Long
    x = 1
End Sub
`);
    const r = findByRefAssignments(proc);
    assert.strictEqual(r.length, 0, 'パラメータなしは空');
    console.log('[PASS] パラメータなしは空配列');
}

console.log('\n✅ byref-assignments: 全テスト通過');

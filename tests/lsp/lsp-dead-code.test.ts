import { LSPServer } from '../../src/lsp/server';
import { assert } from '../../test-libs/test-runner';

const URI = 'file:///test.bas';

function getDiags(src: string) {
    const server = new LSPServer();
    server.didOpen(URI, src);
    return server.getDiagnostics(URI);
}

// 1. 参照のない Private Sub → Warning が出る
{
    const src = [
        'Private Sub UnusedHelper()',
        'End Sub',
        'Sub Main()',
        'End Sub',
    ].join('\n');

    const diags = getDiags(src);
    const warn = diags.find((d: any) => d.message.includes('UnusedHelper'));
    assert.ok(warn, 'UnusedHelper に警告が出る');
    assert.strictEqual(warn.severity, 2, 'severity = 2 (warning)');
    assert.strictEqual(warn.source, 'vba-runner', 'source = vba-runner');
    assert.strictEqual(warn.range.start.line, 0, '行 0 を指す');
    console.log('[PASS] Private 未参照 Sub に Warning:', warn.message);
}

// 2. 参照がある Private Sub → 警告なし
{
    const src = [
        'Private Sub Helper()',
        'End Sub',
        'Sub Main()',
        '    Helper()',
        'End Sub',
    ].join('\n');

    const diags = getDiags(src);
    const warn = diags.find((d: any) => d.message.includes('Helper'));
    assert.ok(!warn, '参照がある Private Sub に警告なし');
    console.log('[PASS] 参照あり Private Sub: 警告なし');
}

// 3. 参照のない Public Sub → エントリーポイント候補のため警告なし
{
    const src = [
        'Sub EntryPoint()',
        'End Sub',
    ].join('\n');

    const diags = getDiags(src);
    const warn = diags.find((d: any) => d.message.includes('EntryPoint'));
    assert.ok(!warn, 'Public 未参照 Sub に警告なし（エントリーポイント候補）');
    console.log('[PASS] Public 未参照 Sub: 警告なし');
}

// 4. パースエラーと Dead code 警告が共存する
{
    const src = [
        'Private Sub Dead()',
        'End Sub',
        '@@@bad',
        'Sub Main()',
        'End Sub',
    ].join('\n');

    const diags = getDiags(src);
    const parseErr = diags.find((d: any) => d.severity === 1);
    const deadWarn = diags.find((d: any) => d.message.includes('Dead'));
    assert.ok(parseErr, 'パースエラーが出る');
    assert.ok(deadWarn, 'Dead code 警告も出る');
    console.log('[PASS] パースエラーと Dead code 警告の共存');
}

// 5. Private Function も検出される
{
    const src = [
        'Private Function Unused() As Long',
        '    Unused = 0',
        'End Function',
        'Sub Main()',
        'End Sub',
    ].join('\n');

    const diags = getDiags(src);
    const warn = diags.find((d: any) => d.message.includes('Unused'));
    assert.ok(warn, 'Private Function にも警告');
    console.log('[PASS] Private Function: Dead code 警告');
}

console.log('\n✅ Dead code Detection: 全テスト通過');

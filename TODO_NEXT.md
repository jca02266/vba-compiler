# 次期開発ロードマップ (Phase 2: VSCode IDE統合)

本プロジェクトの付加価値は **VBAを実際に実行できること** と **レガシーVBAのリファクタリング支援** にあります。
この2点を軸に、VSCode拡張機能としての完成度を高めていきます。

> Phase 1（MS-VBAL 仕様書に列挙された構文要素・標準ライブラリ関数の実装）は完了。
> 仕様書本文に記載されたランタイム挙動の細部は [`TODO.md` の「VBA ランタイム挙動」](TODO.md#vba-ランタイム挙動) でトラッキング中。

---

## Step 1 — 繋ぐだけで使えるもの（実装済み・未接続）

実装とテストは存在するが、extension.ts に登録されていないだけの機能。

- [ ] **Diagnostics（構文エラーの破線表示）**
  - [x] Parser が `ast.diagnostics[]` にエラーを収集（エラー耐性パース）。テスト: `lsp-diagnostics.test.ts`
  - [ ] DiagnosticsProvider クラスの作成（`ast.diagnostics` → `vscode.Diagnostic[]` 変換）
  - [ ] extension.ts への登録（`createDiagnosticCollection`、ドキュメント変更時に更新）
- [ ] **DocumentSymbol（アウトライン表示）**
  - [x] SymbolProvider 実装済み。テスト: `lsp-symbol-provider.test.ts`
  - [ ] extension.ts への登録（`registerDocumentSymbolProvider`）
- [ ] **DAP（デバッガ）の VSCode 接続**
  - [x] ステップ実行・ブレークポイント・変数表示 実装済み。テスト: `lsp-debugger.test.ts`, `lsp-debug-adapter.test.ts`
  - [ ] extension.ts への登録（`registerDebugAdapterDescriptorFactory`）

---

## Step 2 — リファクタリングに必須の LSP 機能

これがないと「安全に変更できる」とは言えない。シンボルテーブルが既にあるので実装コストは比較的低い。

- [ ] **Find All References（Shift+F12）**
  - 「この Sub はどこから呼ばれているか」を一覧表示。削除・改名の影響範囲を把握するために必須
  - `textDocument/references` の実装と extension.ts への登録
- [ ] **Rename Symbol（F2）**
  - Sub / 変数の名前を全ファイル横断で一括変更
  - `textDocument/rename` の実装と extension.ts への登録

---

## Step 3 — 「実行できる」固有の価値を活かす機能

他のVBAエディタとの差別化。VBAを実行できるこのプロジェクトならではの機能。

- [ ] **Code Lens（各 Sub/Function 上のインライン情報）**
  ```
  ▶ Run  |  3 references  |  未テスト
  Function CalcTotal(a As Long, b As Long) As Long
  ```
  - `▶ Run` ボタン：クリックで即実行し結果を Output Panel に表示
  - `N references`：参照元数を表示（`0 references` は削除候補）
  - `未テスト / テスト済み`：`Test_` プロシージャの有無で表示
- [ ] **Dead code 検出**
  - 参照が 0 件のプロシージャを Diagnostics として警告（レガシーVBAには未使用関数が大量にある）
- [ ] **VBA to TypeScript トランスパイラ**
  - 構築済みの AST から TypeScript コードを自動生成する Generator の実装
  - リファクタリングの最終ゴール：VBA からのモダン言語移行支援

---

## 完了済みの機能

### VSCode 拡張機能（extension.ts に接続済み）
- [x] Tolerant Parsing — Lexer 列番号、Parser `loc` 位置情報、エラー耐性パース
- [x] Hover — シンボルのシグネチャ表示。テスト: `lsp-hover.test.ts`
- [x] Go to Definition（F12）— シンボル宣言位置へジャンプ。テスト: `lsp-definition.test.ts`
- [x] Completion — 標準関数・変数補完。テスト: `lsp-completion.test.ts`
- [x] Test Explorer — `Test_` プロシージャの自動検出と実行。テスト: `lsp-test-discovery.test.ts`, `lsp-test-runner.test.ts`

### テスト基盤（TypeScript テストランナー）
- [x] Spy / Mock API（`MsgBox`, `Shell` 等の副作用検証）。テスト: `spy-mock-api.test.ts`
- [x] Time Mocking（`Now` / `Date` の固定）。テスト: `time-mocking.test.ts`
- [x] In-Memory FS（`MemoryFileSystem`）の基礎実装

### Extension インフラ
- [x] `src/extension.ts`（Desktop 用 Node.js 版）
- [x] `language-configuration.json`, `syntaxes/vba.tmLanguage.json`
- [x] `package.json`（VSCode extension metadata）、`.vscodeignore`

---

## 保留中・将来課題

### パッケージング
- [ ] VSCode Marketplace への公開（アイコン整備、`.vsix` ビルド検証）
- [ ] Web Extension 化（evaluator が Node.js `path` に依存しており、先にブラウザ対応が必要）

### Web UI（デモサイト）の改善
- [ ] `Dir` 関数の完全実装（ディレクトリ列挙）
- [ ] `Kill` ワイルドカード対応（例: `Kill "*.txt"`）
- [ ] コールスタックの出力・`Erl` 関数サポート

---

**最終ビジョン**:
VBA開発者が VBE を開かずに VSCode だけで、LSP 補完を受けながらコードを書き、
参照・リネームで安全にリファクタリングし、テストを即座に実行し、
ステップデバッグで挙動を確認できる「モダンなVBA開発環境」を実現する。

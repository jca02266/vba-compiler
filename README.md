# VBA Web Execution Environment

Excelがない環境で、作成したVBAコードの動作確認、リファクタリング、および自動テスト（ユニットテスト）を実行するためのプロジェクト（VBAパーサーおよびAST評価器）です。React + TypeScript のWeb UIと、CLIによるテストランナー環境を備えています。

## 本プロジェクトの目的
- **環境非依存での動作確認**: Excel (Windows/Mac) に依存せず、モダンなブラウザ上で直接VBAの構文とロジックを実行・検証できます。
- **リファクタリングの支援**: 巨大なVBAコードベースから純粋な関数・サブルーチンを安全に切り出し、モジュールを整理するための検証基盤を提供します。
- **ユニットテストの実行**: TypeScriptを利用したテストランナーを通じて、抽出されたVBA関数に対して直接モックデータやアサーションを評価し、プログラムによるテスト自動化を可能にします。

## ディレクトリ構成
- `src/compiler/` - TypeScriptで書かれたVBA用Lexer、Parser、Evaluatorのコアエンジン群。
- `src/` - React を用いたWeb IDE（エディタ、コンソール出力）のUIソース。
- `tests/vba/` - リファクタリング対象の元のVBAスクリプトおよび、整理抽出された関数が含まれるVBAファイル群 (`test.vba`, `test_refactored.vba`)。
- `tests/ts/` - Node.js 上で評価器を直接バインディングし、自動テストを走らせるTypeScriptのテストランナー・スイート。

## 使い方（UI画面の起動）
ブラウザ上でVBAエディタ環境を立ち上げ、`Debug.Print` クラスの動作などを確認します。

```bash
# 依存関係のインストール
npm install

# 開発用ローカルサーバーの起動 (http://localhost:5173/)
npm run dev
```

## 自動テストの実行
CLIからTypeScriptのテストランナーをコールし、`tests/vba/` 内のモジュールを評価してアサーションを実行します。

```bash
# バンドルしてCJSスクリプトを実行
npx esbuild tests/ts/run.ts --bundle --outfile=tests/ts/run.cjs --platform=node && node tests/ts/run.cjs
```

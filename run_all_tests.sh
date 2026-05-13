#!/bin/bash

# プロジェクト内のすべてのテストを一括実行するスクリプト
#  - tests/spec/             : VBA コンパイラの仕様テスト
#  - tests/test-libs-tests/  : test-libs/ (VBATest 等) のテスト
#  - sample/tests/ts/        : サンプル VBA コードのテスト

# 失敗時に停止
set -e

echo "--- Starting all tests ---"

for f in tests/spec/*.test.ts tests/test-libs-tests/*.test.ts sample/tests/ts/*.test.ts; do
  if [ -f "$f" ]; then
    out="${f%.ts}.cjs"
    echo "Running $f..."
    ./node_modules/.bin/esbuild "$f" --bundle --outfile="$out" --platform=node
    node "$out"
    # 中間ファイルを削除したい場合はコメント解除
    # rm "$out"
  fi
done

echo "--- All tests completed successfully! ---"

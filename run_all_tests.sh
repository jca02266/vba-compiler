#!/bin/bash

# プロジェクト内のすべてのテストを一括実行するスクリプト
#  - tests/spec/             : VBA コンパイラの仕様テスト
#  - tests/test-libs-tests/  : test-libs/ (VBATest 等) のテスト
#  - sample/tests/ts/        : サンプル VBA コードのテスト

echo "--- Starting all tests ---"

# テスト失敗フラグ
TESTS_FAILED=0

for f in tests/spec/*.test.ts tests/test-libs-tests/*.test.ts sample/tests/ts/*.test.ts; do
  if [ -f "$f" ]; then
    out="${f%.ts}.cjs"
    echo "Running $f..."

    # esbuild でバンドル
    if ! ./node_modules/.bin/esbuild "$f" --bundle --outfile="$out" --platform=node; then
      echo "❌ Build failed: $f"
      TESTS_FAILED=1
      continue
    fi

    # テスト実行
    if ! node "$out"; then
      echo "❌ Test failed: $f"
      TESTS_FAILED=1
    fi

    # 中間ファイルを削除したい場合はコメント解除
    # rm "$out"
  fi
done

echo ""
if [ $TESTS_FAILED -eq 0 ]; then
  echo "--- All tests completed successfully! ---"
  exit 0
else
  echo "--- Some tests FAILED ---"
  exit 1
fi

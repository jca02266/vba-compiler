#!/bin/bash

# tests/spec/ 配下のすべてのテストを一括実行するスクリプト

echo "--- Starting all spec tests ---"

# テスト失敗フラグ
TESTS_FAILED=0

for f in tests/spec/*.test.ts; do
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
done

echo ""
if [ $TESTS_FAILED -eq 0 ]; then
  echo "--- All spec tests completed successfully! ---"
  exit 0
else
  echo "--- Some spec tests FAILED ---"
  exit 1
fi

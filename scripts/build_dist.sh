#!/usr/bin/env bash
# 构建单文件可执行程序（使用 Bun，自带运行时；目标机器无需安装 node）
# 产物：dist/tag2bookmark-{darwin-arm64,darwin-x64,windows-x64(.exe)}
set -e
cd "$(dirname "$0")/.."
BUN="${BUN:-$HOME/.bun/bin/bun}"
command -v "$BUN" >/dev/null 2>&1 || BUN=bun
[ -x "$BUN" ] || { echo "未找到 bun，请先安装：https://bun.sh"; exit 1; }
mkdir -p dist
echo "== 构建 darwin-arm64 =="
"$BUN" build cli/tag2bookmark.cjs --compile --target=bun-darwin-arm64 --outfile dist/tag2bookmark-darwin-arm64
echo "== 构建 darwin-x64 =="
"$BUN" build cli/tag2bookmark.cjs --compile --target=bun-darwin-x64 --outfile dist/tag2bookmark-darwin-x64
echo "== 构建 windows-x64 =="
"$BUN" build cli/tag2bookmark.cjs --compile --target=bun-windows-x64 --outfile dist/tag2bookmark-windows-x64.exe
# macOS 需要 ad-hoc 签名，否则系统会以"未签名"直接终止
if [ "$(uname -s)" = "Darwin" ]; then
  echo "== ad-hoc 签名 mac 产物 =="
  codesign --force --sign - dist/tag2bookmark-darwin-arm64 >/dev/null 2>&1 || true
  codesign --force --sign - dist/tag2bookmark-darwin-x64 >/dev/null 2>&1 || true
fi
echo "== 完成 =="
ls -lh dist/

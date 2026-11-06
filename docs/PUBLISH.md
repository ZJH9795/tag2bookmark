# 发布到 GitHub 的步骤

本目录已是一个可直接初始化为 Git 仓库的完整项目（含 .gitignore，不会带出任何内部素材）。

## 1. 初始化并推送

```bash
cd github-release
git init -b main
git add .
git commit -m "feat: Tag2Bookmark v1.0.0 - convert ${placeholder} text into Word bookmarks"
# 在 GitHub 上新建空仓库（不要勾选 README/LICENSE），然后：
git remote add origin git@github.com:<你的用户名>/tag2bookmark.git
git push -u origin main
```

## 2. 构建并上传 Release 二进制

```bash
bash scripts/build_dist.sh        # 产出 dist/ 三个平台二进制
# GitHub 网页：Releases → Draft a new release → tag v1.0.0
# 上传 dist/tag2bookmark-darwin-arm64、tag2bookmark-darwin-x64、tag2bookmark-windows-x64.exe
```

> Windows exe 未做代码签名，首次运行可能被 SmartScreen 提示「更多信息 → 仍要运行」；
> macOS 二进制已做 ad-hoc 签名，若被隔离属性拦截可执行 `xattr -d com.apple.quarantine <文件>`。

## 3. 发布前自检

```bash
python3 scripts/scan_sensitive.py   # 敏感信息扫描（必须通过）
npm install
npm test                            # 样例自测（必须全过）
```

## 4. 版本建议

- 语义化版本（SemVer）：`v1.0.0` 起；
- 每次改动更新 `package.json` 的 version 与 Release Notes。

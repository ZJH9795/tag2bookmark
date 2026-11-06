# Tag2Bookmark（标签转书签）

把 Word 文档里的 `${字段名}` 占位符批量转换成 Word 书签的通用工具。

适用于模板填充、公文批处理等场景：很多模板引擎（Aspose.Words、poi-tl、Word 域、WPS 二次开发等）按书签名定位插入点，把模板中的 `${xxx}` 文本一次性变成书签，后续系统就能按名取数、精准填充。

**纯本地处理，不上传任何文件；不修改文档的保护/加密状态。**

## 功能

- **三连处理**（可在界面单独开关）：
  1. **占位符唯一化**：同名占位符按出现顺序改写为 `${名}`、`${名_1}`、`${名_2}`…（全文档连续编号）；
  2. **转书签**：在每个占位符起始处创建书签，命名规则见下；
  3. **删除标签**：清除占位符文本本身（书签保留）。
- 覆盖范围：正文、页眉、页脚（含文本框内段落）；支持跨 run 的拆分占位符（如 `${合同编` + `号}`）。
- 输入格式：`.docx / .docm / .dotx / .dotm`；CLI 版额外支持 `.doc`（97-2003）**保真管道**。
- 批量处理、逐文件处理明细日志、打包下载。

## 书签命名规则

1. 保留中文、字母、数字、下划线，其余字符替换为 `_`；
2. 合并连续 `_`、去掉首尾 `_`；清空后兜底为 `Placeholder`；
3. 数字开头加前缀 `B_`；
4. 长度截断至 40 字符；
5. 与文档已有书签重名时自动追加 `_1`、`_2` 递增后缀。

## 两种使用方式

### 1. 网页版（开箱即用）

双击 `web/tag2bookmark.html`（单文件，JSZip 已内嵌，离线可用）：
拖入文件 → 选择处理步骤 → 开始处理 → 下载或打包下载。全程浏览器本地完成。

### 2. CLI / 桌面版（支持 .doc 与全保真）

```bash
# 源码运行（需要 node ≥ 18）
node cli/tag2bookmark.cjs 文件1.doc 文件2.doc        # 或目录

# 也可以直接用编译好的单文件可执行程序（Release 页下载，或自行构建）
./dist/tag2bookmark-macos-arm64 文件1.doc
dist\\tag2bookmark-windows-x64.exe 文件1.doc
```

CLI 对 `.doc` 采用**保真管道**：
`doc --(文档引擎)--> docx --(本工具三连)--> docx --(文档引擎)--> doc`。
格式转换由本机文档引擎完成，样式全程无损；输出命名为 `原名_已转书签.doc`（重名自动加时间戳，不会覆盖任何文件）。

**引擎自适配与平台差异：**

| 平台 | .doc 转换引擎 | .doc 输出书签保留 |
|---|---|---|
| Windows | 优先 WPS / Microsoft Word（COM 自动化），其次 LibreOffice | ✅ 原生引擎，完整保留（含表格内书签） |
| macOS / Linux | LibreOffice | ⚠️ LibreOffice 的 docx→doc 会丢弃**表格内**书签（其已知限制）。工具会自动校验：发现书签丢失时**自动改输出 `.docx`**（样式保真、书签完整）并给出警告；如确需 `.doc` 可加 `--force-doc`（接受书签丢失风险） |

> 也就是说：需要「.doc 格式 + 书签完整」的正式交付，请使用 **Windows 版（WPS/Word）**；macOS 上建议直接交付 `.docx`（现代系统均支持）。

依赖（仅 CLI 的 .doc 保真管道需要）：
- Windows：本机已装的 WPS 或 Microsoft Word（推荐，零额外安装）；或 [LibreOffice](https://www.libreoffice.org/)
- macOS / Linux：[LibreOffice](https://www.libreoffice.org/)
- 处理 `.docx/.docm` 时无任何外部依赖。

## 构建单文件可执行程序（可选）

利用 [Bun](https://bun.sh) 把 CLI 编译成自带运行时的单文件程序，目标机器**不需要**安装 node：

```bash
bun install
bash scripts/build_dist.sh        # 产出 dist/tag2bookmark-{darwin-arm64,darwin-x64,windows-x64}(.exe)
```

## CLI 选项

```
--steps unique,bookmarks,deleteTags   处理步骤组合（默认全三步）
--out <目录>                          输出目录（默认源文件同目录）
--force-doc                           .doc 回写即使丢失书签也强制输出 .doc
-h, --help                            帮助
```

## 关于文档保护与解锁密码

- 本工具 **不读取、不校验、不修改** 任何文档保护设置（限制编辑/窗体保护原样保留在输出文件中）；
- 带「打开密码」的加密文件无法直接处理，请先用 Word/WPS 输入密码打开，另存为不受保护的 `.docx` 后再处理（工具会给出明确提示）。
- 参考宏 `macros/` 中包含一个「解除保护 → 转换 → 恢复保护」的示例实现，其中密码由**运行时输入**，仓库中不存在任何硬编码密码。

## 目录结构

```
web/        网页版单文件工具与其构建源码
cli/        Node/Bun 版 CLI（含 .doc 保真管道）与无浏览器转换内核
macros/     WPS JSA 参考宏（脱敏版，密码运行时输入）
samples/    程序生成的通用测试样例与生成脚本
scripts/    构建 / 验证脚本
docs/       实现口径与验收说明
```

## 开源许可

本项目代码以 [MIT](LICENSE) 发布。依赖组件的许可声明见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)：
JSZip（MIT / GPLv3 双授权，本项目使用 MIT 分支）、@xmldom/xmldom（MIT）、Bun（MIT，用于可选的单文件编译）。

.doc ↔ .docx 双向转换依赖 separate 的文档引擎（LibreOffice 为 MPL 2.0，WPS/Word 为闭源商业软件），
本项目不打包、不分发、不修改其任何代码，仅以命令行/COM 方式调用用户本机已安装的程序。

## 免责声明

本工具仅做「文本占位符 → 书签」的内容改写，不校验文档业务内容；请在正式使用前用样例文档核对输出是否符合你的模板引擎要求。

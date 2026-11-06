# Third-Party Licenses（第三方组件许可声明）

本仓库的原创代码以 MIT 许可证发布（见 [LICENSE](LICENSE)）。下列第三方组件随本项目分发或参与可选构建，依据其许可证要求在此保留版权与许可声明。

## JSZip

- 用途：OOXML（docx/docm）包的读写；`web/` 版内嵌其压缩版源码。
- 许可证：**MIT 或 GPLv3 双授权**。本项目按其许可条款**选择 MIT 分支**使用，不适用 GPLv3。
- 版权：
```
Copyright (c) 2009-2026 Stuart Knightley, David Duponchel, Franz Buchinger, António Afonso
```
- MIT 文本（JPEG/JSZip 声明原文）：
```
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
- 官方仓库：<https://github.com/Stuk/jszip>

## @xmldom/xmldom

- 用途：`cli/` 无浏览器转换内核的 XML DOM 解析与序列化（替代浏览器 DOMParser）。
- 许可证：**MIT**（ актуальная версия 0.8.x 采用 MIT）。
- 版权：
```
Copyright (c) 2019-2023 broofa <jason.denis@gmail.com>
Copyright (c) 2019-2023 Christoph Dürr <contact@chronoid.com> and contributors
（原 jsdom-wrapped xmlparser 系列作者 jindw 等，见包内 LICENSE文件）
```
- MIT 义务：随发行版保留上述版权与许可声明（本文件即为履行）。
- 官方仓库：<https://github.com/xmldom/xmldom>

## Bun（可选参与：单文件可执行程序构建）

- 用途：`scripts/build_dist.sh` 使用 Bun 的 `bun build --compile` 产出自带运行时的单文件二进制。
- 许可证：**MIT**。编译出的二进制内嵌 Bun 运行时，因此从 Release 下载二进制的用户同样受其 MIT 条款保障并保留声明（本文件即为随附声明）。
- 版权：
```
Copyright (c) 2021-present, Jarred Sumner and Bun contributors
```
- MIT 义务：同上，随发行版保留本声明即可。
- 官方仓库：<https://github.com/oven-sh/bun>
- 注意：Bun 运行时内部聚合了若干第三方组件，其完整清单与厂商声明见 Bun 官方仓库 [`LICENSE`](https://github.com/oven-sh/bun/blob/main/LICENSE) 与 [`THIRD_PARTY_LICENSES.md`](https://github.com/oven-sh/bun/blob/main/THIRD_PARTY_LICENSES.md)。为完整履行该等第三方组件的声明义务，发布 Release 二进制时本文件应与该官方文件一同附带（或在 Release 说明中给出链接）。

## 外部文档引擎（不作分发）

CLI 的 `.doc` 保真管道调用用户本机已安装的文档引擎执行 doc↔docx 转换：

- LibreOffice —— MPL 2.0。本项目不打包、不修改、不反向集成其任何代码，仅以命令行方式调用；MPL 2.0 对此类独立进程调用无源码义务。
- WPS Office / Microsoft Word —— 闭源商业软件。仅通过其公开的 COM 自动化接口调用用户已安装的副本。

因此本仓库不包含、也不需要附带上述引擎任何部分的源码或二进制。

#!/usr/bin/env node
/**
 * Tag2Bookmark CLI（通用版）
 *
 * 三连处理：占位符唯一化 → 书签（占位符起始处）→ 删除标签
 * 输入：.docx/.docm/.dotx/.dotm 直接处理；.doc 走保真管道（内部引擎 doc→docx→三连→doc）
 * 输出：与源文件同目录，命名 `原名_已转书签.<原扩展名>`（.doc 保持 .doc）；重名自动加时间戳，绝不覆盖
 *
 * 用法：
 *   node tag2bookmark.mjs 文件1.docx [文件2.doc ...]     # 参数可为文件或目录
 *   node tag2bookmark.mjs --steps unique,bookmarks       # 自定义步骤（默认全三步）
 *   node tag2bookmark.mjs --out 输出目录 文件1.doc       # 指定输出目录
 *
 * 依赖：node ≥ 18；处理 .doc 时需要本机文档引擎（Windows: WPS/Word/LibreOffice，macOS/Linux: LibreOffice）
 */
'use strict';

const path = require('node:path');
const fs = require('node:fs');
const os = require('node:os');
const { spawn } = require('node:child_process');

const HERE = __dirname;
let _core = null;
function coreMod() { if (!_core) _core = require('./node_core.cjs'); return _core; }

/* JSZip：优先 npm 内联（bun 编译可打包），其次仓库内 vendored 文件 */
function loadJSZip() {
  try { return require('jszip'); } catch (e) {}
  const vendored = path.join(HERE, '..', 'web', 'dep', 'jszip.min.js');
  if (fs.existsSync(vendored)) return require(vendored);
  throw new Error('缺少依赖 jszip：请执行 npm install，或使用官方发行版二进制');
}

/* @xmldom/xmldom */
function loadXmldom() {
  try { return require('@xmldom/xmldom'); } catch (e) {
    throw new Error('缺少依赖 @xmldom/xmldom：请执行 npm install，或使用官方发行版二进制');
  }
}

/* ================= 文档引擎（保真管道） ================= */
function findSoffice() {
  const candidates = process.platform === 'darwin'
    ? ['/Applications/LibreOffice.app/Contents/MacOS/soffice']
    : process.platform === 'win32'
      ? ['C:\\Program Files\\LibreOffice\\program\\soffice.exe',
         'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe']
      : ['/usr/bin/soffice', '/opt/libreoffice/program/soffice'];
  for (const c of candidates) { try { if (fs.existsSync(c)) return c; } catch (e) {} }
  return null;
}

function run(cmd, args, timeoutMs) {
  return new Promise((resolve) => {
    const child = spawn(cmd, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let out = '', err = '';
    const to = setTimeout(() => {
      try { child.kill(); } catch (e) {}
      resolve({ ok: false, err: 'timeout' });
    }, timeoutMs || 240000);
    child.stdout.on('data', (d) => { out += d; });
    child.stderr.on('data', (d) => { err += d; });
    child.on('error', (e) => { clearTimeout(to); resolve({ ok: false, err: String((e && e.message) || e) }); });
    child.on('close', (code) => { clearTimeout(to); resolve({ ok: code === 0, code, out, err }); });
  });
}

let _tmp = null;
function tmpDir() {
  if (!_tmp) _tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'tag2bookmark-'));
  return _tmp;
}

/** 通过 PowerShell（-EncodedCommand，UTF-16LE Base64）调用 WPS/Word COM 做格式转换
 *  format: 'docx'（目标 .docx）| 'doc'（目标 .doc）
 *  返回输出文件路径或 null
 */
async function windowsComConvert(srcPath, dstPath, format) {
  // 12 = wdFormatXMLDocument(.docx)；0 = wdFormatDocument97(.doc)
  const fmt = format === 'docx' ? 12 : 0;
  const ps = [
    "$ErrorActionPreference='Stop'",
    "$src='" + srcPath.replace(/'/g, "''") + "'",
    "$dst='" + dstPath.replace(/'/g, "''") + "'",
    "$app=$null",
    "foreach($p in @('KWPS.Application','Kwps.Application','WPS.Application','Word.Application')){",
    "  try { $app=New-Object -ComObject $p; break } catch {}",
    "}",
    "if($app -eq $null){ exit 2 }",
    "try {",
    "  $app.Visible=$false",
    "  try { $app.DisplayAlerts=0 } catch {}",
    "  $doc=$app.Documents.Open($src,$false,$true)",
    "  try { $doc.SaveAs2($dst," + fmt + ") } catch { $doc.SaveAs($dst," + fmt + ") }",
    "  $doc.Close($false)",
    "  $app.Quit()",
    "  exit 0",
    "} catch { try{$app.Quit()}catch{}; exit 3 }",
  ].join('\r\n');
  const b64 = Buffer.from(ps, 'utf16le').toString('base64');
  const r = await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', b64], 240000);
  return (r.ok && fs.existsSync(dstPath)) ? dstPath : null;
}

async function docToDocx(srcPath) {
  const outdir = path.join(tmpDir(), 'in' + Date.now());
  fs.mkdirSync(outdir, { recursive: true });
  const dst = path.join(outdir, path.basename(srcPath).replace(/\.doc$/i, '') + '.docx');
  if (process.platform === 'win32') {
    const viaCom = await windowsComConvert(srcPath, dst, 'docx');
    if (viaCom) return viaCom;
  }
  const soffice = findSoffice();
  if (!soffice) return null;
  const r = await run(soffice, ['--headless', '-env:UserInstallation=file://' + path.join(tmpDir(), 'lo'),
    '--convert-to', 'docx', '--outdir', outdir, srcPath]);
  if (!r.ok) return null;
  if (fs.existsSync(dst)) return dst;
  const base = path.basename(srcPath).replace(/\.doc$/i, '').toLowerCase();
  for (const g of fs.readdirSync(outdir)) {
    if (g.replace(/\.docx$/i, '').toLowerCase() === base) return path.join(outdir, g);
  }
  return null;
}

async function docxToDoc(srcPath) {
  const outdir = path.join(tmpDir(), 'out' + Date.now());
  fs.mkdirSync(outdir, { recursive: true });
  const dst = path.join(outdir, path.basename(srcPath).replace(/\.docx$/i, '') + '.doc');
  if (process.platform === 'win32') {
    const viaCom = await windowsComConvert(srcPath, dst, 'doc');
    if (viaCom) return viaCom;
  }
  const soffice = findSoffice();
  if (!soffice) return null;
  const r = await run(soffice, ['--headless', '-env:UserInstallation=file://' + path.join(tmpDir(), 'lo'),
    '--convert-to', 'doc:MS Word 97', '--outdir', outdir, srcPath]);
  if (!r.ok) return null;
  if (fs.existsSync(dst)) return dst;
  const base = path.basename(srcPath).replace(/\.docx$/i, '').toLowerCase();
  for (const g of fs.readdirSync(outdir)) {
    if (g.replace(/\.doc$/i, '').toLowerCase() === base) return path.join(outdir, g);
  }
  return null;
}

/** 零宽书签垫锚：给零宽书签包 1 空格（引擎回写 doc 时零宽书签会丢，范围书签可保留） */
async function padBookmarks(docxBytes) {
  const zip = await globalThis.JSZip.loadAsync(docxBytes);
  const re = /(<w:bookmarkStart [^>]*\/>)\s*(<w:bookmarkEnd [^>]*\/>)/g;
  let touched = 0;
  for (const name of Object.keys(zip.files)) {
    if (!/^word\/(document|header[^/]*|footer[^/]*)\.xml$/.test(name)) continue;
    const txt = await zip.file(name).async('string');
    const out = txt.replace(re, (m, s1, s2) => {
      touched++;
      return s1 + '<w:r><w:t xml:space="preserve"> </w:t></w:r>' + s2;
    });
    if (out !== txt) zip.file(name, out);
  }
  const buf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  return { bytes: buf, count: touched };
}

/* ---- 书签回写校验：把输出 doc 再转回 docx，核对书签是否保住 ---- */
async function verifyDocBookmarks(docPath, expectNames) {
  try {
    const outdir = path.join(tmpDir(), 'verify' + Date.now());
    fs.mkdirSync(outdir, { recursive: true });
    const back = await docToDocxFromAny(docPath, outdir);
    if (!back) return { ok: null, kept: 0, reason: '无法回读校验' };
    const zip = await globalThis.JSZip.loadAsync(new Uint8Array(fs.readFileSync(back)));
    let names = [];
    for (const n of Object.keys(zip.files)) {
      if (!/^word\/(document|header[^/]*|footer[^/]*)\.xml$/.test(n)) continue;
      const xml = await zip.file(n).async('string');
      const re = /<w:bookmarkStart [^>]*w:name="([^"]+)"/g;
      let m;
      while ((m = re.exec(xml)) !== null) names.push(m[1]);
    }
    const expectSet = new Set(expectNames || []);
    const kept = names.filter((x) => expectSet.has(x)).length;
    return { ok: kept >= expectNames.length, kept: kept, total: expectNames.length, names: names };
  } catch (e) {
    return { ok: null, kept: 0, reason: String((e && e.message) || e) };
  }
}

/** doc→docx（用于校验，不关心来源）*/
async function docToDocxFromAny(docPath, outdir) {
  if (process.platform === 'win32') {
    const dst = path.join(outdir, path.basename(docPath).replace(/\.doc$/i, '') + '.docx');
    const via = await windowsComConvert(docPath, dst, 'docx');
    if (via) return via;
  }
  const soffice = findSoffice();
  if (!soffice) return null;
  const r = await run(soffice, ['--headless', '-env:UserInstallation=file://' + path.join(tmpDir(), 'lo'),
    '--convert-to', 'docx', '--outdir', outdir, docPath]);
  if (!r.ok) return null;
  const base = path.basename(docPath).replace(/\.doc$/i, '').toLowerCase();
  for (const g of fs.readdirSync(outdir)) {
    if (g.replace(/\.docx$/i, '').toLowerCase() === base) return path.join(outdir, g);
  }
  return null;
}

/* ================= 主流程 ================= */
function nonConflict(p) {
  if (!fs.existsSync(p)) return p;
  const d = new Date();
  const pad = (n) => (n < 10 ? '0' : '') + n;
  return p.replace(/(\.[^.]+)$/, '_' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '_' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()) + '$1');
}

function usage() {
  process.stdout.write([
    'Tag2Bookmark —— 把 Word 里的 ${占位符} 转成书签（通用工具）',
    '',
    '用法：node tag2bookmark.mjs [选项] <文件或目录...>',
    '',
    '选项：',
    '  --steps a,b,c   处理步骤组合：unique / bookmarks / deleteTags（默认全三步）',
    '  --out <目录>    输出目录（默认与源文件同目录）',
    '  -h, --help      显示本帮助',
    '',
    '示例：',
    '  node tag2bookmark.mjs 范本1.docx 范本2.doc',
    '  node tag2bookmark.mjs --steps bookmarks ./范本目录',
  ].join('\n') + '\n');
}

async function main() {
  const argv = process.argv.slice(2);
  const steps = { unique: true, bookmarks: true, deleteTags: true };
  let outDir = null;
  let forceDoc = false;
  const files = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--help' || a === '-h') { usage(); return; }
    if (a === '--steps' && argv[i + 1]) {
      const keys = String(argv[++i]).split(',').map((s) => s.trim());
      steps.unique = keys.includes('unique');
      steps.bookmarks = keys.includes('bookmarks');
      steps.deleteTags = keys.includes('deleteTags');
      continue;
    }
    if (a === '--out' && argv[i + 1]) { outDir = argv[++i]; continue; }
    if (a === '--force-doc') { forceDoc = true; continue; }
    if (a.startsWith('--')) continue;
    files.push(a);
  }
  const expanded = [];
  for (const f of files) {
    const st = fs.statSync(f);
    if (st.isDirectory()) {
      for (const g of fs.readdirSync(f)) {
        if (/\.(doc|docx|docm|dotx|dotm)$/i.test(g) && g.indexOf('_已转书签') < 0) expanded.push(path.join(f, g));
      }
    } else if (/\.(doc|docx|docm|dotx|dotm)$/i.test(f)) {
      expanded.push(f);
    }
  }
  if (!expanded.length) { usage(); return; }

  globalThis.JSZip = loadJSZip();
  loadXmldom();

  const outputs = [];
  const failures = [];
  for (let k = 0; k < expanded.length; k++) {
    const src = expanded[k];
    const ext = (src.match(/\.(doc|docx|docm|dotx|dotm)$/i) || [, 'doc'])[1].toLowerCase();
    const label = '[' + (k + 1) + '/' + expanded.length + '] ' + path.basename(src);
    try {
      if (ext === 'doc') {
        process.stdout.write(label + '：保真管道 doc→docx…\n');
        const mid = await docToDocx(src);
        if (!mid) throw new Error('doc→docx 转换失败：需要本机已安装 WPS/Word（Windows）或 LibreOffice（macOS/Linux）');
        const r1 = await coreMod().processDocumentPackage(new Uint8Array(fs.readFileSync(mid)), steps);
        process.stdout.write('  三连完成（书签 ' + r1.report.summary['书签'] + '）→ 回写 doc…\n');
        const padded = await padBookmarks(r1.bytes);
        const tmpDocx = path.join(tmpDir(), 'result' + k + '.docx');
        fs.writeFileSync(tmpDocx, padded.bytes);
        const viaDoc = await docxToDoc(tmpDocx);
        if (!viaDoc) throw new Error('docx→doc 回写失败：需要本机已安装 WPS/Word（Windows）或 LibreOffice（macOS/Linux）');
        const wantNames = [];
        {
          const zip = await globalThis.JSZip.loadAsync(r1.bytes);
          for (const n of Object.keys(zip.files)) {
            if (!/^word\/(document|header[^/]*|footer[^/]*)\.xml$/.test(n)) continue;
            const xml = await zip.file(n).async('string');
            const re = /<w:bookmarkStart [^>]*w:name="([^"]+)"/g;
            let mm;
            while ((mm = re.exec(xml)) !== null) wantNames.push(mm[1]);
          }
        }
        const verify = await verifyDocBookmarks(viaDoc, wantNames);
        const baseName = path.basename(src).replace(/\.doc$/i, '');
        if (verify.ok === false && !forceDoc) {
          // 当前引擎的 doc 回写会丢书签（典型：LibreOffice 的 docx→doc 对表格内书签的限制）
          const dstDocx = nonConflict(path.join(outDir || path.dirname(src), baseName + '_已转书签.docx'));
          fs.writeFileSync(dstDocx, r1.bytes);
          outputs.push({ src, out: dstDocx, report: r1.report, mode: 'doc 保真回退（书签丢失→输出 docx）' });
          process.stderr.write('⚠ ' + path.basename(src) + '：当前引擎回写 .doc 会丢失表格内书签（' + verify.kept + '/' + verify.total +
            ' 保留），已自动改输出 .docx（样式保真、书签完整）。如需强制输出 .doc 请加 --force-doc，或改用 Windows 版（WPS/Word 引擎）。\n');
        } else {
          const dst = nonConflict(path.join(outDir || path.dirname(src), baseName + '_已转书签.doc'));
          fs.copyFileSync(viaDoc, dst);
          outputs.push({ src, out: dst, report: r1.report, mode: 'doc 保真' + (verify.ok === false ? '（书签 ' + verify.kept + '/' + verify.total + '，--force-doc 强制）' : '') });
        }
      } else {
        const r2 = await coreMod().processDocumentPackage(new Uint8Array(fs.readFileSync(src)), steps);
        const name = path.basename(src).replace(/(\.docx|\.docm|\.dotx|\.dotm)$/i, '') + '_已转书签.' + ext;
        const dst = nonConflict(path.join(outDir || path.dirname(src), name));
        fs.writeFileSync(dst, r2.bytes);
        outputs.push({ src, out: dst, report: r2.report, mode: 'OOXML 直通' });
      }
    } catch (e) {
      const msg = (e && e.message) || String(e);
      failures.push({ src, msg });
      process.stderr.write('✗ ' + path.basename(src) + '：' + msg + '\n');
    }
  }

  process.stdout.write('\n===== 处理完成 =====\n');
  for (const o of outputs) {
    const s = o.report.summary;
    process.stdout.write('✅ ' + path.basename(o.out) + '（' + o.mode + '：占位符 ' + o.report.placeholderTotal +
      '，书签 ' + s['书签'] + '，唯一化 ' + s['唯一化'] + '，删除 ' + s['删除'] + '）\n');
  }
  for (const f of failures) process.stdout.write('✗ ' + path.basename(f.src) + '：' + f.msg + '\n');
  if (_tmp) { try { fs.rmSync(_tmp, { recursive: true, force: true }); } catch (e) {} }
  if (failures.length && !outputs.length) process.exit(1);
}

main().catch((e) => {
  process.stderr.write('错误：' + ((e && e.message) || e) + '\n');
  process.exit(1);
});

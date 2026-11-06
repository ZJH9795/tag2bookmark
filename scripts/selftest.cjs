#!/usr/bin/env node
/**
 * 自测：用 samples/generated 的通用样例跑 CLI，并断言处理结果
 * 用法：node scripts/selftest.mjs
 */
'use strict';
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const ROOT = path.resolve(__dirname, '..');
const GEN = path.join(ROOT, 'samples', 'generated');
const CLI = path.join(ROOT, 'cli', 'tag2bookmark.cjs');

function fail(msg) { console.error('✗ ' + msg); process.exitCode = 1; }
function ok(msg) { console.log('✓ ' + msg); }

if (!fs.existsSync(path.join(GEN, '01_basic.docx'))) {
  console.log('生成样例…');
  const r = spawnSync('python3', [path.join(ROOT, 'samples', 'generate_samples.py')], { stdio: 'inherit' });
  if (r.status !== 0) { console.error('样例生成失败'); process.exit(1); }
}

const work = fs.mkdtempSync(path.join(os.tmpdir(), 't2b-selftest-'));
for (const f of fs.readdirSync(GEN)) {
  if (/\.docx$/.test(f) && !f.includes('已转书签')) fs.copyFileSync(path.join(GEN, f), path.join(work, f));
}

const res = spawnSync(process.execPath, [CLI, work], { encoding: 'utf8', timeout: 180000 });
console.log(res.stdout, res.stderr || '');
if (res.status !== 0) { console.error('CLI 执行失败'); process.exit(1); }

const JSZip = require(path.join(ROOT, 'web', 'dep', 'jszip.min.js'));
const { DOMParser } = require('@xmldom/xmldom');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function bookmarksOf(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const names = [];
  for (const n of Object.keys(zip.files)) {
    if (!/^word\/(document|header[^/]*|footer[^/]*)\.xml$/.test(n)) continue;
    const xml = await zip.file(n).async('string');
    const re = /<w:bookmarkStart [^>]*w:name="([^"]+)"/g;
    let m;
    while ((m = re.exec(xml)) !== null) names.push(m[1]);
  }
  return names;
}

async function main() {
  const cases = [
    { f: '01_basic_已转书签.docx', want: ['name', 'name_1', 'name_2'] },
    { f: '02_header_footer_已转书签.docx', want: ['name', 'name_1', 'name_2'] },
    { f: '03_cross_run_已转书签.docx', want: ['contractNo'] },
    { f: '04_table_已转书签.docx', want: ['supplierName', 'buyerName', 'signDate'] },
    { f: '05_protected_已转书签.docx', want: ['name'] },
  ];
  for (const c of cases) {
    const p = path.join(work, c.f);
    if (!fs.existsSync(p)) { fail('缺少输出 ' + c.f); continue; }
    const got = await bookmarksOf(p);
    const sortA = JSON.stringify([...got].sort());
    const sortB = JSON.stringify([...c.want].sort());
    if (sortA === sortB) ok(c.f + ' 书签集合正确（' + got.length + ' 个）');
    else fail(c.f + ' 书签不符：' + sortA + ' != ' + sortB);
  }
  fs.rmSync(work, { recursive: true, force: true });
  console.log('\n自测完成' + (process.exitCode ? '（有失败）' : '：全部通过'));
}
main().catch((e) => { console.error(e); process.exit(1); });

/**
 * 标签转书签 · 核心逻辑（Web 版）
 * 口径基准：WPS JSA 宏「标签转书签自动化」
 *   - 步骤1 唯一化：全文档连续编号（正文+页眉+页脚共用计数），首次不变，重复者 ${名_1} ${名_2}…
 *   - 步骤2 书签：落在占位符起始处；cleanBookmarkName 规则照抄宏
 *   - 步骤3 删除：占位符文本清空；书签为 XML 元素不受影响
 *   - 保护设置、docm 宏工程等非文本部件原样保留
 * 依赖：浏览器 DOMParser / XMLSerializer / JSZip（全局），无其他依赖
 */
'use strict';

var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
var PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
var OFC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// 与宏 Find 通配符 "\$\{[!\}]{1,}\}" 等价：${ 到第一个 } 之间至少 1 个字符（} 不可出现，{ 可出现）
// \u0000 是 tab/br 分隔符（见 collectParaText），需排除以匹配宏「不跨 tab/br」的语义
function phRegexExecAll(text) {
  var out = [];
  var re = /\$\{([^}\u0000]{1,})\}/g;
  var m;
  while ((m = re.exec(text)) !== null) {
    out.push({ text: m[0], name: m[1], start: m.index, end: m.index + m[0].length });
    if (m.index === re.lastIndex) re.lastIndex++; // 防御
  }
  return out;
}

/** 书签名清洗 —— 逐行照抄宏 cleanBookmarkName（含宏中的冗余第二行 replace） */
function cleanBookmarkName(s) {
  if (!s || String(s).trim() === '') return 'Placeholder';
  var cleaned = String(s).replace(/[^\w\u4e00-\u9fa5]/g, '_');
  cleaned = cleaned.replace(/[\[\]<>\\:;\*?"\/|]/g, '_');
  if (/^[0-9]/.test(cleaned)) cleaned = 'B_' + cleaned;
  cleaned = cleaned.replace(/_+/g, '_');
  cleaned = cleaned.replace(/^_+|_+$/g, '');
  if (cleaned === '') cleaned = 'Placeholder';
  if (cleaned.length > 40) cleaned = cleaned.substring(0, 40);
  return cleaned;
}

function isTextPart(name) {
  if (name === 'word/document.xml') return true;
  return /^word\/(header|footer)[^\/]*\.xml$/.test(name);
}

function parseXml(text, partName) {
  var doc = new DOMParser().parseFromString(text, 'application/xml');
  var errs = doc.getElementsByTagName('parsererror');
  if (errs.length) throw new Error('XML 解析失败（可能不是有效的 docx/docm）: ' + partName);
  return doc;
}

/** 序列化并恢复原 <?xml ...?> 声明头（standalone / encoding 原样保留） */
function serializeXml(doc, originalText) {
  var out = new XMLSerializer().serializeToString(doc);
  var m = /^<\?xml[\s\S]{0,300}?\?>/.exec(originalText);
  if (m) out = out.replace(/^<\?xml[\s\S]{0,300}?\?>/, m[0]);
  return out;
}

function listParas(docEl) {
  // 所有段落（含文本框/表格内嵌套段落）都是独立处理单元，按文档序；
  // 外层段落不得重复收取嵌套段落内的文本 —— 由 collectParaText 的 insideNestedParagraph 保证。
  return Array.prototype.slice.call(docEl.getElementsByTagNameNS(W_NS, 'p'));
}

/** 该节点是否处于比 parentDoc 段落更深的另一个 w:p 内（防止外层段落把这些文本再收一遍） */
function insideNestedParagraph(el, ownerPara) {
  var par = el.parentNode;
  while (par && par !== ownerPara && par.nodeType === Node.ELEMENT_NODE) {
    if ((par.localName || '') === 'p' && par.namespaceURI === W_NS) return true;
    par = par.parentNode;
  }
  return false;
}

/**
 * 部件处理顺序：贴近宏的节遍历——
 *   1) document.xml 里 headerReference（文档序）→ 2) footerReference（文档序）
 *   3) 未被引用的页眉/页脚文件按字典序殿后（等价超集，宏不处理它们）
 */
function resolvePartOrder(textPartNames, bodyDoc, relsDoc) {
  var referenced = [];
  var seen = {};
  var pathSet = {};
  textPartNames.forEach(function (n) { pathSet[n] = true; });
  function pushRef(rid) {
    if (!relsDoc || !rid) return;
    var rels = relsDoc.getElementsByTagNameNS(PKG_REL_NS, 'Relationship');
    for (var i = 0; i < rels.length; i++) {
      if (rels[i].getAttribute('Id') === rid) {
        var tgt = rels[i].getAttribute('Target') || '';
        var p = tgt.charAt(0) === '/' ? tgt.slice(1) : 'word/' + tgt;
        if (isTextPart(p) && pathSet[p] && !seen[p]) { seen[p] = true; referenced.push(p); }
        return;
      }
    }
  }
  try {
    var hrefs = bodyDoc.getElementsByTagNameNS(W_NS, 'headerReference');
    for (var m = 0; m < hrefs.length; m++) pushRef(hrefs[m].getAttributeNS(OFC_REL_NS, 'id'));
    var frefs = bodyDoc.getElementsByTagNameNS(W_NS, 'footerReference');
    for (var n = 0; n < frefs.length; n++) pushRef(frefs[n].getAttributeNS(OFC_REL_NS, 'id'));
  } catch (e) { /* rels 缺失等异常 → 走字典序兜底 */ }
  var ordered = referenced.slice();
  textPartNames.filter(function (n) { return !seen[n]; }).sort().forEach(function (p) { ordered.push(p); });
  return ordered;
}

/* ============ 段内文本拼装（跨 run 合并匹配的基础） ============ */

/**
 * 收集段落内按文档序排列的文本片段：
 *   w:t → 文本字符；w:tab / w:br → 分隔符 \u0000（打断匹配，宏 Find 同样不跨这些）；
 *   w:instrText、w:delText 等不参与（宏 Find 不匹配域代码/修订删除文本）。
 * 返回 { S: 拼接文本, map: [ {el, idx} | null, ... 与 S 等长 ] }
 */
function collectParaText(p) {
  var items = [];
  var ts = p.getElementsByTagNameNS(W_NS, 't');
  for (var i = 0; i < ts.length; i++) {
    if (ts[i] !== p && insideNestedParagraph(ts[i], p)) continue;
    items.push({ el: ts[i], kind: 't' });
  }
  var tabs = p.getElementsByTagNameNS(W_NS, 'tab');
  for (var j = 0; j < tabs.length; j++) {
    if (tabs[j] !== p && insideNestedParagraph(tabs[j], p)) continue;
    items.push({ el: tabs[j], kind: 'brk' });
  }
  var brs = p.getElementsByTagNameNS(W_NS, 'br');
  for (var k = 0; k < brs.length; k++) {
    if (brs[k] !== p && insideNestedParagraph(brs[k], p)) continue;
    items.push({ el: brs[k], kind: 'brk' });
  }
  items.sort(function (a, b) {
    return (a.el.compareDocumentPosition(b.el) & Node.DOCUMENT_POSITION_FOLLOWING) ? -1 : 1;
  });
  var S = '';
  var map = [];
  for (var n = 0; n < items.length; n++) {
    var it = items[n];
    if (it.kind === 't') {
      var v = it.el.textContent;
      for (var c = 0; c < v.length; c++) { S += v.charAt(c); map.push({ el: it.el, idx: c }); }
    } else {
      S += '\u0000';
      map.push(null);
    }
  }
  return { S: S, map: map };
}

/** 每个 w:t 的文本统一重写（合并 text 子节点），首尾含空白时补 xml:space="preserve" */
function setTextContent(el, str) {
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(el.ownerDocument.createTextNode(str));
  if (str.length && (/\s/.test(str.charAt(0)) || /\s/.test(str.charAt(str.length - 1)))) {
    el.setAttribute('xml:space', 'preserve');
  }
}

/**
 * 将一组编辑落到段落 DOM：
 *   edits = [{start, end, replacement}]，区间互不重叠（调用方保证）
 * replacement 只写进命中起点所在的 w:t，其余涉及的 w:t 对应字符清空
 */
function applyEdits(map, edits) {
  var perEl = new Map(); // el -> [{s, e, repl}]（基于该 el 的原 textContent 索引）
  for (var e = 0; e < edits.length; e++) {
    var ed = edits[e];
    var firstEl = null;
    var involved = new Map(); // el -> {min, max}
    for (var i = ed.start; i < ed.end; i++) {
      var m = map[i];
      if (!m) continue;
      var o = involved.get(m.el);
      if (!o) { o = { min: m.idx, max: m.idx }; involved.set(m.el, o); }
      if (m.idx < o.min) o.min = m.idx;
      if (m.idx > o.max) o.max = m.idx;
      if (!firstEl) firstEl = m.el;
    }
    if (!firstEl) continue;
    var keys = Array.from(involved.keys());
    for (var q = 0; q < keys.length; q++) {
      var el = keys[q];
      var rg = involved.get(el);
      var list = perEl.get(el);
      if (!list) { list = []; perEl.set(el, list); }
      list.push({ s: rg.min, e: rg.max + 1, repl: (el === firstEl ? (ed.replacement == null ? '' : ed.replacement) : '') });
    }
  }
  perEl.forEach(function (list, el) {
    list.sort(function (a, b) { return a.s - b.s; });
    var v = el.textContent; // 原 textContent（本轮尚未写回）
    var out = '';
    var pos = 0;
    for (var x = 0; x < list.length; x++) {
      var seg = list[x];
      out += v.slice(pos, seg.s) + seg.repl;
      pos = seg.e;
    }
    out += v.slice(pos);
    setTextContent(el, out);
  });
}

/** 在段落内某个字符位置前插入零宽书签（bookmarkStart+bookmarkEnd 相邻） */
function insertBookmarkAt(p, charIndex, map, name, state) {
  var m = map[charIndex];
  if (!m) return null;
  var anchor = m.el;
  // 取包含该 w:t 的「段落直接子元素」作为锚（w:r / w:hyperlink / w:ins / w:sdt …）
  var cur = anchor;
  while (cur.parentNode && cur.parentNode !== p && cur.parentNode.nodeType === Node.ELEMENT_NODE) {
    cur = cur.parentNode;
  }
  if (cur === p) cur = anchor.parentNode || anchor;
  var parent = cur.parentNode;
  if (!parent) return null;
  var id = String(state.nextId++);
  var bStart = p.ownerDocument.createElementNS(W_NS, 'w:bookmarkStart');
  bStart.setAttribute('w:id', id);
  bStart.setAttribute('w:name', name);
  var bEnd = p.ownerDocument.createElementNS(W_NS, 'w:bookmarkEnd');
  bEnd.setAttribute('w:id', id);
  parent.insertBefore(bStart, cur);
  parent.insertBefore(bEnd, cur);
  return name;
}

/* ============ 包级处理主入口 ============ */

/**
 * 处理一个 docx/docm（fileBytes: Uint8Array）
 * options: { unique, bookmarks, deleteTags } 均默认 true
 * 返回 { bytes: Uint8Array(处理后), report: {...} }
 */
async function processDocumentPackage(fileBytes, options) {
  var opts = { unique: true, bookmarks: true, deleteTags: true };
  if (options) for (var key in options) if (options[key] !== undefined) opts[key] = options[key];

  var zip = await JSZip.loadAsync(fileBytes);
  var names = Object.keys(zip.files);
  if (!zip.file('word/document.xml')) {
    throw new Error('不是有效的 Word OOXML 文档（缺少 word/document.xml）');
  }

  var textPartNames = names.filter(isTextPart);
  // 顺序：document → headerReference 文档序 → footerReference 文档序 → 未引用殿后（贴近宏节遍历）
  var docXmlText = await zip.file('word/document.xml').async('string');
  var bodyDoc = parseXml(docXmlText, 'word/document.xml');
  var relsDoc = null;
  var relsFile = zip.file('word/_rels/document.xml.rels');
  if (relsFile) {
    try { relsDoc = parseXml(await relsFile.async('string'), 'word/_rels/document.xml.rels'); } catch (e) { relsDoc = null; }
  }
  textPartNames = resolvePartOrder(textPartNames, bodyDoc, relsDoc);

  var parts = [];
  for (var pi = 0; pi < textPartNames.length; pi++) {
    var path = textPartNames[pi];
    var txt = await zip.file(path).async('string');
    var doc = parseXml(txt, path);
    var label;
    if (path === 'word/document.xml') label = '正文';
    else if (path.indexOf('word/header') === 0) label = '页眉-' + path.replace(/^word\//, '').replace(/\.xml$/, '');
    else label = '页脚-' + path.replace(/^word\//, '').replace(/\.xml$/, '');
    parts.push({ path: path, label: label, doc: doc, originalText: txt });
  }

  // 既有书签：名字集合（含页眉页脚内）+ 最大 id
  var usedNames = [];
  var maxId = 0;
  for (var a = 0; a < parts.length; a++) {
    var pd = parts[a];
    var bs = pd.doc.getElementsByTagNameNS(W_NS, 'bookmarkStart');
    for (var b1 = 0; b1 < bs.length; b1++) {
      var nm = bs[b1].getAttribute('w:name');
      if (nm && usedNames.indexOf(nm) < 0) usedNames.push(nm);
      var idv = parseInt(bs[b1].getAttribute('w:id'), 10);
      if (isFinite(idv) && idv > maxId) maxId = idv;
    }
    var be = pd.doc.getElementsByTagNameNS(W_NS, 'bookmarkEnd');
    for (var b2 = 0; b2 < be.length; b2++) {
      var idv2 = parseInt(be[b2].getAttribute('w:id'), 10);
      if (isFinite(idv2) && idv2 > maxId) maxId = idv2;
    }
  }

  var state = { counters: {}, usedNames: usedNames.slice(), nextId: maxId + 1, logs: [] };
  var logs = state.logs;

  // 每个部件的段落列表（文档序，含表格、文本框内的 w:p）
  var paraGroups = parts.map(function (pd) {
    return { pd: pd, paras: listParas(pd.doc) };
  });

  /* ---- Pass 1: 唯一化（全文档连续计数；首次出现保持原文不改写） ---- */
  if (opts.unique) {
    for (var g1 = 0; g1 < paraGroups.length; g1++) {
      var grp1 = paraGroups[g1];
      for (var p1 = 0; p1 < grp1.paras.length; p1++) {
        var par1 = grp1.paras[p1];
        var c1 = collectParaText(par1);
        var hits1 = phRegexExecAll(c1.S);
        var edits1 = [];
        for (var h1 = 0; h1 < hits1.length; h1++) {
          var hit1 = hits1[h1];
          var base = String(hit1.name).replace(/^\s+|\s+$/g, '');
          var cnt = Object.prototype.hasOwnProperty.call(state.counters, base) ? state.counters[base] : 0;
          if (cnt > 0) {
            var nn = '${' + base + '_' + cnt + '}';
            edits1.push({ start: hit1.start, end: hit1.end, replacement: nn });
            logs.push({ part: grp1.pd.label, action: '唯一化', detail: hit1.text + ' → ' + nn });
          }
          state.counters[base] = cnt + 1;
        }
        if (edits1.length) applyEdits(c1.map, edits1);
      }
    }
  }

  /* ---- Pass 2: 转书签（占位符起始处；重名自动 _1/_2 递增） ---- */
  if (opts.bookmarks) {
    for (var g2 = 0; g2 < paraGroups.length; g2++) {
      var grp2 = paraGroups[g2];
      for (var p2 = 0; p2 < grp2.paras.length; p2++) {
        var par2 = grp2.paras[p2];
        var c2 = collectParaText(par2);
        var hits2 = phRegexExecAll(c2.S);
        for (var h2 = hits2.length - 1; h2 >= 0; h2--) { // 倒序插入，锚点引用不受影响
          var hit2 = hits2[h2];
          var clean = cleanBookmarkName(hit2.name);
          var finalName = clean;
          if (state.usedNames.indexOf(finalName) >= 0) {
            var k = 1;
            while (state.usedNames.indexOf(clean + '_' + k) >= 0) k++;
            finalName = clean + '_' + k;
          }
          state.usedNames.push(finalName);
          insertBookmarkAt(par2, hit2.start, c2.map, finalName, state);
          logs.push({ part: grp2.pd.label, action: '书签', detail: finalName + ' ← ' + hit2.text });
        }
      }
    }
  }

  /* ---- Pass 3: 删除标签文本（书签是元素，不受文本删除影响） ---- */
  if (opts.deleteTags) {
    for (var g3 = 0; g3 < paraGroups.length; g3++) {
      var grp3 = paraGroups[g3];
      for (var p3 = 0; p3 < grp3.paras.length; p3++) {
        var par3 = grp3.paras[p3];
        var c3 = collectParaText(par3);
        var hits3 = phRegexExecAll(c3.S);
        if (!hits3.length) continue;
        var edits3 = [];
        for (var h3 = 0; h3 < hits3.length; h3++) {
          edits3.push({ start: hits3[h3].start, end: hits3[h3].end, replacement: '' });
          logs.push({ part: grp3.pd.label, action: '删除', detail: hits3[h3].text });
        }
        applyEdits(c3.map, edits3);
      }
    }
  }

  /* ---- 序列化回 zip（其余部件原样保留） ---- */
  for (var si = 0; si < parts.length; si++) {
    var spd = parts[si];
    var outXml = serializeXml(spd.doc, spd.originalText);
    zip.file(spd.path, outXml);
  }
  var blob = await zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });

  var perPart = {};
  for (var li = 0; li < logs.length; li++) {
    var L = logs[li];
    perPart[L.part] = (perPart[L.part] || 0) + 1;
  }
  var bookmarksMade = 0, uniqueMade = 0, deletedMade = 0;
  for (var lj = 0; lj < logs.length; lj++) {
    if (logs[lj].action === '书签') bookmarksMade++;
    else if (logs[lj].action === '唯一化') uniqueMade++;
    else if (logs[lj].action === '删除') deletedMade++;
  }
  var report = {
    placeholderTotal: Object.keys(state.counters).reduce(function (acc, k) { return acc + state.counters[k]; }, 0),
    logs: logs,
    summary: { 书签: bookmarksMade, 唯一化: uniqueMade, 删除: deletedMade }
  };
  return { bytes: blob, report: report };
}

/* ============ 文件格式探测与顶层入口（含 .doc / 加密检测） ============ */

function sniffFormat(bytes) {
  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04) return 'ooxml';
  if (bytes.length >= 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0
      && bytes[4] === 0xA1 && bytes[5] === 0xB1 && bytes[6] === 0x1A && bytes[7] === 0xE1) return 'ole2';
  return 'unknown';
}

/** 处理单个上传文件（bytes/名字），返回 {ok, outBytes?, report?, errorKind?, message?} */
async function processSingleFile(bytes, origName, options) {
  var fmt = sniffFormat(bytes);
  if (fmt === 'ole2') {
    // .doc（Word 97-2003）与加密 OOXML 共用 OLE2 头：能解析则转为 docx 输出，并给出另存指引兜底
    if (typeof processDocFile === 'function') {
      try {
        var opts2 = { unique: true, bookmarks: true, deleteTags: true };
        if (options) for (var key in options) if (options[key] !== undefined) opts2[key] = options[key];
        var dr = await processDocFile(bytes, opts2);
        return { ok: true, outBytes: dr.bytes, report: dr.report };
      } catch (e) {
        if (e && e.formatIntercept) {
          return { ok: false, errorKind: 'legacy-or-encrypted', message: (e && e.message) ? e.message : String(e) };
        }
        return { ok: false, errorKind: 'doc-parse-error', message: (e && e.message) ? e.message : String(e) };
      }
    }
    return {
      ok: false,
      errorKind: 'legacy-or-encrypted',
      message: '该文件是 .doc（97-2003）格式或已被「打开密码」加密，浏览器无法直接处理。请先用 Word/WPS 打开并「另存为」.docx 或 .docm 后重试。'
    };
  }
  if (fmt !== 'ooxml') {
    return { ok: false, errorKind: 'unknown-format', message: '无法识别的文件格式，请上传 .docx / .docm / .dotx / .dotm / .doc' };
  }
  try {
    var r = await processDocumentPackage(bytes, options);
    return { ok: true, outBytes: r.bytes, report: r.report };
  } catch (e) {
    return { ok: false, errorKind: 'process-error', message: (e && e.message) ? e.message : String(e) };
  }
}

/** 输出文件命名：docx/docm 同扩展名；.doc 一律转为 .docx 输出 */
function outputFileName(name) {
  var m = /^(.*?)(\.(docx|docm|dotx|dotm))$/i.exec(name);
  if (m) return m[1] + '_已转书签' + m[2];
  var d = /^(.*?)(\.doc)$/i.exec(name);
  if (d) return d[1] + '_已转书签.docx';
  return name + '_已转书签.docx';
}

// 供自动化测试 / 页面复用的统一入口
async function convertFiles(entries, options) {
  var results = [];
  for (var i = 0; i < entries.length; i++) {
    var en = entries[i];
    var bytes;
    if (en.b64 != null) {
      var bin = atob(en.b64);
      bytes = new Uint8Array(bin.length);
      for (var j = 0; j < bin.length; j++) bytes[j] = bin.charCodeAt(j);
    } else {
      bytes = en.bytes;
    }
    var r = await processSingleFile(bytes, en.name, options);
    results.push({
      name: en.name,
      ok: r.ok,
      errorKind: r.errorKind,
      message: r.message,
      report: r.report,
      outName: r.ok ? outputFileName(en.name) : null,
      outB64: r.ok ? uint8ToB64(r.outBytes) : null
    });
  }
  return results;
}

function uint8ToB64(u8) {
  var CHUNK = 0x8000;
  var parts = [];
  for (var i = 0; i < u8.length; i += CHUNK) {
    parts.push(String.fromCharCode.apply(null, u8.subarray(i, i + CHUNK)));
  }
  return btoa(parts.join(''));
}

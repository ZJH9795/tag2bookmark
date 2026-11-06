/**
 * 无浏览器转换内核（node/bun 通用，零 DOM 依赖、require 版）
 * 由 core.js 的浏览器版移植：三连口径完全一致（唯一化→书签@占位符起始处→删除标签；全文档连续编号）
 * 依赖：globalThis.JSZip（由宿主注入）、@xmldom/xmldom
 */
'use strict';

var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
var PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
var OFC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
var DP_FOLLOWING = 4; // Node.DOCUMENT_POSITION_FOLLOWING

var _XMLDOM = null;
function xdom() {
  if (_XMLDOM) return _XMLDOM;
  var m = require('@xmldom/xmldom');
  _XMLDOM = m.DOMParser ? m : (m.default || m);
  return _XMLDOM;
}

function phRegexExecAll(text) {
  var out = [];
  var re = /\$\{([^}\u0000]{1,})\}/g;
  var m;
  while ((m = re.exec(text)) !== null) {
    out.push({ text: m[0], name: m[1], start: m.index, end: m.index + m[0].length });
    if (m.index === re.lastIndex) re.lastIndex++;
  }
  return out;
}

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
  var D = xdom();
  var errs = [];
  var parser = new D.DOMParser({
    onError: function (msg) { errs.push(typeof msg === 'string' ? msg : (msg && msg.message) || ''); }
  });
  var doc = parser.parseFromString(text, 'application/xml');
  if (errs.length) throw new Error('XML 解析失败（' + partName + '）: ' + errs[0]);
  if (!doc || !doc.documentElement) throw new Error('XML 解析失败（' + partName + '）: 结构为空');
  return doc;
}

function serializeXml(doc, originalText) {
  var D = xdom();
  var out = new D.XMLSerializer().serializeToString(doc);
  var m = /^<\?xml[\s\S]{0,300}?\?>/.exec(originalText);
  if (m) {
    if (/^<\?xml/.test(out)) out = out.replace(/^<\?xml[\s\S]{0,300}?\?>/, m[0]);
    else out = m[0] + '\r\n' + out;
  }
  return out;
}

/** 该节点是否处于比 ownerPara 更深的另一个 w:p 内（防止外层段落重复收取文本框内文本） */
function insideNestedParagraph(el, ownerPara) {
  var par = el.parentNode;
  while (par && par !== ownerPara && par.nodeType === 1) {
    var ln = par.localName != null ? par.localName : String(par.tagName || '').replace(/^.*:/, '');
    if (ln === 'p' && par.namespaceURI === W_NS) return true;
    par = par.parentNode;
  }
  return false;
}

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
  } catch (e) {}
  var ordered = referenced.slice();
  textPartNames.filter(function (n) { return !seen[n]; }).sort().forEach(function (p) { ordered.push(p); });
  return ordered;
}

function pathToRoot(el) {
  var chain = [];
  var x = el;
  while (x) { chain.push(x); x = x.parentNode; }
  return chain;
}

/** 文档序比较：返回 -1 表示 a 在 b 前，1 表示在后（替代 xmldom 缺失的 compareDocumentPosition） */
function docOrder(a, b) {
  if (a === b) return 0;
  var pa = pathToRoot(a);
  var pb = pathToRoot(b);
  var i = pa.length - 1, j = pb.length - 1;
  var common = null;
  while (i >= 0 && j >= 0 && pa[i] === pb[j]) { common = pa[i]; i--; j--; }
  if (i < 0) return -1; // a 是 b 的祖先：a 在前
  if (j < 0) return 1;
  var childA = pa[i], childB = pb[j];
  if (!common || !common.childNodes) return 1;
  var kids = common.childNodes;
  for (var k = 0; k < kids.length; k++) {
    if (kids[k] === childA) return -1;
    if (kids[k] === childB) return 1;
  }
  return 1;
}

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
    return docOrder(a.el, b.el);
  });
  var S = '';
  var map = [];
  for (var n = 0; n < items.length; n++) {
    var it = items[n];
    if (it.kind === 't') {
      var v = it.el.textContent || '';
      for (var c = 0; c < v.length; c++) { S += v.charAt(c); map.push({ el: it.el, idx: c }); }
    } else {
      S += '\u0000';
      map.push(null);
    }
  }
  return { S: S, map: map };
}

function setTextContent(el, str) {
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(el.ownerDocument.createTextNode(str));
  if (str.length && (/\s/.test(str.charAt(0)) || /\s/.test(str.charAt(str.length - 1)))) {
    el.setAttribute('xml:space', 'preserve');
  }
}

function applyEdits(map, edits) {
  var perEl = new Map();
  for (var e = 0; e < edits.length; e++) {
    var ed = edits[e];
    var firstEl = null;
    var involved = new Map();
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
    var v = el.textContent || '';
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

function insertBookmarkAt(p, charIndex, map, name, state) {
  var m = map[charIndex];
  if (!m) return null;
  var anchor = m.el;
  var cur = anchor;
  while (cur.parentNode && cur.parentNode !== p && cur.parentNode.nodeType === 1) {
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

/** 包级处理主入口（与浏览器版同口径） */
async function processDocumentPackage(fileBytes, options) {
  var opts = { unique: true, bookmarks: true, deleteTags: true };
  if (options) for (var key in options) if (options[key] !== undefined) opts[key] = options[key];
  var zip = await JSZip.loadAsync(fileBytes);
  var names = Object.keys(zip.files);
  if (!zip.file('word/document.xml')) {
    throw new Error('不是有效的 Word OOXML 文档（缺少 word/document.xml）');
  }
  var textPartNames = names.filter(isTextPart);
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
    var pth = textPartNames[pi];
    var txt = await zip.file(pth).async('string');
    var doc = parseXml(txt, pth);
    var label;
    if (pth === 'word/document.xml') label = '正文';
    else if (pth.indexOf('word/header') === 0) label = '页眉-' + pth.replace(/^word\//, '').replace(/\.xml$/, '');
    else label = '页脚-' + pth.replace(/^word\//, '').replace(/\.xml$/, '');
    parts.push({ path: pth, label: label, doc: doc, originalText: txt });
  }

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

  var paraGroups = parts.map(function (pd) {
    var ps = pd.doc.getElementsByTagNameNS(W_NS, 'p');
    var arr = [];
    for (var i = 0; i < ps.length; i++) arr.push(ps[i]);
    return { pd: pd, paras: arr };
  });

  /* Pass 1: 唯一化（全文档连续计数） */
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

  /* Pass 2: 转书签（起始处；重名 _1 递增） */
  if (opts.bookmarks) {
    for (var g2 = 0; g2 < paraGroups.length; g2++) {
      var grp2 = paraGroups[g2];
      for (var p2 = 0; p2 < grp2.paras.length; p2++) {
        var par2 = grp2.paras[p2];
        var c2 = collectParaText(par2);
        var hits2 = phRegexExecAll(c2.S);
        for (var h2 = hits2.length - 1; h2 >= 0; h2--) {
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

  /* Pass 3: 删除标签文本 */
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

function outputFileName(name) {
  var m = /^(.*?)(\.(docx|docm|dotx|dotm))$/i.exec(name);
  if (m) return m[1] + '_已转书签' + m[2];
  return name + '_已转书签.docx';
}

/** 单文件入口（保真管道只喂它 OOXML） */
async function processSingleFile(bytes, origName, options) {
  if (!(bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04)) {
    throw new Error('非 OOXML 输入（保真管道第一步应已完成 doc→docx）');
  }
  var r = await processDocumentPackage(bytes, options);
  return { bytes: r.bytes, report: r.report, outName: outputFileName(origName) };
}

module.exports = { processDocumentPackage: processDocumentPackage, processSingleFile: processSingleFile, cleanBookmarkName: cleanBookmarkName };

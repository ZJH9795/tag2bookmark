/**
 * .doc（Word 97-2003 二进制）解析与 docx 重建
 * 结构依据：MS-DOC（Word Binary File Format）；校准样本：法务部范本（WPS 存）、
 * Word 原生存档、LibreOffice 转换档（build/probe_doc.py 可复查结构）
 *
 * v1 覆盖：
 *   - CFB 复合文档、WordDocument/1Table|0Table 流
 *   - FIB（fEncrypted 拦截）、CLX piece table（8bit/16bit piece）
 *   - 正文分段（0x0D/0x07/0x0C）、\t→w:tab、0x0B→w:br、域指令剔除、控制符清理
 *   - PAPX FKPs：fInTable/fTtp（表格 cell/行末判定）、对齐 jc
 *   - 页眉/页脚 story：plcfHdd CP 或 Fc(相对 CP) 有效性自动降级；引用段剔除；PAGE 域归页脚
 * v1 限制（输出明告）：
 *   - 字体/字号/颜色/编号样式不保留；first/even 页眉页脚不同内容时并入默认页眉页脚
 *   - 单元格宽度按列数均分；合并单元格还原为普通单元格
 *   - 图片/OLE 对象不保留
 * 依赖：JSZip（全局）；输出为"干净 docx"，三连（唯一化/书签/删除）由 core.js 的 docx 通道完成
 */
'use strict';

/* ---------- 基础读 ---------- */
function u16(b, o) { return b[o] | (b[o + 1] << 8); }
function u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

var CF_END = 0xFFFFFFFE;

/** 复合文档（CFB/OLE2）最小读取器 */
function CompoundFile(data) {
  if (data.length < 512) throw new Error('文件太小，不是有效的 .doc');
  this.data = data;
  this.ssz = 1 << u16(data, 30);
  this.mssz = 1 << u16(data, 32);
  this.cutoff = u32(data, 56);

  var self = this;
  function sec(s) {
    var off = 512 + s * self.ssz;
    if (s < 0 || off < 512 || off + self.ssz > data.length) return new Uint8Array(self.ssz);
    return data.subarray(off, off + self.ssz);
  }
  this.sec = sec;

  // DIFAT → FAT 扇区表
  var fatSectors = [];
  for (var i = 0; i < 109; i++) {
    var v = u32(data, 76 + i * 4);
    if (v !== CF_END && v !== 0xFFFFFFFF) fatSectors.push(v);
  }
  var difat = u32(data, 68);
  var guard = 0;
  while (difat !== CF_END && difat !== 0xFFFFFFFF && guard++ < 4096) {
    var blk = sec(difat);
    var n = self.ssz / 4;
    for (var k = 0; k < n - 1; k++) {
      var vv = u32(blk, k * 4);
      if (vv !== CF_END && vv !== 0xFFFFFFFF) fatSectors.push(vv);
    }
    difat = u32(blk, (n - 1) * 4);
  }
  this.fat = [];
  for (var f = 0; f < fatSectors.length && f < 4000; f++) {
    var fatSec = sec(fatSectors[f]);
    for (var w = 0; w < self.ssz / 4; w++) this.fat.push(u32(fatSec, w * 4));
  }

  // miniFAT（注意：数组必须为空初始化，逐项 push，保证 miniFat[i] 与 mini sector i 对齐）
  this.miniFat = [];
  var miniFirst = u32(data, 60);
  var miniCount = u32(data, 64);
  if (miniFirst !== CF_END && miniFirst !== 0xFFFFFFFF && miniCount > 0) {
    var mblk = this.chainBytes(miniFirst, miniCount * this.ssz);
    for (var m = 0; m + 4 <= mblk.length; m += 4) this.miniFat.push(u32(mblk, m));
  }

  // 目录（线性扫描全部扇区，取 type==2 的流）
  this.entries = [];
  var dirFirst = u32(data, 48);
  var s = dirFirst;
  guard = 0;
  while (s !== CF_END && s !== 0xFFFFFFFF && guard++ < 4096 && this.entries.length < 4096) {
    var page = sec(s);
    for (var o = 0; o + 128 <= this.ssz; o += 128) {
      var e = page.subarray(o, o + 128);
      var nl = u16(e, 64);
      var nm = '';
      if (nl >= 2 && nl <= 64) {
        var cs = [];
        for (var c = 0; c + 1 < nl - 1; c += 2) cs.push(String.fromCharCode(e[c] | (e[c + 1] << 8)));
        nm = cs.join('');
      }
      this.entries.push({ name: nm, type: e[66], start: u32(e, 116), size: u32(e, 120) });
    }
    s = this.fat[s] || CF_END;
  }
  var root = this.entries[0];
  // mini stream 读取做尺寸钳制（加密包等异常 root.size 防爆）
  var rootSize = (root && root.size > 0) ? Math.min(root.size, data.length * 2) : 0;
  this.miniStream = rootSize ? this.chainBytes(root.start, rootSize) : new Uint8Array(0);
}

CompoundFile.prototype.chainBytes = function (first, size) {
  var out = new Uint8Array(size);
  var pos = 0;
  var s = first;
  var guard = 0;
  while (s !== CF_END && s !== 0xFFFFFFFF && guard++ < 100000 && pos < size) {
    var page = this.sec(s);
    var take = Math.min(page.length, size - pos);
    if (take <= 0) break;
    out.set(page.subarray(0, take), pos);
    pos += take;
    s = this.fat[s] || CF_END;
  }
  return out.subarray(0, pos);
};

CompoundFile.prototype.miniChainBytes = function (first, size) {
  var out = new Uint8Array(size);
  var pos = 0;
  var s = first;
  var guard = 0;
  while (s !== CF_END && s !== 0xFFFFFFFF && this.miniStream.length && guard++ < 100000 && pos < size) {
    var from = s * this.mssz;
    if (from + this.mssz > this.miniStream.length) break;
    var take = Math.min(this.mssz, size - pos);
    out.set(this.miniStream.subarray(from, from + take), pos);
    pos += take;
    s = this.miniFat[s] || CF_END;
  }
  return out.subarray(0, pos);
};

CompoundFile.prototype.stream = function (name) {
  for (var i = 0; i < this.entries.length; i++) {
    var e = this.entries[i];
    if (e.type === 2 && e.name === name) {
      if (e.size > 0 && e.size < this.cutoff && i > 0) return this.miniChainBytes(e.start, e.size);
      return this.chainBytes(e.start, e.size);
    }
  }
  return null;
};

/* ---------- Word 解析 ---------- */

/** CLX → piece table；返回 {T, cps, pieces, cp2fc} */
function buildTextFromClx(wd, tbl, fcClx, lcbClx) {
  if (!lcbClx || !tbl) throw new Error('doc 缺少 piece table（CLX），无法解析文本');
  var pos = fcClx;
  var end = fcClx + lcbClx;
  var plc = null;
  while (pos < end && pos < tbl.length) {
    var t = tbl[pos];
    if (t === 1) { var cb = tbl[pos + 1]; pos += 2 + cb; continue; }
    if (t === 2) {
      var plen = u32(tbl, pos + 1);
      plc = tbl.subarray(pos + 5, Math.min(pos + 5 + plen, tbl.length));
      break;
    }
    throw new Error('doc 的 CLX 结构异常（记录类型 0x' + t.toString(16) + '）');
  }
  if (!plc) throw new Error('doc 的 CLX 中没有 piece table');
  var rawN = (plc.length - 4) / 12;
  if (rawN < 1 || Math.abs(rawN - Math.round(rawN)) > 1e-9) throw new Error('doc 的 piece table 长度异常');
  var n = Math.round(rawN);
  var cps = [];
  for (var i = 0; i <= n; i++) cps.push(u32(plc, i * 4));
  var pieces = [];
  var epos = (n + 1) * 4;
  for (var k = 0; k < n; k++) {
    var pcd = plc.subarray(epos + k * 8, epos + k * 8 + 8);
    var rawFc = u32(pcd, 2);
    var compressed = !!(rawFc & 0x40000000);
    pieces.push({ cpStart: cps[k], cpEnd: cps[k + 1], compressed: compressed, fc: rawFc & 0x3FFFFFFF });
  }
  var chars = [];
  for (var q = 0; q < pieces.length; q++) {
    var p = pieces[q];
    var ln = p.cpEnd - p.cpStart;
    var bytesLen = p.compressed ? ln : ln * 2;
    var raw = wd.subarray(p.fc, p.fc + bytesLen);
    if (raw.length < bytesLen) throw new Error('doc 文本流越界（piece ' + q + '）');
    if (p.compressed) {
      for (var a = 0; a < ln; a++) chars.push(String.fromCharCode(raw[a]));
    } else {
      for (var b = 0; b < ln; b++) chars.push(String.fromCharCode(raw[b * 2] | (raw[b * 2 + 1] << 8)));
    }
  }
  var T = chars.join('');
  function cp2fc(cp) {
    for (var m = 0; m < pieces.length; m++) {
      var p2 = pieces[m];
      if (cp >= p2.cpStart && cp < p2.cpEnd) {
        return p2.fc + (cp - p2.cpStart) * (p2.compressed ? 1 : 2);
      }
    }
    return -1;
  }
  return { T: T, cps: cps, pieces: pieces, cp2fc: cp2fc };
}

/** ===== 段落属性（PAPX 特征扫描） =====
 * MS-DOC sprm 实测/规范核验值：
 *   fInTable 0x2416 / fTtp 0x2417 / Jc 0x2403
 *   左缩进 0x840F、右缩进 0x840E、首行缩进 0x8411（2B twips）
 *   行距 0x6412（2B）、段前 0xA413、段后 0xA414（2B twips）
 *   表格行定义 0xD609（TDefTable 变长）
 */
function defaultParaFlags() {
  return { fInTable: false, fTtp: false, jc: null, dxaLeft: null, dxaRight: null, dxaLeft1: null, dyaLine: null, dyaBefore: null, dyaAfter: null, tdCols: null };
}
function saneTwips(v) { return (v > -4000 && v < 22000) ? v : null; }
function saneDya(v) { return (v >= 0 && v <= 4000) ? v : null; }

function makePapxLookup(tbl, wd, fcBtePapx, lcbBtePapx) {
  if (!lcbBtePapx || lcbBtePapx < 8 || !tbl) return function () { return null; };
  var rawV = (lcbBtePapx - 4) / 8;
  if (Math.abs(rawV - Math.round(rawV)) > 1e-9) return function () { return null; };
  var n = Math.round(rawV);
  if (n < 1) return function () { return null; };
  var pages = [];
  for (var q = 0; q < n; q++) {
    var pn = u32(tbl, fcBtePapx + (n + 1) * 4 + q * 4) & 0x3FFFFF;
    var pageOff = pn * 512;
    if (pageOff + 512 <= wd.length) pages.push(wd.subarray(pageOff, pageOff + 512));
  }
  function findCrun(pg) {
    var c = pg[511];
    if (c < 1 || c > 0x1D) return -1;
    if ((c + 1) * 4 + c * 13 > 511) return -1;
    var prev = -1;
    for (var r = 0; r <= c; r++) {
      var v = u32(pg, r * 4);
      if (r > 0 && v <= prev) return -1;
      prev = v;
    }
    return c;
  }
  function papxAt(pg, byteOff) {
    var F = defaultParaFlags();
    if (byteOff < 0 || byteOff + 1 >= pg.length) return F;
    var cb = pg[byteOff];
    var grpprlLen, grpprlStartOff;
    if (cb === 0) {
      var cb2 = pg[byteOff + 1];
      if (cb2 < 1) return F;
      grpprlLen = 2 * cb2 - 2;
      grpprlStartOff = byteOff + 3;
    } else {
      grpprlLen = 2 * cb - 2;
      grpprlStartOff = byteOff + 3;
    }
    if (grpprlLen <= 0 || grpprlStartOff + grpprlLen > pg.length) return F;
    var endP = grpprlStartOff + grpprlLen;
    function i16(o) { var v = u16(pg, o); return v >= 0x8000 ? v - 0x10000 : v; }
    for (var i = grpprlStartOff; i + 2 < endP; i++) {
      var b0 = pg[i], b1 = pg[i + 1], b2 = pg[i + 2];
      if (b0 === 0x16 && b1 === 0x24) { if (b2 === 0x01) F.fInTable = true; }
      else if (b0 === 0x17 && b1 === 0x24) { if (b2 === 0x01) F.fTtp = true; }
      else if ((b0 === 0x03 || b0 === 0x61) && b1 === 0x24) { F.jc = b2; }
      else if (b0 === 0x0F && b1 === 0x84 && i + 3 < endP) { var v1 = saneTwips(i16(i + 2)); if (v1 !== null) F.dxaLeft = v1; }
      else if (b0 === 0x0E && b1 === 0x84 && i + 3 < endP) { var v2 = saneTwips(i16(i + 2)); if (v2 !== null) F.dxaRight = v2; }
      else if (b0 === 0x11 && b1 === 0x84 && i + 3 < endP) { var v3 = saneTwips(i16(i + 2)); if (v3 !== null) F.dxaLeft1 = v3; }
      else if (b0 === 0x12 && b1 === 0x64 && i + 3 < endP) { var v4 = saneDya(u16(pg, i + 2)); if (v4 !== null) F.dyaLine = v4; }
      else if (b0 === 0x13 && b1 === 0xA4 && i + 3 < endP) { var v5 = saneDya(u16(pg, i + 2)); if (v5 !== null) F.dyaBefore = v5; }
      else if (b0 === 0x14 && b1 === 0xA4 && i + 3 < endP) { var v6 = saneDya(u16(pg, i + 2)); if (v6 !== null) F.dyaAfter = v6; }
      else if (b0 === 0x08 && b1 === 0xD6) {
        // sprmTDefTable 0xD608：变长操作符规范特例「无长度字节前缀」，
        // 操作数直接 = TDefTableOperand: cb(2B, "剩余字节数+1") + nCols(1B) + rgdxaCenter[(nCols+1)×XAS(2B)]
        if (i + 5 < endP) {
          var tdCb = u16(pg, i + 2);
          var nCols = pg[i + 4];
          if (nCols >= 1 && nCols <= 63 && tdCb >= (nCols + 1) * 2 + 1 && i + 2 + tdCb <= endP + 2) {
            var okT = true, prevXT = -1e9, colsT = [];
            for (var cT = 0; cT <= nCols; cT++) {
              if (i + 5 + cT * 2 + 1 >= pg.length) { okT = false; break; }
              var xvT = u16(pg, i + 5 + cT * 2);
              if (xvT >= 0x8000) xvT -= 0x10000; // XAS 为 16 位有符号 twips
              if (xvT <= prevXT) { okT = false; break; }
              prevXT = xvT; colsT.push(xvT);
            }
            if (okT) F.tdCols = colsT;
          }
        }
      }
    }
    return F;
  }
  return function (fcByte) {
    if (fcByte < 0) return null;
    for (var i = 0; i < pages.length; i++) {
      var pg = pages[i];
      var crun = findCrun(pg);
      if (crun < 0) continue;
      var lastEnd = u32(pg, crun * 4);
      if (fcByte >= lastEnd) continue;
      var best = -1;
      for (var r = 0; r < crun; r++) {
        var v = u32(pg, r * 4);
        if (v <= fcByte) best = r; else break;
      }
      if (best < 0) continue;
      var bOffset = pg[(crun + 1) * 4 + best * 13];
      if (!bOffset) return defaultParaFlags();
      return papxAt(pg, bOffset * 2);
    }
    return null;
  };
}

/** ===== 字符属性（CHPX FKP 特征扫描） =====
 *  Bold 0x0835 / Italic 0x0836 / Strike 0x0837 / Sz 0x4A43(2B half-pt)
 *  Ico 0x2A42 / CCv 0x6870(4B COLORREF) / 字体 ftc0 0x4A4F / ftc1 0x4A50 / ftc2 0x4A51
 */
function makeChpxLookup(tbl, wd, fcBteChpx, lcbBteChpx) {
  if (!lcbBteChpx || lcbBteChpx < 8 || !tbl) return function () { return null; };
  var rawV = (lcbBteChpx - 4) / 8;
  if (Math.abs(rawV - Math.round(rawV)) > 1e-9) return function () { return null; };
  var n = Math.round(rawV);
  if (n < 1) return function () { return null; };
  var pages = [];
  for (var q = 0; q < n; q++) {
    var pn = u32(tbl, fcBteChpx + (n + 1) * 4 + q * 4) & 0x3FFFFF;
    var pageOff = pn * 512;
    if (pageOff + 512 <= wd.length) pages.push(wd.subarray(pageOff, pageOff + 512));
  }
  function findCrun(pg) {
    // CHPX FKP：[rgfc (crun+1)个u32][rgb crun×1B(offset/2)]...[crun@511]
    var c = pg[511];
    if (c < 1 || c > 0x65) return -1;
    if ((c + 1) * 4 + c > 511) return -1;
    var prev = -1;
    for (var r = 0; r <= c; r++) {
      var v = u32(pg, r * 4);
      if (r > 0 && v <= prev) return -1;
      prev = v;
    }
    return c;
  }
  function emptyProps() { return { b: false, i: false, strike: false, sz: null, ico: null, rgb: null, ftc0: null, ftc1: null, ftc2: null }; }
  function chpxAt(pg, byteOff) {
    var P = emptyProps();
    if (byteOff < 0 || byteOff + 1 >= pg.length) return P;
    var cb = pg[byteOff];
    var start, len;
    if (cb === 0) { len = pg[byteOff + 1]; start = byteOff + 2; } else { len = cb; start = byteOff + 1; }
    if (len <= 0 || start + len > pg.length) return P;
    var endP = start + len;
    for (var i = start; i + 2 < endP; i++) {
      var b0 = pg[i], b1 = pg[i + 1], b2 = pg[i + 2];
      if (b0 === 0x35 && b1 === 0x08) P.b = b2 === 0x01;
      else if (b0 === 0x36 && b1 === 0x08) P.i = b2 === 0x01;
      else if (b0 === 0x37 && b1 === 0x08) P.strike = b2 === 0x01;
      else if (b0 === 0x43 && b1 === 0x4A && i + 3 < endP) { var szv = u16(pg, i + 2); if (szv > 0 && szv <= 360) P.sz = szv; }
      else if (b0 === 0x42 && b1 === 0x2A) P.ico = b2;
      else if (b0 === 0x70 && b1 === 0x68 && i + 5 < endP) {
        var cr = u32(pg, i + 2); // 0x00BBGGRR
        P.rgb = { r: cr & 0xFF, g: (cr >> 8) & 0xFF, b: (cr >> 16) & 0xFF };
      }
      else if (b0 === 0x4F && b1 === 0x4A && i + 3 < endP) { var f0 = u16(pg, i + 2); if (f0 < 4096) P.ftc0 = f0; }
      else if (b0 === 0x50 && b1 === 0x4A && i + 3 < endP) { var f1 = u16(pg, i + 2); if (f1 < 4096) P.ftc1 = f1; }
      else if (b0 === 0x51 && b1 === 0x4A && i + 3 < endP) { var f2 = u16(pg, i + 2); if (f2 < 4096) P.ftc2 = f2; }
    }
    return P;
  }
  var cache = { fc0: -1, fc1: -1, props: null, pageIx: 0 };
  return function (fcByte) {
    if (fcByte < 0) return null;
    if (fcByte >= cache.fc0 && fcByte < cache.fc1 && cache.props) return cache.props;
    var total = pages.length;
    for (var s = 0; s < total; s++) {
      var pi = (cache.pageIx + s) % total;
      var pg = pages[pi];
      var crun = findCrun(pg);
      if (crun < 0) continue;
      var lastEnd = u32(pg, crun * 4);
      if (fcByte >= lastEnd) continue;
      var best = -1;
      for (var r = 0; r < crun; r++) {
        var v = u32(pg, r * 4);
        if (v <= fcByte) best = r; else break;
      }
      if (best < 0) continue;
      var bOff = pg[(crun + 1) * 4 + best];
      cache.fc0 = u32(pg, best * 4);
      cache.fc1 = u32(pg, (best + 1) * 4);
      cache.pageIx = pi;
      cache.props = bOff ? chpxAt(pg, bOff * 2) : emptyProps();
      return cache.props;
    }
    return null;
  };
}

/** 字体表 SttbfFfn → {索引: 字体名} */
function parseFfnChain(tbl, p0, cData, end) {
  var fonts = {};
  var p = p0;
  for (var i = 0; i < cData && p < end; i++) {
    var cbFfnM1 = tbl[p];
    var total = cbFfnM1 + 1;
    if (total < 20 || p + total > end) break;
    // FFN: cbFfnM1(1) prq(1) wWeight(2) chs(1) ixchSzAlt(1) panose(10) fs(24) → xszFfn(UTF-16)
    var nameOff = p + 40;
    var nm = [];
    for (var c = nameOff; c + 1 < p + total; c += 2) {
      var ch = tbl[c] | (tbl[c + 1] << 8);
      if (ch === 0 || ch < 0x20 || ch > 0xFFEF) break;
      nm.push(String.fromCharCode(ch));
    }
    if (nm.length >= 1) fonts[i] = nm.join('');
    p += total;
  }
  return fonts;
}
function parseFontTable(tbl, fc, lcb) {
  try {
    if (!lcb || fc + 6 > tbl.length) return {};
    var end = Math.min(fc + lcb, tbl.length);
    var fExtend = u16(tbl, fc);
    if (fExtend === 0xFFFF) {
      var cData1 = u16(tbl, fc + 2);
      var f1 = parseFfnChain(tbl, fc + 6, cData1, end);
      if (Object.keys(f1).length) return f1;
    }
    // WPS/部分转换器：4 字节头（cData 2B + cbExtra 2B）
    var cData2 = u16(tbl, fc);
    if (cData2 >= 1 && cData2 <= 600) {
      var f2 = parseFfnChain(tbl, fc + 4, cData2, end);
      if (Object.keys(f2).length) return f2;
    }
    // 最后尝试标准 6B 头（0xFFFF 分支）
    var cData3 = u16(tbl, fc + 2);
    var f3 = parseFfnChain(tbl, fc + 6, cData3, end);
    return f3;
  } catch (e) { return {}; }
}

/** 节属性 SEPX：页面尺寸/边距（取第一节）
 *  XaPage 0xB01F / YaPage 0xB020 / DxaLeft 0xB021 / DxaRight 0xB022（4B）
 *  DyaTop 0x9023 / DyaBottom 0x9024（2B）/ DyaHdr 0xB017 / DyaFtr 0xB018（4B）
 */
function parseSepx(tbl, wd, fcPlcfsed, lcbPlcfsed) {
  var def = { w: 11906, h: 16838, left: 1800, right: 1800, top: 1440, bottom: 1440, hdr: 851, ftr: 992 };
  try {
    if (!lcbPlcfsed || lcbPlcfsed < 16) return def;
    var n = Math.round((lcbPlcfsed - 4) / 16);
    if (n < 1) return def;
    var sedOff = fcPlcfsed + (n + 1) * 4;
    var fcSepx = u32(tbl, sedOff + 2); // SED: fn(2) fcSepx(4) ...
    if (!fcSepx || fcSepx + 4 > wd.length) return def;
    var cb = u16(wd, fcSepx);
    var start = fcSepx + 2, end = start + cb;
    if (cb < 2 || end > wd.length) return def;
    var out = { w: null, h: null, left: null, right: null, top: null, bottom: null, hdr: null, ftr: null };
    function sane(v, lo, hi) { return (v >= lo && v <= hi) ? v : null; }
    for (var i = start; i + 2 < end; i++) {
      var b0 = wd[i], b1 = wd[i + 1];
      // 实测（WPS/Word97 存档）：这些节属性 sprm 的 operand 均为 2 字节（twips）
      if (b0 === 0x1F && b1 === 0xB0 && i + 3 < end) { var x = u16(wd, i + 2); if (sane(x, 3000, 40000)) out.w = x; }
      else if (b0 === 0x20 && b1 === 0xB0 && i + 3 < end) { var y = u16(wd, i + 2); if (sane(y, 3000, 40000)) out.h = y; }
      else if (b0 === 0x21 && b1 === 0xB0 && i + 3 < end) { var l = u16(wd, i + 2); if (sane(l, 0, 6000)) out.left = l; }
      else if (b0 === 0x22 && b1 === 0xB0 && i + 3 < end) { var r = u16(wd, i + 2); if (sane(r, 0, 6000)) out.right = r; }
      else if (b0 === 0x23 && b1 === 0x90 && i + 3 < end) { var t = u16(wd, i + 2); if (sane(t, 0, 6000)) out.top = t; }
      else if (b0 === 0x24 && b1 === 0x90 && i + 3 < end) { var bm = u16(wd, i + 2); if (sane(bm, 0, 6000)) out.bottom = bm; }
      else if (b0 === 0x17 && b1 === 0xB0 && i + 3 < end) { var hd = u16(wd, i + 2); if (sane(hd, 0, 6000)) out.hdr = hd; }
      else if (b0 === 0x18 && b1 === 0xB0 && i + 3 < end) { var ft = u16(wd, i + 2); if (sane(ft, 0, 6000)) out.ftr = ft; }
    }
    if (out.w) def.w = out.w;
    if (out.h) def.h = out.h;
    if (out.left !== null) def.left = out.left;
    if (out.right !== null) def.right = out.right;
    if (out.top !== null) def.top = out.top;
    if (out.bottom !== null) def.bottom = out.bottom;
    if (out.hdr !== null) def.hdr = out.hdr;
    if (out.ftr !== null) def.ftr = out.ftr;
  } catch (e) { }
  return def;
}

/** 文本区间 → 段落结构（段落属性 + 字符 run 属性；表格行/cell 标记；域指令剔除） */
function extractStructure(T, startCp, endCp, papxAt, cp2fc, chpxAt) {
  var paras = [];
  var cur = { text: '', runs: [], segStart: startCp, fieldDepth: 0, fieldState: 0, lastRunKey: null, runProps: null };
  function flagsFor(segStartCp, markCp) {
    var info = defaultParaFlags();
    var fc1 = cp2fc(segStartCp);
    if (fc1 >= 0) { var g = papxAt(fc1); if (g) info = g; }
    if (markCp !== segStartCp) {
      var fc2 = cp2fc(markCp);
      if (fc2 >= 0 && fc2 !== fc1) {
        var g2 = papxAt(fc2);
        if (g2) {
          if (!info.fInTable && g2.fInTable) info.fInTable = true;
          if (!info.fTtp && g2.fTtp) info.fTtp = true;
          if (info.tdCols == null && g2.tdCols) info.tdCols = g2.tdCols;
          if (info.jc == null && g2.jc != null) info.jc = g2.jc;
          if (info.dxaLeft == null && g2.dxaLeft != null) info.dxaLeft = g2.dxaLeft;
          if (info.dxaRight == null && g2.dxaRight != null) info.dxaRight = g2.dxaRight;
          if (info.dxaLeft1 == null && g2.dxaLeft1 != null) info.dxaLeft1 = g2.dxaLeft1;
          if (info.dyaLine == null && g2.dyaLine != null) info.dyaLine = g2.dyaLine;
          if (info.dyaBefore == null && g2.dyaBefore != null) info.dyaBefore = g2.dyaBefore;
          if (info.dyaAfter == null && g2.dyaAfter != null) info.dyaAfter = g2.dyaAfter;
        }
      }
    }
    return info;
  }
  function pushRunChar(ch, props) {
    var key = props ? [props.b ? 1 : 0, props.i ? 1 : 0, props.strike ? 1 : 0,
      props.sz || '', props.ico == null ? '' : props.ico,
      props.rgb ? (props.rgb.r + ',' + props.rgb.g + ',' + props.rgb.b) : '',
      props.ftc0 == null ? '' : props.ftc0].join('|') : 'plain';
    if (key !== cur.lastRunKey) {
      cur.runs.push({ text: '', props: props || null });
      cur.lastRunKey = key;
    }
    cur.runs[cur.runs.length - 1].text += ch;
  }
  function flush(markCp, info) {
    paras.push({
      text: cur.text, runs: cur.runs, mark: 'cell', isCell: !!info.fInTable, isRowEnd: !!info.fTtp,
      jc: info.jc, dxaLeft: info.dxaLeft, dxaRight: info.dxaRight, dxaLeft1: info.dxaLeft1,
      dyaLine: info.dyaLine, dyaBefore: info.dyaBefore, dyaAfter: info.dyaAfter, tdCols: info.tdCols
    });
    cur.text = ''; cur.runs = []; cur.lastRunKey = null; cur.segStart = markCp + 1;
  }
  for (var i = startCp; i < endCp && i < T.length; i++) {
    var code = T.charCodeAt(i);
    if (code === 0x0D || code === 0x07 || code === 0x0C) {
      var segStart = cur.segStart;
      var info = flagsFor(segStart, i);
      flush(i, info);
      continue;
    }
    if (code === 0x09) { pushRunChar('\t', cur.runProps); cur.text += '\t'; continue; }
    if (code === 0x0B) { pushRunChar('\n', cur.runProps); cur.text += '\n'; continue; }
    if (code < 0x20) {
      if (code === 0x13) { cur.fieldDepth++; cur.fieldState = 0; continue; }
      if (code === 0x14) { cur.fieldState = 1; continue; }
      if (code === 0x15) { if (cur.fieldDepth) cur.fieldDepth--; cur.fieldState = 0; continue; }
      if (cur.fieldDepth) {
        if (cur.fieldState === 0) continue;
        if (code === 0x02 || code === 0x05 || code === 0x06 || code === 0x08) continue;
        pushRunChar(T.charAt(i), cur.runProps); cur.text += T.charAt(i);
        continue;
      }
      continue;
    }
    if (chpxAt) {
      var fcx = cp2fc(i);
      if (fcx >= 0) {
        var p = chpxAt(fcx);
        if (p) cur.runProps = p;
      }
    }
    pushRunChar(T.charAt(i), cur.runProps);
    cur.text += T.charAt(i);
  }
  if (cur.text !== '') {
    paras.push({
      text: cur.text, runs: cur.runs, mark: 'p', isCell: false, isRowEnd: false,
      jc: null, dxaLeft: null, dxaRight: null, dxaLeft1: null, dyaLine: null, dyaBefore: null, dyaAfter: null, tdCols: null
    });
  }
  return paras;
}


/* ---------- 文本 → XML ---------- */
function xmlEsc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}
var DOC_FONT_MAP = {};
var ICO_MAP = ['auto', '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];
function xmlAttr(s) { return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function fontNameOf(props) {
  var ix = props.ftc0 != null ? props.ftc0 : (props.ftc1 != null ? props.ftc1 : null);
  if (ix == null) return null;
  return DOC_FONT_MAP[ix] || null;
}
function colorToHex(props) {
  function hx(n) { var s = (n || 0).toString(16).toUpperCase(); return s.length < 2 ? '0' + s : s; }
  if (props.rgb) return hx(props.rgb.r) + hx(props.rgb.g) + hx(props.rgb.b);
  if (props.ico != null && props.ico !== 0 && props.ico < ICO_MAP.length) return ICO_MAP[props.ico];
  return null;
}
function runXml(run) {
  var props = run.props || {};
  var rPr = '';
  if (props.b) rPr += '<w:b/><w:bCs/>';
  if (props.i) rPr += '<w:i/><w:iCs/>';
  if (props.strike) rPr += '<w:strike/>';
  if (props.sz) rPr += '<w:sz w:val="' + props.sz + '"/><w:szCs w:val="' + props.sz + '"/>';
  var colorHex = colorToHex(props);
  if (colorHex && colorHex !== 'auto') rPr += '<w:color w:val="' + colorHex + '"/>';
  var fontName = fontNameOf(props);
  if (fontName) rPr += '<w:rFonts w:ascii="' + xmlAttr(fontName) + '" w:hAnsi="' + xmlAttr(fontName) + '" w:eastAsia="' + xmlAttr(fontName) + '"/>';
  var parts = String(run.text || '').split(/([\t\n])/);
  var body = '';
  for (var i = 0; i < parts.length; i++) {
    if (parts[i] === '\t') body += '<w:tab/>';
    else if (parts[i] === '\n') body += '<w:br/>';
    else if (parts[i] !== '') body += '<w:t xml:space="preserve">' + xmlEsc(parts[i]) + '</w:t>';
  }
  if (!body) return '';
  return '<w:r>' + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') + body + '</w:r>';
}
function paraXml(p) {
  var pr = '';
  if (p.jc != null) {
    var map = { 0: 'left', 1: 'center', 2: 'right', 3: 'both', 4: 'distribute' };
    pr += '<w:jc w:val="' + (map[p.jc] || 'left') + '"/>';
  }
  var ind = '';
  if (p.dxaLeft != null) ind += ' w:left="' + p.dxaLeft + '"';
  if (p.dxaRight != null) ind += ' w:right="' + p.dxaRight + '"';
  if (p.dxaLeft1 != null) ind += ' w:firstLine="' + p.dxaLeft1 + '"';
  if (ind) pr += '<w:ind' + ind + '/>';
  var sp = '';
  if (p.dyaBefore != null) sp += ' w:before="' + p.dyaBefore + '"';
  if (p.dyaAfter != null) sp += ' w:after="' + p.dyaAfter + '"';
  if (p.dyaLine != null) sp += ' w:line="' + p.dyaLine + '" w:lineRule="auto"';
  if (sp) pr += '<w:spacing' + sp + '/>';
  var prXml = pr ? '<w:pPr>' + pr + '</w:pPr>' : '';
  var body = '';
  if (p.runs && p.runs.length) {
    for (var r = 0; r < p.runs.length; r++) body += runXml(p.runs[r]);
  } else {
    body = runXml({ text: p.text });
  }
  return '<w:p>' + prXml + body + '</w:p>';
}
function tcXml(p, widthTwips) {
  var w = widthTwips ? ('<w:tcW w:w="' + widthTwips + '" w:type="dxa"/>') : '<w:tcW w:type="pct" w:w="5000"/>';
  return '<w:tc><w:tcPr>' + w + '</w:tcPr>' + paraXml(p) + '</w:tc>';
}

/* ---------- doc → 干净 docx → 三连 ---------- */

/**
 * 处理 .doc：解析 → 生成干净 docx → 复用 core.processDocumentPackage 完成三连
 * 返回 { bytes, report }（report 附加 docWarnings）
 */
async function processDocFile(fileBytes, options) {
  if (typeof JSZip === 'undefined') throw new Error('内部错误：JSZip 未加载');
  var cfb;
  try { cfb = new CompoundFile(fileBytes); }
  catch (e) {
    var cfErr = new Error('该文件是 OLE2 复合格式但不是有效的 .doc：' + (e.message || e));
    cfErr.formatIntercept = true;
    throw cfErr;
  }
  var names = (cfb.entries || []).map(function (e) { return e.name; });
  if (names.indexOf('EncryptedPackage') >= 0) {
    var encErr = new Error('该文件已被「打开密码」加密：请用 Word/WPS 输入密码打开后另存为 .docx 再处理');
    encErr.formatIntercept = true;
    throw encErr;
  }
  if (names.indexOf('WordDocument') < 0) {
    throw new Error('不是 Word 97-2003 文档（无 WordDocument 流）。若是 Excel/PPT 请用对应工具处理');
  }
  var wd = cfb.stream('WordDocument');
  if (!wd || wd.length < 64) {
    throw new Error('.doc 的 WordDocument 流读取失败（文件损坏）');
  }
  var tbl = cfb.stream(u16(wd, 10) & 0x0200 ? '1Table' : '0Table')
    || cfb.stream('1Table') || cfb.stream('0Table');
  if (!tbl) throw new Error('.doc 的表格流（0Table/1Table）缺失或损坏');
  var wIdent = u16(wd, 0);
  if (wIdent !== 0xA5EC && wIdent !== 0xA5DB && wIdent !== 0xA6DB) {
    throw new Error('不是 Word 97-2003 文档（wIdent=0x' + wIdent.toString(16) + '）');
  }
  var flags = u16(wd, 10);
  if (flags & 0x0100) {
    throw new Error('该 .doc 已加密（打开密码保护）：请用 Word/WPS 输入密码打开后另存为 .docx 再处理');
  }
  var csw = u16(wd, 32);
  var cslwOff = 34 + csw * 2;
  var cslw = u16(wd, cslwOff);
  if (cslw < 11) throw new Error('doc FIB 结构异常（cslw=' + cslw + '）');
  var lwOff = cslwOff + 2;
  var ccpText = u32(wd, lwOff + 12);
  var ccpFtn = u32(wd, lwOff + 16);
  var ccpHdd = u32(wd, lwOff + 20);
  if (ccpText <= 0) throw new Error('doc 无正文文本（ccpText=0）');

  var fclcbOff = lwOff + cslw * 4;
  function pair(ix) {
    var o = fclcbOff + 2 + ix * 8;
    return { fc: u32(wd, o), lcb: u32(wd, o + 4) };
  }
  var clxPair = pair(33);
  var btePapxPair = pair(13);
  var bteChpxPair = pair(12);
  var sttbFfnPair = pair(15);
  var sedPair = pair(6);
  var hddPair = pair(11);

  var txt = buildTextFromClx(wd, tbl, clxPair.fc, clxPair.lcb);
  var T = txt.T;
  if (T.length === 0) throw new Error('doc 解析出空文本（piece table 覆盖为 0）');
  var warnings = [];

  var papxAt = makePapxLookup(tbl, wd, btePapxPair.fc, btePapxPair.lcb);
  var chpxAt = makeChpxLookup(tbl, wd, bteChpxPair.fc, bteChpxPair.lcb);
  DOC_FONT_MAP = parseFontTable(tbl, sttbFfnPair.fc, sttbFfnPair.lcb);
  var pageSetup = parseSepx(tbl, wd, sedPair.fc, sedPair.lcb);

  // ----- 正文 -----
  var bodyParas = extractStructure(T, 0, Math.min(ccpText, T.length), papxAt, txt.cp2fc, chpxAt);
  // 图片/对象检测提示（输出不保留图片）
  var bodyText = T.substring(0, Math.min(ccpText, T.length));
  if (/[\u0001\u0008]/.test(bodyText)) {
    warnings.push('检测到图片或嵌入对象：doc 模式不保留图片内容，请在输出文档中人工补图');
  }

  // ----- 页眉/页脚 -----
  var headers = [];
  var footers = [];
  if (ccpHdd > 0 && T.length >= ccpText + ccpFtn + ccpHdd) {
    var zoneStart = ccpText + ccpFtn;
    var zoneLen = ccpHdd;
    var nStories = (hddPair.lcb > 4) ? Math.floor((hddPair.lcb - 4) / 8) : 0;
    var bounds = null;
    if (nStories > 0) {
      var hcp = [];
      for (var hx = 0; hx <= nStories; hx++) hcp.push(u32(tbl, hddPair.fc + hx * 4));
      var cpOk = hcp[0] === 0 && hcp[nStories] === zoneLen;
      for (var hy = 1; hy <= nStories; hy++) if (hcp[hy] < hcp[hy - 1]) cpOk = false;
      if (cpOk) {
        bounds = hcp;
      } else {
        var fcs = [];
        for (var fz = 0; fz < nStories; fz++) fcs.push(u32(tbl, hddPair.fc + (nStories + 1) * 4 + fz * 4) & 0x3FFFFFFF);
        var fcOk = fcs[0] === 0;
        for (var fy = 1; fy < nStories; fy++) if (fcs[fy] < fcs[fy - 1]) fcOk = false;
        if (fcOk && fcs[nStories - 1] <= zoneLen) {
          bounds = fcs.slice();
          bounds.push(zoneLen);
        }
      }
    }
    var stories = [];
    if (bounds) {
      for (var sb = 0; sb < bounds.length - 1; sb++) {
        if (bounds[sb + 1] > bounds[sb]) stories.push({ start: zoneStart + bounds[sb], end: zoneStart + bounds[sb + 1] });
      }
      if (!stories.length) stories.push({ start: zoneStart, end: zoneStart + zoneLen });
    } else {
      stories.push({ start: zoneStart, end: zoneStart + zoneLen });
      warnings.push('该文档的页眉/页脚位置信息不完整（常见于 WPS/转换器另存），已将页眉区内容整体并入默认页眉');
    }
    for (var s2 = 0; s2 < stories.length; s2++) {
      var st = stories[s2];
      var paras2 = extractStructure(T, st.start, st.end, papxAt, txt.cp2fc, chpxAt);
      var storyIsFooter = false;
      var storyHasContent = false;
      for (var q2 = 0; q2 < paras2.length; q2++) {
        if (/PAGE\s*\b|PAGE\\|\\\\PAGE/.test(paras2[q2].text || '')) storyIsFooter = true;
        if ((paras2[q2].text || '').replace(/[\t\n ]/g, '') !== '') storyHasContent = true;
      }
      if (!storyHasContent) continue;
      for (var w2 = 0; w2 < paras2.length; w2++) {
        var p2 = paras2[w2];
        if (!p2.text || p2.text.replace(/[\t\n ]/g, '') === '') continue;
        (storyIsFooter ? footers : headers).push(p2);
      }
    }
  }

  // ----- 组装干净 docx -----
  function bodyXml(items) {
    var out = [];
    var i = 0;
    while (i < items.length) {
      var it = items[i];
      if (it.isCell) {
        var row = [];
        while (i < items.length && items[i].isCell && !items[i].isRowEnd) { row.push(items[i]); i++; }
        var colWidths = null;
        if (i < items.length && items[i].isRowEnd) {
          colWidths = items[i].tdCols || null;
          i++;
        }
        if (!colWidths) {
          for (var rr = 0; rr < row.length; rr++) if (row[rr].tdCols) { colWidths = row[rr].tdCols; break; }
        }
        var gridXml = '', totalW = 0, okCols = colWidths && colWidths.length === row.length + 1;
        if (okCols) {
          for (var g = 0; g < row.length; g++) {
            var cw = colWidths[g + 1] - colWidths[g];
            if (cw <= 0) cw = 2000;
            gridXml += '<w:gridCol w:w="' + cw + '"/>';
            totalW += cw;
          }
        }
        var cellsXml = '';
        for (var c = 0; c < row.length; c++) {
          var cw2 = okCols ? (colWidths[c + 1] - colWidths[c]) : null;
          cellsXml += tcXml(row[c], cw2);
        }
        var tblPr = '<w:tblPr><w:tblW w:type="' + (totalW ? 'dxa' : 'pct') + '" w:w="' + (totalW || 5000) + '"/>' +
          '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>';
        out.push('<w:tbl>' + tblPr + (gridXml ? '<w:tblGrid>' + gridXml + '</w:tblGrid>' : '') + '<w:tr>' + cellsXml + '</w:tr></w:tbl>');
      } else if (it.isRowEnd) {
        i++; // 孤立 row-end 段无内容，跳过
      } else {
        out.push(paraXml(it));
        i++;
      }
    }
    return out.join('');
  }
  var hdXml = headers.map(paraXml).join('');
  var ftXml = footers.map(paraXml).join('');
  var docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
    bodyXml(bodyParas) +
    '<w:sectPr><w:pgSz w:w="' + pageSetup.w + '" w:h="' + pageSetup.h + '"/>' +
    '<w:pgMar w:top="' + pageSetup.top + '" w:right="' + pageSetup.right + '" w:bottom="' + pageSetup.bottom + '" w:left="' + pageSetup.left + '" w:header="' + pageSetup.hdr + '" w:footer="' + pageSetup.ftr + '" w:gutter="0"/><w:cols w:space="425"/>' +
    (hdXml ? '<w:headerReference w:type="default" r:id="rIdHdr1"/>' : '') +
    (ftXml ? '<w:footerReference w:type="default" r:id="rIdFtr1"/>' : '') +
    '</w:sectPr></w:body></w:document>';
  var hdrXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + hdXml + '</w:hdr>';
  var ftrXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + ftXml + '</w:ftr>';
  var ctXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    (hdXml ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : '') +
    (ftXml ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : '') +
    '</Types>';
  var relsR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    (hdXml ? '<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' : '') +
    (ftXml ? '<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' : '') +
    '</Relationships>';
  var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>';

  var zip = new JSZip();
  zip.file('[Content_Types].xml', ctXml);
  zip.file('_rels/.rels', relsR);
  zip.file('word/document.xml', docXml);
  zip.file('word/_rels/document.xml.rels', docRels);
  zip.file('word/styles.xml', stylesXml);
  if (hdXml) zip.file('word/header1.xml', hdrXml);
  if (ftXml) zip.file('word/footer1.xml', ftrXml);
  var plainBytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });

  var r = await processDocumentPackage(plainBytes, options);
  if (r && r.report && warnings.length) r.report.docWarnings = warnings;
  return r;
}

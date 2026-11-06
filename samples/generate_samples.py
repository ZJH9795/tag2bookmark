#!/usr/bin/env python3
"""生成通用测试样例（不含任何真实业务内容）
生成 5 个 docx：
  01_basic.docx          正文重复占位符（唯一化）
  02_header_footer.docx  正文+页眉+页脚（跨区全局编号）
  03_cross_run.docx      占位符被拆到多个 run（格式切换处）
  04_table.docx          表格内占位符
  05_protected.docx      带窗体保护（documentProtection，密码为演示哈希，不涉及真实密码）
"""
import os
import zipfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/content-types'
PREL = 'http://schemas.openxmlformats.org/package/2006/relationships'
BAT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
HDR = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
FTR = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
SET = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'


def head():
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def p(text):
    return '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text


def document(body, sectpr='', extra_ns=''):
    return (head() + '<w:document xmlns:w="%s" xmlns:r="%s"%s><w:body>%s%s</w:body></w:document>'
            % (W, R, extra_ns, body, sectpr))


def sectpr(hrefs=(), frefs=()):
    inner = ''.join('<w:headerReference w:type="%s" r:id="%s"/>' % t for t in hrefs)
    inner += ''.join('<w:footerReference w:type="%s" r:id="%s"/>' % t for t in frefs)
    return '<w:sectPr>%s<w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' % inner


def header(body):
    return head() + '<w:hdr xmlns:w="%s">%s</w:hdr>' % (W, body)


def footer(body):
    return head() + '<w:ftr xmlns:w="%s">%s</w:ftr>' % (W, body)


def rels(entries):
    return (head() + '<Relationships xmlns="%s">%s</Relationships>'
            % (PREL, ''.join('<Relationship Id="%s" Type="%s/%s" Target="%s"/>' % e for e in entries)))


def ctypes(main=MAIN, overrides=()):
    base = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="%s"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="%s"/>' % (PKG, main))
    base += ''.join('<Override PartName="%s" ContentType="%s"/>' % o for o in overrides)
    return base + '</Types>'


def base_files():
    return {
        '[Content_Types].xml': ctypes(),
        '_rels/.rels': rels([('rId1', BAT, 'officeDocument', 'word/document.xml')]),
    }


def write_pkg(outdir, name, files):
    with zipfile.ZipFile(os.path.join(outdir, name), 'w', zipfile.ZIP_DEFLATED) as z:
        for k in files:
            z.writestr(k, files[k])
    print('gen:', name)


def main():
    outdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'generated')
    os.makedirs(outdir, exist_ok=True)

    # 01 基础：重复占位符
    body = p('${name} 条款A') + p('${name} 条款B') + p('${name} 条款C')
    write_pkg(outdir, '01_basic.docx', {**base_files(), 'word/document.xml': document(body).encode()})

    # 02 页眉页脚
    body = p('${name} 主文')
    files = base_files()
    files['[Content_Types].xml'] = ctypes(overrides=[('/word/header1.xml', HDR), ('/word/footer1.xml', FTR)])
    files['word/document.xml'] = document(body, sectpr((('default', 'rIdH1'),), (('default', 'rIdF1'),))).encode()
    files['word/_rels/document.xml.rels'] = rels([
        ('rIdH1', BAT, 'header', 'header1.xml'), ('rIdF1', BAT, 'footer', 'footer1.xml')]).encode()
    files['word/header1.xml'] = header(p('页眉 ${name}')).encode()
    files['word/footer1.xml'] = footer(p('页脚 ${name}')).encode()
    write_pkg(outdir, '02_header_footer.docx', files)

    # 03 跨 run
    body = ('<w:p><w:r><w:t>${contract</w:t></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">No} 汇总</w:t></w:r></w:p>')
    write_pkg(outdir, '03_cross_run.docx', {**base_files(), 'word/document.xml': document(body).encode()})

    # 04 表格
    cell1 = '<w:tc><w:tcPr/><w:p><w:r><w:t>供方：${supplierName}</w:t></w:r></w:p></w:tc>'
    cell2 = '<w:tc><w:tcPr/><w:p><w:r><w:t>需方：${buyerName}</w:t></w:r></w:p></w:tc>'
    body = ('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            '<w:tr>' + cell1 + cell2 + '</w:tr></w:tbl>' + p('${signDate} 签署'))
    write_pkg(outdir, '04_table.docx', {**base_files(), 'word/document.xml': document(body).encode()})

    # 05 保护
    body = p('${name} 受保护内容')
    files = base_files()
    files['[Content_Types].xml'] = ctypes(overrides=[('/word/settings.xml', SET)])
    files['word/document.xml'] = document(body).encode()
    files['word/_rels/document.xml.rels'] = rels([('rIdS1', BAT, 'settings', 'settings.xml')]).encode()
    files['word/settings.xml'] = (head() + '<w:settings xmlns:w="%s">'
        '<w:documentProtection w:password="DEMO-HASH" w:edit="forms" w:enforcement="1"/>'
        '</w:settings>' % W).encode()
    write_pkg(outdir, '05_protected.docx', files)


if __name__ == '__main__':
    main()

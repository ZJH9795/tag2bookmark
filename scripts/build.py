#!/usr/bin/env python3
"""构建网页版单文件工具：web/template.html + web/dep/jszip.min.js + web/doc_reader.js + web/core.js → web/tag2bookmark.html"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(HERE, '..', 'web')

def read(p):
    with open(p, encoding='utf-8') as f:
        return f.read()

tpl = read(os.path.join(WEB, 'template.html'))
jszip = read(os.path.join(WEB, 'dep', 'jszip.min.js'))
doc_reader = read(os.path.join(WEB, 'doc_reader.js'))
core = read(os.path.join(WEB, 'core.js'))

note = '${'
for name, src in (('jszip', jszip), ('doc_reader', doc_reader), ('core', core)):
    assert '</script' not in src.lower(), name + ' contains </script'
    assert '`' not in src, name + ' contains backtick (conflicts with ' + note + '} placeholders)'

out = tpl
for marker, src in (('/*__JSZIP_SRC__*/', jszip), ('/*__DOC_READER_SRC__*/', doc_reader), ('/*__CORE_SRC__*/', core)):
    assert marker in out, 'template missing marker ' + marker
    out = out.replace(marker, src, 1)

dest = os.path.join(WEB, 'tag2bookmark.html')
with open(dest, 'w', encoding='utf-8') as f:
    f.write(out)
print('built:', dest, '(%.1f KB)' % (os.path.getsize(dest) / 1024))

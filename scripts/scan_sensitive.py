#!/usr/bin/env python3
"""敏感信息扫描：确保开源仓库不含项目信息、密码、内部路径等
用法：python3 scripts/scan_sensitive.py
退出码 0 = 干净；1 = 命中禁词
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# 禁词/禁模式：项目名、账号、密码、内部路径、真实样例名
BLOCK_PATTERNS = [
    (r'\{\{兴发\}\}', '项目代号'),
    (r'\{\{xingfa\}\}', '项目代号2'),
    (r'xf3110', '硬编码密码'),
    (r'zhoujianhua', '用户账号'),
    (r'/Users/zhoujianhua', '本机绝对路径'),
    (r'\{\{PASSWORD_[a-z0-9]+\}\}', '密码占位符'),
    (r'\{\{SECRET_[a-z0-9]+\}\}', '密钥占位符'),
    (r'real_fangwu|real_gongcheng|real_baozhuang', '真实范本名'),
    (r'标签转书签-网页版', '内部产品旧名'),
]
TEXT_EXT = {'.md', '.js', '.cjs', '.mjs', '.py', '.html', '.json', '.txt', '.sh', '.yml', '.yaml'}
SKIP_DIRS = {'node_modules', '.git', 'dist', 'generated'}

hits = []
for dirpath, dirnames, filenames in os.walk(ROOT):
    dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
    for fn in filenames:
        ext = os.path.splitext(fn)[1].lower()
        if ext not in TEXT_EXT:
            continue
        p = os.path.join(dirpath, fn)
        if os.path.abspath(p) == os.path.abspath(__file__):
            continue
        try:
            with open(p, encoding='utf-8', errors='ignore') as f:
                content = f.read()
        except Exception:
            continue
        for pat, label in BLOCK_PATTERNS:
            for m in re.finditer(pat, content):
                line = content[:m.start()].count('\n') + 1
                hits.append((os.path.relpath(p, ROOT), line, label, m.group(0)))

if hits:
    print('发现敏感命中 %d 处：' % len(hits))
    for h in hits:
        print('  %s:%d [%s] %s' % h)
    sys.exit(1)
print('扫描通过：仓库无敏感词（%d 项规则）' % len(BLOCK_PATTERNS))

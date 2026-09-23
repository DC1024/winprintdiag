# -*- coding: utf-8 -*-
"""docs/index.html 结构静态校验：
1) 标签配对（忽略注释 / style / script 内容、void 元素、自闭合）
2) 中英双语键数量一致性（data-lang-zh 与 data-lang-en）
3) 常用属性拼写陷阱扫描（如 x2= / width=0 之类的残肢属性）
运行：python tests/htmlcheck.py
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TARGET = os.path.join(ROOT, 'docs', 'index.html')

VOID = {'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
        'link', 'meta', 'param', 'source', 'track', 'wbr'}

TAG_RE = re.compile(r'<(/?)([a-zA-Z][a-zA-Z0-9]*)((?:"[^"]*"|\'[^\']*\'|[^>"\'])*?)(/?)>')


def strip_noise(html: str) -> str:
    """把注释/style/script/属性内文本替换成等长空白，保持偏移与行号不变。"""
    out = list(html)

    def blank(m):
        for i in range(m.start(), m.end()):
            if out[i] != '\n':
                out[i] = ' '
        return ''

    patterns = [
        r'<!--.*?-->',
        r'<style\b.*?</style>',
        r'<script\b.*?</script>',
    ]
    for p in patterns:
        for m in re.finditer(p, html, re.S | re.I):
            blank(m)
    return ''.join(out)


def main() -> int:
    with io.open(TARGET, 'r', encoding='utf-8') as f:
        html = f.read()

    res = []
    res.append('文件: %s' % TARGET)
    res.append('大小: %d 字符 / %d 行' % (len(html), html.count('\n') + 1))

    clean = strip_noise(html)

    def line_of(pos: int) -> int:
        return clean.count('\n', 0, pos) + 1

    stack = []
    errors = []
    net = {}

    for m in TAG_RE.finditer(clean):
        closing = m.group(1) == '/'
        name = m.group(2).lower()
        self_close = m.group(4) == '/'
        net.setdefault(name, 0)
        if name in VOID or self_close:
            continue
        if not closing:
            net[name] += 1
            stack.append((name, m.start()))
        else:
            net[name] -= 1
            if not stack:
                errors.append('多余闭合 </%s> @第 %d 行' % (name, line_of(m.start())))
            else:
                top_name, top_pos = stack[-1]
                if top_name != name:
                    errors.append('不匹配: 期望 </%s>（开于第 %d 行）却遇到 </%s> @第 %d 行'
                                  % (top_name, line_of(top_pos), name, line_of(m.start())))
                    stack.pop()
                else:
                    stack.pop()

    res.append('')
    res.append('=== 未闭合标签 ===')
    if not stack:
        res.append('  （无，全部闭合）')
    else:
        for n, p in stack:
            res.append('  未闭合 <%s> @第 %d 行' % (n, line_of(p)))

    res.append('')
    res.append('=== 错配 / 多余闭合 ===')
    if not errors:
        res.append('  （无）')
    else:
        for e in errors:
            res.append('  ' + e)

    res.append('')
    res.append('=== 净差（应为 0）===')
    bad = []
    for k in sorted(net):
        if net[k] != 0:
            res.append('  %-12s %+d' % (k, net[k]))
            bad.append(k)
    if not bad:
        res.append('  （全部为 0）')

    # 只统计正文标记里的双语键：CSS 选择器 / JS 里的同名标识不算
    body_html = clean
    if '<body' in body_html:
        body_html = body_html[body_html.index('<body'):]
    zh = len(re.findall(r'\bdata-lang-zh\b', body_html))
    en = len(re.findall(r'\bdata-lang-en\b', body_html))
    res.append('')
    res.append('=== 双语键 ===')
    res.append('  data-lang-zh = %d' % zh)
    res.append('  data-lang-en = %d' % en)
    if zh == en:
        res.append('  配对正常')
    else:
        res.append('  !! 数量不一致，差值 %+d' % (zh - en))

    res.append('')
    res.append('=== 属性残肢扫描 ===')
    traps = [
        (r'<(?:rect|circle|ellipse|polygon|polyline|image)\b[^>]*\bx2\s*=', 'rect/circle 等元素上出现 x2= （该元素不支持 x2，疑似属性写残）'),
        (r'<(?:line|text)\b[^>]*\bwidth2\s*=', 'line/text 上出现 width2= （疑似 width= 被切断）'),
        (r'\bdata-lang-(?!zh\b|en\b)\w+', '未知的 data-lang-* 键'),
        (r'href\s*=\s*"(?!#|https?:|mailto:|data:|\./|/)[^"]*"', '非标准的 href（既非锚点也非绝对地址/data URI）'),
        (r'<(?:div|span|p|section)\b[^>]*\bviewBox\s*=', 'HTML 元素上误用 viewBox（应为 SVG 元素）'),
    ]
    hits = 0
    for pat, desc in traps:
        for m in re.finditer(pat, html):
            hits += 1
            res.append('  第 %d 行: %s' % (html.count('\n', 0, m.start()) + 1, desc))
    if hits == 0:
        res.append('  （未发现可疑写法）')

    res.append('')
    res.append('=== 外链依赖（应为 0，宣传页须零外部资源）===')
    ext = re.findall(r'(?:src|href)\s*=\s*"(https?://[^"]+)"', html)
    ext = [u for u in ext if 'github.com' not in u and 'dc1024.github.io' not in u]
    if not ext:
        res.append('  （无外部资源引用）')
    else:
        for u in sorted(set(ext)):
            res.append('  ' + u)

    out = '\n'.join(res)
    with io.open(os.path.join(HERE, 'htmlcheck_result.txt'), 'w', encoding='utf-8') as f:
        f.write(out + '\n')

    ok = (not stack) and (not errors) and (not bad) and (zh == en)
    print('OK' if ok else 'FAIL')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())

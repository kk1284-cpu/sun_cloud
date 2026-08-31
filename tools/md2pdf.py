#!/usr/bin/env python3
"""Markdown → A4のPDF（Chrome headless）

定例報告と同じ体裁にそろえる。LibreOfficeが無い環境なので、
HTMLに起こして Chrome の --print-to-pdf に渡す。

使い方:
    python3 tools/md2pdf.py 入力.md 出力.pdf ["ヘッダーの文字"]

⚠ .page を height:297mm の flex 列にして、フッターを margin-top:auto の
   最終子要素にする。min-height + position:absolute にすると、わずかな溢れで
   フッターが次ページへ押し出され、ページ数が倍に割れる。
"""
import html as H
import io
import os
import re
import subprocess
import sys

CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe'

CSS = """
@page { size: A4; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  font-family: "Yu Gothic", YuGothic, "Hiragino Kaku Gothic ProN", Meiryo, sans-serif;
  color: #1a1a1a; background: #fff;
  -webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.sheet { width: 210mm; padding: 15mm 18mm 12mm; }
.hd { font-size: 8.6pt; font-weight: 700; letter-spacing: .12em; color: #1B6B3A;
      border-bottom: 1.6pt solid #1B6B3A; padding-bottom: 2.4mm; margin-bottom: 6mm; }
.hd .dia { color: #C9932B; }
h1 { font-size: 19pt; font-weight: 700; line-height: 1.34; margin: 0 0 4mm;
     letter-spacing: -.01em; }
h1 .bul { color: #1B6B3A; }
h2 { font-size: 13.4pt; font-weight: 700; margin: 7mm 0 3mm; padding-top: 2.6mm;
     border-top: .8pt solid #d5d9d3; break-after: avoid; }
h3 { font-size: 10.8pt; font-weight: 700; margin: 4.6mm 0 2mm; color: #24312a;
     break-after: avoid; }
p  { font-size: 9.2pt; line-height: 1.78; margin: 0 0 2.8mm; }
ul, ol { margin: 0 0 3mm; padding-left: 5.6mm; }
li { font-size: 9.2pt; line-height: 1.74; margin-bottom: 1mm; }
strong { font-weight: 700; }
code { font-family: Consolas, "Courier New", monospace; font-size: 8.4pt;
       background: #f2f4f0; padding: .3mm 1.1mm; border-radius: 1mm; }
pre { background: #f6f8f4; border: .6pt solid #d5d9d3; border-left: 2.4pt solid #1B6B3A;
      padding: 2.8mm 3.4mm; margin: 0 0 3.4mm; break-inside: avoid; }
pre code { background: none; padding: 0; font-size: 8.2pt; line-height: 1.6; }
table { width: 100%; border-collapse: collapse; font-size: 8.4pt; margin: 0 0 3.6mm;
        break-inside: avoid; }
th, td { border-bottom: .6pt solid #e4e7e2; padding: 1.6mm 2.1mm; text-align: left;
         vertical-align: top; line-height: 1.62; }
thead th { background: #24312a; color: #fff; font-size: 8pt; border-bottom: none; }
tbody tr:last-child td { border-bottom: none; }
hr { border: none; border-top: .8pt solid #d5d9d3; margin: 6mm 0; }
blockquote { border-left: 2.4pt solid #C9932B; background: #fdfaf3; margin: 0 0 3.4mm;
             padding: 2.6mm 3.4mm; }
blockquote p { margin: 0; font-size: 8.8pt; }
.foot { border-top: .6pt solid #d5d9d3; padding-top: 2.4mm; margin-top: 8mm;
        font-size: 7.4pt; color: #8a8f88; }
"""

INLINE = [
    (re.compile(r'`([^`]+)`'),           lambda m: '<code>%s</code>' % H.escape(m.group(1))),
    (re.compile(r'\*\*([^*]+)\*\*'),     lambda m: '<strong>%s</strong>' % m.group(1)),
    (re.compile(r'(?<!\*)\*([^*\n]+)\*'), lambda m: '<strong>%s</strong>' % m.group(1)),
    (re.compile(r'\[([^\]]+)\]\(([^)]+)\)'),
     lambda m: '<a href="%s">%s</a>' % (H.escape(m.group(2)), m.group(1))),
    (re.compile(r'(?<!["=])(https?://[^\s<>"]+)'),
     lambda m: '<a href="%s">%s</a>' % (m.group(1), m.group(1))),
]


def inline(s):
    s = H.escape(s)
    s = s.replace('&lt;br&gt;', '<br>')
    for pat, fn in INLINE:
        s = pat.sub(fn, s)
    return s


def cells(line):
    return [c.strip() for c in line.strip().strip('|').split('|')]


def render(md):
    out, i, lines = [], 0, md.split('\n')
    while i < len(lines):
        ln = lines[i]

        if ln.startswith('```'):
            buf, i = [], i + 1
            while i < len(lines) and not lines[i].startswith('```'):
                buf.append(lines[i]); i += 1
            out.append('<pre><code>%s</code></pre>' % H.escape('\n'.join(buf)))
            i += 1; continue

        if re.match(r'^\|.*\|\s*$', ln) and i + 1 < len(lines) \
                and re.match(r'^\|[\s:\-|]+\|\s*$', lines[i + 1]):
            head = cells(ln); i += 2
            body = []
            while i < len(lines) and re.match(r'^\|.*\|\s*$', lines[i]):
                body.append(cells(lines[i])); i += 1
            t = ['<table><thead><tr>'] + ['<th>%s</th>' % inline(c) for c in head] + ['</tr></thead><tbody>']
            for r in body:
                t.append('<tr>' + ''.join('<td>%s</td>' % inline(c) for c in r) + '</tr>')
            t.append('</tbody></table>')
            out.append(''.join(t)); continue

        m = re.match(r'^(#{1,4})\s+(.*)$', ln)
        if m:
            lv = min(len(m.group(1)), 3)
            # 「# 見出し」は h1、それ以外は h2/h3
            out.append('<h%d>%s</h%d>' % (lv, inline(m.group(2)), lv)); i += 1; continue

        if re.match(r'^---+\s*$', ln):
            out.append('<hr>'); i += 1; continue

        if ln.startswith('> '):
            buf = []
            while i < len(lines) and lines[i].startswith('> '):
                buf.append(lines[i][2:]); i += 1
            out.append('<blockquote><p>%s</p></blockquote>' % inline('<br>'.join(buf)))
            continue

        if re.match(r'^\s*[-*]\s+', ln) or re.match(r'^\s*\d+\.\s+', ln):
            ordered = bool(re.match(r'^\s*\d+\.\s+', ln))
            buf = []
            while i < len(lines) and (re.match(r'^\s*[-*]\s+', lines[i])
                                      or re.match(r'^\s*\d+\.\s+', lines[i])):
                buf.append(re.sub(r'^\s*(?:[-*]|\d+\.)\s+', '', lines[i])); i += 1
            tag = 'ol' if ordered else 'ul'
            out.append('<%s>%s</%s>' % (tag, ''.join('<li>%s</li>' % inline(b) for b in buf), tag))
            continue

        if ln.strip():
            buf = []
            while i < len(lines) and lines[i].strip() and not re.match(
                    r'^(#{1,4}\s|\||```|>\s|---+\s*$|\s*[-*]\s|\s*\d+\.\s)', lines[i]):
                buf.append(lines[i]); i += 1
            out.append('<p>%s</p>' % inline('<br>'.join(buf)))
            continue
        i += 1
    return '\n'.join(out)


def main():
    src, dst = sys.argv[1], sys.argv[2]
    header = sys.argv[3] if len(sys.argv) > 3 else 'クローバ経営研究所 × 有限会社高村'
    md = io.open(src, encoding='utf-8').read()

    body = render(md)
    # 最初の h1 に飾りの●を付ける（render後に入れる。inline()を通すとエスケープされる）
    body = body.replace('<h1>', '<h1><span class="bul">●</span> ', 1)
    doc = ('<meta charset="utf-8"><style>%s</style><div class="sheet">'
           '<div class="hd">%s <span class="dia">◆</span></div>%s'
           '<div class="foot">有限会社高村 SUNグループ ／ 株式会社クローバ経営研究所</div>'
           '</div>' % (CSS, H.escape(header), body))

    tmp = os.path.splitext(dst)[0] + '_tmp.html'
    io.open(tmp, 'w', encoding='utf-8').write(doc)
    r = subprocess.run([CHROME, '--headless=new', '--disable-gpu', '--no-sandbox',
                        '--no-pdf-header-footer', '--virtual-time-budget=8000',
                        '--print-to-pdf=' + dst,
                        'file:///' + os.path.abspath(tmp).replace('\\', '/')],
                       capture_output=True, text=True, errors='replace')
    os.remove(tmp)
    if not os.path.exists(dst):
        print('!! 失敗'); print((r.stderr or '')[-500:]); return 1
    try:
        import fitz
        d = fitz.open(dst)
        mx = max((w[3] for p in d for w in p.get_text('words')), default=0)
        print('  %s  %d ページ / %.1f KB / 最大y %.0fpt'
              % (os.path.basename(dst), len(d), os.path.getsize(dst) / 1024, mx))
        d.close()
    except Exception:
        print('  %s  %.1f KB' % (os.path.basename(dst), os.path.getsize(dst) / 1024))
    return 0


if __name__ == '__main__':
    sys.exit(main())

#!/usr/bin/env python3
"""トライアルメンバーリスト → メンバーシート貼付用TSV

🔴 このリポジトリに名簿の実体は置かない。
   職員の実名と個人のメールアドレスは個人情報にあたる（契約第7条）。
   入力は手元のファイルから読み、出力もリポジトリの外に書く。

入力（どちらか）
  1. 引数で渡したTSV        python3 tools/build_members.py 入力.tsv 出力.tsv
  2. 顧客の原本を手で保存したもの
     https://docs.google.com/spreadsheets/d/1tU41gWTs_jIEfZ0zTb4jKJN_zTvJEtrB34xU8E23Ub8/edit
     「部署 / 氏名 / メールアドレス」の3列をタブ区切りで保存する

入力の形（メール欄は `表示名 <アドレス>` でも素のアドレスでもよい）
    輝ららデイ<TAB>日前由佳理<TAB>日前由佳理 <xxxx@example.com>

出力
    メンバーシートの8列。そのまま貼れる。
    メンバーID / 氏名 / 部署 / 累計EXP / 報告数 / 最終報告日 / 配信停止 / 社員番号
"""
import io
import os
import re
import sys

# 原本の部署 → 評価基準の部署（4つ）
# ⚠ 判定は「その人の部署の基準の中から選ぶ」ので、ここがずれると判定もずれる。
#    部署名はトライアルメンバーリストの記載を優先する（松山判断 2026-08-31）。
DEPT_MAP = {
    '輝ららデイ':         '輝らら',
    '輝らら':             '輝らら',
    'ネットワーク事業課':   '薬局・ネットワーク',
    '薬局':               '薬局・ネットワーク',
    'DX推進':             '本部',
    '経営推進':            '本部',
    'ジュニアクラブ':       'サンフラワーズ',   # ⚠ 要確認。障がい事業所として暫定
    # 変換後の名前をそのまま渡されても通す（作り直しても壊れないように）
    '薬局・ネットワーク':   '薬局・ネットワーク',
    'サンフラワーズ':       'サンフラワーズ',
    '本部':               '本部',
}

MAIL = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}')
HEADER = ['メンバーID', '氏名', '部署', '累計EXP', '報告数', '最終報告日', '配信停止', '社員番号']


def parse(path):
    rows = []
    for line in io.open(path, encoding='utf-8'):
        line = line.rstrip('\n')
        if not line.strip():
            continue
        cols = [c.strip() for c in line.split('\t')]
        if len(cols) < 3:
            continue
        dept, name, raw = cols[0], cols[1], cols[2]
        if dept in ('部署', '') or name in ('氏名', ''):
            continue          # ヘッダー行・区切り行
        m = MAIL.search(raw)
        rows.append({
            'dept_src': dept,
            'dept': DEPT_MAP.get(dept, ''),
            'name': name.replace(' ', '').replace('　', ''),
            'mail': m.group(0).lower() if m else '',
        })
    return rows


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    src, dst = sys.argv[1], sys.argv[2]
    if not os.path.exists(src):
        print('入力が見つかりません: %s' % src)
        return 1

    rows = parse(src)
    lines = ['\t'.join(HEADER)]
    for r in rows:
        lines.append('\t'.join([r['mail'], r['name'], r['dept'], '0', '0', '', '', '']))
    io.open(dst, 'w', encoding='utf-8', newline='\n').write('\n'.join(lines) + '\n')

    no_mail = [r for r in rows if not r['mail']]
    no_dept = [r for r in rows if not r['dept']]
    print('人数 %d名' % len(rows))
    print('メール取得もれ %d件 / 部署未対応 %d件' % (len(no_mail), len(no_dept)))
    for r in no_mail + no_dept:
        print('  🔴 %s（%s）' % (r['name'], r['dept_src']))

    seen = {}
    for r in rows:
        seen.setdefault(r['mail'], []).append(r['name'])
    for k, v in seen.items():
        if len(v) > 1:
            print('  ⚠ メール重複 %s' % v)

    print()
    print('%-22s %-20s %s' % ('原本の部署', '判定に使う部署', '人数'))
    agg = {}
    for r in rows:
        key = (r['dept_src'], r['dept'])
        agg[key] = agg.get(key, 0) + 1
    for (a, b), n in sorted(agg.items(), key=lambda x: -x[1]):
        print('%-22s %-20s %2d名%s' % (a, b, n, '' if b else '  🔴未対応'))
    print()
    print('書き出し: %s' % os.path.normpath(dst))
    print('⚠ 出力には個人情報が入る。リポジトリに置かないこと。')
    return 0


if __name__ == '__main__':
    sys.exit(main())

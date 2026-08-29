#!/usr/bin/env python3
"""gas/ 配下のファイルが Drive の原本と一致しているかを確かめる。

Drive から取り出したファイルは、転送の途中で1バイト単位の取りこぼしが起きうる。
実際に コード.gs で1バイトの欠落が起き、日本語コメントが壊れた
（「APIキーの診断」→「APIキーの訽??」）。以後は必ずこの3点を通す。

  1. バイト数が manifest.tsv の期待値と完全に一致するか
  2. UTF-8 として厳密にデコードできるか（壊れると必ずここで落ちる）
  3. JavaScript として構文が通るか（node --check。node があるときだけ）

使い方:
    python3 tools/verify_gas.py
終了コード 0 なら全件一致。1 なら不一致あり。
"""

import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MANIFEST = os.path.join(ROOT, "gas", "manifest.tsv")


def load_manifest():
    rows = []
    with open(MANIFEST, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line or line.startswith("#"):
                continue
            name, drive_id, size, state = line.split("\t")
            rows.append((name, drive_id, int(size), state))
    return rows


def check_syntax(path):
    """node --check は拡張子で判定するので .js に写してから渡す。"""
    if not shutil.which("node"):
        return None
    tmp = os.path.join(tempfile.mkdtemp(), "check.js")
    shutil.copy(path, tmp)
    r = subprocess.run(["node", "--check", tmp], capture_output=True)
    return r.returncode == 0


def main():
    ng = 0
    pending = 0
    for name, _drive_id, size, state in load_manifest():
        path = os.path.join(ROOT, "gas", name)

        if state in ("pending", "excluded"):
            if state == "pending":
                pending += 1
            if os.path.exists(path):
                label = "未移送" if state == "pending" else "意図して入れない"
                print(f"NG   {name}: manifest では{label}のにファイルがある")
                ng += 1
            elif state == "excluded":
                print(f"--   {name}: 意図して入れていない")
            continue

        if not os.path.exists(path):
            print(f"NG   {name}: 移送済みのはずだがファイルがない")
            ng += 1
            continue

        data = open(path, "rb").read()
        problems = []

        # config.gs.sample は APP_TOKEN を伏せ字にしている分だけ短い
        expected = size - 4 if name == "config.gs.sample" else size
        if len(data) != expected:
            problems.append(f"サイズ {len(data)} != {expected}")

        try:
            data.decode("utf-8")
        except UnicodeDecodeError as e:
            problems.append(f"UTF-8として壊れている（{e.start}バイト目）")

        if name.endswith(".gs"):
            ok = check_syntax(path)
            if ok is False:
                problems.append("JavaScriptとして構文が通らない")

        if problems:
            print(f"NG   {name}: " + " / ".join(problems))
            ng += 1
        else:
            print(f"OK   {name}")

    print()
    print(f"不一致 {ng}件 / 未移送 {pending}件")
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())

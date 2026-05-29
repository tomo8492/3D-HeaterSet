#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
M_SanjigenDaisha.bas (UTF-8) → M_SanjigenDaisha_sjis.bas (Shift-JIS / CP932, CRLF) 生成。

日本語環境の VBE「ファイル > ファイルのインポート」は .bas を Shift-JIS として
読み込むため、UTF-8版をインポートすると文字化けする。インポート用の Shift-JIS版を
このスクリプトで生成して同梱する。

★ 重要: M_SanjigenDaisha.bas(UTF-8) を編集したら必ず本スクリプトを実行し、
        _sjis.bas を再生成して同期させること(同期ズレ・文字化け防止)。

使い方:
    python gen_sjis.py
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "M_SanjigenDaisha.bas")
DST = os.path.join(HERE, "M_SanjigenDaisha_sjis.bas")


def _cp932_ok(ch: str) -> bool:
    try:
        ch.encode("cp932")
        return True
    except UnicodeEncodeError:
        return False


def main() -> int:
    if not os.path.exists(SRC):
        print(f"NG: ソースが見つかりません: {SRC}")
        return 1

    with open(SRC, encoding="utf-8") as f:
        text = f.read()

    bad = sorted({ch for ch in text if not _cp932_ok(ch)})
    if bad:
        print("NG: CP932(Shift-JIS)に変換できない文字があります:")
        for ch in bad:
            print(f"    U+{ord(ch):04X} {ch!r}")
        print("    → これらの文字を使わないよう .bas を修正してください。")
        return 1

    # VBE 標準形式に合わせ改行を CRLF へ正規化
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")

    with open(DST, "wb") as f:
        f.write(text.encode("cp932"))

    # 検証(復号して内容一致と改行を確認)
    data = open(DST, "rb").read()
    lone_lf = sum(1 for i, b in enumerate(data) if b == 0x0A and (i == 0 or data[i - 1] != 0x0D))
    ok = (data.decode("cp932").replace("\r\n", "\n") == text.replace("\r\n", "\n")) and lone_lf == 0
    print(f"生成: {DST} ({len(data)} bytes, CRLF={data.count(bytes([13, 10]))}, 孤立LF={lone_lf})")
    print("検証: " + ("OK(内容一致・改行CRLF)" if ok else "NG(要確認)"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

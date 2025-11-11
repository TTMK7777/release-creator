#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
改善版VBAモジュールをShift_JISに変換
"""

import os
import sys

def convert_to_sjis(input_file, output_file):
    """UTF-8のVBAファイルをShift_JIS (CP932) に変換"""
    try:
        # UTF-8で読み込み (BOMを削除)
        with open(input_file, 'r', encoding='utf-8-sig') as f:
            content = f.read()

        # CP932で表現できない文字を置換
        replacements = {
            '\u2713': '[OK]',      # ✓ → [OK]
            '\u2714': '[OK]',      # ✔ → [OK]
            '\u2717': '[NG]',      # ✗ → [NG]
            '\u2718': '[NG]',      # ✘ → [NG]
            '\xae': '(R)',         # ® → (R)
        }

        for char, replacement in replacements.items():
            content = content.replace(char, replacement)

        # Shift_JIS (CP932) で保存
        with open(output_file, 'w', encoding='cp932') as f:
            f.write(content)

        print(f"[OK] {os.path.basename(output_file)} を作成しました")
        print(f"     サイズ: {len(content)} 文字")
        return True

    except Exception as e:
        print(f"[ERROR] 変換失敗: {e}")
        return False

if __name__ == "__main__":
    # vba_modulesディレクトリのパス
    script_dir = os.path.dirname(os.path.abspath(__file__))
    vba_dir = os.path.join(os.path.dirname(script_dir), "vba_modules")

    # 改善版モジュールのリスト
    modules = [
        ("Module3_Image_Improved.bas", "Module3_Image_Improved_SJIS.bas"),
        ("Module4_Word_Improved.bas", "Module4_Word_Improved_SJIS.bas"),
    ]

    print(f"=== 改善版VBAモジュール変換 (UTF-8 → Shift_JIS) ===")
    print(f"対象ディレクトリ: {vba_dir}")
    print()

    for input_name, output_name in modules:
        input_path = os.path.join(vba_dir, input_name)
        output_path = os.path.join(vba_dir, output_name)

        if os.path.exists(input_path):
            print(f"変換中: {input_name}")
            convert_to_sjis(input_path, output_path)
            print()
        else:
            print(f"[SKIP] ファイルが見つかりません: {input_name}")
            print()

    print("変換完了")

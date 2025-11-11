#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Module1_Main_OptimizedをShift_JISに変換
"""

import os

def convert_to_sjis(input_file, output_file):
    """UTF-8のVBAファイルをShift_JIS (CP932) に変換"""
    try:
        # UTF-8で読み込み
        with open(input_file, 'r', encoding='utf-8-sig') as f:
            content = f.read()

        # CP932で表現できない文字を置換
        replacements = {
            '\u2713': '[OK]',
            '\u2714': '[OK]',
            '\u2717': '[NG]',
            '\u2718': '[NG]',
            '\xae': '(R)',
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
    script_dir = os.path.dirname(os.path.abspath(__file__))
    vba_dir = os.path.join(os.path.dirname(script_dir), "vba_modules")

    input_path = os.path.join(vba_dir, "Module1_Main_Optimized.bas")
    output_path = os.path.join(vba_dir, "Module1_Main_Optimized_SJIS.bas")

    print("=== Module1_Main_Optimized 変換 (UTF-8 → Shift_JIS) ===")
    print(f"対象ディレクトリ: {vba_dir}")
    print()

    if os.path.exists(input_path):
        print(f"変換中: Module1_Main_Optimized.bas")
        convert_to_sjis(input_path, output_path)
        print()
    else:
        print(f"[ERROR] ファイルが見つかりません: Module1_Main_Optimized.bas")

    print("変換完了")

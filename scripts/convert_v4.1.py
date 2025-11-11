#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
v4.1ファイルをShift_JISに変換
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
    utf8_dir = os.path.join(vba_dir, "archive", "utf8")

    print("=== v4.1 変換 (UTF-8 → Shift_JIS) ===")
    print(f"UTF-8ディレクトリ: {utf8_dir}")
    print(f"出力ディレクトリ: {vba_dir}")
    print()

    # Module1_Main_Optimized_v4.1.bas
    input_path1 = os.path.join(utf8_dir, "Module1_Main_Optimized_v4.1.bas")
    output_path1 = os.path.join(vba_dir, "Module1_Main_Optimized_v4.1_SJIS.bas")

    if os.path.exists(input_path1):
        print(f"変換中: Module1_Main_Optimized_v4.1.bas")
        convert_to_sjis(input_path1, output_path1)
        print()
    else:
        print(f"[ERROR] ファイルが見つかりません: {input_path1}")

    # Module2_Data_Complete.bas (完全版)
    input_path2 = os.path.join(utf8_dir, "Module2_Data_Complete.bas")
    output_path2 = os.path.join(vba_dir, "Module2_Data_Complete_SJIS.bas")

    if os.path.exists(input_path2):
        print(f"変換中: Module2_Data_Complete.bas")
        convert_to_sjis(input_path2, output_path2)
        print()
    else:
        print(f"[ERROR] ファイルが見つかりません: {input_path2}")

    # Module3_Image_Improved_v2.1.bas
    input_path3 = os.path.join(utf8_dir, "Module3_Image_Improved_v2.1.bas")
    output_path3 = os.path.join(vba_dir, "Module3_Image_Improved_v2.1_SJIS.bas")

    if os.path.exists(input_path3):
        print(f"変換中: Module3_Image_Improved_v2.1.bas")
        convert_to_sjis(input_path3, output_path3)
        print()
    else:
        print(f"[ERROR] ファイルが見つかりません: {input_path3}")

    print("変換完了")

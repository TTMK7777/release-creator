#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
VBAファイルをShift_JIS (CP932) エンコーディングに変換するスクリプト

Excel VBAは日本語環境でCP932を期待するため、
UTF-8で保存されたVBAファイルを変換する必要がある。
"""

import os

def convert_to_sjis(input_file):
    """UTF-8のVBAファイルをShift_JIS (CP932) に変換"""

    # 出力ファイル名を生成 (元のファイル名を保持)
    base_name = os.path.splitext(input_file)[0]
    output_file = base_name + "_SJIS.bas"

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
            '\xae': '(R)',         # ® → (R) ※VBAでは Chr(&HAE) で表現可能
        }

        for char, replacement in replacements.items():
            content = content.replace(char, replacement)

        # Shift_JIS (CP932) で保存 (BOMなし)
        with open(output_file, 'w', encoding='cp932') as f:
            f.write(content)

        print(f"[OK] {os.path.basename(output_file)} を作成しました")
        print(f"     入力: {input_file}")
        print(f"     出力: {output_file}")
        print(f"     サイズ: {len(content)} 文字")

        return True

    except Exception as e:
        print(f"[ERROR] 変換失敗: {e}")
        return False

if __name__ == "__main__":
    # 現在のディレクトリ内の全.basファイルを変換
    current_dir = os.path.dirname(os.path.abspath(__file__))

    bas_files = [f for f in os.listdir(current_dir) if f.endswith('.bas') and not f.endswith('_SJIS.bas')]

    print(f"=== VBA ファイル変換 (UTF-8 → Shift_JIS) ===")
    print(f"対象ディレクトリ: {current_dir}")
    print(f"対象ファイル数: {len(bas_files)}")
    print()

    for bas_file in bas_files:
        convert_to_sjis(os.path.join(current_dir, bas_file))
        print()

    print("変換完了")

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
GPT-4oが生成したModule3/4を抽出するスクリプト
"""

import json
import re
import os

def extract_vba_code(text):
    """VBAコードブロックを抽出"""
    # ```vb または ```vba のコードブロックを検索
    pattern = r'```(?:vba|vb)\n(.*?)```'
    matches = re.findall(pattern, text, re.DOTALL)
    return matches

def main():
    report_file = r'C:\Users\t-tsuji\.claude\orchestrator\reports\orchestrator_report_v3_20251110_182155.json'

    print("=" * 60)
    print("GPT-4o Module3/4 抽出")
    print("=" * 60)

    with open(report_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # GPT-4oの出力を取得
    phase2 = data['results']['phase_2_results']
    gpt_output = phase2['gpt4o']['output']

    print(f"GPT-4o出力長: {len(gpt_output)} 文字")
    print()

    # VBAコードを抽出
    vba_blocks = extract_vba_code(gpt_output)
    print(f"検出されたVBAコードブロック: {len(vba_blocks)}")
    print()

    # モジュール名を検出して保存
    output_dir = r'C:\Users\t-tsuji\AIアプリ開発\release-creator\vba_modules'

    for i, code in enumerate(vba_blocks, 1):
        # Attribute VB_Nameからモジュール名を抽出
        name_match = re.search(r'Attribute VB_Name = "([^"]+)"', code)
        if name_match:
            module_name = name_match.group(1)
        else:
            module_name = f"Module_GPT4o_{i}"

        output_file = os.path.join(output_dir, f"{module_name}.bas")

        # UTF-8で保存
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(code)

        print(f"[OK] 保存: {module_name}.bas ({len(code)} 文字)")

        # CP932版も生成
        try:
            # BOMを削除、特殊文字を置換
            content = code.replace('\ufeff', '')
            content = content.replace('✓', '[OK]')
            content = content.replace('✔', '[OK]')
            content = content.replace('✗', '[NG]')
            content = content.replace('✘', '[NG]')
            content = content.replace('®', '(R)')

            output_file_sjis = os.path.join(output_dir, f"{module_name}_SJIS.bas")
            with open(output_file_sjis, 'w', encoding='cp932') as f:
                f.write(content)

            print(f"     → SJIS版: {module_name}_SJIS.bas")

        except Exception as e:
            print(f"     [WARNING] SJIS変換失敗: {e}")

    print()
    print("=" * 60)
    print(f"抽出完了! 合計 {len(vba_blocks)} モジュール")
    print(f"出力先: {output_dir}")
    print("=" * 60)

if __name__ == '__main__':
    main()

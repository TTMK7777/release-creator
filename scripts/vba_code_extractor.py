#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBAコード抽出ツール - Multi-AI Orchestratorレポートから抽出
"""

import json
import re
import os

def extract_vba_code(json_path, output_dir):
    """OrchestratorレポートからVBAコードを抽出"""

    print("=" * 60)
    print("VBAコード抽出開始")
    print("=" * 60)

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 実装結果を取得
    results = data.get('results', {})
    phase2 = results.get('phase_2_results', {})

    # Claude Sonnet 4.5の結果
    claude_result = phase2.get('claude_sonnet_4.5', {}).get('output', '')
    print(f"[OK] Claude Sonnet 4.5の結果を取得 ({len(claude_result)} 文字)")

    # GPT-4oの結果
    gpt4o_result = phase2.get('gpt4o', {}).get('output', '')
    print(f"[OK] GPT-4oの結果を取得 ({len(gpt4o_result)} 文字)")

    # 結合
    combined_text = claude_result + "\n\n" + gpt4o_result

    # VBAコードブロックを抽出（vbaまたはvb）
    vba_pattern = r'```(?:vba|vb)\n(.*?)```'
    vba_blocks = re.findall(vba_pattern, combined_text, re.DOTALL)

    print(f"\n見つかったVBAコードブロック数: {len(vba_blocks)}")

    # モジュールごとに分類
    modules = {}

    for i, code in enumerate(vba_blocks):
        # モジュール名を特定
        module_name = None

        if 'Module1_Main' in code or 'Attribute VB_Name = "Module1_Main"' in code:
            module_name = 'Module1_Main'
        elif 'Module2_Data' in code or 'Attribute VB_Name = "Module2_Data"' in code:
            module_name = 'Module2_Data'
        elif 'Module3_Image' in code or 'Attribute VB_Name = "Module3_Image"' in code:
            module_name = 'Module3_Image'
        elif 'Module4_Word' in code or 'Attribute VB_Name = "Module4_Word"' in code:
            module_name = 'Module4_Word'
        elif 'Module5_Utils' in code or 'Attribute VB_Name = "Module5_Utils"' in code:
            module_name = 'Module5_Utils'
        else:
            module_name = f'Module_Unknown_{i+1}'

        # より長いコードを優先（完全版）
        if module_name not in modules or len(code) > len(modules.get(module_name, '')):
            modules[module_name] = code

    # ファイルに保存
    os.makedirs(output_dir, exist_ok=True)

    for module_name, code in modules.items():
        output_path = os.path.join(output_dir, f'{module_name}.bas')

        with open(output_path, 'w', encoding='utf-8-sig') as f:  # BOM付きUTF-8
            f.write(code)

        print(f"[OK] 保存完了: {module_name}.bas ({len(code)} 文字)")

    # QAレビュー結果も保存
    phase3 = results.get('phase_3_results', {})
    gemini_result = phase3.get('gemini_2.5_reviewer', {}).get('output', '')

    qa_path = os.path.join(output_dir, 'QA_Review.md')
    with open(qa_path, 'w', encoding='utf-8') as f:
        f.write("# Gemini 2.5 Flash - コードレビューレポート\n\n")
        f.write(gemini_result)

    print(f"[OK] QAレビュー保存: QA_Review.md")

    print("\n" + "=" * 60)
    print(f"抽出完了! 合計 {len(modules)} モジュール")
    print(f"出力先: {output_dir}")
    print("=" * 60)

    return modules

if __name__ == '__main__':
    # 最新のレポートファイルを自動検出
    import glob
    reports_dir = r'C:\Users\t-tsuji\.claude\orchestrator\reports'
    json_files = glob.glob(os.path.join(reports_dir, 'orchestrator_report_v3_*.json'))

    if not json_files:
        print("エラー: レポートファイルが見つかりません")
        exit(1)

    # 最新のファイルを取得
    json_path = max(json_files, key=os.path.getmtime)
    print(f"使用するレポート: {os.path.basename(json_path)}\n")

    output_dir = r'C:\Users\t-tsuji\AIアプリ開発\release-creator\vba_modules'

    modules = extract_vba_code(json_path, output_dir)

    # モジュール一覧を表示
    print("\n抽出されたモジュール:")
    for module_name in sorted(modules.keys()):
        print(f"  - {module_name}")

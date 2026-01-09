#!/usr/bin/env python3
"""
ドキュメント自動生成/更新スクリプト
version.json からREADME.mdとHANDOVER.mdを自動生成/更新

使用方法:
    python scripts/generate_docs.py [--force]

オプション:
    --force: 既存ファイルを強制的に上書き（デフォルトは更新のみ）

必要ファイル:
    - version.json (プロジェクトルート)

セキュリティ注意:
    - version.jsonにAPIキー、トークン、パスワード等の機密情報を含めないこと
    - env_varsには環境変数の名前のみを記載し、値は絶対に含めないこと
"""

import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

# 機密情報を含む可能性のあるキーワード（検出用）
SENSITIVE_KEYWORDS = [
    "api_key", "apikey", "secret", "token", "password", "passwd",
    "credential", "private_key", "access_key", "auth"
]


def check_sensitive_data(data: dict, path: str = "") -> list:
    """機密情報が含まれていないかチェック"""
    warnings = []
    for key, value in data.items():
        current_path = f"{path}.{key}" if path else key
        key_lower = key.lower()

        # キー名に機密キーワードが含まれているかチェック
        for keyword in SENSITIVE_KEYWORDS:
            if keyword in key_lower and value:
                warnings.append(f"Warning: '{current_path}' may contain sensitive data")

        # 値が辞書の場合、再帰的にチェック
        if isinstance(value, dict):
            warnings.extend(check_sensitive_data(value, current_path))
        # 値がリストの場合、各要素をチェック
        elif isinstance(value, list):
            for i, item in enumerate(value):
                if isinstance(item, dict):
                    warnings.extend(check_sensitive_data(item, f"{current_path}[{i}]"))

    return warnings


def load_version_json(path: str = "version.json") -> dict:
    """version.jsonを読み込む（機密情報チェック付き）"""
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 機密情報チェック
    warnings = check_sensitive_data(data)
    for warning in warnings:
        print(f"[WARN] {warning}")

    if warnings:
        print("[WARN] version.json may contain sensitive data.")
        print("[WARN] Store API keys in .env, not in version.json.")

    return data


def generate_readme(data: dict) -> str:
    """README.mdを生成"""
    name = data.get("name", "Project")
    version = data.get("version", "1.0")
    edition = data.get("edition", "")
    description = data.get("description", "")
    highlights = data.get("highlights", [])
    features = data.get("features", [])
    changelog = data.get("changelog", [])
    quick_start = data.get("quick_start", {})
    tech_stack = data.get("tech_stack", [])

    version_str = f"v{version}"
    if edition:
        version_str += f" ({edition})"

    # README生成
    lines = []
    lines.append(f"# {name}")
    lines.append("")
    lines.append(f"![Version](https://img.shields.io/badge/version-{version}-blue)")
    # Windows互換のフォーマット（%-m は Linux のみ）
    date_str = datetime.now().strftime('%Y-%m-%d').replace('-0', '-')
    lines.append(f"![Updated](https://img.shields.io/badge/updated-{date_str}-green)")
    lines.append("")

    if description:
        lines.append(f"> {description}")
        lines.append("")

    # ハイライト
    if highlights:
        lines.append("## Highlights")
        lines.append("")
        for h in highlights:
            lines.append(f"- {h}")
        lines.append("")

    # 機能
    if features:
        lines.append("## Features")
        lines.append("")
        for f in features:
            lines.append(f"- {f}")
        lines.append("")

    # クイックスタート
    if quick_start.get("install") or quick_start.get("run"):
        lines.append("## Quick Start")
        lines.append("")
        if quick_start.get("install"):
            lines.append("```bash")
            lines.append(f"# Install")
            lines.append(quick_start["install"])
            lines.append("```")
            lines.append("")
        if quick_start.get("run"):
            lines.append("```bash")
            lines.append(f"# Run")
            lines.append(quick_start["run"])
            lines.append("```")
            lines.append("")
        if quick_start.get("env_vars"):
            lines.append("### Environment Variables")
            lines.append("")
            lines.append("```bash")
            for env in quick_start["env_vars"]:
                lines.append(f"{env}=your_value")
            lines.append("```")
            lines.append("")

    # 技術スタック
    if tech_stack:
        lines.append("## Tech Stack")
        lines.append("")
        for tech in tech_stack:
            lines.append(f"- {tech}")
        lines.append("")

    # 変更履歴
    if changelog:
        lines.append("## Changelog")
        lines.append("")
        for entry in changelog[:5]:  # 最新5件
            v = entry.get("version", "?")
            date = entry.get("date", "")
            changes = entry.get("changes", [])
            lines.append(f"### v{v} ({date})")
            for c in changes:
                lines.append(f"- {c}")
            lines.append("")

    # フッター
    lines.append("---")
    lines.append("")
    lines.append(f"最終更新: {datetime.now().strftime('%Y-%m-%d')}")
    lines.append(f"バージョン: {version_str}")
    lines.append("")

    return "\n".join(lines)


def generate_handover(data: dict) -> str:
    """HANDOVER.mdを生成"""
    name = data.get("name", "Project")
    version = data.get("version", "1.0")
    edition = data.get("edition", "")
    description = data.get("description", "")
    features = data.get("features", [])
    changelog = data.get("changelog", [])
    quick_start = data.get("quick_start", {})
    tech_stack = data.get("tech_stack", [])
    project_type = data.get("project_type", "Unknown")

    version_str = f"v{version}"
    if edition:
        version_str += f" ({edition})"

    lines = []
    lines.append(f"# {name} - 引継ぎ資料")
    lines.append("")
    lines.append(f"**バージョン**: {version_str}")
    lines.append(f"**最終更新**: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"**プロジェクトタイプ**: {project_type}")
    lines.append("")
    lines.append("---")
    lines.append("")

    # 概要
    lines.append("## 1. プロジェクト概要")
    lines.append("")
    if description:
        lines.append(description)
    else:
        lines.append("（説明なし）")
    lines.append("")

    # 主要機能
    lines.append("## 2. 主要機能")
    lines.append("")
    if features:
        for i, f in enumerate(features, 1):
            lines.append(f"{i}. {f}")
    else:
        lines.append("- （未定義）")
    lines.append("")

    # 技術スタック
    lines.append("## 3. 技術スタック")
    lines.append("")
    if tech_stack:
        lines.append("| カテゴリ | 技術 |")
        lines.append("|----------|------|")
        for tech in tech_stack:
            lines.append(f"| - | {tech} |")
    else:
        lines.append("- （未定義）")
    lines.append("")

    # セットアップ手順
    lines.append("## 4. セットアップ手順")
    lines.append("")
    if quick_start.get("install"):
        lines.append("### インストール")
        lines.append("")
        lines.append("```bash")
        lines.append(quick_start["install"])
        lines.append("```")
        lines.append("")
    if quick_start.get("run"):
        lines.append("### 実行")
        lines.append("")
        lines.append("```bash")
        lines.append(quick_start["run"])
        lines.append("```")
        lines.append("")
    if quick_start.get("env_vars"):
        lines.append("### 環境変数")
        lines.append("")
        lines.append("| 変数名 | 説明 | 必須 |")
        lines.append("|--------|------|------|")
        for env in quick_start["env_vars"]:
            lines.append(f"| `{env}` | - | Yes |")
        lines.append("")

    if not quick_start.get("install") and not quick_start.get("run"):
        lines.append("（セットアップ手順未定義）")
        lines.append("")

    # 変更履歴
    lines.append("## 5. 変更履歴")
    lines.append("")
    if changelog:
        for entry in changelog:
            v = entry.get("version", "?")
            date = entry.get("date", "")
            changes = entry.get("changes", [])
            lines.append(f"### v{v} ({date})")
            lines.append("")
            for c in changes:
                lines.append(f"- {c}")
            lines.append("")
    else:
        lines.append("- 初期リリース")
        lines.append("")

    # 注意事項
    lines.append("## 6. 注意事項・既知の問題")
    lines.append("")
    lines.append("- （特になし）")
    lines.append("")

    # 連絡先
    lines.append("## 7. 連絡先")
    lines.append("")
    lines.append("- 担当者: （要設定）")
    lines.append("- リポジトリ: （GitHubリンク）")
    lines.append("")

    # フッター
    lines.append("---")
    lines.append("")
    lines.append("*この資料は version.json から自動生成されています。*")
    lines.append(f"*生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")
    lines.append("")

    return "\n".join(lines)


def update_file_if_exists(filepath: str, new_content: str, force: bool = False) -> str:
    """
    ファイルが存在しない場合は新規生成、存在する場合は更新

    Returns:
        "created": 新規作成
        "updated": 更新
        "unchanged": 変更なし
    """
    if not os.path.exists(filepath):
        # 新規作成
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(new_content)
        return "created"

    # 既存ファイルを読み込み
    with open(filepath, "r", encoding="utf-8") as f:
        existing_content = f.read()

    # 内容が同じなら何もしない
    if existing_content.strip() == new_content.strip():
        return "unchanged"

    # 更新
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(new_content)
    return "updated"


def main():
    """メイン処理"""
    # コマンドライン引数の処理
    force = "--force" in sys.argv

    # version.json を読み込み
    data = load_version_json("version.json")

    if not data:
        print("Warning: version.json not found or empty. Using defaults.")
        data = {
            "name": Path.cwd().name,
            "version": "1.0",
            "description": "",
        }

    print(f"Project: {data.get('name', 'Unknown')}")
    print(f"Version: v{data.get('version', '?')}")
    print("-" * 40)

    # README.md 生成/更新
    readme_content = generate_readme(data)
    readme_status = update_file_if_exists("README.md", readme_content, force)
    status_icon = {"created": "[NEW]", "updated": "[UPD]", "unchanged": "[SKIP]"}
    print(f"{status_icon[readme_status]} README.md: {readme_status}")

    # HANDOVER.md 生成/更新
    handover_content = generate_handover(data)
    handover_status = update_file_if_exists("HANDOVER.md", handover_content, force)
    print(f"{status_icon[handover_status]} HANDOVER.md: {handover_status}")

    print("-" * 40)
    print("Done!")


if __name__ == "__main__":
    main()

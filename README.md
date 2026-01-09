# Release Creator - プレスリリース自動生成システム

![Version](https://img.shields.io/badge/version-7.10-blue)
![Updated](https://img.shields.io/badge/updated-2026-1-9-green)

> Streamlit WebアプリでExcelランキングデータからプレスリリース用の表を自動生成

## Highlights

- 全215ランキング対応検証済み
- ハイブリッド自動検出（sort-nav優先 + レガシーフォールバック）
- Word出力機能
- Streamlit Cloud デプロイ済み

## Features

- Excel (.xlsx) ファイルのアップロード
- 総合ランキング/評価項目/部門別の自動解析
- プレスリリース用表の自動生成
- トレンドグラフ表示
- 画像ダウンロード機能
- 複数部門対応（ネット証券、FX、クレジットカード等）
- 社名正規化・エイリアス対応

## Quick Start

```bash
# Install
pip install -r streamlit-app/requirements.txt
```

```bash
# Run
cd streamlit-app && streamlit run app.py
```

## Tech Stack

- Python 3.11+
- Streamlit
- pandas
- openpyxl
- python-docx

## Changelog

### v7.10 (2025-12)
- 最新安定版

### v7.0 (2025-12)
- ハイブリッド自動検出
- 社名正規化・エイリアス対応
- Word出力機能
- 全215ランキング対応検証済み

### v6.2 (2025-12-02)
- コードレビュー・リファクタリング
- SVODジャンル別部門名抽出修正

### v6.0 (2025-12-02)
- ネット証券部門対応
- トレンドグラフをタブ上部に配置

### v5.0 (2025-11-25)
- Streamlit Cloud対応
- UI全面改善

---

最終更新: 2026-01-09
バージョン: v7.10

# Release Creator - プレスリリース自動生成システム

![Version](https://img.shields.io/badge/version-8.2-blue)
![Updated](https://img.shields.io/badge/updated-2026--02--26-green)

> Streamlit WebアプリでExcelランキングデータからプレスリリース用の表を自動生成

## Highlights

- 全215ランキング対応検証済み
- ハイブリッド自動検出（sort-nav優先 + レガシーフォールバック）
- Word出力機能
- Streamlit Cloud デプロイ済み
- **ポータブル配布対応**（Python未インストールPCで動作）
- **未公表ローカルデータ参照**（共有フォルダ経由で公開前データを安全参照）

## Features

- Excel (.xlsx) ファイルのアップロード
- 総合ランキング/評価項目/部門別の自動解析
- プレスリリース用表の自動生成
- トレンドグラフ表示
- 画像ダウンロード機能
- 複数部門対応（ネット証券、FX、クレジットカード等）
- 社名正規化・エイリアス対応

## Quick Start

### 開発者向け
```bash
pip install -r streamlit-app/requirements.txt
cd streamlit-app && streamlit run app.py
```

### ポータブル版ビルド
```bash
python build/build_portable.py
# → dist/release-creator-portable/ にパッケージ生成
```

### エンドユーザー向け
共有フォルダの `ReleaseCreator.bat` をダブルクリック（Python不要）

## Tech Stack

- Python 3.11+
- Streamlit
- pandas
- openpyxl
- python-docx

## Changelog

### v8.2 / LocalDataReader v1.2 (2026-02-26)
- 未公表ローカルデータ参照機構を追加
- `local_data_reader.py`: config.json マッピング + Excel ヘッダー自動検出
- 社内Excel形式（`合計` スコア列）に対応
- `LOCAL_DATA_PATH` 環境変数で共有フォルダを指定可能

### Portable v1.0 (2026-02)
- Python embeddable ポータブル配布パッケージ
- 自動インストール + 差分更新 + デスクトップショートカット
- ポート 8501-8510 動的フォールバック

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

最終更新: 2026-02-26
バージョン: v8.2 / Portable v1.0 / LocalDataReader v1.2

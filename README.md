# Release Creator - プレスリリース自動生成システム

Streamlit WebアプリでExcelランキングデータからプレスリリース用の表を自動生成します。

**現バージョン**: v6.0
**アプリURL**: https://release-creator.streamlit.app/

---

## 機能

- Excel (.xlsx) ファイルのアップロード
- 総合ランキング/評価項目/部門別の自動解析
- プレスリリース用表の自動生成
- トレンドグラフ表示（評価項目・部門タブ上部）
- 画像ダウンロード機能
- 複数部門対応（ネット証券、FX、クレジットカード等）

---

## クイックスタート

### ローカル実行

```bash
cd streamlit-app
streamlit run app.py
```

または、`起動.bat`をダブルクリック

### 依存パッケージ

```bash
pip install -r streamlit-app/requirements.txt
```

---

## プロジェクト構成

```
release-creator/
├── README.md                 # このファイル
├── 起動.bat                  # ルートからの起動用
│
├── streamlit-app/            # メインアプリケーション
│   ├── app.py               # Streamlitアプリ本体
│   ├── scraper.py           # Excel解析ロジック
│   ├── HANDOVER.md          # 引継ぎ資料・バージョン履歴
│   ├── requirements.txt     # Python依存関係
│   ├── 起動.bat             # アプリ起動バッチ
│   └── test/                # テスト用Excelファイル
│
└── _archive/                 # 旧VBA版（アーカイブ）
    ├── docs/
    ├── scripts/
    ├── vba_modules/
    └── テンプレート/
```

---

## バージョン履歴

| Version | Date | Changes |
|---------|------|---------|
| v6.0 | 2025-12-02 | ネット証券部門対応、トレンドグラフをタブ上部に配置 |
| v5.9 | 2025-12-01 | 部門タブ改善、エラーハンドリング強化 |
| v5.0 | 2025-11-25 | Streamlit Cloud対応、UI全面改善 |

詳細は `streamlit-app/HANDOVER.md` を参照

---

## 開発情報

- **フレームワーク**: Streamlit
- **言語**: Python 3.11+
- **デプロイ**: Streamlit Cloud

---

**最終更新**: 2025年12月2日

# AI統合コードレビュー・リファクタリング報告書

**実施日**: 2025年12月25日
**対象**: オリコン顧客満足度調査 TOPICSサポートシステム
**レビュー担当AI**: Gemini AI (コードレビュー), Perplexity (ベストプラクティス調査)
**統括**: Claude Opus 4.5

---

## 目次

1. [レビュー概要](#レビュー概要)
2. [重大な問題（Critical Issues）](#重大な問題critical-issues)
3. [バグ・潜在的エラー](#バグ潜在的エラー)
4. [パフォーマンス問題](#パフォーマンス問題)
5. [セキュリティ懸念](#セキュリティ懸念)
6. [コード品質・保守性](#コード品質保守性)
7. [リファクタリング提案](#リファクタリング提案)
8. [ベストプラクティス調査結果](#ベストプラクティス調査結果)
9. [対応優先度マトリクス](#対応優先度マトリクス)
10. [総評](#総評)

---

## レビュー概要

### 対象ファイル

| ファイル | 行数 | 主な機能 |
|----------|------|----------|
| `app.py` | 2,268行 | メインアプリケーション、UI |
| `analyzer.py` | 1,376行 | TOPICS分析ロジック |
| `scraper.py` | 1,653行 | オリコンサイトスクレイパー |
| `release_generator.py` | 822行 | プレスリリース生成 |
| `site_analyzer.py` | 564行 | サイト構造分析 |
| `validator.py` | 562行 | データ検証 |
| `word_generator.py` | 907行 | Word文書生成 |
| `image_generator.py` | 582行 | 画像生成 |
| `release_tab.py` | 846行 | プレスリリースタブUI |
| `company_master.py` | 404行 | 企業マスタ管理 |

### 設計上の意図的な選択（問題なし）

> **重要**: 年度の型混在（`int`/`str`、例: `2024` と `"2014-2015"`）は**オリコンサイトのデータ形式に合わせた意図的な設計**です。`_year_sort_key` ヘルパー関数で正しく対応済み。

---

## 重大な問題（Critical Issues）

### 1. リソースリーク（scraper.py, site_analyzer.py）

**問題**: `requests.Session()` を作成後、`app.py` で明示的に `close()` を呼び出していない

**影響**: 長時間実行時にソケットリークでメモリ消費増加

**ファイル**:
- `scraper.py` L272-287: セッション作成
- `site_analyzer.py` L286-299: セッション作成
- `app.py` L1139: scraperインスタンス作成

**現状コード**:
```python
# app.py L1139
scraper = OriconScraper(ranking_slug, ranking_name)
# 処理...
# ❌ scraper.close() が呼ばれていない
```

**推奨修正**:
```python
# with文でコンテキストマネージャー使用
with OriconScraper(ranking_slug, ranking_name) as scraper:
    # 処理...
    # ✅ 自動的にclose()が呼ばれる
```

**備考**: v7.10で `close()` メソッドとコンテキストマネージャーは実装済み。使用側の修正のみ必要。

---

## バグ・潜在的エラー

### 2. スコア0の誤判定（release_generator.py）

**ファイル**: `release_generator.py` L234-238

**問題**:
```python
score = item.get('score')
row.append(f"{score:.2f}点" if score else "-")  # ❌ score=0の場合"-"になる
```

**修正案**:
```python
score = item.get('score')
row.append(f"{score:.2f}点" if score is not None else "-")
```

### 3. Excelパース時の型エラー（app.py）

**ファイル**: `app.py` L377付近

**問題**:
```python
year = int(row[year_col])  # ❌ "2024年"のような文字列でValueError
```

**修正案**:
```python
year_str = str(row[year_col]).replace('年', '').strip()
year = int(year_str)
```

### 4. NonePointerの可能性（analyzer.py）

**ファイル**: `analyzer.py` L327

**問題**:
```python
top_score = data[0].get("score")  # ❌ data[0]がNone/不正な辞書の可能性
```

**修正案**:
```python
if not data or not isinstance(data, list) or not data[0]:
    continue
top_score = data[0].get("score")
```

---

## パフォーマンス問題

### 5. 重複コード（analyzer.py）

**ファイル**: `analyzer.py` L1153-1290

**問題**: `_analyze_item_consecutive_wins` と `_analyze_dept_consecutive_wins` がほぼ同じロジック（150行×2）

**推奨**: 共通関数に統合
```python
def _analyze_consecutive_wins_generic(self, data_dict: Dict, category: str) -> List[Dict]:
    """項目別/部門別共通の連続記録分析"""
    # 共通ロジック
```

### 6. メモリ効率の問題（app.py）

**ファイル**: `app.py` L47-218（`create_excel_export`）

**問題**: 全データをメモリ上でDataFrameに変換してからExcel出力

**備考**: `ExcelWriter` を使用しているので一部対応済み。大規模データでは要注意。

---

## セキュリティ懸念

### 7. 情報漏洩リスク（app.py）

**ファイル**: `app.py` L1248-1254

**現状**: エラー詳細をUIに折りたたみ表示
```python
with st.expander("🔍 エラー詳細（開発者向け）", expanded=False):
    st.code(error_detail, language="python")
```

**評価**: ⚠️ **要注意** - `expanded=False` で対策済みだが、本番環境では環境変数で非表示にすることを推奨

**推奨修正**:
```python
if os.environ.get("SHOW_DEBUG_INFO", "false").lower() == "true":
    with st.expander("🔍 エラー詳細（開発者向け）", expanded=False):
        st.code(error_detail, language="python")
```

### 8. 入力検証の不足（app.py）

**問題**: アップロードファイルの検証が不十分
- ファイルサイズ制限なし
- シート名の事前検証なし

**推奨**:
```python
MAX_FILE_SIZE_MB = 10
if uploaded_file.size > MAX_FILE_SIZE_MB * 1024 * 1024:
    st.error(f"ファイルサイズは{MAX_FILE_SIZE_MB}MB以下にしてください")
    return
```

---

## コード品質・保守性

### 9. マジックナンバー（analyzer.py）

**未定数化の値**:
- L879: `2` (連続記録の最小年数)
- L982: `2.0` (注目すべき得点差)
- L1034: `0.6` (独占と判定する割合)

**推奨**:
```python
# analyzer.py 冒頭に追加
MIN_CONSECUTIVE_YEARS = 2      # 連続記録の最小年数
MIN_NOTABLE_SCORE_DIFF = 2.0   # 注目すべき得点差
DOMINANCE_THRESHOLD = 0.6      # 独占と判定する割合
```

### 10. 複雑度が高い関数

| ファイル | 関数 | 行数 | 推奨 |
|----------|------|------|------|
| `app.py` | `parse_uploaded_excel` | 245行 | 分割推奨 |
| `analyzer.py` | `_calc_consecutive_wins` | 93行 | 分割推奨 |

**推奨分割例**:
```python
def parse_uploaded_excel(self, ...):
    header_row = self._detect_header_row(df)
    columns = self._detect_columns(df, header_row)
    return self._extract_data(df, columns)
```

### 11. 型ヒント不足

**現状**: 一部の関数にのみ型ヒント

**推奨**: 段階的に追加
```python
from typing import Dict, List, Union, Optional, Any

YearType = Union[int, str]
RankingData = Dict[YearType, List[Dict[str, Any]]]

def __init__(
    self,
    overall_data: RankingData,
    item_data: Dict[str, RankingData],
    ...
):
```

---

## リファクタリング提案

### 12. データクラス活用

**現状**: 辞書で複雑なデータ構造を管理

**推奨**:
```python
from dataclasses import dataclass
from typing import Union, List

@dataclass
class ConsecutiveWinRecord:
    company: str
    start_year: Union[int, str]
    end_year: Union[int, str]
    years: int
    years_list: List[Union[int, str]]
    is_current: bool
```

### 13. 設定ファイルの外部化

**現状**: `SUBDOMAIN_MAP`, `DEPT_PATTERNS` 等がコード内に埋め込み

**推奨**: YAML設定ファイル
```yaml
# config/scraper_config.yaml
subdomain_map:
  online-english: juken
  _agent: career
```

---

## ベストプラクティス調査結果

### Streamlit パフォーマンス最適化（2024-2025）

| 機能 | 用途 | 適用状況 |
|------|------|----------|
| `@st.cache_data` | データキャッシュ | ✅ 適用推奨 |
| `@st.cache_resource` | リソースキャッシュ | ✅ 適用推奨 |
| `st.fragment` (1.33.0+) | 部分再レンダリング | 検討推奨 |
| Session State | 状態管理 | ✅ 使用中 |

### BeautifulSoup エラーハンドリング

**推奨パターン**:
```python
from requests.exceptions import Timeout, HTTPError, ConnectionError

try:
    response = session.get(url, timeout=10)
    response.raise_for_status()
except Timeout:
    logger.error("タイムアウト")
except HTTPError as e:
    logger.error(f"HTTPエラー: {e.response.status_code}")
except ConnectionError:
    logger.error("接続エラー")
```

### Python例外処理モダンパターン

**避けるべき**:
```python
except Exception as e:  # ❌ 広すぎる
```

**推奨**:
```python
except (ValueError, TypeError) as e:  # ✅ 具体的
```

---

## 対応優先度マトリクス

| 優先度 | 項目 | 内容 | 推定工数 |
|--------|------|------|----------|
| 🔴 **P0** | #1 | リソースリーク対策（`with`文化） | 2h |
| 🟠 **P1** | #2 | スコア0の誤判定修正 | 30min |
| 🟠 **P1** | #3 | Excelパース型エラー | 1h |
| 🟠 **P1** | #4 | NonePointer対策 | 1h |
| 🟡 **P2** | #5 | 重複コード統合 | 4h |
| 🟡 **P2** | #9 | マジックナンバー定数化 | 2h |
| 🔵 **P3** | #10 | 複雑関数分割 | 6h |
| 🔵 **P3** | #11 | 型ヒント追加 | 8h |
| ⚪ **P4** | #12 | データクラス化 | 10h |
| ⚪ **P4** | #13 | 設定ファイル外部化 | 6h |

---

## 総評

### 強み

- ✅ **年度型混在問題**: `_year_sort_key` で適切に対応済み（意図的な設計）
- ✅ **同点1位処理**: 全TOPICSで正しく考慮
- ✅ **社名エイリアス**: 外部ファイル化で保守性向上
- ✅ **エラーハンドリング**: 基本的な対策は実装済み
- ✅ **ドキュメント**: バージョン履歴・docstringが充実

### 改善点

- ⚠️ **リソース管理**: セッションclose漏れ（即対応推奨）
- ⚠️ **型不整合**: 一部の境界条件（0点、None）
- ⚠️ **複雑度**: 200行超の関数が存在
- ⚠️ **型ヒント不足**: 長期的な保守性への影響

### 推奨対応順序

1. **即時対応（P0）**: リソースリーク修正
2. **1週間以内（P1）**: 型エラー・NonePointer修正
3. **2週間以内（P2）**: 重複コード統合、マジックナンバー定数化
4. **継続的（P3-P4）**: 型ヒント追加、リファクタリング

---

## 付録：レビュー実施AI情報

| 役割 | AI | 担当内容 |
|------|-----|----------|
| コードレビュー | Gemini AI | バグ検出、パフォーマンス分析、リファクタリング提案 |
| ベストプラクティス調査 | Perplexity (sonar-reasoning-pro) | 2024-2025年最新技術情報の収集 |
| 統括・最終判断 | Claude Opus 4.5 | レビュー結果の統合、優先度判定 |

---

**レポート作成日時**: 2025-12-25 14:10 JST
**次回レビュー推奨**: P0/P1対応完了後

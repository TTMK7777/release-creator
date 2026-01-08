# -*- coding: utf-8 -*-
"""
オリコン顧客満足度®調査 TOPICSサポートシステム
Streamlit版 - バージョンはHANDOVER.mdで管理
- 総合ランキングタブに1位獲得回数ランキングを追加（評価項目別・部門別と同様）
- 評価項目別・部門別タブに1位獲得回数ランキングを追加
- 年度検出ロジック修正: 更新日を年度基準として使用（調査期間は不使用）
- 同率順位対応: 同点の場合「同率1位」として分析（2社/3社以上対応）
- 順位抽出: icon-rankクラス優先、評価項目別テーブル除外
- 年度列の誤検出を防止（回答者数（最新年）等を除外）
- 年度値の妥当性チェック（動的年度範囲対応）
- オリコン内部Excelフォーマット対応（ヘッダー行自動検出）
- セキュリティ改善: トレースバック情報の非公開化
- 動的年度検出: トップページから実際の発表年度を自動判定
"""

# バージョン情報
__version__ = "β版"

import logging

# ロギング設定（本番環境ではログレベルをINFOに変更推奨）
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

import os
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from scraper import OriconScraper
from analyzer import TopicsAnalyzer, HistoricalAnalyzer, _year_sort_key

# プレスリリース生成・正誤チェックモジュール (v8.0追加)
try:
    from release_tab import render_release_tab, RELEASE_FEATURES_AVAILABLE
except ImportError as e:
    logger.warning(f"プレスリリース機能モジュールが見つかりません: {e}")
    RELEASE_FEATURES_AVAILABLE = False

# アップロード機能の有効化フラグ（環境変数で制御）
# Streamlit Cloud: Secrets で ENABLE_UPLOAD_FEATURE = "true" を設定
# ローカル: 環境変数 ENABLE_UPLOAD_FEATURE=true を設定
ENABLE_UPLOAD = os.environ.get("ENABLE_UPLOAD_FEATURE", "false").lower() == "true"


def create_excel_export(ranking_name, overall_data, item_data, dept_data, historical_data, used_urls=None):
    """取得データをExcelファイルにエクスポート"""
    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # === シート1: サマリー ===
        summary_data = []
        records = historical_data.get("historical_records", {})
        summary = records.get("summary", {})

        if summary.get("max_consecutive"):
            mc = summary["max_consecutive"]
            summary_data.append(["最長連続1位", mc["company"], f"{mc['years']}年連続", f"{mc['start_year']}〜{mc['end_year']}"])
        if summary.get("all_time_high"):
            ath = summary["all_time_high"]
            summary_data.append(["過去最高得点", ath["company"], f"{ath['score']}点", f"{ath['year']}年"])
        if summary.get("most_wins"):
            mw = summary["most_wins"]
            summary_data.append(["最多1位獲得", mw["company"], f"{mw['wins']}回", f"{mw['total_years']}年中"])

        if summary_data:
            df_summary = pd.DataFrame(summary_data, columns=["記録", "企業名", "数値", "詳細"])
            df_summary.to_excel(writer, sheet_name="サマリー", index=False)

        # === シート2: 総合ランキング（全年度） ===
        all_overall = []
        for year in sorted(overall_data.keys(), key=_year_sort_key, reverse=True):
            for item in overall_data[year]:
                all_overall.append({
                    "年度": year,
                    "順位": item.get("rank"),
                    "企業名": item.get("company"),
                    "得点": item.get("score")
                })
        if all_overall:
            pd.DataFrame(all_overall).to_excel(writer, sheet_name="総合ランキング", index=False)

        # === シート3: 経年比較（ピボット） ===
        companies = set()
        for year_data in overall_data.values():
            for item in year_data:
                companies.add(item.get("company", ""))

        pivot_data = []
        for company in sorted(companies):
            if not company:
                continue
            row = {"企業名": company}
            for year in sorted(overall_data.keys(), key=_year_sort_key):
                score = None
                rank = None
                for item in overall_data.get(year, []):
                    if item.get("company") == company:
                        score = item.get("score")
                        rank = item.get("rank")
                        break
                row[f"{year}年_得点"] = score if score is not None else ""
                row[f"{year}年_順位"] = rank if rank is not None else ""
            pivot_data.append(row)
        if pivot_data:
            pd.DataFrame(pivot_data).to_excel(writer, sheet_name="経年比較", index=False)

        # === シート4: 連続1位記録 ===
        consecutive = records.get("consecutive_wins", [])
        if consecutive:
            df_cons = pd.DataFrame([
                {
                    "企業名": r["company"],
                    "連続年数": r["years"],
                    "開始年": r["start_year"],
                    "終了年": r["end_year"],
                    "継続中": "○" if r.get("is_current") else ""
                }
                for r in consecutive
            ])
            df_cons.to_excel(writer, sheet_name="連続1位記録", index=False)

        # === シート5: 1位獲得回数 ===
        most_wins = records.get("most_wins", [])
        if most_wins:
            df_wins = pd.DataFrame([
                {
                    "企業名": r["company"],
                    "1位回数": r["wins"],
                    "総年数": r["total_years"],
                    "獲得率": f"{r['wins']/r['total_years']*100:.1f}%" if r['total_years'] > 0 else "0.0%",
                    "獲得年": ", ".join(map(str, r["years"]))
                }
                for r in most_wins
            ])
            df_wins.to_excel(writer, sheet_name="1位獲得回数", index=False)

        # === シート6: 過去最高得点 ===
        highest = records.get("highest_scores", [])
        if highest:
            df_high = pd.DataFrame([
                {
                    "順位": i,
                    "企業名": r["company"],
                    "得点": r["score"],
                    "年度": r["year"],
                    "その年の順位": r["rank"]
                }
                for i, r in enumerate(highest[:20], 1)
            ])
            df_high.to_excel(writer, sheet_name="過去最高得点", index=False)

        # === シート7〜: 評価項目別 ===
        for item_name, year_data in item_data.items():
            if isinstance(year_data, dict):
                item_rows = []
                for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                    for item in year_data.get(year, []):
                        item_rows.append({
                            "年度": year,
                            "順位": item.get("rank"),
                            "企業名": item.get("company"),
                            "得点": item.get("score")
                        })
                if item_rows:
                    sheet_name = f"項目_{item_name[:20]}"
                    sheet_name = sheet_name.replace("/", "_").replace("\\", "_")[:31]
                    pd.DataFrame(item_rows).to_excel(writer, sheet_name=sheet_name, index=False)

        # === 部門別 ===
        for dept_name, year_data in dept_data.items():
            if isinstance(year_data, dict):
                dept_rows = []
                for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                    for item in year_data.get(year, []):
                        dept_rows.append({
                            "年度": year,
                            "順位": item.get("rank"),
                            "企業名": item.get("company"),
                            "得点": item.get("score")
                        })
                if dept_rows:
                    sheet_name = f"部門_{dept_name[:20]}"
                    sheet_name = sheet_name.replace("/", "_").replace("\\", "_")[:31]
                    pd.DataFrame(dept_rows).to_excel(writer, sheet_name=sheet_name, index=False)

        # === 参考資料（URL）シート ===
        if used_urls:
            url_rows = []
            for item in used_urls.get("overall", []):
                url_rows.append({
                    "カテゴリ": "総合ランキング",
                    "年度/項目": item.get("year", ""),
                    "URL": item.get("url", ""),
                    "ステータス": "成功" if item.get("status") == "success" else "失敗"
                })
            for item in used_urls.get("items", []):
                url_rows.append({
                    "カテゴリ": "評価項目別",
                    "年度/項目": item.get("name", ""),
                    "URL": item.get("url", ""),
                    "ステータス": "成功" if item.get("status") == "success" else "失敗"
                })
            for item in used_urls.get("departments", []):
                url_rows.append({
                    "カテゴリ": "部門別",
                    "年度/項目": item.get("name", ""),
                    "URL": item.get("url", ""),
                    "ステータス": "成功" if item.get("status") == "success" else "失敗"
                })
            if url_rows:
                pd.DataFrame(url_rows).to_excel(writer, sheet_name="参考資料URL", index=False)

    output.seek(0)
    return output.getvalue()


def parse_uploaded_excel(uploaded_file, specified_year=None):
    """アップロードされたExcelファイルを解析してデータを抽出

    対応フォーマット:
    1. 標準フォーマット（年度列あり）
    2. オリコン内部フォーマット（年度なし、ヘッダー行が3行目以降）
    3. 評価項目シート（1列目に評価項目名）
    4. 部門別シート（ヘッダー行の上にカテゴリ名）

    Args:
        uploaded_file: アップロードされたファイル
        specified_year: ユーザーが指定した年度（Noneの場合はファイル名から推測）
    """
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names

        overall_data = {}
        item_data = {}
        dept_data = {}

        # 年度を決定（ユーザー指定 > ファイル名から推測 > 現在年）
        if specified_year:
            inferred_year = specified_year
        else:
            filename = uploaded_file.name if hasattr(uploaded_file, 'name') else ""
            import re
            year_match = re.search(r'20\d{2}', filename)
            if year_match:
                inferred_year = int(year_match.group())
            else:
                inferred_year = datetime.now().year

        for sheet_name in sheet_names:
            # スキップするシート
            skip_sheets = ['継続利用意向', '推奨意向', '作業用']
            if any(skip in sheet_name for skip in skip_sheets):
                continue

            # まずヘッダーなしで読み込んでヘッダー行を検出
            df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None)

            # ヘッダー行を検出（"順位"を含み、"企業"または"ランキング"を含む行）
            # ID列は必須条件から除外（汎用性向上）
            header_row = None
            category_name = None  # 部門別シートのカテゴリ名
            for idx, row in df_raw.iterrows():
                row_str = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if '順位' in row_str and ('企業' in row_str or 'ランキング' in row_str or '会社' in row_str):
                    header_row = idx
                    # ヘッダー行の上にカテゴリ名がある場合（部門別シート）
                    # 構造: Row0=FX, Row1=シート名, Row2=カテゴリ名, Row3=n＝, Row4=ヘッダー
                    if idx >= 2:
                        # Row2を優先的に確認（通常カテゴリ名がある場所）
                        for cat_idx in [2, idx - 2, idx - 1]:
                            if cat_idx < 0 or cat_idx >= idx:
                                continue
                            cat_row = df_raw.iloc[cat_idx]
                            cat_val = cat_row.iloc[0] if pd.notna(cat_row.iloc[0]) else None
                            if cat_val:
                                cat_str = str(cat_val)
                                # 除外条件: n＝、シート名、FX、nan
                                if (cat_str not in ['nan', 'NaN', sheet_name, 'FX', '評価項目']
                                    and 'n＝' not in cat_str
                                    and 'n=' not in cat_str
                                    and cat_str != sheet_name.replace('別', '')):
                                    category_name = cat_str
                                    break
                    break

            if header_row is None:
                continue

            # ヘッダー行を指定して読み込み
            df = pd.read_excel(xl, sheet_name=sheet_name, header=header_row)

            # 年度列があるかチェック（誤検出を防ぐため厳密に）
            year_col = None
            year_exclude_patterns = ['回答者数', '最新年', '回答', '者数', '前年', '昨年', '今年', '毎年']
            for col in df.columns:
                col_str = str(col)
                if any(pattern in col_str for pattern in year_exclude_patterns):
                    continue
                if col_str == '年度' or '年度' in col_str:
                    year_col = col
                    break
                elif col_str == '年':
                    year_col = col
                    break
                elif len(col_str) == 5 and col_str.endswith('年') and col_str[:4].isdigit():
                    year_col = col
                    break

            # 企業名列を探す
            company_col = None
            for col in df.columns:
                col_str = str(col)
                if 'ランキング対象企業' in col_str or '企業名' in col_str:
                    company_col = col
                    break
            if company_col is None:
                for col in df.columns:
                    col_str = str(col)
                    if '企業' in col_str or '会社' in col_str:
                        company_col = col
                        break

            # 順位列を探す（"順位"という列名を優先）
            rank_col = None
            for col in df.columns:
                col_str = str(col)
                if col_str == '順位':
                    rank_col = col
                    break
            if rank_col is None:
                for col in df.columns:
                    col_str = str(col)
                    if '順位' in col_str:
                        rank_col = col
                        break

            # 得点列を探す（優先順位: スコア > 合計 > 得点）
            score_col = None
            for col in df.columns:
                col_str = str(col)
                if col_str == 'スコア' or 'スコア' in col_str:
                    score_col = col
                    break
            if score_col is None:
                for col in df.columns:
                    if str(col) == '合計':
                        score_col = col
                        break
            if score_col is None:
                for col in df.columns:
                    col_str = str(col)
                    if '得点' in col_str or '点数' in col_str:
                        score_col = col
                        break

            # 評価項目列を探す（1列目が評価項目名の場合）
            eval_item_col = None
            first_col = df.columns[0] if len(df.columns) > 0 else None
            first_col_str = str(first_col) if first_col is not None else ""

            # 評価項目シートの判定: 1列目が順位/IDでなく、評価項目名っぽい場合
            if first_col_str not in ['順位', 'ID', '年度', 'rank', ''] and first_col_str == '評価項目':
                eval_item_col = first_col
            elif '評価項目' in sheet_name and first_col_str not in ['順位', 'ID', '年度', 'rank', '']:
                eval_item_col = first_col

            if company_col and (rank_col or score_col):
                for _, row in df.iterrows():
                    # 年度の取得（v7.9: "2024年"形式に対応）
                    if year_col and pd.notna(row.get(year_col)):
                        try:
                            year_str = str(row[year_col]).replace('年', '').strip()
                            year = int(year_str)
                            # 動的年度範囲: 2000年から現在年+5年まで
                            current_year = datetime.now().year
                            max_year = current_year + 5
                            if year < 2000 or year > max_year:
                                year = inferred_year
                        except (ValueError, TypeError):
                            year = inferred_year
                    else:
                        year = inferred_year

                    # 企業名の取得
                    company = str(row[company_col]) if pd.notna(row.get(company_col)) else ""
                    if not company or company.lower() in ['nan', 'none', '']:
                        continue

                    # 順位の取得
                    try:
                        rank_val = row.get(rank_col) if rank_col else None
                        rank = int(rank_val) if rank_val is not None and pd.notna(rank_val) else None
                    except (ValueError, TypeError):
                        rank = None

                    # 得点の取得
                    try:
                        score_val = row.get(score_col) if score_col else None
                        score = float(score_val) if score_val is not None and pd.notna(score_val) else None
                    except (ValueError, TypeError):
                        score = None

                    # 評価項目名の取得
                    eval_item_name = None
                    if eval_item_col:
                        try:
                            val = row.get(eval_item_col)
                            eval_item_name = str(val) if pd.notna(val) and str(val) not in ['nan', 'None', '評価項目'] else None
                        except:
                            eval_item_name = None

                    # シート種別に応じてデータを格納
                    # 1. 総合ランキング系
                    if '総合' in sheet_name or '対象企業' in sheet_name:
                        if year not in overall_data:
                            overall_data[year] = []
                        overall_data[year].append({
                            "rank": rank,
                            "company": company,
                            "score": score
                        })

                    # 2. 評価項目シート（1列目に項目名がある）
                    elif eval_item_name and ('評価項目' in sheet_name or eval_item_col):
                        if eval_item_name not in item_data:
                            item_data[eval_item_name] = {}
                        if year not in item_data[eval_item_name]:
                            item_data[eval_item_name][year] = []
                        item_data[eval_item_name][year].append({
                            "rank": rank,
                            "company": company,
                            "score": score
                        })

                    # 3. 部門別シート（業態別、投資スタイル別、利用チャート別、レベル別、サポート別、部門_XXX）
                    # v7.10: '部門'を追加（エクスポートされたExcelの"部門_XXX"シートを認識）
                    elif any(x in sheet_name for x in ['業態', '投資スタイル', '利用チャート', 'チャート', 'レベル', 'サポート', '別', '部門']):
                        # カテゴリ名があればそれを使用、なければシート名から抽出
                        # v7.10: "部門_XXX"形式に対応（例: "部門_男性" → "男性"）
                        if category_name:
                            dept_name = category_name
                        elif sheet_name.startswith('部門_'):
                            dept_name = sheet_name[3:]  # "部門_" を除去
                        else:
                            dept_name = sheet_name.replace('別', '')
                        if dept_name not in dept_data:
                            dept_data[dept_name] = {}
                        if year not in dept_data[dept_name]:
                            dept_data[dept_name][year] = []
                        dept_data[dept_name][year].append({
                            "rank": rank,
                            "company": company,
                            "score": score
                        })

        return overall_data, item_data, dept_data, None
    except Exception as e:
        import traceback
        # セキュリティ対策: トレースバック詳細はログのみに出力（ユーザーには非表示）
        logger.error(f"Excel解析エラー: {str(e)}\n{traceback.format_exc()}")
        # ユーザーには一般的なエラーメッセージのみ表示
        return None, None, None, f"Excelファイルの解析中にエラーが発生しました: {str(e)}"


def merge_data(uploaded_data, scraped_data):
    """アップロードデータとスクレイピングデータを統合（アップロードデータ優先）"""
    merged = {}

    # スクレイピングデータをベースに
    for year, data in scraped_data.items():
        merged[year] = data

    # アップロードデータで上書き（優先）
    for year, data in uploaded_data.items():
        merged[year] = data

    return merged


def merge_nested_data(uploaded_data, scraped_data):
    """評価項目別・部門別データを統合"""
    merged = {}

    # スクレイピングデータをベースに
    for key, year_data in scraped_data.items():
        if key not in merged:
            merged[key] = {}
        if isinstance(year_data, dict):
            for year, data in year_data.items():
                merged[key][year] = data

    # アップロードデータで上書き（優先）
    for key, year_data in uploaded_data.items():
        if key not in merged:
            merged[key] = {}
        if isinstance(year_data, dict):
            for year, data in year_data.items():
                merged[key][year] = data

    return merged


def detect_name_changes(used_urls, category="items"):
    """
    同じslug（item_slug/dept_path）でページタイトルが異なるものを検出し、名称変更履歴を返す

    Args:
        used_urls: スクレイパーから取得したURL情報（page_title, item_slug/dept_path, year を含む）
        category: "items"（評価項目）または "departments"（部門）

    Returns:
        dict: {
            現在の名称（リンクテキスト）: {
                "changes": [{from_name, to_name, change_year}, ...],
                "latest_name": ページタイトルから取得した最新名称
            }
        }
    """
    if not used_urls:
        return {}

    url_items = used_urls.get(category, [])
    if not url_items:
        return {}

    # slug（item_slug または dept_path）でグループ化
    slug_key = "item_slug" if category == "items" else "dept_path"

    # slug → [(page_title, year, link_name), ...] のマッピング
    slug_map = {}
    for item in url_items:
        status = item.get("status", "")
        if status != "success":
            continue

        slug = item.get(slug_key)
        page_title = item.get("page_title")
        year = item.get("year")
        link_name = item.get("name", "").replace(f"({year}年)", "").strip() if year else item.get("name", "")

        if not slug or not year:
            continue

        if slug not in slug_map:
            slug_map[slug] = []
        slug_map[slug].append({
            "page_title": page_title,
            "year": year,
            "link_name": link_name
        })

    # 名称変更を検出
    name_changes = {}
    for slug, items in slug_map.items():
        # 年度でソート（古い順）
        items_sorted = sorted(items, key=lambda x: _year_sort_key(x["year"]))

        if len(items_sorted) < 2:
            continue

        # page_titleの変化を追跡（Noneは除外）
        unique_titles = []
        for item in items_sorted:
            title = item["page_title"]
            year = item["year"]
            if title is None:
                continue
            if not unique_titles or unique_titles[-1][0] != title:
                unique_titles.append((title, year))

        # 最新のリンク名称をキーとして使用
        latest_link_name = items_sorted[-1]["link_name"]
        latest_page_title = None
        for item in reversed(items_sorted):
            if item["page_title"]:
                latest_page_title = item["page_title"]
                break

        if len(unique_titles) > 1:
            # 名称変更があった場合
            changes = []
            for i, (title, year) in enumerate(unique_titles[:-1]):
                next_title = unique_titles[i + 1][0]
                next_year = unique_titles[i + 1][1]
                changes.append({
                    "from_name": title,
                    "to_name": next_title,
                    "change_year": next_year
                })
            name_changes[latest_link_name] = {
                "changes": changes,
                "latest_name": latest_page_title
            }
        elif latest_page_title and latest_page_title != latest_link_name:
            # 名称変更はないが、最新名称がリンク名と異なる場合も記録
            name_changes[latest_link_name] = {
                "changes": [],
                "latest_name": latest_page_title
            }

    return name_changes


def display_historical_summary(records, prefix=""):
    """歴代記録・連続記録のサマリーを表示"""
    if not records:
        return

    summary = records.get("summary", {})
    if summary:
        col1, col2, col3 = st.columns(3)
        with col1:
            if summary.get("max_consecutive"):
                mc = summary["max_consecutive"]
                st.metric(
                    f"{prefix}🥇 最長連続1位",
                    f"{mc['company']}",
                    f"{mc['years']}年連続 ({mc['start_year']}〜{mc['end_year']})"
                )
        with col2:
            if summary.get("all_time_high"):
                ath = summary["all_time_high"]
                st.metric(
                    f"{prefix}📈 過去最高得点",
                    f"{ath['score']}点",
                    f"{ath['company']} ({ath['year']}年)"
                )
        with col3:
            if summary.get("most_wins"):
                mw = summary["most_wins"]
                st.metric(
                    f"{prefix}🏆 最多1位獲得",
                    f"{mw['company']}",
                    f"{mw['wins']}回 / {mw['total_years']}年中"
                )


def display_consecutive_wins_compact(records):
    """連続1位記録をコンパクトに表示"""
    consecutive = records.get("consecutive_wins", [])
    # 複数回1位を獲得したもののみ対象（1年だけの受賞は除外）
    consecutive_filtered = [r for r in consecutive if r.get("years", 0) >= 2]
    if consecutive_filtered:
        st.markdown("**🥇 連続1位記録（上位10件）**")
        cons_df = pd.DataFrame([
            {
                "企業名": r["company"],
                "連続年数": f"{r['years']}年",
                "期間": f"{r['start_year']}〜{r['end_year']}",
                "継続中": "✅" if r.get("is_current") else ""
            }
            for r in consecutive_filtered[:10]
        ])
        st.dataframe(cons_df, use_container_width=True, hide_index=True)


# ページ設定
st.set_page_config(
    page_title="オリコン顧客満足度®調査 TOPICSサポートシステム",
    page_icon="📰",
    layout="wide"
)

# タイトル
st.title("📰 オリコン顧客満足度®調査 TOPICSサポートシステム")
st.markdown("オリコン顧客満足度調査の経年結果を調査。連続記録や1位獲得回数の参照に活用いただけます。")
st.warning("⚠️ **注意事項**: 情報の正確性は担当者が必ず確認してください。")

# サイドバー
st.sidebar.header("⚙️ 設定")

# ランキング選択（約200カテゴリ対応）
ranking_options = {
    # ========================================
    # 保険
    # ========================================
    "自動車保険（ダイレクト型）": "_insurance",
    "自動車保険（代理店型）": "_insurance@type02",
    "自動車保険（FP推奨）": "_insurance@type03",
    "バイク保険（ダイレクト型）": "_bike",
    "バイク保険（代理店型）": "_bike@type02",
    "自転車保険": "bicycle-insurance",
    "火災保険": "fire-insurance",
    "海外旅行保険": "travel-insurance",
    "ペット保険": "_pet",
    "生命保険": "life-insurance",
    "医療保険": "medical_insurance",
    "がん保険": "cancer-insurance",
    "学資保険": "educational-insurance",
    "保険相談ショップ": "_hokenshop",
    # ========================================
    # 金融・投資
    # ========================================
    "ネット証券（顧客満足度）": "_certificate",
    "ネット証券（FP評価）": "_certificate@type02",
    "NISA（証券会社）": "_nisa",
    "NISA（銀行）": "_nisa@type02",
    "NISA（FP評価）": "_nisa@type03",
    "iDeCo（証券会社）": "ideco",
    "iDeCo（FP評価）": "ideco@type02",
    "ネット銀行": "_netbank/bank",
    "インターネットバンキング": "_netbank/banking",
    "住宅ローン": "_housingloan",
    "住宅ローン（FP評価）": "_housingloan@type02",
    "外貨預金": "foreign-currency-deposits",
    "FX（顧客満足度）": "_fx",
    "FX（FP評価）": "_fx@type02",
    "カードローン（銀行系）": "card-loan",
    "カードローン（ノンバンク）": "card-loan/nonbank",
    "クレジットカード（一般）": "credit-card/general",
    "クレジットカード（年会費無料）": "credit-card/free-annual",
    "クレジットカード（ゴールド）": "credit-card/gold-card",
    "キャッシュレス決済アプリ": "smartphone-payment",
    "暗号資産（現物取引）": "cryptocurrency/cash-transaction",
    "暗号資産（証拠金取引）": "cryptocurrency/margin-transaction",
    "ロボアドバイザー": "robo-advisor",
    # ========================================
    # 住宅・不動産
    # ========================================
    "不動産仲介 売却（マンション）": "estate-agency-sell/mansion",
    "不動産仲介 売却（戸建て）": "estate-agency-sell/kodate",
    "不動産仲介 売却（土地）": "estate-agency-sell/land",
    "不動産仲介 購入（マンション）": "estate-agency-buy/mansion",
    "不動産仲介 購入（戸建て）": "estate-agency-buy/kodate",
    "分譲マンション管理会社（首都圏）": "mansion-maintenance/syutoken",
    "分譲マンション管理会社（東海）": "mansion-maintenance/tokai",
    "分譲マンション管理会社（近畿）": "mansion-maintenance/kinki",
    "分譲マンション管理会社（九州）": "mansion-maintenance/kyusyu",
    "マンション大規模修繕": "mansion-large-repair",
    "不動産仲介 賃貸": "rental-housing",
    "賃貸情報サイト": "rental-housing/website",
    "賃貸マンション": "rental-condominiums",
    "リフォーム（フルリフォーム）": "_reform/large",
    "リフォーム（戸建て）": "_reform/kodate",
    "リフォーム（マンション）": "_reform/mansion",
    "新築分譲マンション（首都圏）": "new-condominiums/syutoken",
    "新築分譲マンション（東海）": "new-condominiums/tokai",
    "新築分譲マンション（近畿）": "new-condominiums/kinki",
    "新築分譲マンション（九州）": "new-condominiums/kyusyu",
    "ハウスメーカー（注文住宅）": "house-maker",
    "建売住宅PB（東北）": "new-ready-built-house/powerbuilder/tohoku",
    "建売住宅PB（北関東）": "new-ready-built-house/powerbuilder/kitakanto",
    "建売住宅PB（首都圏）": "new-ready-built-house/powerbuilder/syutoken",
    "建売住宅PB（東海）": "new-ready-built-house/powerbuilder/tokai",
    "建売住宅PB（近畿）": "new-ready-built-house/powerbuilder/kinki",
    "建売住宅PB（九州）": "new-ready-built-house/powerbuilder/kyusyu",
    # ========================================
    # 生活サービス
    # ========================================
    "ウォーターサーバー": "_waterserver",
    "ウォーターサーバー（浄水型）": "_waterserver/purifier",
    "家事代行サービス": "housekeeping",
    "ハウスクリーニング": "house-cleaning",
    "コインランドリー": "laundromat",
    "食材宅配（首都圏）": "food-delivery/syutoken",
    "食材宅配（東海）": "food-delivery/tokai",
    "食材宅配（近畿）": "food-delivery/kinki",
    "ミールキット（首都圏）": "food-delivery/meal-kit/syutoken",
    "ミールキット（東海）": "food-delivery/meal-kit/tokai",
    "ミールキット（近畿）": "food-delivery/meal-kit/kinki",
    "ネットスーパー": "net-super",
    "フードデリバリーサービス": "food-delivery-service",
    "ふるさと納税サイト": "hometown-tax-website",
    "トランクルーム（レンタル収納）": "trunk-room/rental-storage-space",
    "トランクルーム（コンテナ）": "trunk-room/container",
    "トランクルーム（宅配型）": "trunk-room/delivery",
    "引越し会社": "_move",
    "カーシェアリング": "carsharing",
    "レンタカー": "rent-a-car",
    "格安レンタカー": "rent-a-car/reasonable",
    "車買取会社": "_carbuyer",
    "中古車情報サイト": "used-car-sell",
    "バイク販売店": "bike-sell",
    "車検": "vehicle-inspection",
    "カーメンテナンスサービス": "car-maintenance",
    "カフェ": "cafe",
    "定額制動画配信サービス": "svod",
    "動画配信（ジャンル別：映画）": "svod/genre",
    "動画配信（ジャンル別：洋画）": "svod/genre/foreign-film",
    "動画配信（ジャンル別：国内ドラマ）": "svod/genre/japanese-drama",
    "動画配信（ジャンル別：海外ドラマ）": "svod/genre/foreign-drama",
    "動画配信（ジャンル別：韓国ドラマ）": "svod/genre/korean-drama",
    "動画配信（ジャンル別：アニメ）": "svod/genre/anime",
    "動画配信（ジャンル別：バラエティ）": "svod/genre/variety",
    "動画配信（ジャンル別：ドキュメンタリー）": "svod/genre/documentary",
    "動画配信（ジャンル別：スポーツ）": "svod/genre/sports",
    "動画配信（ジャンル別：オリジナル）": "svod/genre/original",
    "動画配信（ジャンル別：キッズ）": "svod/genre/kids",
    "子ども写真スタジオ": "kids-photo-studio",
    "電子書籍サービス": "ebook",
    "電子コミックサービス": "manga-apps",
    "マンガアプリ（オリジナル）": "manga-apps/original",
    "マンガアプリ（出版社）": "manga-apps/publisher",
    "ブランド品買取（店舗）": "brand-sell",
    "電力会社（小売）": "electricity/retailing",
    "子ども見守りGPS": "child-gps",
    # ========================================
    # 通信
    # ========================================
    "携帯キャリア": "mobile-carrier",
    "キャリア格安ブランド": "mobile-carrier/reasonable",
    "格安SIM": "mvno",
    "格安SIM（SIMのみ）": "mvno/sim",
    "格安スマホ": "mvno/sp",
    "プロバイダ": "_internet",
    "プロバイダ（北海道）": "_internet/hokkaido",
    "プロバイダ（東北）": "_internet/tohoku",
    "プロバイダ（関東）": "_internet/kanto",
    "プロバイダ（甲信越・北陸）": "_internet/koshinetsu-hokuriku",
    "プロバイダ（東海）": "_internet/tokai",
    "プロバイダ（近畿）": "_internet/kinki",
    "プロバイダ（中国）": "_internet/chugoku",
    "プロバイダ（四国）": "_internet/shikoku",
    "プロバイダ（九州・沖縄）": "_internet/kyusyu-okinawa",
    # ========================================
    # 教育（塾・受験）※juken.oricon.co.jp
    # ========================================
    "大学受験 塾・予備校（首都圏）": "_college/syutoken",
    "大学受験 塾・予備校（東海）": "_college/tokai",
    "大学受験 塾・予備校（近畿）": "_college/kinki",
    "大学受験 個別指導塾（首都圏）": "college-individual/syutoken",
    "大学受験 個別指導塾（東海）": "college-individual/tokai",
    "大学受験 個別指導塾（近畿）": "college-individual/kinki",
    "大学受験 難関大学特化型（首都圏）": "_college/elite",
    "大学受験 映像授業": "college-video",
    "高校受験 塾（北海道）": "highschool/hokkaido",
    "高校受験 塾（東北）": "highschool/tohoku",
    "高校受験 塾（北関東）": "highschool/kitakanto",
    "高校受験 塾（首都圏）": "highschool/syutoken",
    "高校受験 塾（甲信越・北陸）": "highschool/koshinetsu-hokuriku",
    "高校受験 塾（東海）": "highschool/tokai",
    "高校受験 塾（近畿）": "highschool/kinki",
    "高校受験 塾（中国・四国）": "highschool/chugoku-shikoku",
    "高校受験 塾（九州・沖縄）": "highschool/kyusyu",
    "高校受験 個別指導塾（北海道）": "highschool-individual/hokkaido",
    "高校受験 個別指導塾（東北）": "highschool-individual/tohoku",
    "高校受験 個別指導塾（北関東）": "highschool-individual/kitakanto",
    "高校受験 個別指導塾（首都圏）": "highschool-individual/syutoken",
    "高校受験 個別指導塾（甲信越・北陸）": "highschool-individual/koshinetsu-hokuriku",
    "高校受験 個別指導塾（東海）": "highschool-individual/tokai",
    "高校受験 個別指導塾（近畿）": "highschool-individual/kinki",
    "高校受験 個別指導塾（中国・四国）": "highschool-individual/chugoku-shikoku",
    "高校受験 個別指導塾（九州・沖縄）": "highschool-individual/kyusyu",
    "中学受験 塾（首都圏）": "_junior/syutoken",
    "中学受験 塾（東海）": "_junior/tokai",
    "中学受験 塾（近畿）": "_junior/kinki",
    "中学受験 個別指導塾": "_junior/individual",
    "公立中高一貫校対策 塾（首都圏）": "public-junior/syutoken",
    "公立中高一貫校対策 塾（東海）": "public-junior/tokai",
    "公立中高一貫校対策 塾（近畿）": "public-junior/kinki",
    # ========================================
    # 教育（通信・英語・資格）※juken.oricon.co.jp
    # ========================================
    "通信教育（高校生）": "online-study/highschool",
    "通信教育（中学生）": "online-study/junior-hs",
    "通信教育（小学生）": "online-study/elementary",
    "家庭教師": "tutor",
    "補習塾": "supplementary-school",
    "幼児・小学生 学習教室": "kids-school/intellectual",
    "子ども英語教室（幼児）": "kids-english/preschooler",
    "子ども英語教室（小学生）": "kids-english/grade-schooler",
    "英会話スクール": "_english",
    "オンライン英会話": "online-english",
    "通信講座（FP）": "cc/fp",
    "通信講座（医療事務）": "cc/mo",
    "通信講座（宅建）": "cc/takken",
    "通信講座（簿記）": "cc/bookkeeping",
    "通信講座（TOEIC）": "cc/toeic",
    "通信講座（社会保険労務士）": "cc/labor-and-social-security",
    "通信講座（ケアマネジャー）": "cc/care-manager",
    "通信講座（公務員）": "cc/public-officer",
    "通信講座（ITパスポート）": "cc/it-certification",
    "資格スクール（FP）": "license/fp",
    "資格スクール（宅建）": "license/takken",
    "資格スクール（簿記）": "license/bookkeeping",
    "資格スクール（社会保険労務士）": "license/labor-and-social-security",
    # ========================================
    # スポーツ・フィットネス
    # ========================================
    "キッズスイミングスクール（幼児）": "kids-swimming/preschooler",
    "キッズスイミングスクール（小学生）": "kids-swimming/grade-schooler",
    "フィットネスクラブ": "_fitness",
    "24時間ジム": "_fitness/24hours",
    "パーソナルトレーニング": "_fitness/service",
    # ========================================
    # 転職・人材 ※career.oricon.co.jp
    # ========================================
    "就活サイト": "new-graduates-hiring-website",
    "逆求人型就活サービス": "reversed-job-offer",
    "アルバイト情報サイト": "arbeit",
    "転職サイト": "job-change",
    "転職サイト（女性）": "job-change_woman",
    "転職スカウトサービス": "job-change_scout",
    "転職エージェント": "_agent",
    "看護師転職": "_agent_nurse",
    "介護転職": "_agent_nursing",
    "ハイクラス・ミドルクラス転職": "_agent_hi-and-middle-class",
    "派遣会社": "_staffing",
    "工場・製造業派遣": "_staffing_manufacture",
    "派遣求人サイト": "temp-staff",
    "求人情報サービス": "employment",
    # ========================================
    # トラベル
    # ========================================
    "旅行予約サイト（国内）": "bargain-hotels-website",
    "旅行予約サイト（海外）": "bargain-airline-website",
    "ツアー比較サイト": "tours-website",
    # ========================================
    # 美容・ウエディング
    # ========================================
    "ブライダルエステ": "_esthe/bridal",
    "フェイシャルエステ": "_esthe/facial",
    "痩身・ボディエステ": "_esthe/slim",
    "サロン検索予約サイト": "salon-website",
    "ハウスウエディング": "wedding-produce",
    "結婚相談所": "_marriage",
    # ========================================
    # 小売・レジャー
    # ========================================
    "家電量販店": "electronics-retail-store",
    "ドラッグストア": "drug-store",
    "映画館": "movie-theater",
    "カラオケボックス": "karaoke",
    "テーマパーク": "theme-park",
    # ========================================
    # その他
    # ========================================
    "カスタム入力": "custom"
}

# 検索型ランキング選択
all_rankings = list(ranking_options.keys())

search_keyword = st.sidebar.text_input(
    "🔍 ランキングを検索・選択",
    placeholder="例：保険、転職、英会話、塾"
)

# フィルタリング & 選択UI
if search_keyword:
    filtered_rankings = [r for r in all_rankings if search_keyword.lower() in r.lower()]

    if filtered_rankings:
        st.sidebar.caption(f"✅ {len(filtered_rankings)}件ヒット")
        selected_ranking = st.sidebar.radio(
            "選択してください",
            filtered_rankings,
            label_visibility="collapsed"
        )
    else:
        st.sidebar.warning(f"「{search_keyword}」に一致するランキングがありません")
        st.sidebar.caption("キーワードを変えて再検索してください")
        selected_ranking = None
else:
    # 未入力時はカテゴリ別に表示
    st.sidebar.caption(f"📊 全{len(all_rankings)}件 - キーワードで絞り込めます")

    category_options = {
        "保険": [r for r in all_rankings if "保険" in r],
        "金融・投資": [r for r in all_rankings if any(k in r for k in ["証券", "NISA", "iDeCo", "銀行", "ローン", "カード", "FX", "暗号", "ロボ", "外貨", "決済"])],
        "住宅・不動産": [r for r in all_rankings if any(k in r for k in ["不動産", "マンション", "住宅", "リフォーム", "ハウス", "建売"])],
        "生活サービス": [r for r in all_rankings if any(k in r for k in ["ウォーター", "家事", "クリーニング", "食材", "ミール", "スーパー", "デリバリー", "ふるさと", "トランク", "引越", "カー", "車", "バイク", "カフェ", "動画", "写真", "電子", "マンガ", "ブランド", "電力", "見守り", "ランドリー"])],
        "通信": [r for r in all_rankings if any(k in r for k in ["携帯", "キャリア", "SIM", "スマホ", "プロバイダ"])],
        "教育・塾": [r for r in all_rankings if any(k in r for k in ["受験", "塾", "予備校", "指導", "通信教育", "家庭教師", "補習", "学習", "英語", "英会話", "通信講座", "資格", "スイミング"])],
        "転職・人材": [r for r in all_rankings if any(k in r for k in ["就活", "求人", "アルバイト", "転職", "派遣", "エージェント"])],
        "トラベル・美容・その他": [r for r in all_rankings if any(k in r for k in ["旅行", "ツアー", "エステ", "サロン", "ウエディング", "結婚", "家電", "ドラッグ", "映画", "カラオケ", "テーマパーク", "フィットネス", "ジム", "パーソナル"])],
    }

    selected_category = st.sidebar.selectbox(
        "カテゴリを選択",
        ["-- カテゴリを選択 --"] + list(category_options.keys()) + ["カスタム入力"]
    )

    if selected_category == "-- カテゴリを選択 --":
        selected_ranking = None
        st.sidebar.info("カテゴリを選択するか、上の検索ボックスでキーワード検索してください")
    elif selected_category == "カスタム入力":
        selected_ranking = "カスタム入力"
    else:
        category_rankings = category_options.get(selected_category, [])
        if category_rankings:
            selected_ranking = st.sidebar.radio(
                f"{selected_category}のランキング",
                category_rankings,
                label_visibility="collapsed"
            )
        else:
            selected_ranking = None

# ランキング選択の処理
if selected_ranking is None:
    ranking_slug = None
    ranking_name = None
elif selected_ranking == "カスタム入力":
    ranking_slug = st.sidebar.text_input(
        "ランキングのURL名",
        placeholder="例: mobile-carrier"
    )
    ranking_name = st.sidebar.text_input(
        "ランキング名（表示用）",
        placeholder="例: 携帯キャリア"
    )
else:
    ranking_slug = ranking_options[selected_ranking]
    ranking_name = selected_ranking

# 年度選択
# 注意: current_yearはWebスクレイピングの最新年度（オリコンサイトで公開されている最新）
# アップロードデータの年度は別途指定可能
# datetime.now().year を使用して動的に現在年を取得
current_year = datetime.now().year  # Webサイトの最新年度（動的）
start_year = 2006

year_option = st.sidebar.radio(
    "過去データ取得範囲",
    ["直近3年", "直近5年", "全年度（2006年〜）", "カスタム範囲"],
    index=2  # デフォルト: 全年度
)

if year_option == "直近3年":
    year_range = (current_year - 2, current_year)
elif year_option == "直近5年":
    year_range = (current_year - 4, current_year)
elif year_option == "全年度（2006年〜）":
    year_range = (start_year, current_year)
else:
    year_range = st.sidebar.slider(
        "年度範囲を選択",
        min_value=start_year,
        max_value=current_year,
        value=(current_year - 4, current_year)
    )

# セッション状態の初期化
if 'results_data' not in st.session_state:
    st.session_state.results_data = None

# 実行ボタン（過去データ取得範囲の直下に配置）
run_button = st.sidebar.button("🚀 TOPICS出し実行", type="primary", use_container_width=True)

# ファイルアップロード（オプション）- 環境変数で有効化時のみ表示
uploaded_file = None
upload_year = None

if ENABLE_UPLOAD:
    st.sidebar.markdown("---")
    with st.sidebar.expander("📁 最新データのアップロード（オプション・非推奨）", expanded=False):
        st.caption("⚠️ 通常はWebから自動取得されるため、アップロードは不要です。未公開の最新データを含める場合のみ使用してください。")
        uploaded_file = st.file_uploader(
            "最新のランキングExcelをアップロード",
            type=["xlsx", "xls"],
            help="最新のランキング資料をアップロードすると、過去データと統合して分析します",
            key="excel_uploader"
        )

        # アップロードデータの年度指定
        if uploaded_file:
            st.success(f"✅ {uploaded_file.name}")
            upload_year = st.number_input(
                "📅 アップロードデータの年度",
                min_value=2006,
                max_value=2030,
                value=2026,
                help="アップロードしたファイルのデータ年度を指定してください（例: 2026年発表データなら2026）"
            )
            st.info(f"📌 **{upload_year}年**のデータとしてアップロードファイルを使用し、それ以外の年度はWebから取得して統合します")

# 実行ボタン処理
if run_button:

    if not ranking_slug:
        st.error("ランキングのURL名を入力してください")
    else:
        # 実行開始時にセッション状態をリセット（前回結果が残らないように）
        st.session_state.results_data = None

        # プログレスバー
        progress_bar = st.progress(0)
        status_text = st.empty()

        # デバッグログ表示エリア
        debug_expander = st.expander("🔍 デバッグログ", expanded=False)
        debug_logs = []

        def log(message):
            debug_logs.append(message)
            with debug_expander:
                st.text("\n".join(debug_logs))

        try:
            uploaded_overall = {}
            uploaded_item = {}
            uploaded_dept = {}
            uploaded_years = set()

            # Step 1: アップロードファイルがあれば解析
            if uploaded_file:
                status_text.text("📁 アップロードファイルを解析中...")
                progress_bar.progress(10)

                uploaded_overall, uploaded_item, uploaded_dept, error = parse_uploaded_excel(uploaded_file, upload_year)

                if error:
                    st.error(f"ファイル解析エラー: {error}")
                    st.stop()

                if uploaded_overall is None:
                    uploaded_overall = {}
                if uploaded_item is None:
                    uploaded_item = {}
                if uploaded_dept is None:
                    uploaded_dept = {}

                uploaded_years = set(uploaded_overall.keys())
                log(f"[OK] ファイル解析完了: {uploaded_file.name}")
                log(f"  - 総合ランキング: {len(uploaded_overall)}年分")
                log(f"  - 含まれる年度: {sorted(uploaded_years, key=_year_sort_key)}")
                for year, data in uploaded_overall.items():
                    log(f"    {year}年: {len(data)}社")
                    if data:
                        top = data[0]
                        log(f"      1位: {top.get('company')} ({top.get('score')}点)")
                log(f"  - 評価項目別: {len(uploaded_item)}項目")
                for item_name in list(uploaded_item.keys())[:3]:
                    log(f"    [{item_name}]")
                log(f"  - 部門別: {len(uploaded_dept)}部門")
                for dept_name in list(uploaded_dept.keys())[:3]:
                    log(f"    [{dept_name}]")

            # Step 2: Webスクレイピングで過去データを取得（v7.9: with文でリソース管理）
            status_text.text("🌐 Webから過去データを取得中...")
            progress_bar.progress(20)

            log(f"[INFO] スクレイパー初期化: {ranking_slug} ({ranking_name})")

            # スクレイピング対象年度を決定
            # - アップロードデータに含まれる年度は除外
            # - Webサイトの最新年度（current_year）を超える年度は除外
            scrape_years = []
            effective_end_year = min(year_range[1], current_year)  # Webサイトの最新年度を超えない
            for y in range(year_range[0], effective_end_year + 1):
                if y not in uploaded_years:
                    scrape_years.append(y)

            log(f"[INFO] 年度範囲設定: {year_range[0]}〜{year_range[1]}")
            log(f"[INFO] Webサイト最新年度: {current_year}")
            log(f"[INFO] アップロード年度: {sorted(uploaded_years, key=_year_sort_key) if uploaded_years else 'なし'}")

            if scrape_years:
                log(f"[INFO] スクレイピング対象年度: {scrape_years}")
                scrape_range = (min(scrape_years, key=_year_sort_key), max(scrape_years, key=_year_sort_key))
            else:
                log(f"[INFO] アップロードデータで全年度カバー済み、スクレイピングをスキップ")
                scrape_range = None

            scraped_overall = {}
            scraped_item = {}
            scraped_dept = {}
            used_urls = None
            update_date = None

            # with文でスクレイパーを使用（自動的にセッションをクローズ）
            with OriconScraper(ranking_slug, ranking_name) as scraper:
                subpath_info = f" + subpath: {scraper.subpath}" if scraper.subpath else ""
                log(f"[INFO] URL prefix: {scraper.url_prefix}{subpath_info}")

                if scrape_range:
                    status_text.text(f"📊 総合ランキングを取得中... ({scrape_range[0]}年〜{scrape_range[1]}年)")
                    progress_bar.progress(30)

                    scraped_overall = scraper.get_overall_rankings(scrape_range)
                    # アップロード済み年度を除外
                    scraped_overall = {y: d for y, d in scraped_overall.items() if y not in uploaded_years}
                    log(f"[OK] 総合ランキング: {len(scraped_overall)}年分取得")
                    for year, data in scraped_overall.items():
                        log(f"  - {year}年: {len(data)}社")
                    progress_bar.progress(45)

                    status_text.text(f"📋 評価項目別データを取得中...")
                    scraped_item = scraper.get_evaluation_items(scrape_range)
                    log(f"[OK] 評価項目別: {len(scraped_item)}項目")
                    progress_bar.progress(60)

                    status_text.text(f"🏷️ 部門別データを取得中...")
                    scraped_dept = scraper.get_departments(scrape_range)
                    log(f"[OK] 部門別: {len(scraped_dept)}部門")
                    progress_bar.progress(70)

                    used_urls = scraper.used_urls

                # 更新日を取得（推奨TOPICSタブで使用）
                update_date = scraper.get_update_date()

            # Step 3: データ統合
            status_text.text("🔄 データを統合中...")
            progress_bar.progress(75)

            overall_data = merge_data(uploaded_overall, scraped_overall)
            item_data = merge_nested_data(uploaded_item, scraped_item)
            dept_data = merge_nested_data(uploaded_dept, scraped_dept)

            log(f"[OK] データ統合完了")
            log(f"  - 総合ランキング: {len(overall_data)}年分（統合後）")
            log(f"    └ アップロード: {len(uploaded_overall)}年分")
            log(f"    └ スクレイピング: {len(scraped_overall)}年分")

            # Step 4: 分析実行（v5.8: 部門別データも渡す）
            status_text.text("🔍 TOPICS分析中...")
            analyzer = TopicsAnalyzer(overall_data, item_data, ranking_name, dept_data)
            topics = analyzer.analyze()
            progress_bar.progress(85)

            # Step 5: 歴代記録・得点推移分析
            status_text.text("📈 歴代記録・得点推移を分析中...")
            historical_analyzer = HistoricalAnalyzer(overall_data, item_data, dept_data, ranking_name)
            historical_data = historical_analyzer.analyze_all()
            # 評価項目別・部門別の1位獲得回数を計算
            item_most_wins = historical_analyzer.calc_item_most_wins()
            dept_most_wins = historical_analyzer.calc_dept_most_wins()
            progress_bar.progress(95)

            # 完了
            status_text.text("✅ 完了!")
            progress_bar.progress(100)

            # Step 6: 名称変更を検出
            item_name_changes = detect_name_changes(used_urls, "items") if used_urls else {}
            dept_name_changes = detect_name_changes(used_urls, "departments") if used_urls else {}

            # セッション状態に結果を保存
            st.session_state.results_data = {
                'ranking_name': ranking_name,
                'overall_data': overall_data,
                'item_data': item_data,
                'dept_data': dept_data,
                'historical_data': historical_data,
                'topics': topics,
                'used_urls': used_urls,
                'uploaded_years': list(uploaded_years),
                'scraped_years': list(scraped_overall.keys()) if scraped_overall else [],
                'item_most_wins': item_most_wins,
                'dept_most_wins': dept_most_wins,
                'item_name_changes': item_name_changes,
                'dept_name_changes': dept_name_changes,
                'update_date': update_date  # 調査概要の更新日（年, 月）
            }

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            logger.error(f"処理エラー: {str(e)}\n{error_detail}")
            st.error(f"エラーが発生しました。入力データやネットワーク接続を確認してください。")
            # デバッグ用: エラー詳細を折りたたみ表示（v7.9: 環境変数で制御）
            # 本番環境ではSHOW_DEBUG_INFO=falseに設定してセキュリティを向上
            if os.environ.get("SHOW_DEBUG_INFO", "false").lower() == "true":
                with st.expander("🔍 エラー詳細（開発者向け）", expanded=False):
                    st.code(error_detail, language="python")

# 結果表示（セッション状態から）
if st.session_state.results_data:
    data = st.session_state.results_data
    ranking_name = data['ranking_name']
    overall_data = data['overall_data']
    item_data = data['item_data']
    dept_data = data['dept_data']
    historical_data = data['historical_data']
    topics = data['topics']
    used_urls = data.get('used_urls')
    uploaded_years = data.get('uploaded_years', [])
    scraped_years = data.get('scraped_years', [])
    item_most_wins = data.get('item_most_wins', {})
    dept_most_wins = data.get('dept_most_wins', {})
    item_name_changes = data.get('item_name_changes', {})
    dept_name_changes = data.get('dept_name_changes', {})
    update_date = data.get('update_date')  # (year, month) のタプル

    # 結果表示
    st.success(f"✅ {ranking_name}のTOPICS出しが完了しました")

    # データソース情報
    if uploaded_years or scraped_years:
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            if uploaded_years:
                st.info(f"📁 **アップロードデータ**: {sorted(uploaded_years, key=_year_sort_key)}年")
        with col_info2:
            if scraped_years:
                st.info(f"🌐 **Webスクレイピング**: {sorted(scraped_years, key=_year_sort_key)}年")

    # Excelダウンロードボタン（大きく目立つように）
    st.markdown("---")
    excel_data = create_excel_export(
        ranking_name,
        overall_data,
        item_data,
        dept_data,
        historical_data,
        used_urls
    )

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.download_button(
            label="📥 全データをExcelでダウンロード",
            data=excel_data,
            file_name=f"{ranking_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="excel_download_main"
        )
    st.markdown("---")

    # 最新年度を取得
    latest_year = max(overall_data.keys(), key=_year_sort_key) if overall_data else None
    # 更新日から年月を取得（調査概要の更新日ベース、取得できない場合は現在日時を使用）
    if update_date:
        update_year, update_month = update_date
    else:
        update_year = latest_year if latest_year else datetime.now().year
        update_month = datetime.now().month

    # タブで結果表示（新しい構成）
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        f"⭐ 推奨TOPICS（{update_year}年{update_month}月時点）" if update_year else "⭐ 推奨TOPICS",
        "🏆 歴代記録・得点推移",
        "📊 総合ランキング",
        "📋 評価項目別",
        "🏷️ 部門別",
        "📝 プレスリリース作成",
        "📎 参考資料"
    ])

    with tab1:
        st.header(f"⭐ 推奨TOPICS（{update_year}年{update_month}月時点）" if update_year else "⭐ 推奨TOPICS")

        # v5.9: カテゴリ別にTOPICSを分類して表示
        recommended_topics = topics.get("recommended", [])

        # カテゴリ別に分類
        overall_topics = [t for t in recommended_topics if t.get("category") == "総合ランキング"]
        item_topics = [t for t in recommended_topics if t.get("category") == "評価項目別"]
        dept_topics = [t for t in recommended_topics if t.get("category") == "部門別"]

        # カテゴリ未設定のものは総合に分類（後方互換）
        other_categorized = [t for t in recommended_topics if t.get("category") not in ["総合ランキング", "評価項目別", "部門別"]]
        overall_topics.extend(other_categorized)

        # 総合ランキング
        if overall_topics:
            st.subheader("📊 総合ランキング")
            for i, topic in enumerate(overall_topics, 1):
                st.markdown(f"**{i}. {topic['title']}**")
            st.divider()

        # 評価項目別
        if item_topics:
            st.subheader("📋 評価項目別")
            for i, topic in enumerate(item_topics, 1):
                st.markdown(f"**{i}. {topic['title']}**")
            st.divider()

        # 部門別
        if dept_topics:
            st.subheader("🏷️ 部門別")
            for i, topic in enumerate(dept_topics, 1):
                st.markdown(f"**{i}. {topic['title']}**")
            st.divider()

        if topics.get("other"):
            st.subheader("📝 その他のTOPICS候補")
            for topic in topics["other"]:
                st.markdown(f"- {topic}")

        # 見出し案セクション（推奨TOPICSタブ内に統合）
        st.divider()
        st.subheader("🎯 見出し案")
        for i, headline in enumerate(topics.get("headlines", []), 1):
            st.markdown(f"**パターン{i}**: {headline}")

        # コピー用テキスト（カテゴリ別に整理）
        st.subheader("📋 コピー用テキスト")
        copy_lines = [f"【推奨TOPICS（{update_year}年{update_month}月時点）】" if update_year else "【推奨TOPICS】"]

        if overall_topics:
            copy_lines.append("\n■ 総合ランキング")
            copy_lines.extend([f"・{t['title']}" for t in overall_topics])
        if item_topics:
            copy_lines.append("\n■ 評価項目別")
            copy_lines.extend([f"・{t['title']}" for t in item_topics])
        if dept_topics:
            copy_lines.append("\n■ 部門別")
            copy_lines.extend([f"・{t['title']}" for t in dept_topics])

        copy_lines.append("\n【見出し案】")
        copy_lines.extend([f"パターン{i}: {h}" for i, h in enumerate(topics.get("headlines", []), 1)])

        copy_text = "\n".join(copy_lines)
        st.text_area("コピー用", copy_text, height=350, label_visibility="collapsed")

    with tab2:
        st.header("🏆 歴代記録・得点推移")
        records = historical_data.get("historical_records", {})
        trends = historical_data.get("score_trends", {})

        if records:
            # サマリー表示
            display_historical_summary(records)
            st.divider()

            # 2カラムレイアウト
            col_left, col_right = st.columns(2)

            with col_left:
                # 連続1位記録（2年以上のみ）
                st.subheader("🥇 連続1位記録")
                consecutive = records.get("consecutive_wins", [])
                if consecutive:
                    # 2年以上の記録のみ表示
                    consecutive_filtered = [r for r in consecutive if r.get("years", 0) >= 2]
                    if consecutive_filtered:
                        cons_df = pd.DataFrame([
                            {
                                "企業名": r["company"],
                                "連続年数": f"{r['years']}年",
                                "期間": f"{r['start_year']}〜{r['end_year']}",
                                "継続中": "✅" if r.get("is_current") else ""
                            }
                            for r in consecutive_filtered[:10]
                        ])
                        st.dataframe(cons_df, use_container_width=True, hide_index=True)

                # 過去最高得点
                st.subheader("📈 過去最高得点TOP10")
                highest = records.get("highest_scores", [])
                if highest:
                    high_df = pd.DataFrame([
                        {
                            "順位": i,
                            "企業名": r["company"],
                            "得点": f"{r['score']}点",
                            "年度": f"{r['year']}年",
                            "その年の順位": f"{r['rank']}位"
                        }
                        for i, r in enumerate(highest[:10], 1)
                    ])
                    st.dataframe(high_df, use_container_width=True, hide_index=True)

            with col_right:
                # 最多1位獲得
                st.subheader("🏆 1位獲得回数ランキング")
                most_wins = records.get("most_wins", [])
                if most_wins:
                    wins_df = pd.DataFrame([
                        {
                            "企業名": r["company"],
                            "1位回数": f"{r['wins']}回",
                            "獲得年": ", ".join(map(str, r["years"]))
                        }
                        for r in most_wins[:10]
                    ])
                    st.dataframe(wins_df, use_container_width=True, hide_index=True)

                # 年度別1位の推移
                st.subheader("🥇 年度別1位の推移")
                top_by_year = trends.get("top_score_by_year", {})
                if top_by_year:
                    top_df = pd.DataFrame([
                        {
                            "年度": year,
                            "1位企業": top_by_year[year].get("company", "-") if isinstance(top_by_year.get(year), dict) else "-",
                            "得点": f"{top_by_year[year].get('score', '-')}点" if isinstance(top_by_year.get(year), dict) else "-"
                        }
                        for year in sorted(top_by_year.keys(), key=_year_sort_key, reverse=True)
                    ])
                    st.dataframe(top_df, use_container_width=True, hide_index=True)

        st.divider()

        # 得点推移グラフ
        if trends and trends.get("years"):
            years = trends["years"]

            # 年度別平均得点
            st.subheader("📊 年度別平均得点の推移")
            avg_scores = trends.get("average_scores", {})
            if avg_scores:
                avg_df = pd.DataFrame([
                    {"年度": year, "平均得点": score}
                    for year, score in sorted(avg_scores.items(), key=lambda x: _year_sort_key(x[0]))
                ])
                import altair as alt
                # 動的Y軸範囲（折れ線グラフはzero=Falseが有効）
                score_values = list(avg_scores.values())
                y_min = max(0, min(score_values) - 3)
                y_max = max(score_values) + 3
                chart = alt.Chart(avg_df).mark_line(point=True).encode(
                    x=alt.X('年度:O', title='年度'),
                    y=alt.Y('平均得点:Q', title='平均得点', scale=alt.Scale(domain=[y_min, y_max]))
                ).properties(height=300)
                st.altair_chart(chart, use_container_width=True)

            # 上位企業の得点推移
            st.subheader("📈 上位企業の得点推移")
            top_companies = trends.get("top_companies", [])[:5]
            companies_data = trends.get("companies", {})

            if top_companies and companies_data:
                chart_data = []
                for company in top_companies:
                    if company in companies_data:
                        for year in years:
                            score = companies_data[company].get(year, {}).get("score")
                            if score:
                                chart_data.append({
                                    "年度": str(year),
                                    "企業名": company,
                                    "得点": score
                                })

                if chart_data:
                    chart_df = pd.DataFrame(chart_data)
                    # 動的Y軸範囲
                    all_scores = [d["得点"] for d in chart_data]
                    y_min = max(0, min(all_scores) - 3)
                    y_max = max(all_scores) + 3
                    chart = alt.Chart(chart_df).mark_line(point=True).encode(
                        x=alt.X('年度:O', title='年度'),
                        y=alt.Y('得点:Q', title='得点', scale=alt.Scale(domain=[y_min, y_max])),
                        color=alt.Color('企業名:N', title='企業名'),
                        tooltip=['年度', '企業名', '得点']
                    ).properties(height=400)
                    st.altair_chart(chart, use_container_width=True)

            # v7.3: 評価項目別・部門別 平均得点推移を「上位企業の得点推移」の下に移動
            # 評価項目別 平均得点推移（縦棒グラフ）- trendsの有無に関わらず表示
            st.divider()
            st.subheader("📊 評価項目別 平均得点推移")
            if item_data:
                # 各評価項目の年度別平均得点を計算
                item_avg_data = []
                for item_name, year_data in item_data.items():
                    if isinstance(year_data, dict):
                        for year, data in year_data.items():
                            # 0点も有効な値として扱う（Noneのみを除外）
                            scores = [d.get("score") for d in data if d.get("score") is not None]
                            if scores:
                                item_avg_data.append({
                                    "評価項目": item_name[:15],  # 長すぎる項目名を短縮
                                    "年度": str(year),
                                    "平均得点": round(sum(scores) / len(scores), 2)
                                })

                if item_avg_data:
                    item_avg_df = pd.DataFrame(item_avg_data)
                    # 最新3年度に絞る
                    latest_years = sorted(item_avg_df["年度"].unique(), key=_year_sort_key, reverse=True)[:3]
                    item_avg_df = item_avg_df[item_avg_df["年度"].isin(latest_years)]

                    # グループ化縦棒グラフ（年度ごとに横並び）- mark_rectで非0基点
                    import altair as alt
                    all_scores = item_avg_df["平均得点"].tolist()
                    y_min = max(0, min(all_scores) - 5)
                    y_max = max(all_scores) + 2
                    item_avg_df["基点"] = y_min
                    chart = alt.Chart(item_avg_df).mark_rect(width=12).encode(
                        x=alt.X('年度:N', title=None, axis=alt.Axis(labelAngle=0)),
                        y=alt.Y('基点:Q', title='平均得点', scale=alt.Scale(domain=[y_min, y_max])),
                        y2=alt.Y2('平均得点:Q'),
                        color=alt.Color('年度:N', title='年度'),
                        column=alt.Column('評価項目:N', title=None, header=alt.Header(labelOrient='bottom')),
                        tooltip=['評価項目', '年度', '平均得点']
                    ).properties(width=60, height=400)
                    st.altair_chart(chart)
                else:
                    st.info("評価項目別データにスコアが含まれていません")
            else:
                st.info("評価項目別データがありません")

            # 部門別 平均得点推移（縦棒グラフ）- trendsの有無に関わらず表示
            st.subheader("📊 部門別 平均得点推移")
            if dept_data:
                dept_avg_data = []
                for dept_name, year_data in dept_data.items():
                    if isinstance(year_data, dict):
                        for year, data in year_data.items():
                            scores = [d.get("score") for d in data if d.get("score") is not None]
                            if scores:
                                dept_avg_data.append({
                                    "部門": dept_name[:15],
                                    "年度": str(year),
                                    "平均得点": round(sum(scores) / len(scores), 2)
                                })

                if dept_avg_data:
                    dept_avg_df = pd.DataFrame(dept_avg_data)
                    latest_years = sorted(dept_avg_df["年度"].unique(), key=_year_sort_key, reverse=True)[:3]
                    dept_avg_df = dept_avg_df[dept_avg_df["年度"].isin(latest_years)]

                    # グループ化縦棒グラフ（年度ごとに横並び）- mark_rectで非0基点
                    import altair as alt
                    all_scores = dept_avg_df["平均得点"].tolist()
                    y_min = max(0, min(all_scores) - 5)
                    y_max = max(all_scores) + 2
                    dept_avg_df["基点"] = y_min
                    chart = alt.Chart(dept_avg_df).mark_rect(width=12).encode(
                        x=alt.X('年度:N', title=None, axis=alt.Axis(labelAngle=0)),
                        y=alt.Y('基点:Q', title='平均得点', scale=alt.Scale(domain=[y_min, y_max])),
                        y2=alt.Y2('平均得点:Q'),
                        color=alt.Color('年度:N', title='年度'),
                        column=alt.Column('部門:N', title=None, header=alt.Header(labelOrient='bottom')),
                        tooltip=['部門', '年度', '平均得点']
                    ).properties(width=60, height=400)
                    st.altair_chart(chart)
                else:
                    st.info("部門別データにスコアが含まれていません")
            else:
                st.info("部門別データがありません")

            # 評価項目別の連続1位
            st.subheader("📋 評価項目別 連続1位記録")
            item_trends = historical_data.get("item_trends", {})
            if item_trends:
                item_records = []
                for item_name, data in item_trends.items():
                    for streak in data.get("consecutive_wins", []):
                        if streak.get("years", 0) >= 2:
                            item_records.append({
                                "評価項目": item_name,
                                "企業名": streak["company"],
                                "連続年数": f"{streak['years']}年",
                                "期間": f"{streak['start']}〜{streak['end']}",
                                "継続中": "✅" if streak.get("is_current") else ""
                            })
                if item_records:
                    item_records.sort(key=lambda x: -int(x["連続年数"].replace("年", "")))
                    st.dataframe(pd.DataFrame(item_records[:15]), use_container_width=True, hide_index=True)

            # 部門別の連続1位
            st.subheader("🏷️ 部門別 連続1位記録")
            dept_trends = historical_data.get("dept_trends", {})
            if dept_trends:
                dept_records = []
                for dept_name, data in dept_trends.items():
                    for streak in data.get("consecutive_wins", []):
                        if streak.get("years", 0) >= 2:
                            dept_records.append({
                                "部門": dept_name,
                                "企業名": streak["company"],
                                "連続年数": f"{streak['years']}年",
                                "期間": f"{streak['start']}〜{streak['end']}",
                                "継続中": "✅" if streak.get("is_current") else ""
                            })
                if dept_records:
                    dept_records.sort(key=lambda x: -int(x["連続年数"].replace("年", "")))
                    st.dataframe(pd.DataFrame(dept_records[:15]), use_container_width=True, hide_index=True)

    with tab3:
        st.header("📊 総合ランキング（経年詳細）")

        # トップに歴代記録を表示
        records = historical_data.get("historical_records", {})
        if records:
            display_historical_summary(records)
            display_consecutive_wins_compact(records)
            st.divider()

        # 総合ランキング1位獲得回数ランキング
        most_wins = records.get("most_wins", []) if records else []
        if most_wins:
            st.subheader("🏆 総合ランキング 1位獲得回数ランキング")
            # 最新年度を取得
            all_years = set()
            for r in most_wins:
                all_years.update(r.get("years", []))
            latest_year = max(all_years, key=_year_sort_key) if all_years else None

            overall_wins_data = []
            for r in most_wins:
                if r.get("wins", 0) > 0:
                    # 継続中フラグ: 最新年度も1位なら✅
                    is_current = latest_year in r.get("years", []) if latest_year else False
                    wins = r.get("wins", 0)
                    overall_wins_data.append({
                        "企業名": r.get("company", ""),
                        "1位回数": wins,  # ソート用に数値で保持
                        "継続中": "✅" if is_current else "",
                        "獲得年": ", ".join(map(str, r.get("years", [])))
                    })
            if overall_wins_data:
                # 1位回数の多い順にソート
                overall_wins_data.sort(key=lambda x: -x["1位回数"])
                # 表示用に回数を文字列に変換
                for d in overall_wins_data:
                    d["1位回数"] = f"{d['1位回数']}回"
                st.dataframe(pd.DataFrame(overall_wins_data), use_container_width=True, hide_index=True)

        # v7.3: 総合ランキング TOP10得点の経年推移を「1位獲得回数ランキング」の下に移動
        if overall_data and len(overall_data) > 1:
            st.subheader("📊 得点の経年推移（TOP10企業）")
            # 最新年度のTOP10企業を取得
            latest_year_for_chart = max(overall_data.keys(), key=_year_sort_key)
            latest_top10_for_chart = sorted(overall_data[latest_year_for_chart], key=lambda x: x.get("score") or 0, reverse=True)[:10]
            top10_companies_for_chart = [d.get("company") for d in latest_top10_for_chart if d.get("company")]

            line_data_for_chart = []
            for year in sorted(overall_data.keys(), key=_year_sort_key):
                for item in overall_data[year]:
                    company = item.get("company")
                    score = item.get("score")
                    if company in top10_companies_for_chart and score is not None:
                        line_data_for_chart.append({
                            "年度": str(year),
                            "得点": score,
                            "企業名": company[:15]  # 長い企業名を短縮
                        })
            if line_data_for_chart and len(line_data_for_chart) > 1:
                import altair as alt
                line_df_for_chart = pd.DataFrame(line_data_for_chart)
                # 動的Y軸範囲
                all_scores_for_chart = [d["得点"] for d in line_data_for_chart]
                y_min_for_chart = max(0, min(all_scores_for_chart) - 3)
                y_max_for_chart = max(all_scores_for_chart) + 3
                chart = alt.Chart(line_df_for_chart).mark_line(point=True).encode(
                    x=alt.X('年度:O', title='年度'),
                    y=alt.Y('得点:Q', title='得点', scale=alt.Scale(domain=[y_min_for_chart, y_max_for_chart])),
                    color=alt.Color('企業名:N', title='企業名'),
                    tooltip=['年度', '企業名', '得点']
                ).properties(height=400, title="総合ランキング 得点の経年推移（TOP10企業）")
                st.altair_chart(chart, use_container_width=True)

        st.divider()

        if overall_data:
            # 年度ごとに全データを表示（アップロードデータをマーク）
            for year in sorted(overall_data.keys(), key=_year_sort_key, reverse=True):
                source_mark = "📁" if year in uploaded_years else "🌐"
                # 該当年度のURLを取得
                year_url = None
                if used_urls:
                    for url_item in used_urls.get("overall", []):
                        if url_item.get("year") == year and url_item.get("status") == "success":
                            year_url = url_item.get("url", "")
                            break
                # expanderのタイトル（URLはクリック可能にするため中に表示）
                expander_title = f"{source_mark} {year}年"
                with st.expander(expander_title, expanded=(year == max(overall_data.keys(), key=_year_sort_key))):
                    # URLを表の上にクリック可能なリンクとして表示
                    if year_url:
                        st.markdown(f"🔗 **参照URL**: [{year_url}]({year_url})")
                    df = pd.DataFrame(overall_data[year])
                    # v7.3: 空白列名、数字のみの列名、Unnamed列を除外
                    valid_cols = [col for col in df.columns
                                  if col and str(col).strip()
                                  and not str(col).strip().isdigit()
                                  and not str(col).startswith('Unnamed')]
                    df = df[valid_cols]
                    st.dataframe(df, use_container_width=True, hide_index=True)

                    # 該当年度の縦棒グラフ（得点上位10社）
                    year_data_sorted = sorted(overall_data[year], key=lambda x: x.get("score") or 0, reverse=True)[:10]
                    if year_data_sorted and any(d.get("score") for d in year_data_sorted):
                        import altair as alt
                        bar_data = []
                        for d in year_data_sorted:
                            if d.get("score") is not None and d.get("company"):
                                bar_data.append({
                                    "企業名": d["company"][:12],  # 長い企業名を短縮
                                    "得点": d["score"]
                                })
                        if bar_data:
                            bar_df = pd.DataFrame(bar_data)
                            # mark_rectで非0基点の棒グラフを実装（差分を見やすく）
                            scores = [d["得点"] for d in bar_data]
                            y_min = max(0, min(scores) - 5)  # 最小値-5を基点に
                            y_max = max(scores) + 2
                            bar_df["基点"] = y_min
                            chart = alt.Chart(bar_df).mark_rect(width=25).encode(
                                x=alt.X('企業名:N', sort=alt.EncodingSortField(field='得点', order='descending'), title=None, axis=alt.Axis(labelAngle=-45)),
                                y=alt.Y('基点:Q', title='得点', scale=alt.Scale(domain=[y_min, y_max])),
                                y2=alt.Y2('得点:Q'),
                                color=alt.Color('得点:Q', scale=alt.Scale(scheme='blues'), legend=None),
                                tooltip=['企業名', '得点']
                            ).properties(height=300, title=f"{year}年 得点上位10社")
                            st.altair_chart(chart, use_container_width=True)

            # 経年比較テーブル
            st.subheader("📈 経年比較（全社得点推移）")

            companies = set()
            for year_data in overall_data.values():
                for item in year_data:
                    companies.add(item.get("company", ""))

            comparison_data = []
            for company in sorted(companies):
                row = {"企業名": company}
                for year in sorted(overall_data.keys(), key=_year_sort_key):
                    score = "-"
                    rank = "-"
                    for item in overall_data[year]:
                        if item.get("company") == company:
                            score = item.get("score", "-")
                            rank = item.get("rank", "-")
                            break
                    row[f"{year}年得点"] = score
                    row[f"{year}年順位"] = rank
                comparison_data.append(row)

            if comparison_data:
                st.dataframe(pd.DataFrame(comparison_data), use_container_width=True)

    with tab4:
        st.header("📋 評価項目別ランキング（経年）")

        # 評価項目別1位獲得回数ランキング
        if item_most_wins:
            st.subheader("🏆 評価項目別 1位獲得回数ランキング")
            # 最新年度を取得
            all_years = set()
            for wins_list in item_most_wins.values():
                for r in wins_list:
                    all_years.update(r.get("years", []))
            latest_year = max(all_years, key=_year_sort_key) if all_years else None

            item_wins_data = []
            for item_name, wins_list in item_most_wins.items():
                for r in wins_list[:3]:  # 各項目上位3社
                    if r["wins"] > 0:
                        # 継続中フラグ: 最新年度も1位なら✅
                        is_current = latest_year in r.get("years", []) if latest_year else False
                        item_wins_data.append({
                            "評価項目": item_name,
                            "企業名": r["company"],
                            "1位回数": r['wins'],  # ソート用に数値で保持
                            "獲得率": f"{r['wins']/r['total_years']*100:.1f}%" if r['total_years'] > 0 else "0.0%",
                            "継続中": "✅" if is_current else "",
                            "獲得年": ", ".join(map(str, r["years"]))
                        })
            if item_wins_data:
                # 1位回数の多い順にソート
                item_wins_data.sort(key=lambda x: -x["1位回数"])
                # 表示用に回数を文字列に変換
                for d in item_wins_data:
                    d["1位回数"] = f"{d['1位回数']}回"
                st.dataframe(pd.DataFrame(item_wins_data), use_container_width=True, hide_index=True)
            st.divider()

        # トップに評価項目別の連続1位記録
        item_trends = historical_data.get("item_trends", {})
        if item_trends:
            st.subheader("📋 評価項目別 連続1位記録（上位10件）")
            item_records = []
            for item_name, data in item_trends.items():
                for streak in data.get("consecutive_wins", []):
                    # 複数回1位を獲得したもののみ対象（1年だけの受賞は除外）
                    if streak.get("years", 0) >= 2:
                        item_records.append({
                            "評価項目": item_name,
                            "企業名": streak["company"],
                            "連続年数": f"{streak['years']}年",
                            "期間": f"{streak['start']}〜{streak['end']}",
                            "継続中": "✅" if streak.get("is_current") else ""
                        })
            if item_records:
                item_records.sort(key=lambda x: -int(x["連続年数"].replace("年", "")))
                st.dataframe(pd.DataFrame(item_records[:10]), use_container_width=True, hide_index=True)
            st.divider()

        # v7.3: 「評価項目別 得点の経年推移（TOP10企業）」セクションを削除

        if item_data:
            for item_name, year_data in item_data.items():
                # 最新名称を取得（名称変更情報があれば使用）
                display_name = item_name
                name_change_info = item_name_changes.get(item_name)
                if name_change_info and name_change_info.get("latest_name"):
                    display_name = name_change_info["latest_name"]

                with st.expander(f"📌 {display_name}", expanded=False):
                    # v6.1: レイアウト変更 - 名称変更 → 1位の推移 → 経年推移 → 年数/URL の順

                    # 1. 名称変更があれば注記を表示
                    if name_change_info and name_change_info.get("changes"):
                        for change in name_change_info["changes"]:
                            st.info(f"📝 **名称変更**: {change['change_year']}年より「{change['from_name']}」→「{change['to_name']}」に変更")

                    if isinstance(year_data, dict) and len(year_data) > 1:
                        # 2. 1位の推移（名称変更の直後に配置）
                        st.markdown("**📈 1位の推移**")
                        history = []
                        for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                            year_list = year_data.get(year)
                            if year_list and isinstance(year_list, list) and len(year_list) > 0:
                                top = year_list[0]
                                if top and isinstance(top, dict):
                                    history.append({
                                        "年度": year,
                                        "1位": top.get("company", "-"),
                                        "得点": top.get("score", "-")
                                    })
                        if history:
                            st.dataframe(pd.DataFrame(history), use_container_width=True)

                        # 3. 経年変化の折れ線グラフ（TOP10企業の得点推移）
                        st.markdown("**📊 得点の経年推移（TOP10企業）**")
                        # 最新年度のTOP10企業を取得
                        latest_yr = max(year_data.keys(), key=_year_sort_key)
                        latest_top10 = sorted(year_data[latest_yr], key=lambda x: x.get("score") or 0, reverse=True)[:10]
                        top10_companies = [d.get("company") for d in latest_top10 if d.get("company")]

                        line_data = []
                        for yr in sorted(year_data.keys(), key=_year_sort_key):
                            for item in year_data[yr]:
                                company = item.get("company")
                                score = item.get("score")
                                if company in top10_companies and score is not None:
                                    line_data.append({
                                        "年度": str(yr),
                                        "得点": score,
                                        "企業名": company[:15]
                                    })
                        if line_data and len(line_data) > 1:
                            import altair as alt
                            line_df = pd.DataFrame(line_data)
                            # 動的Y軸範囲
                            all_scores = [d["得点"] for d in line_data]
                            y_min = max(0, min(all_scores) - 3)
                            y_max = max(all_scores) + 3
                            chart = alt.Chart(line_df).mark_line(point=True).encode(
                                x=alt.X('年度:O', title='年度'),
                                y=alt.Y('得点:Q', title='得点', scale=alt.Scale(domain=[y_min, y_max])),
                                color=alt.Color('企業名:N', title='企業名'),
                                tooltip=['年度', '企業名', '得点']
                            ).properties(height=300, title=f"{item_name} 得点の経年推移（TOP10企業）")
                            st.altair_chart(chart, use_container_width=True)

                        st.divider()

                        # 4. 各年度データ（年数/URL）
                        for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                            # 該当年度のURLを取得
                            year_url = None
                            if used_urls:
                                for url_item in used_urls.get("items", []):
                                    search_name = f"{item_name}({year}年)"
                                    if url_item.get("name") == search_name and url_item.get("status") == "success":
                                        year_url = url_item.get("url", "")
                                        break
                            # 年度の横にURL表示
                            if year_url:
                                st.markdown(f"**{year}年** 🔗 {year_url}")
                            else:
                                st.markdown(f"**{year}年**")
                            df = pd.DataFrame(year_data[year])
                            # v7.3: 空白列名、数字のみの列名、Unnamed列を除外
                            valid_cols = [col for col in df.columns
                                          if col and str(col).strip()
                                          and not str(col).strip().isdigit()
                                          and not str(col).startswith('Unnamed')]
                            df = df[valid_cols]
                            st.dataframe(df, use_container_width=True, hide_index=True)

                    elif isinstance(year_data, dict):
                        # 1年分のみのデータ
                        for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                            year_url = None
                            if used_urls:
                                for url_item in used_urls.get("items", []):
                                    search_name = f"{item_name}({year}年)"
                                    if url_item.get("name") == search_name and url_item.get("status") == "success":
                                        year_url = url_item.get("url", "")
                                        break
                            if year_url:
                                st.markdown(f"**{year}年** 🔗 {year_url}")
                            else:
                                st.markdown(f"**{year}年**")
                            df = pd.DataFrame(year_data[year])
                            # v7.3: 空白列名、数字のみの列名、Unnamed列を除外
                            valid_cols = [col for col in df.columns
                                          if col and str(col).strip()
                                          and not str(col).strip().isdigit()
                                          and not str(col).startswith('Unnamed')]
                            df = df[valid_cols]
                            st.dataframe(df, use_container_width=True, hide_index=True)
                    else:
                        df = pd.DataFrame(year_data)
                        st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("評価項目別データは取得できませんでした")

    with tab5:
        st.header("🏷️ 部門別ランキング（経年）")

        # 部門別1位獲得回数ランキング
        if dept_most_wins:
            st.subheader("🏆 部門別 1位獲得回数ランキング")
            # 最新年度を取得
            all_years = set()
            for wins_list in dept_most_wins.values():
                for r in wins_list:
                    all_years.update(r.get("years", []))
            latest_year = max(all_years, key=_year_sort_key) if all_years else None

            dept_wins_data = []
            for dept_name, wins_list in dept_most_wins.items():
                for r in wins_list[:3]:  # 各部門上位3社
                    if r["wins"] > 0:
                        # 継続中フラグ: 最新年度も1位なら✅
                        is_current = latest_year in r.get("years", []) if latest_year else False
                        dept_wins_data.append({
                            "部門": dept_name,
                            "企業名": r["company"],
                            "1位回数": r['wins'],  # ソート用に数値で保持
                            "獲得率": f"{r['wins']/r['total_years']*100:.1f}%" if r['total_years'] > 0 else "0.0%",
                            "継続中": "✅" if is_current else "",
                            "獲得年": ", ".join(map(str, r["years"]))
                        })
            if dept_wins_data:
                # 1位回数の多い順にソート
                dept_wins_data.sort(key=lambda x: -x["1位回数"])
                # 表示用に回数を文字列に変換
                for d in dept_wins_data:
                    d["1位回数"] = f"{d['1位回数']}回"
                st.dataframe(pd.DataFrame(dept_wins_data), use_container_width=True, hide_index=True)
            st.divider()

        # トップに部門別の連続1位記録
        dept_trends = historical_data.get("dept_trends", {})
        if dept_trends:
            st.subheader("🏷️ 部門別 連続1位記録（上位10件）")
            dept_records = []
            for dept_name, data in dept_trends.items():
                for streak in data.get("consecutive_wins", []):
                    # 複数回1位を獲得したもののみ対象（1年だけの受賞は除外）
                    if streak.get("years", 0) >= 2:
                        dept_records.append({
                            "部門": dept_name,
                            "企業名": streak["company"],
                            "連続年数": f"{streak['years']}年",
                            "期間": f"{streak['start']}〜{streak['end']}",
                            "継続中": "✅" if streak.get("is_current") else ""
                        })
            if dept_records:
                dept_records.sort(key=lambda x: -int(x["連続年数"].replace("年", "")))
                st.dataframe(pd.DataFrame(dept_records[:10]), use_container_width=True, hide_index=True)
            st.divider()

        # v7.3: 「部門別 得点の経年推移（TOP10企業）」セクションを削除

        if dept_data:
            for dept_name, year_data in dept_data.items():
                # 最新名称を取得（名称変更情報があれば使用）
                display_name = dept_name
                name_change_info = dept_name_changes.get(dept_name)
                if name_change_info and name_change_info.get("latest_name"):
                    display_name = name_change_info["latest_name"]

                with st.expander(f"📌 {display_name}", expanded=False):
                    # v6.1: レイアウト変更 - 名称変更 → 1位の推移 → 経年推移 → 年数/URL の順（評価項目別と同じ）

                    # 1. 名称変更があれば注記を表示
                    if name_change_info and name_change_info.get("changes"):
                        for change in name_change_info["changes"]:
                            st.info(f"📝 **名称変更**: {change['change_year']}年より「{change['from_name']}」→「{change['to_name']}」に変更")

                    if isinstance(year_data, dict) and len(year_data) > 1:
                        # 2. 1位の推移（名称変更の直後に配置）
                        st.markdown("**📈 1位の推移**")
                        history = []
                        for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                            year_list = year_data.get(year)
                            if year_list and isinstance(year_list, list) and len(year_list) > 0:
                                top = year_list[0]
                                if top and isinstance(top, dict):
                                    history.append({
                                        "年度": year,
                                        "1位": top.get("company", "-"),
                                        "得点": top.get("score", "-")
                                    })
                        if history:
                            st.dataframe(pd.DataFrame(history), use_container_width=True)

                        # 3. 経年変化の折れ線グラフ（TOP10企業の得点推移）
                        st.markdown("**📊 得点の経年推移（TOP10企業）**")
                        # 最新年度のTOP10企業を取得
                        latest_yr = max(year_data.keys(), key=_year_sort_key)
                        latest_top10 = sorted(year_data[latest_yr], key=lambda x: x.get("score") or 0, reverse=True)[:10]
                        top10_companies = [d.get("company") for d in latest_top10 if d.get("company")]

                        line_data = []
                        for yr in sorted(year_data.keys(), key=_year_sort_key):
                            for item in year_data[yr]:
                                company = item.get("company")
                                score = item.get("score")
                                if company in top10_companies and score is not None:
                                    line_data.append({
                                        "年度": str(yr),
                                        "得点": score,
                                        "企業名": company[:15]
                                    })
                        if line_data and len(line_data) > 1:
                            import altair as alt
                            line_df = pd.DataFrame(line_data)
                            # 動的Y軸範囲
                            all_scores = [d["得点"] for d in line_data]
                            y_min = max(0, min(all_scores) - 3)
                            y_max = max(all_scores) + 3
                            chart = alt.Chart(line_df).mark_line(point=True).encode(
                                x=alt.X('年度:O', title='年度'),
                                y=alt.Y('得点:Q', title='得点', scale=alt.Scale(domain=[y_min, y_max])),
                                color=alt.Color('企業名:N', title='企業名'),
                                tooltip=['年度', '企業名', '得点']
                            ).properties(height=300, title=f"{dept_name} 得点の経年推移（TOP10企業）")
                            st.altair_chart(chart, use_container_width=True)

                        st.divider()

                        # 4. 各年度データ（年数/URL）
                        for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                            # 該当年度のURLを取得
                            year_url = None
                            if used_urls:
                                for url_item in used_urls.get("departments", []):
                                    search_name = f"{dept_name}({year}年)"
                                    if url_item.get("name") == search_name and url_item.get("status") == "success":
                                        year_url = url_item.get("url", "")
                                        break
                            # 年度の横にURL表示
                            if year_url:
                                st.markdown(f"**{year}年** 🔗 {year_url}")
                            else:
                                st.markdown(f"**{year}年**")
                            df = pd.DataFrame(year_data[year])
                            # v7.3: 空白列名、数字のみの列名、Unnamed列を除外
                            valid_cols = [col for col in df.columns
                                          if col and str(col).strip()
                                          and not str(col).strip().isdigit()
                                          and not str(col).startswith('Unnamed')]
                            df = df[valid_cols]
                            st.dataframe(df, use_container_width=True, hide_index=True)

                    elif isinstance(year_data, dict):
                        # 1年分のみのデータ
                        for year in sorted(year_data.keys(), key=_year_sort_key, reverse=True):
                            year_url = None
                            if used_urls:
                                for url_item in used_urls.get("departments", []):
                                    search_name = f"{dept_name}({year}年)"
                                    if url_item.get("name") == search_name and url_item.get("status") == "success":
                                        year_url = url_item.get("url", "")
                                        break
                            if year_url:
                                st.markdown(f"**{year}年** 🔗 {year_url}")
                            else:
                                st.markdown(f"**{year}年**")
                            df = pd.DataFrame(year_data[year])
                            # v7.3: 空白列名、数字のみの列名、Unnamed列を除外
                            valid_cols = [col for col in df.columns
                                          if col and str(col).strip()
                                          and not str(col).strip().isdigit()
                                          and not str(col).startswith('Unnamed')]
                            df = df[valid_cols]
                            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("部門別データは存在しないか取得できませんでした")

    with tab6:
        # プレスリリース作成タブ (v8.0追加)
        if RELEASE_FEATURES_AVAILABLE:
            render_release_tab(
                ranking_name=ranking_name,
                overall_data=overall_data,
                item_data=item_data,
                dept_data=dept_data,
                historical_data=historical_data
            )
        else:
            st.warning("プレスリリース機能のモジュールが見つかりません")
            st.info("release_tab.py, validator.py, release_generator.py, company_master.py が必要です")

    with tab7:
        st.header("📎 参考資料（使用したURL）")

        if used_urls:
            # 総合ランキングURL
            st.subheader("📊 総合ランキング")
            overall_urls = used_urls.get("overall", [])
            if overall_urls:
                url_df = pd.DataFrame([
                    {
                        "年度": item.get("year", ""),
                        "ステータス": "✅ 成功" if item.get("status") == "success" else "❌ 失敗",
                        "URL": item.get("url", "")
                    }
                    for item in overall_urls
                ])
                # URLをクリック可能なリンクとして表示
                st.dataframe(
                    url_df,
                    column_config={
                        "URL": st.column_config.LinkColumn("URL", display_text="🔗 リンクを開く")
                    },
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("総合ランキングのURLデータがありません")

            st.divider()

            # 評価項目別URL
            st.subheader("📋 評価項目別ランキング")
            item_urls = used_urls.get("items", [])
            if item_urls:
                url_df = pd.DataFrame([
                    {
                        "項目名": item.get("name", ""),
                        "ステータス": "✅ 成功" if item.get("status") == "success" else "❌ 失敗",
                        "URL": item.get("url", "")
                    }
                    for item in item_urls
                ])
                st.dataframe(
                    url_df,
                    column_config={
                        "URL": st.column_config.LinkColumn("URL", display_text="🔗 リンクを開く")
                    },
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("評価項目別のURLデータがありません")

            st.divider()

            # 部門別URL
            st.subheader("🏷️ 部門別ランキング")
            dept_urls = used_urls.get("departments", [])
            if dept_urls:
                url_df = pd.DataFrame([
                    {
                        "部門名": item.get("name", ""),
                        "ステータス": "✅ 成功" if item.get("status") == "success" else "❌ 失敗",
                        "URL": item.get("url", "")
                    }
                    for item in dept_urls
                ])
                st.dataframe(
                    url_df,
                    column_config={
                        "URL": st.column_config.LinkColumn("URL", display_text="🔗 リンクを開く")
                    },
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("部門別のURLデータがありません")
        else:
            if uploaded_years and not scraped_years:
                st.info("📁 アップロードデータのみを使用したため、参考URLはありません")
            else:
                st.info("参考資料（URL情報）がありません")

        # データソース
        st.divider()
        st.markdown("**📌 データソース**: [オリコン顧客満足度ランキング](https://life.oricon.co.jp/)")

# フッター
st.sidebar.divider()
st.sidebar.markdown("---")
st.sidebar.markdown("📌 **データソース**: life.oricon.co.jp")
st.sidebar.markdown(f"🔧 **バージョン**: {__version__}")


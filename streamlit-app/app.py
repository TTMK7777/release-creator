# -*- coding: utf-8 -*-
"""
オリコン顧客満足度®調査 TOPICSサポートシステム
Streamlit版 v3.5 - 年度列検出ロジック改善版
- 年度列の誤検出を防止（回答者数（最新年）等を除外）
- 年度値の妥当性チェック（2000-2030範囲外は指定年度を使用）
- オリコン内部Excelフォーマット対応（ヘッダー行自動検出）
- 年度列がない場合はファイル名から年度を推測
- 列名の柔軟な検出（ランキング対象企業名、スコア等）
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from scraper import OriconScraper
from analyzer import TopicsAnalyzer, HistoricalAnalyzer


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
        for year in sorted(overall_data.keys(), reverse=True):
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
            for year in sorted(overall_data.keys()):
                score = None
                rank = None
                for item in overall_data.get(year, []):
                    if item.get("company") == company:
                        score = item.get("score")
                        rank = item.get("rank")
                        break
                row[f"{year}年_得点"] = score if score else ""
                row[f"{year}年_順位"] = rank if rank else ""
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
                    "獲得率": f"{r['wins']/r['total_years']*100:.1f}%",
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
                for year in sorted(year_data.keys(), reverse=True):
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
                for year in sorted(year_data.keys(), reverse=True):
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

            # ヘッダー行を検出（"順位"と"ID"を含む行）
            header_row = None
            category_name = None  # 部門別シートのカテゴリ名
            for idx, row in df_raw.iterrows():
                row_str = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if '順位' in row_str and 'ID' in row_str and ('企業' in row_str or 'ランキング' in row_str):
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
                    # 年度の取得
                    if year_col and pd.notna(row.get(year_col)):
                        try:
                            year = int(row[year_col])
                            if year < 2000 or year > 2030:
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

                    # 3. 部門別シート（業態別、投資スタイル別、利用チャート別、レベル別、サポート別）
                    elif any(x in sheet_name for x in ['業態', '投資スタイル', '利用チャート', 'チャート', 'レベル', 'サポート', '別']):
                        # カテゴリ名があればそれを使用、なければシート名
                        dept_name = category_name if category_name else sheet_name.replace('別', '')
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
        return None, None, None, f"{str(e)}\n{traceback.format_exc()}"


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
    if consecutive:
        st.markdown("**🥇 連続1位記録（上位5件）**")
        cons_df = pd.DataFrame([
            {
                "企業名": r["company"],
                "連続年数": f"{r['years']}年",
                "期間": f"{r['start_year']}〜{r['end_year']}",
                "継続中": "✅" if r.get("is_current") else ""
            }
            for r in consecutive[:5]
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
st.warning("⚠️ **注意事項**: Webスクレイピング技術を使用しています。情報の正確性は担当者が必ず確認してください。")
st.markdown("プレスリリースの見出しトピックス候補を自動生成します")

# サイドバー
st.sidebar.header("⚙️ 設定")

# ランキング選択
ranking_options = {
    # === 金融・投資 ===
    "FX（顧客満足度）": "_fx",
    "FX（FP評価）": "_fx@type02",
    "銀行カードローン": "card-loan",
    "ノンバンクカードローン": "card-loan/nonbank",
    "ネット証券（顧客満足度）": "_certificate",
    "ネット証券（FP評価）": "_certificate@type02",
    "iDeCo証券会社": "ideco",
    "NISA（証券会社）": "_nisa",
    "クレジットカード": "creditcard",
    # === 保険 ===
    "自動車保険（ダイレクト型）": "_insurance",
    "自動車保険（代理店型）": "_insurance@type02",
    "自動車保険（FP推奨）": "_insurance@type03",
    "生命保険": "life-insurance",
    "保険ショップ（FP）": "_hokenshop",
    # === 通信 ===
    "携帯キャリア": "mobile-carrier",
    "格安SIM": "mvno",
    "格安SIM（SIMのみ）": "mvno/sim",
    "格安スマホ": "mvno/sp",
    # === 教育（英会話） ===
    "英会話スクール": "english-school",
    "オンライン英会話": "online-english",
    "子ども英語教室（幼児）": "kids-english/preschooler",
    "子ども英語教室（小学生）": "kids-english/grade-schooler",
    # === 教育（学習） ===
    "家庭教師": "tutor",
    "通信教育（高校生）": "online-study/highschool",
    "通信教育（中学生）": "online-study/junior-hs",
    "通信教育（小学生）": "online-study/elementary",
    # === 教育（スポーツ） ===
    "キッズスイミングスクール（幼児）": "kids-swimming/preschooler",
    "キッズスイミングスクール（小学生）": "kids-swimming/grade-schooler",
    # === 転職・人材 ===
    "転職サイト": "recruit",
    "転職エージェント": "_agent",
    "派遣会社（製造業）": "_staffing_manufacture",
    # === 住宅・不動産 ===
    "ハウスメーカー（注文住宅）": "house-maker",
    "建売住宅ビルダー": "new-ready-built-house",
    "建売住宅（パワービルダー）": "new-ready-built-house/powerbuilder",
    "新築分譲マンション": "new-condominiums",
    # === 生活サービス ===
    "引越し会社": "_move",
    "食材宅配": "food-delivery",
    "ミールキット": "food-delivery/meal-kit",
    "子ども見守りGPS": "child-gps",
    # === フィットネス ===
    "フィットネスクラブ": "_fitness",
    "24時間ジム": "_fitness/24hours",
    # === その他 ===
    "カスタム入力": "custom"
}

selected_ranking = st.sidebar.selectbox(
    "ランキングを選択",
    list(ranking_options.keys())
)

# カスタム入力の場合
if selected_ranking == "カスタム入力":
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
current_year = 2025  # Webサイトの最新年度
start_year = 2006

year_option = st.sidebar.radio(
    "過去データ取得範囲",
    ["直近3年", "直近5年", "全年度（2006年〜）", "カスタム範囲"]
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

# ファイルアップロード（オプション）
st.sidebar.markdown("---")
st.sidebar.markdown("### 📁 最新データのアップロード（オプション）")
uploaded_file = st.sidebar.file_uploader(
    "最新のランキングExcelをアップロード",
    type=["xlsx", "xls"],
    help="最新のランキング資料をアップロードすると、過去データと統合して分析します"
)

# アップロードデータの年度指定
upload_year = None
if uploaded_file:
    st.sidebar.success(f"✅ {uploaded_file.name}")
    upload_year = st.sidebar.number_input(
        "📅 アップロードデータの年度",
        min_value=2006,
        max_value=2030,
        value=2026,
        help="アップロードしたファイルのデータ年度を指定してください（例: 2026年発表データなら2026）"
    )
    st.sidebar.info(f"📌 **{upload_year}年**のデータとしてアップロードファイルを使用し、それ以外の年度はWebから取得して統合します")

# セッション状態の初期化
if 'results_data' not in st.session_state:
    st.session_state.results_data = None

# 実行ボタン
if st.sidebar.button("🚀 TOPICS出し実行", type="primary", use_container_width=True):

    if not ranking_slug:
        st.error("ランキングのURL名を入力してください")
    else:
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
                log(f"  - 含まれる年度: {sorted(uploaded_years)}")
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

            # Step 2: Webスクレイピングで過去データを取得
            status_text.text("🌐 Webから過去データを取得中...")
            progress_bar.progress(20)

            log(f"[INFO] スクレイパー初期化: {ranking_slug} ({ranking_name})")
            scraper = OriconScraper(ranking_slug, ranking_name)
            subpath_info = f" + subpath: {scraper.subpath}" if scraper.subpath else ""
            log(f"[INFO] URL prefix: {scraper.url_prefix}{subpath_info}")

            # スクレイピング対象年度を決定
            # - アップロードデータに含まれる年度は除外
            # - Webサイトの最新年度（current_year=2025）を超える年度は除外
            scrape_years = []
            effective_end_year = min(year_range[1], current_year)  # Webサイトの最新年度を超えない
            for y in range(year_range[0], effective_end_year + 1):
                if y not in uploaded_years:
                    scrape_years.append(y)

            log(f"[INFO] 年度範囲設定: {year_range[0]}〜{year_range[1]}")
            log(f"[INFO] Webサイト最新年度: {current_year}")
            log(f"[INFO] アップロード年度: {sorted(uploaded_years) if uploaded_years else 'なし'}")

            if scrape_years:
                log(f"[INFO] スクレイピング対象年度: {scrape_years}")
                scrape_range = (min(scrape_years), max(scrape_years))
            else:
                log(f"[INFO] アップロードデータで全年度カバー済み、スクレイピングをスキップ")
                scrape_range = None

            scraped_overall = {}
            scraped_item = {}
            scraped_dept = {}

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

            used_urls = scraper.used_urls if scrape_range else None

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

            # Step 4: 分析実行
            status_text.text("🔍 TOPICS分析中...")
            analyzer = TopicsAnalyzer(overall_data, item_data, ranking_name)
            topics = analyzer.analyze()
            progress_bar.progress(85)

            # Step 5: 歴代記録・得点推移分析
            status_text.text("📈 歴代記録・得点推移を分析中...")
            historical_analyzer = HistoricalAnalyzer(overall_data, item_data, dept_data, ranking_name)
            historical_data = historical_analyzer.analyze_all()
            progress_bar.progress(95)

            # 完了
            status_text.text("✅ 完了!")
            progress_bar.progress(100)

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
                'scraped_years': list(scraped_overall.keys()) if scraped_overall else []
            }

        except Exception as e:
            st.error(f"エラーが発生しました: {str(e)}")
            st.exception(e)

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

    # 結果表示
    st.success(f"✅ {ranking_name}のTOPICS出しが完了しました")

    # データソース情報
    if uploaded_years or scraped_years:
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            if uploaded_years:
                st.info(f"📁 **アップロードデータ**: {sorted(uploaded_years)}年")
        with col_info2:
            if scraped_years:
                st.info(f"🌐 **Webスクレイピング**: {sorted(scraped_years)}年")

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

    # タブで結果表示（新しい構成）
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "⭐ 推奨TOPICS",
        "🎯 見出し案",
        "🏆 歴代記録・得点推移",
        "📊 総合ランキング",
        "📋 評価項目別",
        "🏷️ 部門別",
        "📎 参考資料"
    ])

    with tab1:
        st.header("⭐ 推奨TOPICS")
        for i, topic in enumerate(topics["recommended"], 1):
            importance = topic.get("importance", "重要")
            st.markdown(f"### {i}. [{importance}] {topic['title']}")
            st.markdown(f"- **根拠**: {topic['evidence']}")
            st.markdown(f"- **インパクト**: {'★' * topic.get('impact', 3)}")
            st.divider()

        if topics.get("other"):
            st.subheader("📊 その他のTOPICS候補")
            for topic in topics["other"]:
                st.markdown(f"- {topic}")

    with tab2:
        st.header("🎯 見出し案")
        for i, headline in enumerate(topics.get("headlines", []), 1):
            st.markdown(f"**パターン{i}**: {headline}")

        # コピー用テキスト
        st.subheader("📋 コピー用テキスト")
        copy_text = "\n".join([
            "【推奨TOPICS】",
            *[f"{i}. {t['title']}" for i, t in enumerate(topics["recommended"], 1)],
            "",
            "【見出し案】",
            *[f"パターン{i}: {h}" for i, h in enumerate(topics.get("headlines", []), 1)]
        ])
        st.text_area("コピー用", copy_text, height=300, label_visibility="collapsed")

    with tab3:
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
                # 連続1位記録
                st.subheader("🥇 連続1位記録")
                consecutive = records.get("consecutive_wins", [])
                if consecutive:
                    cons_df = pd.DataFrame([
                        {
                            "企業名": r["company"],
                            "連続年数": f"{r['years']}年",
                            "期間": f"{r['start_year']}〜{r['end_year']}",
                            "継続中": "✅" if r.get("is_current") else ""
                        }
                        for r in consecutive[:10]
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
                            "獲得率": f"{r['wins']/r['total_years']*100:.1f}%",
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
                            "1位企業": top_by_year[year]["company"],
                            "得点": f"{top_by_year[year]['score']}点"
                        }
                        for year in sorted(top_by_year.keys(), reverse=True)
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
                    for year, score in sorted(avg_scores.items())
                ])
                import altair as alt
                chart = alt.Chart(avg_df).mark_line(point=True).encode(
                    x=alt.X('年度:O', title='年度'),
                    y=alt.Y('平均得点:Q', title='平均得点', scale=alt.Scale(domain=[60, 80]))
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
                    chart = alt.Chart(chart_df).mark_line(point=True).encode(
                        x=alt.X('年度:O', title='年度'),
                        y=alt.Y('得点:Q', title='得点', scale=alt.Scale(domain=[60, 80])),
                        color=alt.Color('企業名:N', title='企業名'),
                        tooltip=['年度', '企業名', '得点']
                    ).properties(height=400)
                    st.altair_chart(chart, use_container_width=True)

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

    with tab4:
        st.header("📊 総合ランキング（経年詳細）")

        # トップに歴代記録を表示
        records = historical_data.get("historical_records", {})
        if records:
            display_historical_summary(records)
            display_consecutive_wins_compact(records)
            st.divider()

        if overall_data:
            # 年度ごとに全データを表示（アップロードデータをマーク）
            for year in sorted(overall_data.keys(), reverse=True):
                source_mark = "📁" if year in uploaded_years else "🌐"
                with st.expander(f"{source_mark} {year}年", expanded=(year == max(overall_data.keys()))):
                    df = pd.DataFrame(overall_data[year])
                    st.dataframe(df, use_container_width=True)

            # 経年比較テーブル
            st.subheader("📈 経年比較（全社得点推移）")

            companies = set()
            for year_data in overall_data.values():
                for item in year_data:
                    companies.add(item.get("company", ""))

            comparison_data = []
            for company in sorted(companies):
                row = {"企業名": company}
                for year in sorted(overall_data.keys()):
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

    with tab5:
        st.header("📋 評価項目別ランキング（経年）")

        # トップに評価項目別の連続1位記録
        item_trends = historical_data.get("item_trends", {})
        if item_trends:
            st.subheader("📋 評価項目別 連続1位記録（上位5件）")
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
                st.dataframe(pd.DataFrame(item_records[:5]), use_container_width=True, hide_index=True)
            st.divider()

        if item_data:
            for item_name, year_data in item_data.items():
                with st.expander(f"📌 {item_name}", expanded=False):
                    if isinstance(year_data, dict):
                        for year in sorted(year_data.keys(), reverse=True):
                            st.markdown(f"**{year}年**")
                            df = pd.DataFrame(year_data[year])
                            st.dataframe(df, use_container_width=True)

                        if len(year_data) > 1:
                            st.markdown("**📈 1位の推移**")
                            history = []
                            for year in sorted(year_data.keys(), reverse=True):
                                if year_data[year]:
                                    top = year_data[year][0]
                                    history.append({
                                        "年度": year,
                                        "1位": top.get("company", "-"),
                                        "得点": top.get("score", "-")
                                    })
                            if history:
                                st.dataframe(pd.DataFrame(history), use_container_width=True)
                    else:
                        df = pd.DataFrame(year_data)
                        st.dataframe(df, use_container_width=True)
        else:
            st.info("評価項目別データは取得できませんでした")

    with tab6:
        st.header("🏷️ 部門別ランキング（経年）")

        # トップに部門別の連続1位記録
        dept_trends = historical_data.get("dept_trends", {})
        if dept_trends:
            st.subheader("🏷️ 部門別 連続1位記録（上位5件）")
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
                st.dataframe(pd.DataFrame(dept_records[:5]), use_container_width=True, hide_index=True)
            st.divider()

        if dept_data:
            for dept_name, year_data in dept_data.items():
                with st.expander(f"📌 {dept_name}", expanded=False):
                    if isinstance(year_data, dict):
                        for year in sorted(year_data.keys(), reverse=True):
                            st.markdown(f"**{year}年**")
                            df = pd.DataFrame(year_data[year])
                            st.dataframe(df, use_container_width=True)

                        if len(year_data) > 1:
                            st.markdown("**📈 1位の推移**")
                            history = []
                            for year in sorted(year_data.keys(), reverse=True):
                                if year_data[year]:
                                    top = year_data[year][0]
                                    history.append({
                                        "年度": year,
                                        "1位": top.get("company", "-"),
                                        "得点": top.get("score", "-")
                                    })
                            if history:
                                st.dataframe(pd.DataFrame(history), use_container_width=True)
        else:
            st.info("部門別データは存在しないか取得できませんでした")

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
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "URL": st.column_config.LinkColumn("URL", display_text="🔗 リンクを開く")
                    }
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
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "URL": st.column_config.LinkColumn("URL", display_text="🔗 リンクを開く")
                    }
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
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "URL": st.column_config.LinkColumn("URL", display_text="🔗 リンクを開く")
                    }
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
st.sidebar.markdown("🔧 **バージョン**: 3.6")

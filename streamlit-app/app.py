# -*- coding: utf-8 -*-
"""
オリコン顧客満足度 TOPICS出しアプリ
Streamlit版 v1.0
"""

import streamlit as st
import pandas as pd
from scraper import OriconScraper
from analyzer import TopicsAnalyzer

# ページ設定
st.set_page_config(
    page_title="オリコン TOPICS出し",
    page_icon="📰",
    layout="wide"
)

# タイトル
st.title("📰 オリコン顧客満足度 TOPICS出し")
st.markdown("プレスリリースの見出しトピックス候補を自動生成します")

# サイドバー
st.sidebar.header("⚙️ 設定")

# ランキング選択
# 注意: URLは rank-xxx または rank_xxx の形式がある
ranking_options = {
    "携帯キャリア": "mobile-carrier",
    "格安SIM": "mvno",
    "FX": "_fx",  # rank_fx
    "銀行カードローン": "card-loan",
    "ノンバンクカードローン": "card-loan/nonbank",
    "ネット証券": "_certificate",  # rank_certificate
    "iDeCo証券会社": "ideco",
    "自動車保険": "car-insurance",
    "生命保険": "life-insurance",
    "クレジットカード": "creditcard",
    "転職サイト": "recruit",
    "英会話スクール": "english-school",
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
current_year = 2024
start_year = 2006  # オリコン顧客満足度調査開始年

year_option = st.sidebar.radio(
    "取得する年度範囲",
    ["直近3年", "直近5年", "全年度（2006年〜）", "カスタム範囲"]
)

if year_option == "直近3年":
    year_range = (current_year - 2, current_year)
elif year_option == "直近5年":
    year_range = (current_year - 4, current_year)
elif year_option == "全年度（2006年〜）":
    year_range = (start_year, current_year)
else:  # カスタム範囲
    year_range = st.sidebar.slider(
        "年度範囲を選択",
        min_value=start_year,
        max_value=current_year,
        value=(current_year - 4, current_year)
    )

# 実行ボタン
if st.sidebar.button("🚀 TOPICS出し実行", type="primary"):

    if not ranking_slug:
        st.error("ランキングのURL名を入力してください")
    else:
        # プログレスバー
        progress_bar = st.progress(0)
        status_text = st.empty()

        # デバッグログ表示エリア
        debug_expander = st.expander("🔍 デバッグログ", expanded=True)
        debug_logs = []

        def log(message):
            debug_logs.append(message)
            with debug_expander:
                st.text("\n".join(debug_logs))

        try:
            # スクレイパー初期化
            log(f"[INFO] スクレイパー初期化: {ranking_slug} ({ranking_name})")
            scraper = OriconScraper(ranking_slug, ranking_name)
            subpath_info = f" + subpath: {scraper.subpath}" if scraper.subpath else ""
            log(f"[INFO] URL prefix: {scraper.url_prefix}{subpath_info}")

            # Step 1: 総合ランキング取得
            status_text.text(f"📊 総合ランキングを取得中... ({year_range[0]}年〜{year_range[1]}年)")
            progress_bar.progress(10)

            overall_data = scraper.get_overall_rankings(year_range)
            log(f"[OK] 総合ランキング: {len(overall_data)}年分取得")
            for year, data in overall_data.items():
                log(f"  - {year}年: {len(data)}社")
            progress_bar.progress(30)

            # Step 2: 評価項目別取得（経年）
            status_text.text(f"📋 評価項目別データを取得中... ({year_range[0]}年〜{year_range[1]}年)")
            item_data = scraper.get_evaluation_items(year_range)
            log(f"[OK] 評価項目別: {len(item_data)}項目")
            for item_name in item_data.keys():
                log(f"  - {item_name}")
            progress_bar.progress(50)

            # Step 3: 部門別取得（経年）
            status_text.text(f"🏷️ 部門別データを取得中... ({year_range[0]}年〜{year_range[1]}年)")
            dept_data = scraper.get_departments(year_range)
            log(f"[OK] 部門別: {len(dept_data)}部門")
            for dept_name in dept_data.keys():
                log(f"  - {dept_name}")
            progress_bar.progress(70)

            # Step 4: 分析実行
            status_text.text("🔍 TOPICS分析中...")
            analyzer = TopicsAnalyzer(overall_data, item_data, ranking_name)
            topics = analyzer.analyze()
            progress_bar.progress(90)

            # 完了
            status_text.text("✅ 完了!")
            progress_bar.progress(100)

            # 結果表示
            st.success(f"✅ {ranking_name}のTOPICS出しが完了しました")

            # タブで結果表示
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "⭐ 推奨TOPICS",
                "📊 総合ランキング（経年）",
                "📋 評価項目別",
                "🏷️ 部門別",
                "🎯 見出し案",
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
                st.header("📊 総合ランキング（経年詳細）")
                if overall_data:
                    # 年度ごとに全データを表示
                    for year in sorted(overall_data.keys(), reverse=True):
                        with st.expander(f"📅 {year}年", expanded=(year == max(overall_data.keys()))):
                            df = pd.DataFrame(overall_data[year])
                            st.dataframe(df, use_container_width=True)

                    # 経年比較テーブル（1位〜4位の推移）
                    st.subheader("📈 経年比較（全社得点推移）")

                    # 企業ごとの経年データを集計
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

            with tab3:
                st.header("📋 評価項目別ランキング（経年）")
                if item_data:
                    for item_name, year_data in item_data.items():
                        with st.expander(f"📌 {item_name}", expanded=False):
                            if isinstance(year_data, dict):
                                # 経年データの場合
                                for year in sorted(year_data.keys(), reverse=True):
                                    st.markdown(f"**{year}年**")
                                    df = pd.DataFrame(year_data[year])
                                    st.dataframe(df, use_container_width=True)

                                # 経年比較（1位の推移）
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
                                # 旧形式（単年データ）の場合
                                df = pd.DataFrame(year_data)
                                st.dataframe(df, use_container_width=True)
                else:
                    st.info("評価項目別データは取得できませんでした")

            with tab4:
                st.header("🏷️ 部門別ランキング（経年）")
                if dept_data:
                    for dept_name, year_data in dept_data.items():
                        with st.expander(f"📌 {dept_name}", expanded=False):
                            if isinstance(year_data, dict):
                                for year in sorted(year_data.keys(), reverse=True):
                                    st.markdown(f"**{year}年**")
                                    df = pd.DataFrame(year_data[year])
                                    st.dataframe(df, use_container_width=True)

                                # 経年比較（1位の推移）
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

            with tab5:
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

            with tab6:
                st.header("📎 参考資料（使用したURL）")

                # 総合ランキングURL
                st.subheader("総合ランキング")
                overall_urls = scraper.used_urls.get("overall", [])
                if overall_urls:
                    for item in overall_urls:
                        status = "✅" if item["status"] == "success" else "❌"
                        st.markdown(f"{status} **{item['year']}年**: [{item['url']}]({item['url']})")

                # 評価項目別URL
                st.subheader("評価項目別ランキング")
                item_urls = scraper.used_urls.get("items", [])
                if item_urls:
                    with st.expander("評価項目別URL一覧", expanded=False):
                        for item in item_urls:
                            status = "✅" if item["status"] == "success" else "❌"
                            st.markdown(f"{status} **{item['name']}**: [{item['url']}]({item['url']})")

                # 部門別URL
                st.subheader("部門別ランキング")
                dept_urls = scraper.used_urls.get("departments", [])
                if dept_urls:
                    with st.expander("部門別URL一覧", expanded=False):
                        for item in dept_urls:
                            status = "✅" if item["status"] == "success" else "❌"
                            st.markdown(f"{status} **{item['name']}**: [{item['url']}]({item['url']})")
                else:
                    st.info("部門別データは存在しませんでした")

                # データソース
                st.divider()
                st.markdown("**データソース**: [オリコン顧客満足度ランキング](https://life.oricon.co.jp/)")

        except Exception as e:
            st.error(f"エラーが発生しました: {str(e)}")
            st.exception(e)

# フッター
st.sidebar.divider()
st.sidebar.markdown("---")
st.sidebar.markdown("📌 **データソース**: life.oricon.co.jp")
st.sidebar.markdown("🔧 **バージョン**: 1.0")

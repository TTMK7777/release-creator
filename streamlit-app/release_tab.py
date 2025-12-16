# -*- coding: utf-8 -*-
"""
プレスリリースタブ モジュール (v1.0)
app.py のタブとして統合するためのヘルパーモジュール

使い方:
1. app.py の import セクションに以下を追加:
   from release_tab import render_release_tab, RELEASE_FEATURES_AVAILABLE

2. タブ定義を更新:
   tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
       "⭐ 推奨TOPICS",
       "🏆 歴代記録・得点推移",
       "📊 総合ランキング",
       "📋 評価項目別",
       "🏷️ 部門別",
       "📝 プレスリリース作成",  # 新規追加
       "📎 参考資料"
   ])

3. 新しいタブの中で:
   with tab6:
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
"""

import streamlit as st
import pandas as pd
import logging
from io import BytesIO
from datetime import datetime
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)

# 機能モジュールのインポート
try:
    from validator import (
        validate_release_data,
        format_validation_report,
        ValidationLevel,
        ValidationResult
    )
    from release_generator import (
        generate_release,
        ReleaseGenerator,
        ReleaseContent
    )
    RELEASE_FEATURES_AVAILABLE = True
except ImportError as e:
    logger.warning(f"プレスリリース機能モジュールが見つかりません: {e}")
    RELEASE_FEATURES_AVAILABLE = False

# Word/画像出力モジュールのインポート
try:
    from word_generator import generate_word_release, WordGenerator
    WORD_AVAILABLE = True
except ImportError as e:
    logger.warning(f"Word出力モジュールが見つかりません: {e}")
    WORD_AVAILABLE = False

try:
    from image_generator import (
        TableImageGenerator,
        generate_ranking_image,
        ExcelTemplateImageGenerator
    )
    IMAGE_AVAILABLE = True
except ImportError as e:
    logger.warning(f"画像出力モジュールが見つかりません: {e}")
    IMAGE_AVAILABLE = False


def render_release_tab(
    ranking_name: str,
    overall_data: Dict,
    item_data: Dict,
    dept_data: Dict,
    historical_data: Dict,
    excel_upload_data: Optional[Dict] = None
):
    """プレスリリースタブをレンダリング

    Args:
        ranking_name: ランキング名
        overall_data: 総合ランキングデータ {year: [entries]}
        item_data: 評価項目別データ
        dept_data: 部門別データ
        historical_data: 歴代記録データ
        excel_upload_data: アップロードされたExcelデータ（任意）
    """
    st.header("📝 プレスリリース作成")

    if not RELEASE_FEATURES_AVAILABLE:
        st.error("プレスリリース機能のモジュールが見つかりません。validator.py と release_generator.py が必要です。")
        return

    # 年度を取得
    available_years = sorted(overall_data.keys(), reverse=True) if overall_data else []
    if not available_years:
        st.warning("ランキングデータがありません。先にTOPICS出しを実行してください。")
        return

    latest_year = available_years[0]

    # タブ内のサブセクション
    sub_tab1, sub_tab2, sub_tab3, sub_tab4, sub_tab5 = st.tabs([
        "✅ 正誤チェック",
        "📊 表の自動生成",
        "📝 文章の自動生成",
        "📄 Word出力",
        "🖼️ 画像出力"
    ])

    # ========================================
    # 1. 正誤チェックタブ
    # ========================================
    with sub_tab1:
        st.subheader("✅ 正誤チェック")
        st.caption("データの正確性を自動検証します")

        # 検証実行ボタン
        if st.button("🔍 正誤チェックを実行", key="run_validation"):
            with st.spinner("検証中..."):
                # ExcelデータとWebデータを分離
                excel_data = excel_upload_data if excel_upload_data else {}
                web_data = overall_data

                # 検証実行
                result = validate_release_data(
                    excel_data=excel_data,
                    web_data=web_data,
                    ranking_name=ranking_name
                )

                # 結果をセッションに保存
                st.session_state['validation_result'] = result

        # 検証結果の表示
        if 'validation_result' in st.session_state:
            result = st.session_state['validation_result']

            # サマリー
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if result.is_valid:
                    st.success(f"✅ 検証OK")
                else:
                    st.error(f"❌ 要修正")
            with col2:
                st.metric("エラー", result.summary.get('ERROR', 0))
            with col3:
                st.metric("警告", result.summary.get('WARNING', 0))
            with col4:
                st.metric("情報", result.summary.get('INFO', 0))

            st.divider()

            # エラー詳細
            errors = result.get_errors()
            if errors:
                st.subheader("❌ エラー（修正が必要）")
                for i, issue in enumerate(errors, 1):
                    with st.expander(f"{i}. [{issue.category}] {issue.message}", expanded=True):
                        if issue.expected:
                            st.write(f"**期待値**: {issue.expected}")
                        if issue.actual:
                            st.write(f"**実際値**: {issue.actual}")
                        if issue.suggestion:
                            st.info(f"💡 提案: {issue.suggestion}")

            # 警告詳細
            warnings = result.get_warnings()
            if warnings:
                st.subheader("⚠️ 警告（確認推奨）")
                for i, issue in enumerate(warnings, 1):
                    with st.expander(f"{i}. [{issue.category}] {issue.message}"):
                        if issue.suggestion:
                            st.info(f"💡 提案: {issue.suggestion}")
                        if issue.context:
                            st.caption(f"詳細: {issue.context}")

            # 情報
            infos = [i for i in result.issues if i.level == ValidationLevel.INFO]
            if infos:
                with st.expander(f"ℹ️ 情報 ({len(infos)}件)"):
                    for issue in infos:
                        st.write(f"- {issue.message}")

            # レポートダウンロード
            st.divider()
            report_text = format_validation_report(result)
            st.download_button(
                label="📄 検証レポートをダウンロード",
                data=report_text,
                file_name=f"validation_report_{ranking_name}_{datetime.now().strftime('%Y%m%d')}.txt",
                mime="text/plain"
            )

    # ========================================
    # 2. 表の自動生成タブ
    # ========================================
    with sub_tab2:
        st.subheader("📊 表の自動生成")
        st.caption("プレスリリース用のランキング表を生成します")

        # オプション
        col1, col2 = st.columns(2)
        with col1:
            target_year = st.selectbox(
                "対象年度",
                available_years,
                index=0,
                key="table_target_year"
            )
            show_score = st.checkbox("得点を表示", value=True, key="show_score")
        with col2:
            display_count = st.slider(
                "表示企業数",
                min_value=3,
                max_value=20,
                value=10,
                key="display_count"
            )
            show_prev_rank = st.checkbox("前年順位を表示", value=False, key="show_prev_rank")

        if st.button("📊 表を生成", key="generate_table"):
            with st.spinner("表を生成中..."):
                # プレスリリース生成
                content = generate_release(
                    ranking_name=ranking_name,
                    year=target_year,
                    overall_data=overall_data,
                    item_data=item_data,
                    dept_data=dept_data,
                    historical_data=historical_data
                )

                st.session_state['release_content'] = content

        # 生成結果の表示
        if 'release_content' in st.session_state:
            content = st.session_state['release_content']

            st.subheader(f"📊 {content.title}")

            # 総合ランキング表
            if content.overall_table is not None and not content.overall_table.empty:
                st.write("**総合ランキング**")
                st.dataframe(content.overall_table, use_container_width=True, hide_index=True)

            # 評価項目別表
            if content.item_tables:
                st.write("**評価項目別ランキング**")
                for item_name, df in content.item_tables.items():
                    if not df.empty:
                        with st.expander(f"📋 {item_name}"):
                            st.dataframe(df, use_container_width=True, hide_index=True)

            # 部門別表
            if content.dept_tables:
                st.write("**部門別ランキング**")
                for dept_name, df in content.dept_tables.items():
                    if not df.empty:
                        with st.expander(f"🏷️ {dept_name}"):
                            st.dataframe(df, use_container_width=True, hide_index=True)

            # Excelダウンロード
            st.divider()
            try:
                generator = ReleaseGenerator(
                    ranking_name=ranking_name,
                    year=content.year,
                    overall_data=overall_data,
                    item_data=item_data,
                    dept_data=dept_data,
                    historical_data=historical_data
                )
                excel_buffer = generator.export_to_excel(content)
                st.download_button(
                    label="📥 Excelでダウンロード",
                    data=excel_buffer,
                    file_name=f"release_{ranking_name}_{content.year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                logger.error(f"Excelエクスポートエラー: {e}")
                st.warning("Excelエクスポートに失敗しました")

    # ========================================
    # 3. 文章の自動生成タブ
    # ========================================
    with sub_tab3:
        st.subheader("📝 文章の自動生成")
        st.caption("テンプレートベースでプレスリリース文章を生成します")

        # 対象年度
        text_target_year = st.selectbox(
            "対象年度",
            available_years,
            index=0,
            key="text_target_year"
        )

        if st.button("📝 文章を生成", key="generate_text"):
            with st.spinner("文章を生成中..."):
                content = generate_release(
                    ranking_name=ranking_name,
                    year=text_target_year,
                    overall_data=overall_data,
                    item_data=item_data,
                    dept_data=dept_data,
                    historical_data=historical_data
                )

                st.session_state['text_content'] = content

        # 生成結果の表示
        if 'text_content' in st.session_state:
            content = st.session_state['text_content']

            # ハイライト（見出し候補）
            if content.highlights:
                st.subheader("🎯 ハイライト（見出し候補）")
                for i, h in enumerate(content.highlights, 1):
                    st.markdown(f"**{i}.** {h}")

            st.divider()

            # 本文
            st.subheader("📝 本文")
            for p in content.paragraphs:
                st.write(p)
                st.write("")  # 段落間の空行

            # コピー用テキスト
            st.divider()
            st.subheader("📋 コピー用テキスト")

            copy_text = f"【{content.title}】\n\n"
            if content.highlights:
                copy_text += "■ ハイライト\n"
                copy_text += "\n".join([f"・{h}" for h in content.highlights])
                copy_text += "\n\n"
            copy_text += "■ 本文\n"
            copy_text += "\n\n".join(content.paragraphs)

            st.text_area(
                "コピー用",
                copy_text,
                height=400,
                label_visibility="collapsed"
            )

            # ダウンロード
            st.download_button(
                label="📄 テキストでダウンロード",
                data=copy_text,
                file_name=f"release_text_{ranking_name}_{content.year}.txt",
                mime="text/plain"
            )

    # ========================================
    # 4. Word出力タブ
    # ========================================
    with sub_tab4:
        st.subheader("📄 Word出力")
        st.caption("Wordテンプレートを使用してプレスリリース文書を生成します")

        if not WORD_AVAILABLE:
            st.warning("Word出力モジュールが見つかりません。word_generator.py が必要です。")
            st.info("python-docx をインストールしてください: `pip install python-docx`")
        else:
            # オプション設定
            col1, col2 = st.columns(2)
            with col1:
                word_target_year = st.selectbox(
                    "対象年度",
                    available_years,
                    index=0,
                    key="word_target_year"
                )
                word_month = st.number_input(
                    "発表月",
                    min_value=1,
                    max_value=12,
                    value=datetime.now().month,
                    key="word_month"
                )
            with col2:
                word_day = st.number_input(
                    "発表日",
                    min_value=1,
                    max_value=31,
                    value=datetime.now().day,
                    key="word_day"
                )
                include_table = st.checkbox(
                    "ランキング表を含める",
                    value=True,
                    key="include_table"
                )

            # TOPICS入力
            st.write("**TOPICS（任意）**")
            topics_text = st.text_area(
                "TOPICS（1行1項目）",
                height=100,
                key="word_topics",
                placeholder="例:\nSBI証券が3年連続1位\n楽天証券と同率1位を達成"
            )
            topics_list = [t.strip() for t in topics_text.split("\n") if t.strip()] if topics_text else []

            # ハイライト入力
            highlight_text = st.text_input(
                "ハイライト（見出し）",
                key="word_highlight",
                placeholder="例: SBI証券が3年連続1位、楽天証券と同率"
            )
            highlights_list = [highlight_text] if highlight_text else []

            if st.button("📄 Word文書を生成", key="generate_word"):
                with st.spinner("Word文書を生成中..."):
                    try:
                        # 総合ランキングデータを取得
                        year_data = overall_data.get(word_target_year, [])

                        # Word生成
                        word_buffer = generate_word_release(
                            ranking_name=ranking_name,
                            year=word_target_year,
                            overall_data=year_data,
                            topics=topics_list,
                            highlights=highlights_list,
                            month=word_month,
                            day=word_day,
                            include_table=include_table
                        )

                        if word_buffer:
                            st.session_state['word_buffer'] = word_buffer
                            st.session_state['word_filename'] = f"release_{ranking_name}_{word_target_year}年{word_month}月.docx"
                            st.success("✅ Word文書の生成が完了しました")
                        else:
                            st.error("Word文書の生成に失敗しました。テンプレートファイルを確認してください。")

                    except Exception as e:
                        logger.error(f"Word生成エラー: {e}")
                        st.error(f"エラーが発生しました: {e}")

            # ダウンロードボタン
            if 'word_buffer' in st.session_state:
                st.divider()
                st.download_button(
                    label="📥 Wordファイルをダウンロード",
                    data=st.session_state['word_buffer'].getvalue(),
                    file_name=st.session_state.get('word_filename', 'release.docx'),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    # ========================================
    # 5. 画像出力タブ
    # ========================================
    with sub_tab5:
        st.subheader("🖼️ 画像出力")
        st.caption("ランキング表を画像として出力します")

        if not IMAGE_AVAILABLE:
            st.warning("画像出力モジュールが見つかりません。image_generator.py が必要です。")
            st.info("matplotlib をインストールしてください: `pip install matplotlib`")
        else:
            # オプション設定
            col1, col2 = st.columns(2)
            with col1:
                img_target_year = st.selectbox(
                    "対象年度",
                    available_years,
                    index=0,
                    key="img_target_year"
                )
                table_type = st.selectbox(
                    "表タイプ",
                    ["総合ランキング", "前年比較", "TOP3強調"],
                    key="table_type"
                )
            with col2:
                display_rows = st.slider(
                    "表示行数",
                    min_value=3,
                    max_value=15,
                    value=10,
                    key="img_display_rows"
                )
                img_show_score = st.checkbox(
                    "得点を表示",
                    value=True,
                    key="img_show_score"
                )

            # 画像スタイル設定
            with st.expander("📐 詳細設定"):
                col1, col2 = st.columns(2)
                with col1:
                    fig_width = st.slider("画像幅", 6, 16, 10, key="fig_width")
                    font_size = st.slider("フォントサイズ", 8, 16, 11, key="font_size")
                with col2:
                    dpi = st.selectbox("解像度(DPI)", [72, 150, 300], index=1, key="dpi")

            if st.button("🖼️ 画像を生成", key="generate_image"):
                with st.spinner("画像を生成中..."):
                    try:
                        # データ取得
                        year_data = overall_data.get(img_target_year, [])
                        prev_year = img_target_year - 1
                        prev_year_data = overall_data.get(prev_year, [])

                        if not year_data:
                            st.warning(f"{img_target_year}年のデータがありません")
                        else:
                            # 画像生成
                            generator = TableImageGenerator(
                                ranking_name=ranking_name,
                                year=img_target_year
                            )

                            if table_type == "総合ランキング":
                                img_buffer = generator.generate_overall_table(
                                    data=year_data[:display_rows],
                                    show_score=img_show_score,
                                    figsize=(fig_width, display_rows * 0.5 + 2),
                                    dpi=dpi
                                )
                            elif table_type == "前年比較":
                                img_buffer = generator.generate_comparison_table(
                                    current_data=year_data[:display_rows],
                                    prev_data=prev_year_data,
                                    prev_year=prev_year,
                                    figsize=(fig_width + 2, display_rows * 0.5 + 2),
                                    dpi=dpi
                                )
                            else:  # TOP3強調
                                img_buffer = generator.generate_top3_highlight(
                                    data=year_data[:3],
                                    figsize=(fig_width, 4),
                                    dpi=dpi
                                )

                            if img_buffer:
                                st.session_state['img_buffer'] = img_buffer
                                st.session_state['img_filename'] = f"ranking_{ranking_name}_{img_target_year}_{table_type}.png"
                                st.success("✅ 画像の生成が完了しました")

                                # プレビュー表示
                                st.image(img_buffer, caption=f"{ranking_name} {img_target_year}年 {table_type}")

                    except Exception as e:
                        logger.error(f"画像生成エラー: {e}")
                        st.error(f"エラーが発生しました: {e}")

            # ダウンロードボタン
            if 'img_buffer' in st.session_state:
                st.divider()
                st.download_button(
                    label="📥 画像をダウンロード",
                    data=st.session_state['img_buffer'].getvalue(),
                    file_name=st.session_state.get('img_filename', 'ranking.png'),
                    mime="image/png"
                )


# ========================================
# スタンドアロン実行用（テスト）
# ========================================
if __name__ == "__main__":
    st.set_page_config(page_title="プレスリリース作成テスト", layout="wide")

    st.title("📝 プレスリリース作成機能テスト")

    # テストデータ
    test_overall = {
        2026: [
            {"rank": 1, "company": "SBI証券", "score": 68.9},
            {"rank": 1, "company": "楽天証券", "score": 68.9},
            {"rank": 3, "company": "マネックス証券", "score": 67.5},
            {"rank": 4, "company": "松井証券", "score": 66.0},
            {"rank": 5, "company": "auカブコム証券", "score": 65.5},
        ],
        2025: [
            {"rank": 1, "company": "SBI証券", "score": 68.5},
            {"rank": 2, "company": "楽天証券", "score": 68.0},
            {"rank": 3, "company": "マネックス証券", "score": 67.0},
        ]
    }

    test_item_data = {
        "取引手数料": {
            2026: [{"rank": 1, "company": "SBI証券", "score": 72.0}]
        }
    }

    render_release_tab(
        ranking_name="ネット証券",
        overall_data=test_overall,
        item_data=test_item_data,
        dept_data={},
        historical_data={}
    )

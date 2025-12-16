# -*- coding: utf-8 -*-
"""
Word出力モジュール (v2.0)
プレスリリースのWord文書を生成

テンプレート形式: {{KEY}} プレースホルダー方式
使用テンプレート: _archive/テンプレート/【テンプレ】プレスリリース_v2.docx
"""

import os
import re
import logging
from io import BytesIO
from typing import Dict, List, Any, Optional
from datetime import datetime

from docx import Document

logger = logging.getLogger(__name__)

# テンプレートパス
TEMPLATE_DIR = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "_archive", "テンプレート"
)
WORD_TEMPLATE_PATH = os.path.join(
    TEMPLATE_DIR,
    "【テンプレ】プレスリリース_v3.docx"
)

# 曜日マッピング
WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


class WordGenerator:
    """Wordプレスリリース生成クラス (v2.0 - {{KEY}}形式対応)"""

    def __init__(
        self,
        ranking_name: str,
        year: int,
        month: int = None,
        day: int = None
    ):
        """
        Args:
            ranking_name: ランキング名（例: "ネット証券"）
            year: 発表年度
            month: 発表月（デフォルト: 現在月）
            day: 発表日（デフォルト: 現在日）
        """
        self.ranking_name = ranking_name
        self.year = year
        self.month = month or datetime.now().month
        self.day = day or datetime.now().day
        self.doc = None

        # 日付計算
        try:
            dt = datetime(year, self.month, self.day)
            self.weekday = WEEKDAY_JP[dt.weekday()]
        except:
            self.weekday = ""

    def load_template(self, template_path: str = None) -> bool:
        """テンプレートを読み込む"""
        path = template_path or WORD_TEMPLATE_PATH

        if not os.path.exists(path):
            logger.error(f"テンプレートが見つかりません: {path}")
            return False

        try:
            self.doc = Document(path)
            logger.info(f"テンプレート読み込み成功: {path}")
            return True
        except Exception as e:
            logger.error(f"テンプレート読み込みエラー: {e}")
            return False

    def _replace_in_paragraph(self, para, replacements: Dict[str, str]):
        """段落内のプレースホルダーを置換"""
        text = para.text
        modified = False

        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"  # {{KEY}} 形式
            if placeholder in text:
                text = text.replace(placeholder, str(value) if value else "")
                modified = True

        if modified and para.runs:
            # 最初のRunに全テキストを設定（書式保持）
            for run in para.runs:
                run.text = ""
            para.runs[0].text = text

    def replace_placeholders(
        self,
        overall_data: List[Dict] = None,
        topics: List[str] = None,
        topic_details: List[str] = None,
        highlights: List[str] = None,
        subheadline: str = None,
        sample_size: int = None,
        min_sample: int = 100,
        company_count: int = None,
        ranking_url: str = None
    ):
        """プレースホルダーを置換

        Args:
            overall_data: 総合ランキングデータ
            topics: TOPICSリスト（タイトル）
            topic_details: TOPICS詳細リスト
            highlights: ハイライト（見出し）リスト
            subheadline: サブ見出し
            sample_size: 回答者数
            min_sample: 規定人数
            company_count: 調査企業数
            ranking_url: ランキングページURL
        """
        if not self.doc:
            logger.error("テンプレートが読み込まれていません")
            return

        # 置換マッピングを構築
        replacements = {
            # 日付関連
            "DATE": f"{self.year}年{self.month}月{self.day}日",
            "YEAR": f"{self.year}年",
            "WEEKDAY": self.weekday,
            "RELEASE_DATE_SLASH": f"{self.year}/{self.month:02d}/{self.day:02d}",

            # ランキング情報
            "RANKING_NAME": self.ranking_name,

            # 見出し
            "HEADLINE": highlights[0] if highlights else "",
            "SUBHEADLINE": subheadline or "",

            # TOPICS
            "TOPIC_1": topics[0] if topics and len(topics) > 0 else "",
            "TOPIC_1_DETAIL": topic_details[0] if topic_details and len(topic_details) > 0 else "",
            "TOPIC_2": topics[1] if topics and len(topics) > 1 else "",
            "TOPIC_2_DETAIL": topic_details[1] if topic_details and len(topic_details) > 1 else "",
            "TOPIC_3": topics[2] if topics and len(topics) > 2 else "",
            "TOPIC_3_DETAIL": topic_details[2] if topic_details and len(topic_details) > 2 else "",

            # 調査概要
            "SAMPLE_SIZE": sample_size or "",
            "MIN_SAMPLE": min_sample,
            "COMPANY_COUNT": company_count or "",
            "RANKING_URL": ranking_url or "",
        }

        # 全段落に対して置換を実行
        for para in self.doc.paragraphs:
            self._replace_in_paragraph(para, replacements)

        # テーブル内のセルも置換
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._replace_in_paragraph(para, replacements)

    def add_ranking_table(
        self,
        overall_data: List[Dict],
        position: int = None,
        title: str = None
    ):
        """ランキング表を追加

        Args:
            overall_data: ランキングデータ
            position: 挿入位置（段落インデックス）
            title: 表のタイトル
        """
        if not self.doc or not overall_data:
            return

        # 表を作成
        table = self.doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'

        # ヘッダー行
        header_cells = table.rows[0].cells
        header_cells[0].text = "順位"
        header_cells[1].text = "企業名"
        header_cells[2].text = "得点"

        # ヘッダーの書式設定
        for cell in header_cells:
            if cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True

        # データ行
        for entry in overall_data[:10]:  # TOP10
            row_cells = table.add_row().cells
            row_cells[0].text = f"{entry.get('rank', '-')}位"
            row_cells[1].text = entry.get('company', '')
            score = entry.get('score')
            row_cells[2].text = f"{score}点" if score is not None else "-"

    def generate(
        self,
        overall_data: List[Dict] = None,
        topics: List[str] = None,
        topic_details: List[str] = None,
        highlights: List[str] = None,
        subheadline: str = None,
        sample_size: int = None,
        min_sample: int = 100,
        company_count: int = None,
        ranking_url: str = None,
        include_table: bool = False
    ) -> Optional[BytesIO]:
        """Word文書を生成

        Args:
            overall_data: 総合ランキングデータ
            topics: TOPICSリスト
            topic_details: TOPICS詳細リスト
            highlights: ハイライトリスト
            subheadline: サブ見出し
            sample_size: 回答者数
            min_sample: 規定人数
            company_count: 調査企業数
            ranking_url: ランキングページURL
            include_table: ランキング表を含めるか

        Returns:
            BytesIOオブジェクト（Wordファイル）
        """
        # テンプレート読み込み
        if not self.load_template():
            return None

        # プレースホルダー置換
        self.replace_placeholders(
            overall_data=overall_data,
            topics=topics,
            topic_details=topic_details,
            highlights=highlights,
            subheadline=subheadline,
            sample_size=sample_size,
            min_sample=min_sample,
            company_count=company_count,
            ranking_url=ranking_url
        )

        # 表を追加（オプション）
        if include_table and overall_data:
            self.add_ranking_table(overall_data)

        # BytesIOに保存
        output = BytesIO()
        self.doc.save(output)
        output.seek(0)

        return output


def generate_word_release(
    ranking_name: str,
    year: int,
    overall_data: List[Dict] = None,
    topics: List[str] = None,
    topic_details: List[str] = None,
    highlights: List[str] = None,
    subheadline: str = None,
    month: int = None,
    day: int = None,
    sample_size: int = None,
    company_count: int = None,
    ranking_url: str = None,
    include_table: bool = False
) -> Optional[BytesIO]:
    """Wordプレスリリースを生成（簡易インターフェース）

    Args:
        ranking_name: ランキング名
        year: 発表年度
        overall_data: 総合ランキングデータ
        topics: TOPICSリスト
        topic_details: TOPICS詳細リスト
        highlights: ハイライトリスト
        subheadline: サブ見出し
        month: 発表月
        day: 発表日
        sample_size: 回答者数
        company_count: 調査企業数
        ranking_url: ランキングページURL
        include_table: 表を含めるか

    Returns:
        BytesIOオブジェクト
    """
    generator = WordGenerator(
        ranking_name=ranking_name,
        year=year,
        month=month,
        day=day
    )

    return generator.generate(
        overall_data=overall_data,
        topics=topics,
        topic_details=topic_details,
        highlights=highlights,
        subheadline=subheadline,
        sample_size=sample_size,
        company_count=company_count,
        ranking_url=ranking_url,
        include_table=include_table
    )


# ========================================
# デバッグ用
# ========================================
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8')

    print("=== Word Generator v2.0 Test ===")

    # テストデータ
    test_data = [
        {"rank": 1, "company": "SBI証券", "score": 68.9},
        {"rank": 1, "company": "楽天証券", "score": 68.9},
        {"rank": 3, "company": "マネックス証券", "score": 67.5},
    ]

    test_topics = [
        "SBI証券と楽天証券が同率1位",
        "SBI証券が3年連続1位を達成",
        "マネックス証券が初のTOP3入り"
    ]

    test_topic_details = [
        "今年度の調査では、SBI証券と楽天証券が68.9点で並び、初の同率1位となりました。",
        "SBI証券は2024年から3年連続で総合1位を獲得。取引手数料の評価が特に高い結果となりました。",
        "マネックス証券は前年5位から3位に躍進。投資情報の充実度が評価されました。"
    ]

    test_highlights = ["SBI証券と楽天証券が同率1位、3年連続の快挙"]

    # 生成
    result = generate_word_release(
        ranking_name="ネット証券",
        year=2026,
        month=1,
        day=15,
        overall_data=test_data,
        topics=test_topics,
        topic_details=test_topic_details,
        highlights=test_highlights,
        subheadline="業界初の同率1位、手数料競争が加速",
        sample_size=5000,
        company_count=15,
        ranking_url="https://cs.oricon.co.jp/rank/netsec/",
        include_table=True
    )

    if result:
        # ファイルに保存（テスト）
        output_path = "test_release_v2.docx"
        with open(output_path, "wb") as f:
            f.write(result.getvalue())
        print(f"[OK] Generated: {output_path}")
    else:
        print("[ERROR] Generation failed")

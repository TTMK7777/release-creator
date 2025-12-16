# -*- coding: utf-8 -*-
"""
Word出力モジュール (v1.0)
プレスリリースのWord文書を生成

使用テンプレート:
- _archive/テンプレート/【テンプレ】20XX年X月発表 『ランキング名』ランキング ニュースリリース（オリコン顧客満足度調査） - コピー.docx
"""

import os
import re
import logging
from io import BytesIO
from typing import Dict, List, Any, Optional
from datetime import datetime
from copy import deepcopy

from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

logger = logging.getLogger(__name__)

# テンプレートパス
TEMPLATE_DIR = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "_archive", "テンプレート"
)
WORD_TEMPLATE_PATH = os.path.join(
    TEMPLATE_DIR,
    "【テンプレ】20XX年X月発表 『ランキング名』ランキング ニュースリリース（オリコン顧客満足度調査） - コピー.docx"
)


class WordGenerator:
    """Wordプレスリリース生成クラス"""

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

    def load_template(self, template_path: str = None) -> bool:
        """テンプレートを読み込む

        Args:
            template_path: テンプレートファイルパス（省略時はデフォルト）

        Returns:
            成功フラグ
        """
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

    def replace_text_in_paragraph(self, paragraph, old_text: str, new_text: str):
        """段落内のテキストを置換（書式を保持）"""
        if old_text in paragraph.text:
            # 全体のテキストを取得
            full_text = paragraph.text
            # 置換
            new_full_text = full_text.replace(old_text, new_text)

            # Runをクリアして再設定
            # 最初のRunの書式を保持
            if paragraph.runs:
                first_run = paragraph.runs[0]
                # 全Runを削除
                for run in paragraph.runs:
                    run.text = ""
                # 最初のRunに新しいテキストを設定
                first_run.text = new_full_text

    def replace_placeholders(
        self,
        overall_data: List[Dict] = None,
        topics: List[str] = None,
        highlights: List[str] = None,
        sample_size: int = None
    ):
        """プレースホルダーを置換

        Args:
            overall_data: 総合ランキングデータ
            topics: TOPICSリスト
            highlights: ハイライト（見出し）リスト
            sample_size: 回答者数
        """
        if not self.doc:
            logger.error("テンプレートが読み込まれていません")
            return

        # 日付置換
        date_str = f"{self.year}年{self.month}月{self.day}日"
        year_str = f"{self.year}年"

        for para in self.doc.paragraphs:
            # 日付
            if "20XX年X月X日" in para.text:
                self.replace_text_in_paragraph(para, "20XX年X月X日", date_str)

            # 年度（タイトル内）
            if "20XX年" in para.text:
                self.replace_text_in_paragraph(para, "20XX年", year_str)

            # ランキング名
            if "『○○○○○』" in para.text:
                self.replace_text_in_paragraph(para, "『○○○○○』", f"『{self.ranking_name}』")
            if "○○○○○" in para.text and "『" not in para.text:
                self.replace_text_in_paragraph(para, "○○○○○", self.ranking_name)

        # 見出しトピックス（パラグラフ5）
        if highlights and len(self.doc.paragraphs) > 5:
            headline_text = highlights[0] if highlights else ""
            para = self.doc.paragraphs[5]
            if "（見出しトピックス）" in para.text or "○○○○○" in para.text:
                # 見出し全体を置換
                new_text = f"（見出しトピックス）{headline_text}"
                if para.runs:
                    for run in para.runs:
                        run.text = ""
                    para.runs[0].text = new_text

        # TOPICS セクション（パラグラフ8-14あたり）
        if topics:
            topics_start = None
            for i, para in enumerate(self.doc.paragraphs):
                if "《TOPICS》" in para.text:
                    topics_start = i + 1
                    break

            if topics_start:
                topic_idx = 0
                for i in range(topics_start, min(topics_start + 10, len(self.doc.paragraphs))):
                    para = self.doc.paragraphs[i]
                    # ■で始まる行（トピックタイトル）
                    if para.text.strip().startswith("■") and topic_idx < len(topics):
                        if para.runs:
                            for run in para.runs:
                                run.text = ""
                            para.runs[0].text = f"■{topics[topic_idx]}"
                        topic_idx += 1
                    # 〇〇〇...で始まる行（トピック詳細）
                    elif "〇〇〇" in para.text:
                        # 詳細は空にするか、別途設定
                        if para.runs:
                            for run in para.runs:
                                run.text = ""

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
        highlights: List[str] = None,
        sample_size: int = None,
        include_table: bool = False
    ) -> Optional[BytesIO]:
        """Word文書を生成

        Args:
            overall_data: 総合ランキングデータ
            topics: TOPICSリスト
            highlights: ハイライトリスト
            sample_size: 回答者数
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
            highlights=highlights,
            sample_size=sample_size
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
    highlights: List[str] = None,
    month: int = None,
    day: int = None,
    include_table: bool = False
) -> Optional[BytesIO]:
    """Wordプレスリリースを生成（簡易インターフェース）

    Args:
        ranking_name: ランキング名
        year: 発表年度
        overall_data: 総合ランキングデータ
        topics: TOPICSリスト
        highlights: ハイライトリスト
        month: 発表月
        day: 発表日
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
        highlights=highlights,
        include_table=include_table
    )


# ========================================
# デバッグ用
# ========================================
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8')

    print("=== Word Generator Test ===")

    # テストデータ
    test_data = [
        {"rank": 1, "company": "SBI証券", "score": 68.9},
        {"rank": 2, "company": "楽天証券", "score": 68.0},
        {"rank": 3, "company": "マネックス証券", "score": 67.5},
    ]

    test_topics = [
        "SBI証券と楽天証券が同率1位",
        "SBI証券が3年連続1位を達成",
        "マネックス証券が初のTOP3入り"
    ]

    test_highlights = ["SBI証券が3年連続1位、楽天証券と同率"]

    # 生成
    result = generate_word_release(
        ranking_name="ネット証券",
        year=2026,
        overall_data=test_data,
        topics=test_topics,
        highlights=test_highlights,
        include_table=True
    )

    if result:
        # ファイルに保存（テスト）
        output_path = "test_release.docx"
        with open(output_path, "wb") as f:
            f.write(result.getvalue())
        print(f"[OK] Generated: {output_path}")
    else:
        print("[ERROR] Generation failed")

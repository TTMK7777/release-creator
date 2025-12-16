# -*- coding: utf-8 -*-
"""
プレスリリース生成モジュール (v1.0)
表・文章の自動生成

機能:
1. 表の自動生成（総合/評価項目/部門別）
2. 文章の自動生成（テンプレート+ルールベース）
3. Word/Excel出力
"""

import os
import logging
from io import BytesIO
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from company_master import normalize_company_name

logger = logging.getLogger(__name__)


# ========================================
# テンプレート定義（ルールベース文章生成用）
# ========================================
RELEASE_TEMPLATES = {
    # 総合ランキング1位の文章テンプレート
    "overall_winner": (
        "{year}年 オリコン顧客満足度®調査「{ranking_name}」ランキングにおいて、"
        "「{company}」が{score}点で総合{rank}位を獲得しました。"
    ),

    # 同率1位の文章テンプレート
    "overall_tie": (
        "{year}年 オリコン顧客満足度®調査「{ranking_name}」ランキングにおいて、"
        "「{companies}」が{score}点で同率{rank}位となりました。"
    ),

    # 連続1位の文章テンプレート
    "consecutive_wins": (
        "「{company}」は{years}年連続で総合1位を獲得しています"
        "（{start_year}年～{end_year}年）。"
    ),

    # 初の1位獲得
    "first_win": (
        "「{company}」は今回初めて総合1位を獲得しました。"
    ),

    # 評価項目別1位
    "item_winner": (
        "評価項目「{item_name}」では「{company}」が{score}点で1位を獲得しました。"
    ),

    # 部門別1位
    "dept_winner": (
        "「{dept_name}」部門では「{company}」が{score}点で1位となりました。"
    ),

    # 順位変動（上昇）
    "rank_up": (
        "「{company}」は前年の{prev_rank}位から{rank}位にランクアップしました。"
    ),

    # 順位変動（下降）
    "rank_down": (
        "「{company}」は前年の{prev_rank}位から{rank}位となりました。"
    ),

    # 得点上昇
    "score_up": (
        "「{company}」の得点は前年比{diff:+.1f}点（{prev_score}点→{score}点）となりました。"
    ),

    # 新規ランクイン
    "new_entry": (
        "「{company}」が今回新たにランクインし、{rank}位（{score}点）を獲得しました。"
    ),

    # 調査概要
    "survey_overview": (
        "本調査は{survey_period}に実施し、{sample_size}名から回答を得ました。"
    ),
}


# ========================================
# データクラス定義
# ========================================
@dataclass
class ReleaseContent:
    """プレスリリースコンテンツ"""
    title: str                      # タイトル
    year: int                       # 発表年度
    ranking_name: str               # ランキング名
    overall_table: pd.DataFrame     # 総合ランキング表
    item_tables: Dict[str, pd.DataFrame]   # 評価項目別表
    dept_tables: Dict[str, pd.DataFrame]   # 部門別表
    paragraphs: List[str]           # 本文段落
    highlights: List[str]           # ハイライト（見出し候補）
    footnotes: List[str]            # 脚注

    def to_dict(self) -> Dict:
        """辞書形式に変換"""
        return {
            "title": self.title,
            "year": self.year,
            "ranking_name": self.ranking_name,
            "overall_table": self.overall_table.to_dict() if self.overall_table is not None else None,
            "item_tables": {k: v.to_dict() for k, v in self.item_tables.items()},
            "dept_tables": {k: v.to_dict() for k, v in self.dept_tables.items()},
            "paragraphs": self.paragraphs,
            "highlights": self.highlights,
            "footnotes": self.footnotes,
        }


# ========================================
# 表生成クラス
# ========================================
class TableGenerator:
    """ランキング表の生成"""

    def __init__(self, display_count: int = 10):
        """
        Args:
            display_count: 表示企業数（デフォルト: TOP10）
        """
        self.display_count = display_count

    def generate_overall_table(
        self,
        data: List[Dict],
        year: int,
        show_score: bool = True,
        show_prev_rank: bool = False,
        prev_data: Optional[List[Dict]] = None
    ) -> pd.DataFrame:
        """総合ランキング表を生成

        Args:
            data: ランキングデータ [{rank, company, score}, ...]
            year: 年度
            show_score: 得点を表示するか
            show_prev_rank: 前年順位を表示するか
            prev_data: 前年のランキングデータ

        Returns:
            DataFrame
        """
        if not data:
            return pd.DataFrame()

        # TOP N に制限
        sorted_data = sorted(data, key=lambda x: x.get("rank", 999))[:self.display_count]

        # 前年順位のマッピング
        prev_ranks = {}
        if prev_data:
            for entry in prev_data:
                company = normalize_company_name(entry.get("company", ""))
                prev_ranks[company] = entry.get("rank")

        rows = []
        for entry in sorted_data:
            rank = entry.get("rank")
            company = entry.get("company", "")
            score = entry.get("score")

            row = {
                "順位": f"{rank}位" if rank else "-",
                "企業名": company,
            }

            if show_score and score is not None:
                row["得点"] = f"{score}点"

            if show_prev_rank:
                normalized = normalize_company_name(company)
                prev_rank = prev_ranks.get(normalized)
                if prev_rank:
                    diff = prev_rank - rank if rank else None
                    if diff is not None:
                        if diff > 0:
                            row["前年比"] = f"↑{diff}"
                        elif diff < 0:
                            row["前年比"] = f"↓{abs(diff)}"
                        else:
                            row["前年比"] = "→"
                    else:
                        row["前年比"] = "-"
                else:
                    row["前年比"] = "NEW"

            rows.append(row)

        return pd.DataFrame(rows)

    def generate_item_table(
        self,
        item_data: Dict[str, List[Dict]],
        year: int
    ) -> Dict[str, pd.DataFrame]:
        """評価項目別ランキング表を生成

        Args:
            item_data: {項目名: [{rank, company, score}, ...]}
            year: 年度

        Returns:
            {項目名: DataFrame}
        """
        result = {}

        for item_name, data in item_data.items():
            if not data:
                continue

            # 年度データを取得
            year_data = data.get(year, []) if isinstance(data, dict) else data

            if not year_data:
                continue

            # TOP5に制限（評価項目は5社程度が適切）
            sorted_data = sorted(year_data, key=lambda x: x.get("rank", 999))[:5]

            rows = []
            for entry in sorted_data:
                rows.append({
                    "順位": f"{entry.get('rank')}位" if entry.get("rank") else "-",
                    "企業名": entry.get("company", ""),
                    "得点": f"{entry.get('score')}点" if entry.get("score") is not None else "-",
                })

            result[item_name] = pd.DataFrame(rows)

        return result

    def generate_dept_table(
        self,
        dept_data: Dict[str, Dict],
        year: int
    ) -> Dict[str, pd.DataFrame]:
        """部門別ランキング表を生成

        Args:
            dept_data: {部門名: {year: [{rank, company, score}, ...]}}
            year: 年度

        Returns:
            {部門名: DataFrame}
        """
        result = {}

        for dept_name, year_data in dept_data.items():
            if not isinstance(year_data, dict):
                continue

            data = year_data.get(year, [])
            if not data:
                continue

            # TOP5に制限
            sorted_data = sorted(data, key=lambda x: x.get("rank", 999))[:5]

            rows = []
            for entry in sorted_data:
                rows.append({
                    "順位": f"{entry.get('rank')}位" if entry.get("rank") else "-",
                    "企業名": entry.get("company", ""),
                    "得点": f"{entry.get('score')}点" if entry.get("score") is not None else "-",
                })

            result[dept_name] = pd.DataFrame(rows)

        return result


# ========================================
# 文章生成クラス
# ========================================
class TextGenerator:
    """プレスリリース文章の生成（テンプレート+ルールベース）"""

    def __init__(self, ranking_name: str, year: int):
        """
        Args:
            ranking_name: ランキング名
            year: 発表年度
        """
        self.ranking_name = ranking_name
        self.year = year
        self.templates = RELEASE_TEMPLATES

    def generate_overall_paragraph(
        self,
        data: List[Dict],
        prev_data: Optional[List[Dict]] = None
    ) -> List[str]:
        """総合ランキングの文章を生成

        Args:
            data: 今年のランキングデータ
            prev_data: 前年のランキングデータ

        Returns:
            段落のリスト
        """
        paragraphs = []

        if not data:
            return paragraphs

        # 1位を取得
        sorted_data = sorted(data, key=lambda x: x.get("rank", 999))
        winners = [d for d in sorted_data if d.get("rank") == 1]

        if not winners:
            return paragraphs

        # 同率1位チェック
        if len(winners) > 1:
            # 同率1位
            companies = "」「".join([w.get("company", "") for w in winners])
            paragraphs.append(
                self.templates["overall_tie"].format(
                    year=self.year,
                    ranking_name=self.ranking_name,
                    companies=companies,
                    score=winners[0].get("score", "-"),
                    rank=1
                )
            )
        else:
            # 単独1位
            winner = winners[0]
            paragraphs.append(
                self.templates["overall_winner"].format(
                    year=self.year,
                    ranking_name=self.ranking_name,
                    company=winner.get("company", ""),
                    score=winner.get("score", "-"),
                    rank=1
                )
            )

        # 前年との比較
        if prev_data:
            prev_dict = {
                normalize_company_name(d.get("company", "")): d
                for d in prev_data
            }

            # 1位が変わったかチェック
            prev_winners = [d for d in prev_data if d.get("rank") == 1]
            if prev_winners:
                prev_winner_companies = {
                    normalize_company_name(w.get("company", ""))
                    for w in prev_winners
                }
                current_winner_companies = {
                    normalize_company_name(w.get("company", ""))
                    for w in winners
                }

                # 新規1位（前年1位ではなかった企業）
                new_winners = current_winner_companies - prev_winner_companies
                for company in new_winners:
                    paragraphs.append(
                        self.templates["first_win"].format(company=company)
                    )

        return paragraphs

    def generate_highlights(
        self,
        overall_data: List[Dict],
        item_data: Dict,
        dept_data: Dict,
        historical_data: Optional[Dict] = None
    ) -> List[str]:
        """ハイライト（見出し候補）を生成

        Args:
            overall_data: 総合ランキングデータ
            item_data: 評価項目別データ
            dept_data: 部門別データ
            historical_data: 歴代記録データ

        Returns:
            ハイライトのリスト
        """
        highlights = []

        # 1位企業
        if overall_data:
            winners = [d for d in overall_data if d.get("rank") == 1]
            if len(winners) > 1:
                companies = "と".join([w.get("company", "") for w in winners])
                highlights.append(f"「{companies}」が同率1位")
            elif winners:
                highlights.append(f"「{winners[0].get('company', '')}」が1位を獲得")

        # 連続記録
        if historical_data:
            consecutive = historical_data.get("historical_records", {}).get("consecutive_wins", [])
            current_streaks = [c for c in consecutive if c.get("is_current")]
            for streak in current_streaks:
                if streak.get("years", 0) >= 2:
                    highlights.append(
                        f"「{streak.get('company', '')}」が{streak.get('years')}年連続1位"
                    )

        # 評価項目の特徴
        if item_data and isinstance(item_data, dict):
            # 全項目1位の企業をチェック
            item_winners = {}
            for item_name, year_data in item_data.items():
                if isinstance(year_data, dict):
                    current = year_data.get(self.year, [])
                    if current:
                        winner = next((d for d in current if d.get("rank") == 1), None)
                        if winner:
                            company = normalize_company_name(winner.get("company", ""))
                            if company not in item_winners:
                                item_winners[company] = []
                            item_winners[company].append(item_name)

            # 複数項目で1位の企業
            for company, items in item_winners.items():
                if len(items) >= 3:
                    highlights.append(f"「{company}」が{len(items)}項目で1位を獲得")

        return highlights

    def generate_item_paragraphs(
        self,
        item_data: Dict,
        top_n: int = 3
    ) -> List[str]:
        """評価項目別の文章を生成

        Args:
            item_data: 評価項目別データ
            top_n: 文章化する項目数

        Returns:
            段落のリスト
        """
        paragraphs = []

        if not item_data or not isinstance(item_data, dict):
            return paragraphs

        count = 0
        for item_name, year_data in item_data.items():
            if count >= top_n:
                break

            if isinstance(year_data, dict):
                current = year_data.get(self.year, [])
                if current:
                    winner = next((d for d in current if d.get("rank") == 1), None)
                    if winner:
                        paragraphs.append(
                            self.templates["item_winner"].format(
                                item_name=item_name,
                                company=winner.get("company", ""),
                                score=winner.get("score", "-")
                            )
                        )
                        count += 1

        return paragraphs

    def generate_dept_paragraphs(
        self,
        dept_data: Dict,
        top_n: int = 3
    ) -> List[str]:
        """部門別の文章を生成

        Args:
            dept_data: 部門別データ
            top_n: 文章化する部門数

        Returns:
            段落のリスト
        """
        paragraphs = []

        if not dept_data or not isinstance(dept_data, dict):
            return paragraphs

        count = 0
        for dept_name, year_data in dept_data.items():
            if count >= top_n:
                break

            if isinstance(year_data, dict):
                current = year_data.get(self.year, [])
                if current:
                    winner = next((d for d in current if d.get("rank") == 1), None)
                    if winner:
                        paragraphs.append(
                            self.templates["dept_winner"].format(
                                dept_name=dept_name,
                                company=winner.get("company", ""),
                                score=winner.get("score", "-")
                            )
                        )
                        count += 1

        return paragraphs


# ========================================
# プレスリリース生成クラス（統合）
# ========================================
class ReleaseGenerator:
    """プレスリリースの統合生成クラス"""

    def __init__(
        self,
        ranking_name: str,
        year: int,
        overall_data: Dict[int, List[Dict]],
        item_data: Dict = None,
        dept_data: Dict = None,
        historical_data: Dict = None
    ):
        """
        Args:
            ranking_name: ランキング名
            year: 発表年度
            overall_data: 総合ランキングデータ {year: [entries]}
            item_data: 評価項目別データ
            dept_data: 部門別データ
            historical_data: 歴代記録データ
        """
        self.ranking_name = ranking_name
        self.year = year
        self.overall_data = overall_data or {}
        self.item_data = item_data or {}
        self.dept_data = dept_data or {}
        self.historical_data = historical_data or {}

        self.table_gen = TableGenerator()
        self.text_gen = TextGenerator(ranking_name, year)

    def generate(self) -> ReleaseContent:
        """プレスリリースを生成

        Returns:
            ReleaseContent
        """
        # 今年と前年のデータ
        current_data = self.overall_data.get(self.year, [])
        prev_year = self.year - 1
        prev_data = self.overall_data.get(prev_year, [])

        # 表の生成
        overall_table = self.table_gen.generate_overall_table(
            data=current_data,
            year=self.year,
            show_score=True,
            show_prev_rank=bool(prev_data),
            prev_data=prev_data
        )

        item_tables = self.table_gen.generate_item_table(
            item_data=self.item_data,
            year=self.year
        )

        dept_tables = self.table_gen.generate_dept_table(
            dept_data=self.dept_data,
            year=self.year
        )

        # 文章の生成
        paragraphs = []

        # 総合ランキング
        paragraphs.extend(
            self.text_gen.generate_overall_paragraph(
                data=current_data,
                prev_data=prev_data
            )
        )

        # 評価項目別
        paragraphs.extend(
            self.text_gen.generate_item_paragraphs(
                item_data=self.item_data,
                top_n=3
            )
        )

        # 部門別
        paragraphs.extend(
            self.text_gen.generate_dept_paragraphs(
                dept_data=self.dept_data,
                top_n=3
            )
        )

        # ハイライト生成
        highlights = self.text_gen.generate_highlights(
            overall_data=current_data,
            item_data=self.item_data,
            dept_data=self.dept_data,
            historical_data=self.historical_data
        )

        # タイトル
        title = f"{self.year}年 オリコン顧客満足度®調査「{self.ranking_name}」ランキング"

        return ReleaseContent(
            title=title,
            year=self.year,
            ranking_name=self.ranking_name,
            overall_table=overall_table,
            item_tables=item_tables,
            dept_tables=dept_tables,
            paragraphs=paragraphs,
            highlights=highlights,
            footnotes=[]
        )

    def export_to_excel(self, content: ReleaseContent) -> BytesIO:
        """Excelファイルとしてエクスポート

        Args:
            content: プレスリリースコンテンツ

        Returns:
            BytesIOオブジェクト
        """
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # フォーマット定義
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#003366',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
            })

            cell_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
            })

            # === 総合ランキングシート ===
            if content.overall_table is not None and not content.overall_table.empty:
                content.overall_table.to_excel(
                    writer,
                    sheet_name="総合ランキング",
                    index=False,
                    startrow=1
                )

                worksheet = writer.sheets["総合ランキング"]
                worksheet.write(0, 0, content.title, header_format)

                # ヘッダーフォーマット適用
                for col_num, value in enumerate(content.overall_table.columns):
                    worksheet.write(1, col_num, value, header_format)

            # === 評価項目別シート ===
            if content.item_tables:
                row_offset = 0
                for item_name, df in content.item_tables.items():
                    if df.empty:
                        continue

                    sheet_name = "評価項目別"
                    if sheet_name not in writer.sheets:
                        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=row_offset + 1)
                        worksheet = writer.sheets[sheet_name]
                        worksheet.write(row_offset, 0, f"■ {item_name}")
                    else:
                        worksheet = writer.sheets[sheet_name]
                        worksheet.write(row_offset, 0, f"■ {item_name}")
                        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=row_offset + 1, header=True)

                    row_offset += len(df) + 3

            # === 部門別シート ===
            if content.dept_tables:
                row_offset = 0
                for dept_name, df in content.dept_tables.items():
                    if df.empty:
                        continue

                    sheet_name = "部門別"
                    if sheet_name not in writer.sheets:
                        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=row_offset + 1)
                        worksheet = writer.sheets[sheet_name]
                        worksheet.write(row_offset, 0, f"■ {dept_name}")
                    else:
                        worksheet = writer.sheets[sheet_name]
                        worksheet.write(row_offset, 0, f"■ {dept_name}")
                        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=row_offset + 1, header=True)

                    row_offset += len(df) + 3

            # === 文章シート ===
            text_df = pd.DataFrame({
                "種類": ["ハイライト"] * len(content.highlights) + ["本文"] * len(content.paragraphs),
                "内容": content.highlights + content.paragraphs
            })
            if not text_df.empty:
                text_df.to_excel(writer, sheet_name="文章", index=False)

        output.seek(0)
        return output


# ========================================
# 便利関数
# ========================================
def generate_release(
    ranking_name: str,
    year: int,
    overall_data: Dict,
    item_data: Dict = None,
    dept_data: Dict = None,
    historical_data: Dict = None
) -> ReleaseContent:
    """プレスリリースを生成（簡易インターフェース）

    Args:
        ranking_name: ランキング名
        year: 発表年度
        overall_data: 総合ランキングデータ
        item_data: 評価項目別データ
        dept_data: 部門別データ
        historical_data: 歴代記録データ

    Returns:
        ReleaseContent
    """
    generator = ReleaseGenerator(
        ranking_name=ranking_name,
        year=year,
        overall_data=overall_data,
        item_data=item_data,
        dept_data=dept_data,
        historical_data=historical_data
    )
    return generator.generate()


# ========================================
# デバッグ用
# ========================================
if __name__ == "__main__":
    # テストデータ
    test_overall = {
        2026: [
            {"rank": 1, "company": "SBI証券", "score": 68.9},
            {"rank": 1, "company": "楽天証券", "score": 68.9},  # 同率1位
            {"rank": 3, "company": "マネックス証券", "score": 67.5},
            {"rank": 4, "company": "松井証券", "score": 66.0},
            {"rank": 5, "company": "auカブコム証券", "score": 65.5},
        ],
        2025: [
            {"rank": 1, "company": "SBI証券", "score": 68.5},
            {"rank": 2, "company": "楽天証券", "score": 68.0},
            {"rank": 3, "company": "マネックス証券", "score": 67.0},
            {"rank": 4, "company": "松井証券", "score": 65.5},
            {"rank": 5, "company": "auカブコム証券", "score": 65.0},
        ]
    }

    test_item_data = {
        "取引手数料": {
            2026: [
                {"rank": 1, "company": "SBI証券", "score": 72.0},
                {"rank": 2, "company": "楽天証券", "score": 71.5},
            ]
        },
        "取引ツール": {
            2026: [
                {"rank": 1, "company": "楽天証券", "score": 70.0},
                {"rank": 2, "company": "SBI証券", "score": 69.5},
            ]
        }
    }

    # 生成実行
    content = generate_release(
        ranking_name="ネット証券",
        year=2026,
        overall_data=test_overall,
        item_data=test_item_data
    )

    print("=== プレスリリース生成結果 ===")
    print(f"\nタイトル: {content.title}")
    print(f"\n■ ハイライト:")
    for h in content.highlights:
        print(f"  - {h}")
    print(f"\n■ 本文:")
    for p in content.paragraphs:
        print(f"  {p}")
    print(f"\n■ 総合ランキング表:")
    print(content.overall_table)

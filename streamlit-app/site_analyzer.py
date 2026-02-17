# -*- coding: utf-8 -*-
"""
SiteStructureAnalyzer - オリコンサイト構造解析モジュール
v1.2 - 2025-12-17

sort-navのTABLE構造を1回の解析で動的に判定し、
総合/評価項目別/部門別/過去年度の情報を一括取得する。

v1.2 - 型安全性・リソース管理改善
- 過去年度検証でUnion[int, str]を正しく処理
- close()メソッド追加でセッションリソースを解放
- コンテキストマネージャー対応

v1.1 - エラーハンドリング強化
- validate()メソッド追加: 構造の妥当性チェック
- StructureValidator クラス追加: 詳細な検証ルール
- 警告・エラーの詳細化
"""

import requests
from requests.adapters import HTTPAdapter
from requests.exceptions import RequestException
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import re
from datetime import datetime
from typing import Dict, List, Optional, NamedTuple, Union
from dataclasses import dataclass, field
import logging

logger = logging.getLogger(__name__)


@dataclass
class DepartmentCategory:
    """部門カテゴリ（例: 年代別、業務内容別など）"""
    name: str  # カテゴリ名（例: "年代別ランキング"）
    departments: Dict[str, str] = field(default_factory=dict)  # {path: name}


@dataclass
class SiteStructure:
    """サイト構造情報"""
    # 基本情報
    url: str
    ranking_name: str = ""

    # 総合ランキング
    has_overall: bool = True

    # 評価項目別
    has_evaluation_items: bool = False
    evaluation_items: Dict[str, str] = field(default_factory=dict)  # {slug: name}

    # 部門別
    has_departments: bool = False
    department_categories: List[DepartmentCategory] = field(default_factory=list)
    departments_flat: Dict[str, str] = field(default_factory=dict)  # 全部門をフラット化

    # 過去年度
    has_past_years: bool = False
    available_years: List[Union[int, str]] = field(default_factory=list)  # 2014-2015形式にも対応
    current_year: Optional[int] = None

    # エラー情報
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    # 検証結果（v1.1追加）
    is_valid: bool = True
    validation_status: str = "未検証"


class StructureValidator:
    """
    サイト構造の妥当性を検証するクラス（v1.1追加）

    検証項目:
    - sort-navの存在
    - 評価項目数の妥当性
    - 部門数の妥当性
    - 年度情報の存在
    """

    # 検証しきい値
    MIN_EVALUATION_ITEMS = 0  # 評価項目なしのランキングもある
    MAX_EVALUATION_ITEMS = 20
    MIN_DEPARTMENTS = 0  # 部門なしのランキングもある
    MAX_DEPARTMENTS = 50
    MIN_YEAR = 2000
    MAX_YEAR = datetime.now().year + 5

    @classmethod
    def validate(cls, structure: SiteStructure) -> SiteStructure:
        """
        構造を検証し、warnings/errorsを設定

        Args:
            structure: 検証対象のSiteStructure

        Returns:
            検証結果が設定されたSiteStructure
        """
        structure.is_valid = True
        structure.validation_status = "OK"

        # エラーがあれば無効
        if structure.errors:
            structure.is_valid = False
            structure.validation_status = "エラーあり"
            return structure

        # 評価項目数の検証
        item_count = len(structure.evaluation_items)
        if item_count > cls.MAX_EVALUATION_ITEMS:
            structure.warnings.append(
                f"評価項目数が多すぎます（{item_count}件、通常は{cls.MAX_EVALUATION_ITEMS}件以下）"
            )

        # 部門数の検証
        dept_count = len(structure.departments_flat)
        if dept_count > cls.MAX_DEPARTMENTS:
            structure.warnings.append(
                f"部門数が多すぎます（{dept_count}件、通常は{cls.MAX_DEPARTMENTS}件以下）"
            )

        # 年度の検証
        if structure.current_year:
            if not (cls.MIN_YEAR <= structure.current_year <= cls.MAX_YEAR):
                structure.warnings.append(
                    f"年度が範囲外です（{structure.current_year}年）"
                )

        # 過去年度の検証（Union[int, str]対応）
        for year in structure.available_years:
            # 文字列年度（例: "2014-2015"）の場合は開始年を基準に検証
            if isinstance(year, str):
                try:
                    start_year = int(year.split("-")[0])
                    if not (cls.MIN_YEAR <= start_year <= cls.MAX_YEAR):
                        structure.warnings.append(f"過去年度が範囲外です（{year}）")
                except (ValueError, IndexError):
                    structure.warnings.append(f"不正な年度形式です（{year}）")
            else:
                if not (cls.MIN_YEAR <= year <= cls.MAX_YEAR):
                    structure.warnings.append(f"過去年度が範囲外です（{year}年）")

        # 警告があれば注意ステータス
        if structure.warnings:
            structure.validation_status = "警告あり"

        return structure

    @classmethod
    def check_structure_change(
        cls,
        current: SiteStructure,
        previous: Optional[SiteStructure]
    ) -> List[str]:
        """
        前回の構造と比較して変更を検知

        Args:
            current: 現在の構造
            previous: 前回の構造（Noneの場合は比較しない）

        Returns:
            変更内容のリスト
        """
        if previous is None:
            return []

        changes = []

        # 評価項目の変更
        curr_items = set(current.evaluation_items.keys())
        prev_items = set(previous.evaluation_items.keys())
        added_items = curr_items - prev_items
        removed_items = prev_items - curr_items
        if added_items:
            changes.append(f"評価項目が追加されました: {added_items}")
        if removed_items:
            changes.append(f"評価項目が削除されました: {removed_items}")

        # 部門の変更
        curr_depts = set(current.departments_flat.keys())
        prev_depts = set(previous.departments_flat.keys())
        added_depts = curr_depts - prev_depts
        removed_depts = prev_depts - curr_depts
        if added_depts:
            changes.append(f"部門が追加されました: {len(added_depts)}件")
        if removed_depts:
            changes.append(f"部門が削除されました: {len(removed_depts)}件")

        # カテゴリ数の変更
        if len(current.department_categories) != len(previous.department_categories):
            changes.append(
                f"部門カテゴリ数が変更されました: "
                f"{len(previous.department_categories)} → {len(current.department_categories)}"
            )

        return changes


class SiteStructureAnalyzer:
    """
    オリコンサイトの構造を動的に解析するクラス

    1回のHTTPリクエストで以下を判定:
    - 総合ランキングの存在有無
    - 評価項目別の存在有無・URL
    - 部門別の存在有無・URL・カテゴリ名
    - 過去年度ページへの遷移可否
    """

    # 除外する見出し
    EXCLUDE_HEADINGS = [
        "TOP",
        "評価項目別", "評価項目",
        "過去のランキング", "過去ランキング",
        "関連ランキング", "関連する"
    ]

    # 評価項目として認識する見出し
    EVALUATION_ITEM_HEADINGS = [
        "評価項目別", "評価項目"
    ]

    # 過去年度として認識する見出し
    PAST_YEAR_HEADINGS = [
        "過去のランキング", "過去ランキング"
    ]

    # 無効な部門名
    INVALID_DEPT_NAMES = [
        # 都道府県名（単体で部門として誤検出されやすい）- 派遣会社等では有効
        # → prefectureパターン対応時に別途処理
        "ランキング", "一覧", "比較", "おすすめ", "検索結果", "教室一覧",
        "2020年", "2021年", "2022年", "2023年", "2024年", "2025年", "2026年", "2027年",
    ]

    # 除外URLパターン
    EXCLUDE_URL_PATTERNS = [
        r"/column",
        r"/special/",
        r"-basic",
        r"/howto",
        r"/how_to",
        r"/recommend",
        r"/compare",
        r"/company/",
        r"/education/",
        r"/school_list/",
        r"/town/",
        r"[?&]pref=",
        r"[?&]area=",
        r"/search",
        r"/ranking-list",
    ]

    def __init__(self, timeout: int = 10, max_retries: int = 3):
        """
        Args:
            timeout: HTTPリクエストのタイムアウト（秒）
            max_retries: リトライ回数
        """
        self.timeout = timeout
        self.session = self._create_session(max_retries)

    def close(self):
        """セッションを閉じてリソースを解放（v1.2追加）"""
        if self.session:
            self.session.close()
            logger.debug("SiteStructureAnalyzerセッションを閉じました")

    def __enter__(self):
        """コンテキストマネージャー対応（v1.2追加）"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """コンテキストマネージャー終了時にセッションを閉じる（v1.2追加）"""
        self.close()
        return False

    def _create_session(self, max_retries: int) -> requests.Session:
        """リトライ機能付きセッションを作成"""
        session = requests.Session()
        retry_strategy = Retry(
            total=max_retries,
            backoff_factor=1,
            status_forcelist=[500, 502, 503, 504],
            allowed_methods=["GET"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })
        return session

    def analyze(self, url: str, url_prefix: str = "") -> SiteStructure:
        """
        サイト構造を解析

        Args:
            url: 解析対象のトップページURL
            url_prefix: URLプレフィックス（例: "rank_staffing"）

        Returns:
            SiteStructure: 解析結果
        """
        result = SiteStructure(url=url)

        try:
            response = self.session.get(url, timeout=self.timeout)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            # sort-navを探す
            sort_nav = soup.find(class_="sort-nav")
            if not sort_nav:
                result.warnings.append("sort-navが見つかりません")
                return result

            # TABLE構造を解析
            table = sort_nav.find("table")
            if table:
                self._analyze_table_structure(table, result, url_prefix)
            else:
                result.warnings.append("sort-nav内にtableが見つかりません")

            # 現在年度を検出
            result.current_year = self._detect_current_year(soup)

        except RequestException as e:
            result.errors.append(f"HTTPエラー: {e}")
            logger.error(f"SiteStructureAnalyzer HTTPエラー: {e}")
        except Exception as e:
            result.errors.append(f"解析エラー: {e}")
            logger.error(f"SiteStructureAnalyzer 解析エラー: {e}")

        # 統計情報をログ出力
        logger.info(
            f"SiteStructure解析完了: "
            f"評価項目={len(result.evaluation_items)}件, "
            f"部門カテゴリ={len(result.department_categories)}件, "
            f"部門合計={len(result.departments_flat)}件, "
            f"過去年度={len(result.available_years)}件"
        )

        # 構造の妥当性検証（v1.1追加）
        result = StructureValidator.validate(result)
        if result.validation_status != "OK":
            logger.warning(f"構造検証結果: {result.validation_status}")

        return result

    def _analyze_table_structure(self, table, result: SiteStructure, url_prefix: str):
        """
        TABLE構造を解析

        <table>
          <tr>
            <th>カテゴリ名</th>
            <td><a href="...">リンク1</a> <a href="...">リンク2</a></td>
          </tr>
        </table>
        """
        for tr in table.find_all("tr"):
            th = tr.find("th")
            if not th:
                continue

            heading_text = th.get_text(strip=True)

            # TOPは総合ランキング（スキップ）
            if heading_text == "TOP":
                result.has_overall = True
                continue

            # 評価項目別
            if any(h in heading_text for h in self.EVALUATION_ITEM_HEADINGS):
                self._extract_evaluation_items(tr, result, url_prefix)
                continue

            # 過去年度
            if any(h in heading_text for h in self.PAST_YEAR_HEADINGS):
                self._extract_past_years(tr, result)
                continue

            # 関連ランキングは除外
            if "関連" in heading_text:
                continue

            # それ以外は部門カテゴリ
            self._extract_department_category(tr, heading_text, result, url_prefix)

    def _extract_evaluation_items(self, tr, result: SiteStructure, url_prefix: str):
        """評価項目を抽出"""
        result.has_evaluation_items = True

        for td in tr.find_all("td"):
            for link in td.find_all("a", href=True):
                href = link.get("href", "")
                link_text = link.get_text(strip=True)

                # evaluation-item パターンを抽出
                match = re.search(r"/evaluation-item/([^/]+)\.html", href)
                if match:
                    slug = match.group(1)
                    result.evaluation_items[slug] = link_text

    def _extract_past_years(self, tr, result: SiteStructure):
        """過去年度を抽出（2014-2015形式にも対応）"""
        result.has_past_years = True

        for td in tr.find_all("td"):
            for link in td.find_all("a", href=True):
                href = link.get("href", "")

                # 年度パターンを抽出（2014-2015形式にも対応）
                match = re.search(r"/(\d{4}(?:-\d{4})?)/?$", href)
                if match:
                    year_str = match.group(1)
                    # 全ての年度を文字列で統一（int/str混在を防止）
                    if "-" in year_str:
                        # 開始年が妥当な範囲かチェック
                        start_year = int(year_str.split("-")[0])
                        if 2000 <= start_year <= 2030:
                            result.available_years.append(year_str)
                    else:
                        year = int(year_str)
                        if 2000 <= year <= 2030:
                            result.available_years.append(str(year))  # 文字列で統一

        # ソート（新しい順）- 文字列と数値が混在するためカスタムキー使用
        def sort_key(y):
            if isinstance(y, str):
                return int(y.split("-")[0])  # 2014-2015 → 2014 で比較
            return y
        result.available_years.sort(key=sort_key, reverse=True)

    def _extract_department_category(self, tr, heading_text: str, result: SiteStructure, url_prefix: str):
        """部門カテゴリを抽出"""
        category = DepartmentCategory(name=heading_text)

        for td in tr.find_all("td"):
            for link in td.find_all("a", href=True):
                href = link.get("href", "")
                link_text = link.get_text(strip=True)

                # 除外パターンチェック
                if any(re.search(pat, href) for pat in self.EXCLUDE_URL_PATTERNS):
                    continue

                # 年度リンクは除外（2014-2015形式にも対応）
                if re.search(r"/\d{4}(?:-\d{4})?/?$", href):
                    continue

                # url_prefixが指定されている場合、それを含むリンクのみ
                if url_prefix and url_prefix not in href:
                    continue

                # パスを抽出
                dept_path = self._extract_dept_path(href, url_prefix)
                if dept_path and self._is_valid_dept_name(link_text):
                    category.departments[dept_path] = link_text
                    result.departments_flat[dept_path] = link_text

        if category.departments:
            result.has_departments = True
            result.department_categories.append(category)

    def _extract_dept_path(self, href: str, url_prefix: str) -> Optional[str]:
        """URLから部門パスを抽出"""
        # フラグメントを除去
        href_clean = href.split('#')[0]

        # パターン: /url_prefix/(.+)
        if url_prefix:
            pattern = rf"/{url_prefix}/(?:\d{{4}}/)?(.+?)(?:\?.*)?$"
            match = re.search(pattern, href_clean)
            if match:
                dept_path = match.group(1)
                if dept_path and not dept_path.rstrip('/').isdigit() and '?' not in dept_path:
                    return dept_path

        return None

    def _is_valid_dept_name(self, dept_name: str) -> bool:
        """部門名の妥当性チェック"""
        if not dept_name or not dept_name.strip():
            return False

        dept_name = dept_name.strip()

        if len(dept_name) > 30:
            return False

        if dept_name in self.INVALID_DEPT_NAMES:
            return False

        if re.match(r"^\d{4}年?$", dept_name):
            return False

        return True

    def _detect_current_year(self, soup) -> Optional[int]:
        """現在年度を検出"""
        text = soup.get_text()

        # パターン1: 最終更新日
        match = re.search(r'(?:最終)?更新日[：:\s]*(\d{4})[-/]\d{1,2}[-/]\d{1,2}', text)
        if match:
            return int(match.group(1))

        # パターン2: タイトル
        match = re.search(r'(\d{4})年\s*オリコン', text)
        if match:
            return int(match.group(1))

        return None

    def get_structure_summary(self, structure: SiteStructure) -> str:
        """構造の概要を文字列で返す"""
        lines = [
            f"URL: {structure.url}",
            f"現在年度: {structure.current_year}",
            f"検証ステータス: {structure.validation_status}",
            f"総合ランキング: {'あり' if structure.has_overall else 'なし'}",
            f"評価項目別: {len(structure.evaluation_items)}件",
            f"部門カテゴリ: {len(structure.department_categories)}件",
            f"部門合計: {len(structure.departments_flat)}件",
            f"過去年度: {len(structure.available_years)}件 {structure.available_years[:5]}...",
        ]

        if structure.department_categories:
            lines.append("\n【部門カテゴリ詳細】")
            for cat in structure.department_categories:
                lines.append(f"  {cat.name}: {len(cat.departments)}件")
                for path, name in list(cat.departments.items())[:3]:
                    lines.append(f"    - {name} ({path})")
                if len(cat.departments) > 3:
                    lines.append(f"    ... 他{len(cat.departments) - 3}件")

        if structure.errors:
            lines.append(f"\n【エラー】{structure.errors}")
        if structure.warnings:
            lines.append(f"【警告】{structure.warnings}")

        return "\n".join(lines)


# テスト用
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    analyzer = SiteStructureAnalyzer()

    # 派遣会社をテスト
    url = "https://career.oricon.co.jp/rank_staffing/"
    structure = analyzer.analyze(url, url_prefix="rank_staffing")
    print(analyzer.get_structure_summary(structure))

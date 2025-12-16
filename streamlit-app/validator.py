# -*- coding: utf-8 -*-
"""
正誤チェックモジュール (v1.0)
プレスリリースの内容を検証

機能:
1. 順位・得点の正確性チェック
2. 企業名表記チェック
3. 連続記録・実績の検証
4. Excel vs Web データのクロスチェック
"""

import logging
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from enum import Enum

from company_master import (
    validate_company_name,
    normalize_company_name,
    find_similar_companies,
    get_official_name
)

logger = logging.getLogger(__name__)


# ========================================
# 検証結果の定義
# ========================================
class ValidationLevel(Enum):
    """検証結果のレベル"""
    OK = "OK"           # 問題なし
    WARNING = "WARNING" # 警告（確認推奨）
    ERROR = "ERROR"     # エラー（修正必要）
    INFO = "INFO"       # 情報


@dataclass
class ValidationIssue:
    """検証で発見された問題"""
    level: ValidationLevel
    category: str       # "ranking", "company", "record", "cross_check"
    field: str          # 問題のあるフィールド
    message: str        # 問題の説明
    expected: Any = None  # 期待値
    actual: Any = None    # 実際の値
    suggestion: str = ""  # 修正提案
    context: Dict = field(default_factory=dict)  # 追加コンテキスト


@dataclass
class ValidationResult:
    """検証結果の全体"""
    is_valid: bool
    issues: List[ValidationIssue]
    summary: Dict[str, int]  # レベルごとの件数

    def __post_init__(self):
        # サマリーを自動計算
        self.summary = {
            "OK": len([i for i in self.issues if i.level == ValidationLevel.OK]),
            "WARNING": len([i for i in self.issues if i.level == ValidationLevel.WARNING]),
            "ERROR": len([i for i in self.issues if i.level == ValidationLevel.ERROR]),
            "INFO": len([i for i in self.issues if i.level == ValidationLevel.INFO]),
        }

    def get_errors(self) -> List[ValidationIssue]:
        """エラーのみ取得"""
        return [i for i in self.issues if i.level == ValidationLevel.ERROR]

    def get_warnings(self) -> List[ValidationIssue]:
        """警告のみ取得"""
        return [i for i in self.issues if i.level == ValidationLevel.WARNING]

    def to_dict(self) -> Dict:
        """辞書形式に変換"""
        return {
            "is_valid": self.is_valid,
            "summary": self.summary,
            "issues": [
                {
                    "level": i.level.value,
                    "category": i.category,
                    "field": i.field,
                    "message": i.message,
                    "expected": i.expected,
                    "actual": i.actual,
                    "suggestion": i.suggestion,
                    "context": i.context
                }
                for i in self.issues
            ]
        }


# ========================================
# バリデータクラス
# ========================================
class ReleaseValidator:
    """プレスリリース内容の検証クラス"""

    def __init__(
        self,
        excel_data: Optional[Dict] = None,
        web_data: Optional[Dict] = None,
        ranking_name: str = ""
    ):
        """
        Args:
            excel_data: Excelから読み込んだデータ（最新年）
            web_data: Webスクレイピングで取得したデータ（過去年）
            ranking_name: ランキング名
        """
        self.excel_data = excel_data or {}
        self.web_data = web_data or {}
        self.ranking_name = ranking_name
        self.issues: List[ValidationIssue] = []

    def validate_all(self) -> ValidationResult:
        """全検証を実行"""
        self.issues = []

        # 1. ランキングデータの検証
        self._validate_ranking_data()

        # 2. 企業名の検証
        self._validate_company_names()

        # 3. 連続記録の検証
        self._validate_records()

        # 4. クロスチェック（Excel vs Web）
        self._cross_check_data()

        # 結果を返す
        has_errors = any(i.level == ValidationLevel.ERROR for i in self.issues)
        return ValidationResult(
            is_valid=not has_errors,
            issues=self.issues,
            summary={}  # __post_init__で計算
        )

    # ========================================
    # 1. ランキングデータの検証
    # ========================================
    def _validate_ranking_data(self):
        """順位・得点の正確性チェック"""
        # Excelデータの検証
        for year, data in self.excel_data.items():
            self._validate_year_data(year, data, source="Excel")

        # Webデータの検証
        for year, data in self.web_data.items():
            self._validate_year_data(year, data, source="Web")

    def _validate_year_data(self, year: int, data: List[Dict], source: str):
        """年度別データの検証"""
        if not data:
            return

        # 順位の連続性チェック
        ranks = [d.get("rank") for d in data if d.get("rank")]
        if ranks:
            expected_ranks = list(range(1, len(ranks) + 1))
            # 同率順位を考慮（同じ順位が複数あってもOK）
            unique_ranks = sorted(set(ranks))
            for i, rank in enumerate(unique_ranks):
                if rank > i + 1 and i + 1 not in unique_ranks:
                    # 順位が飛んでいる場合（同率でない）
                    self.issues.append(ValidationIssue(
                        level=ValidationLevel.WARNING,
                        category="ranking",
                        field=f"{year}年順位",
                        message=f"順位が連続していません（{i+1}位が欠番）",
                        expected=i + 1,
                        actual=rank,
                        context={"source": source, "year": year}
                    ))

        # 得点の妥当性チェック
        for entry in data:
            score = entry.get("score")
            if score is not None:
                # 得点範囲チェック（0-100が一般的）
                if not (0 <= score <= 100):
                    self.issues.append(ValidationIssue(
                        level=ValidationLevel.ERROR,
                        category="ranking",
                        field=f"{year}年得点",
                        message=f"得点が範囲外です: {entry.get('company')} = {score}点",
                        expected="0-100",
                        actual=score,
                        context={"source": source, "year": year, "company": entry.get("company")}
                    ))

        # 1位の得点が2位以上であることを確認
        sorted_data = sorted(
            [d for d in data if d.get("score") is not None],
            key=lambda x: x.get("rank", 999)
        )
        if len(sorted_data) >= 2:
            first = sorted_data[0]
            second = sorted_data[1]
            if first.get("score", 0) < second.get("score", 0):
                self.issues.append(ValidationIssue(
                    level=ValidationLevel.ERROR,
                    category="ranking",
                    field=f"{year}年順位",
                    message=f"1位の得点が2位より低いです",
                    expected=f"1位({first.get('company')}) >= 2位({second.get('company')})",
                    actual=f"{first.get('score')}点 < {second.get('score')}点",
                    context={"source": source, "year": year}
                ))

    # ========================================
    # 2. 企業名の検証
    # ========================================
    def _validate_company_names(self):
        """企業名表記チェック"""
        all_companies = set()

        # 全データから企業名を収集
        for year, data in {**self.excel_data, **self.web_data}.items():
            if isinstance(data, list):
                for entry in data:
                    company = entry.get("company")
                    if company:
                        all_companies.add(company)

        # 各企業名を検証
        for company in all_companies:
            result = validate_company_name(company)

            if not result["is_valid"]:
                # マスタに存在しない企業
                if result["suggestion"]:
                    self.issues.append(ValidationIssue(
                        level=ValidationLevel.WARNING,
                        category="company",
                        field="企業名",
                        message=f"「{company}」はマスタに未登録です",
                        suggestion=f"「{result['suggestion']['name']}」(類似度{result['suggestion']['similarity']:.0%})ではありませんか？",
                        context={"input": company, "normalized": result["normalized"]}
                    ))
                else:
                    self.issues.append(ValidationIssue(
                        level=ValidationLevel.INFO,
                        category="company",
                        field="企業名",
                        message=f"「{company}」はマスタに未登録です（新規企業の可能性）",
                        context={"input": company, "normalized": result["normalized"]}
                    ))
            elif company != result["official_name"]:
                # 表記ゆれ
                self.issues.append(ValidationIssue(
                    level=ValidationLevel.INFO,
                    category="company",
                    field="企業名表記",
                    message=f"表記が正式名称と異なります",
                    expected=result["official_name"],
                    actual=company,
                    suggestion=f"正式名称「{result['official_name']}」への統一を推奨",
                    context={"category": result["category"]}
                ))

    # ========================================
    # 3. 連続記録の検証
    # ========================================
    def _validate_records(self):
        """連続記録・実績の検証"""
        # 全年度のデータを統合
        all_data = {}
        for year, data in {**self.web_data, **self.excel_data}.items():
            if isinstance(data, list) and data:
                all_data[year] = data

        if len(all_data) < 2:
            return  # 2年分以上ないと連続記録は検証できない

        years = sorted(all_data.keys())

        # 連続1位の検証
        self._validate_consecutive_wins(all_data, years)

        # 初登場企業の検出
        self._detect_first_appearances(all_data, years)

    def _validate_consecutive_wins(self, all_data: Dict, years: List[int]):
        """連続1位記録の検証"""
        # 各年の1位を取得
        winners = {}
        for year in years:
            data = all_data.get(year, [])
            for entry in data:
                if entry.get("rank") == 1:
                    winners[year] = normalize_company_name(entry.get("company", ""))
                    break

        if not winners:
            return

        # 連続記録を計算
        current_company = None
        consecutive_count = 0
        consecutive_records = []

        for year in years:
            winner = winners.get(year)
            if winner == current_company:
                consecutive_count += 1
            else:
                if current_company and consecutive_count >= 2:
                    consecutive_records.append({
                        "company": current_company,
                        "years": consecutive_count,
                        "end_year": year - 1
                    })
                current_company = winner
                consecutive_count = 1

        # 最後の連続記録
        if current_company and consecutive_count >= 2:
            consecutive_records.append({
                "company": current_company,
                "years": consecutive_count,
                "end_year": years[-1]
            })

        # 連続記録を情報として追加
        for record in consecutive_records:
            is_current = record["end_year"] == years[-1]
            self.issues.append(ValidationIssue(
                level=ValidationLevel.INFO,
                category="record",
                field="連続1位",
                message=f"{record['company']}が{record['years']}年連続1位" +
                       ("（継続中）" if is_current else ""),
                context={
                    "company": record["company"],
                    "years": record["years"],
                    "end_year": record["end_year"],
                    "is_current": is_current
                }
            ))

    def _detect_first_appearances(self, all_data: Dict, years: List[int]):
        """初登場企業の検出"""
        known_companies = set()
        latest_year = max(years)

        for year in sorted(years):
            data = all_data.get(year, [])
            for entry in data:
                company = normalize_company_name(entry.get("company", ""))
                if company and company not in known_companies:
                    known_companies.add(company)
                    # 最新年の初登場企業を報告
                    if year == latest_year:
                        rank = entry.get("rank")
                        self.issues.append(ValidationIssue(
                            level=ValidationLevel.INFO,
                            category="record",
                            field="初登場",
                            message=f"{company}が{year}年に初登場（{rank}位）",
                            context={
                                "company": company,
                                "year": year,
                                "rank": rank,
                                "score": entry.get("score")
                            }
                        ))

    # ========================================
    # 4. クロスチェック（Excel vs Web）
    # ========================================
    def _cross_check_data(self):
        """ExcelデータとWebデータのクロスチェック"""
        # 重複する年度を検索
        excel_years = set(self.excel_data.keys())
        web_years = set(self.web_data.keys())
        common_years = excel_years & web_years

        for year in common_years:
            excel_entries = self.excel_data.get(year, [])
            web_entries = self.web_data.get(year, [])

            # 企業ごとに比較
            excel_dict = {
                normalize_company_name(e.get("company", "")): e
                for e in excel_entries if e.get("company")
            }
            web_dict = {
                normalize_company_name(e.get("company", "")): e
                for e in web_entries if e.get("company")
            }

            # Excel にあって Web にない企業
            excel_only = set(excel_dict.keys()) - set(web_dict.keys())
            for company in excel_only:
                self.issues.append(ValidationIssue(
                    level=ValidationLevel.WARNING,
                    category="cross_check",
                    field=f"{year}年データ",
                    message=f"「{company}」はExcelにあるがWebにない",
                    context={"source": "Excel only", "year": year}
                ))

            # Web にあって Excel にない企業
            web_only = set(web_dict.keys()) - set(excel_dict.keys())
            for company in web_only:
                self.issues.append(ValidationIssue(
                    level=ValidationLevel.WARNING,
                    category="cross_check",
                    field=f"{year}年データ",
                    message=f"「{company}」はWebにあるがExcelにない",
                    context={"source": "Web only", "year": year}
                ))

            # 両方にある企業の順位・得点を比較
            common_companies = set(excel_dict.keys()) & set(web_dict.keys())
            for company in common_companies:
                excel_entry = excel_dict[company]
                web_entry = web_dict[company]

                # 順位の比較
                excel_rank = excel_entry.get("rank")
                web_rank = web_entry.get("rank")
                if excel_rank is not None and web_rank is not None:
                    if excel_rank != web_rank:
                        self.issues.append(ValidationIssue(
                            level=ValidationLevel.ERROR,
                            category="cross_check",
                            field=f"{year}年順位",
                            message=f"「{company}」の順位が不一致",
                            expected=f"Excel: {excel_rank}位",
                            actual=f"Web: {web_rank}位",
                            context={"company": company, "year": year}
                        ))

                # 得点の比較
                excel_score = excel_entry.get("score")
                web_score = web_entry.get("score")
                if excel_score is not None and web_score is not None:
                    # 小数点以下の誤差を許容（0.1点まで）
                    if abs(excel_score - web_score) > 0.1:
                        self.issues.append(ValidationIssue(
                            level=ValidationLevel.ERROR,
                            category="cross_check",
                            field=f"{year}年得点",
                            message=f"「{company}」の得点が不一致",
                            expected=f"Excel: {excel_score}点",
                            actual=f"Web: {web_score}点",
                            context={"company": company, "year": year}
                        ))


# ========================================
# 便利関数
# ========================================
def validate_release_data(
    excel_data: Optional[Dict] = None,
    web_data: Optional[Dict] = None,
    ranking_name: str = ""
) -> ValidationResult:
    """プレスリリースデータを検証（簡易インターフェース）

    Args:
        excel_data: Excelデータ {year: [entries]}
        web_data: Webデータ {year: [entries]}
        ranking_name: ランキング名

    Returns:
        ValidationResult
    """
    validator = ReleaseValidator(
        excel_data=excel_data,
        web_data=web_data,
        ranking_name=ranking_name
    )
    return validator.validate_all()


def format_validation_report(result: ValidationResult) -> str:
    """検証結果をテキストレポート形式にフォーマット"""
    lines = []
    lines.append("=" * 60)
    lines.append("正誤チェックレポート")
    lines.append("=" * 60)

    # サマリー
    lines.append(f"\n■ 検証結果: {'✅ OK' if result.is_valid else '❌ 要修正'}")
    lines.append(f"  - エラー: {result.summary['ERROR']}件")
    lines.append(f"  - 警告: {result.summary['WARNING']}件")
    lines.append(f"  - 情報: {result.summary['INFO']}件")

    # エラー詳細
    errors = result.get_errors()
    if errors:
        lines.append(f"\n■ エラー ({len(errors)}件) - 修正が必要です")
        lines.append("-" * 40)
        for i, issue in enumerate(errors, 1):
            lines.append(f"{i}. [{issue.category}] {issue.message}")
            if issue.expected:
                lines.append(f"   期待値: {issue.expected}")
            if issue.actual:
                lines.append(f"   実際値: {issue.actual}")
            if issue.suggestion:
                lines.append(f"   提案: {issue.suggestion}")

    # 警告詳細
    warnings = result.get_warnings()
    if warnings:
        lines.append(f"\n■ 警告 ({len(warnings)}件) - 確認を推奨します")
        lines.append("-" * 40)
        for i, issue in enumerate(warnings, 1):
            lines.append(f"{i}. [{issue.category}] {issue.message}")
            if issue.suggestion:
                lines.append(f"   提案: {issue.suggestion}")

    lines.append("\n" + "=" * 60)
    return "\n".join(lines)


# ========================================
# デバッグ用
# ========================================
if __name__ == "__main__":
    # テストデータ
    test_excel_data = {
        2026: [
            {"rank": 1, "company": "SBI証券", "score": 68.9},
            {"rank": 1, "company": "楽天証券", "score": 68.9},  # 同率1位
            {"rank": 3, "company": "マネックス証券", "score": 67.5},
            {"rank": 4, "company": "松井証券", "score": 66.0},
        ]
    }

    test_web_data = {
        2025: [
            {"rank": 1, "company": "SBI証券", "score": 68.5},
            {"rank": 2, "company": "楽天証券株式会社", "score": 68.0},  # 表記ゆれ
            {"rank": 3, "company": "マネックス", "score": 67.0},  # 略称
            {"rank": 4, "company": "松井証券", "score": 65.5},
        ],
        2024: [
            {"rank": 1, "company": "SBI証券", "score": 68.0},
            {"rank": 2, "company": "楽天証券", "score": 67.5},
            {"rank": 3, "company": "マネックス証券", "score": 66.5},
        ]
    }

    # 検証実行
    result = validate_release_data(
        excel_data=test_excel_data,
        web_data=test_web_data,
        ranking_name="ネット証券"
    )

    # レポート出力
    print(format_validation_report(result))

# -*- coding: utf-8 -*-
"""
TOPICS分析ロジック（ルールベース）
v7.6 - コードリファクタリング
- 同点1位検出ロジックを共通メソッド(_count_wins_from_data)に統合
- 重複コード削減（~150行 × 3箇所を1箇所に集約）

v7.5 - 企業マスタ外部ファイル化
- 社名エイリアスを外部JSONファイル（config/company_aliases.json）で管理
- 非エンジニアでもエイリアス追加・編集が可能に

v7.4 - 同点1位対応強化
- 連続1位記録計算で同点1位を考慮（_calc_consecutive_wins, analyze_item_trends, analyze_dept_trends）
- Z会エイリアス追加（Z会の通信教育、Ｚ会の通信教育、Ｚ会 → Z会）

v7.3.1 - 社名正規化の網羅性向上、全角/半角・空白正規化追加
"""

import os
import json
import re
import logging
from typing import Dict, List, Any, Optional
from collections import defaultdict

logger = logging.getLogger(__name__)

# ========================================
# 社名エイリアス定義
# 外部ファイル（config/company_aliases.json）から読み込み
# ファイルがない場合はデフォルト値を使用
# ========================================

# デフォルトエイリアス（外部ファイルがない場合のフォールバック）
DEFAULT_COMPANY_ALIASES = {
    "JACリクルートメント": "JAC Recruitment",
    "ＪＡＣリクルートメント": "JAC Recruitment",
    "Z会の通信教育": "Z会",
    "Ｚ会の通信教育": "Z会",
    "Ｚ会": "Z会",
}


def _load_company_aliases() -> Dict[str, str]:
    """外部ファイルから社名エイリアスを読み込む

    Returns:
        社名エイリアス辞書（旧社名→正規社名）
    """
    config_path = os.path.join(
        os.path.dirname(__file__),
        "config",
        "company_aliases.json"
    )

    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                aliases = data.get("aliases", {})
                logger.info(f"社名エイリアスを読み込みました: {len(aliases)}件")
                return aliases
        except Exception as e:
            logger.warning(f"社名エイリアスファイル読み込みエラー: {e}")
            return DEFAULT_COMPANY_ALIASES
    else:
        logger.info("社名エイリアスファイルが見つかりません、デフォルト値を使用")
        return DEFAULT_COMPANY_ALIASES


# モジュール読み込み時にエイリアスをロード
COMPANY_ALIASES = _load_company_aliases()


def normalize_company_name(company: str) -> str:
    """社名を正規化する（エイリアスを正規社名に変換）

    v7.3.1: 前処理追加（全角/半角変換、空白正規化）

    Args:
        company: 企業名

    Returns:
        正規化された企業名
    """
    if not company:
        return company

    # 前処理: 文字列の正規化
    normalized = company.strip()

    # 全角英数字→半角英数字
    zen_to_han = str.maketrans(
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９',
        'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    )
    normalized = normalized.translate(zen_to_han)

    # 連続する空白を1つに（全角スペースも含む）
    normalized = re.sub(r'[\s　]+', ' ', normalized)

    # エイリアス適用（元の値と正規化後の値の両方でチェック）
    if company in COMPANY_ALIASES:
        return COMPANY_ALIASES[company]
    if normalized in COMPANY_ALIASES:
        return COMPANY_ALIASES[normalized]

    return normalized


class HistoricalAnalyzer:
    """歴代記録・得点推移の分析"""

    def __init__(self, overall_data: Dict, item_data: Dict, dept_data: Dict, ranking_name: str):
        """
        Args:
            overall_data: 総合ランキングデータ {年度: [企業データ]}
            item_data: 評価項目別データ {項目名: {年度: [企業データ]}}
            dept_data: 部門別データ {部門名: {年度: [企業データ]}}
            ranking_name: ランキング名
        """
        self.overall = overall_data
        self.items = item_data
        self.depts = dept_data
        self.ranking_name = ranking_name

    def analyze_all(self) -> Dict[str, Any]:
        """全分析を実行"""
        return {
            "historical_records": self.analyze_historical_records(),
            "score_trends": self.analyze_score_trends(),
            "item_trends": self.analyze_item_trends(),
            "dept_trends": self.analyze_dept_trends(),
        }

    def analyze_historical_records(self) -> Dict[str, Any]:
        """歴代記録を分析"""
        if not self.overall:
            return {}

        years = sorted(self.overall.keys())
        records = {
            "consecutive_wins": [],      # 連続1位記録
            "highest_scores": [],        # 過去最高得点
            "most_wins": [],             # 最多1位獲得
            "first_appearances": [],     # 初登場年
            "summary": {}                # サマリー
        }

        # === 連続1位記録 ===
        consecutive_records = self._calc_consecutive_wins()
        records["consecutive_wins"] = consecutive_records

        # === 過去最高得点 ===
        highest_scores = self._calc_highest_scores()
        records["highest_scores"] = highest_scores[:10]  # 上位10件

        # === 最多1位獲得 ===
        most_wins = self._calc_most_wins()
        records["most_wins"] = most_wins

        # === 初登場年 ===
        first_appearances = self._calc_first_appearances()
        records["first_appearances"] = first_appearances

        # === サマリー ===
        if consecutive_records:
            max_consecutive = max(consecutive_records, key=lambda x: x["years"])
            records["summary"]["max_consecutive"] = max_consecutive

        if highest_scores:
            records["summary"]["all_time_high"] = highest_scores[0]

        if most_wins:
            records["summary"]["most_wins"] = most_wins[0]

        return records

    def _calc_consecutive_wins(self) -> List[Dict]:
        """連続1位記録を計算（発表回数ベース：未発表年度をスキップして連続とカウント）

        修正: v4.7
        - 年度の連続性ではなく「発表回数」を基準にカウント
        - 未発表年度（データがない年度）があっても連続記録は途切れない
        - 例: 2016〜2021年が未発表でも、2015年→2022年で同じ企業が1位なら連続とカウント

        修正: v7.3
        - 社名エイリアス対応: 社名変更があっても連続記録を通算

        修正: v7.4
        - 同点1位対応: 同じ得点の企業はすべて1位としてカウントし、それぞれの連続記録を計算
        - 例: A社とB社が同点1位の年も、両社の連続記録として計算
        """
        if not self.overall:
            return []

        years = sorted(self.overall.keys())
        company_streaks = defaultdict(list)  # 企業ごとの連続1位期間
        company_current = {}  # 企業ごとの現在の連続状態

        for year in years:
            if not self.overall[year]:
                continue

            # v7.4: 同点1位の企業をすべて取得
            top_score = self.overall[year][0].get("score")
            top_companies = set()

            for entry in self.overall[year]:
                company_raw = entry.get("company", "")
                score = entry.get("score")

                # 同点1位の企業をすべて収集（得点が異なれば終了）
                if score is not None and score == top_score:
                    company = normalize_company_name(company_raw)
                    if company:
                        top_companies.add(company)
                elif score is not None and score != top_score:
                    break

            # 各1位企業について連続記録を更新
            for company in top_companies:
                if company in company_current:
                    # 連続継続
                    company_current[company]["count"] += 1
                    company_current[company]["years_list"].append(year)
                else:
                    # 新しい連続記録開始
                    company_current[company] = {
                        "start": year,
                        "count": 1,
                        "years_list": [year]
                    }

            # 1位から外れた企業の連続記録を確定
            for company in list(company_current.keys()):
                if company not in top_companies:
                    streak = company_current.pop(company)
                    company_streaks[company].append({
                        "start": streak["start"],
                        "end": streak["years_list"][-1] if streak["years_list"] else streak["start"],
                        "count": streak["count"],
                        "years_list": streak["years_list"]
                    })

        # 最後の連続記録を確定
        for company, streak in company_current.items():
            company_streaks[company].append({
                "start": streak["start"],
                "end": streak["years_list"][-1] if streak["years_list"] else streak["start"],
                "count": streak["count"],
                "years_list": streak["years_list"]
            })

        # 結果を整形
        results = []
        max_year = max(self.overall.keys())
        for company, streaks in company_streaks.items():
            for streak in streaks:
                if streak["count"] >= 1:
                    results.append({
                        "company": company,
                        "start_year": streak["start"],
                        "end_year": streak["end"],
                        "years": streak["count"],  # 発表回数を「連続年数」として表示
                        "years_list": streak["years_list"],  # 実際の年度リスト
                        "is_current": streak["end"] == max_year
                    })

        # 連続年数（発表回数）でソート
        results.sort(key=lambda x: (-x["years"], -x["end_year"]))
        return results

    def _count_wins_from_year_data(self, year_data: Dict) -> Dict[str, Dict]:
        """年度別データから1位獲得回数を集計する共通ヘルパー（v7.6追加）

        同点1位対応: 1位と同じ得点の企業はすべて1位としてカウント
        社名エイリアス対応: 社名変更があっても受賞回数を通算

        Args:
            year_data: {年度: [企業データ]} の形式

        Returns:
            {企業名: {"count": 回数, "years": [年度]}}
        """
        win_counts = defaultdict(lambda: {"count": 0, "years": []})

        # 辞書でない場合は空の結果を返す
        if not isinstance(year_data, dict):
            return win_counts

        for year, data in year_data.items():
            # データが空またはリストでない場合はスキップ
            if not data or not isinstance(data, list) or len(data) == 0:
                continue

            # 1位の得点を取得
            top_score = data[0].get("score")

            # 同点1位の企業をすべてカウント
            for entry in data:
                company_raw = entry.get("company", "")
                company = normalize_company_name(company_raw)
                score = entry.get("score")

                # 1位と同じ得点の企業は1位としてカウント
                if company and score is not None and score == top_score:
                    win_counts[company]["count"] += 1
                    win_counts[company]["years"].append(year)
                elif score is not None and score != top_score:
                    # 得点が異なったらループ終了
                    break

        return win_counts

    def _calc_highest_scores(self) -> List[Dict]:
        """過去最高得点を計算"""
        all_scores = []

        for year, data in self.overall.items():
            for item in data:
                score = item.get("score")
                company = item.get("company")
                rank = item.get("rank")
                if score and company:
                    all_scores.append({
                        "company": company,
                        "score": score,
                        "year": year,
                        "rank": rank
                    })

        # 得点でソート
        all_scores.sort(key=lambda x: -x["score"])
        return all_scores

    def _calc_most_wins(self) -> List[Dict]:
        """最多1位獲得を計算（総合ランキング）

        修正: v7.6 - 共通ヘルパーメソッドを使用
        修正: v4.5 - 同点1位対応
        修正: v7.3 - 社名エイリアス対応
        """
        win_counts = self._count_wins_from_year_data(self.overall)

        results = [
            {
                "company": company,
                "wins": info["count"],
                "years": sorted(info["years"]),
                "total_years": len(self.overall)
            }
            for company, info in win_counts.items()
        ]

        results.sort(key=lambda x: (-x["wins"], x["company"]))
        return results

    def calc_item_most_wins(self) -> Dict[str, List[Dict]]:
        """評価項目別の1位獲得回数を計算

        修正: v7.6 - 共通ヘルパーメソッドを使用
        修正: v4.5 - 同点1位対応
        修正: v7.3 - 社名エイリアス対応

        Returns:
            {項目名: [{"company": 企業名, "wins": 回数, "years": [年度], "total_years": 総年数}, ...]}
        """
        item_wins = {}

        for item_name, year_data in self.items.items():
            win_counts = self._count_wins_from_year_data(year_data)

            results = [
                {
                    "company": company,
                    "wins": info["count"],
                    "years": sorted(info["years"]),
                    "total_years": len(year_data)
                }
                for company, info in win_counts.items()
            ]
            results.sort(key=lambda x: (-x["wins"], x["company"]))
            item_wins[item_name] = results

        return item_wins

    def calc_dept_most_wins(self) -> Dict[str, List[Dict]]:
        """部門別の1位獲得回数を計算

        修正: v7.6 - 共通ヘルパーメソッドを使用
        修正: v4.5 - 同点1位対応
        修正: v7.3 - 社名エイリアス対応

        Returns:
            {部門名: [{"company": 企業名, "wins": 回数, "years": [年度], "total_years": 総年数}, ...]}
        """
        dept_wins = {}

        for dept_name, year_data in self.depts.items():
            win_counts = self._count_wins_from_year_data(year_data)

            results = [
                {
                    "company": company,
                    "wins": info["count"],
                    "years": sorted(info["years"]),
                    "total_years": len(year_data)
                }
                for company, info in win_counts.items()
            ]
            results.sort(key=lambda x: (-x["wins"], x["company"]))
            dept_wins[dept_name] = results

        return dept_wins

    def _calc_first_appearances(self) -> List[Dict]:
        """初登場年を計算

        修正: v7.3.1
        - 社名正規化を追加（エイリアス対応）
        """
        first_year = {}

        for year in sorted(self.overall.keys()):
            for item in self.overall[year]:
                company_raw = item.get("company", "")
                company = normalize_company_name(company_raw)
                if company and company not in first_year:
                    first_year[company] = {
                        "year": year,
                        "rank": item.get("rank"),
                        "score": item.get("score")
                    }

        results = [
            {
                "company": company,
                "first_year": info["year"],
                "first_rank": info["rank"],
                "first_score": info["score"]
            }
            for company, info in first_year.items()
        ]

        results.sort(key=lambda x: (x["first_year"], x["first_rank"] or 999))
        return results

    def analyze_score_trends(self) -> Dict[str, Any]:
        """総合ランキングの得点推移を分析

        修正: v7.3.1
        - 社名正規化を追加（エイリアス対応）
        """
        if not self.overall:
            return {}

        years = sorted(self.overall.keys())
        trends = {
            "years": years,
            "companies": {},          # 企業別得点推移
            "top_companies": [],      # 上位企業リスト
            "average_scores": {},     # 年度別平均得点
            "top_score_by_year": {},  # 年度別1位得点
        }

        # 全企業を収集（正規化済み）
        all_companies = set()
        for year_data in self.overall.values():
            for item in year_data:
                company = normalize_company_name(item.get("company", ""))
                all_companies.add(company)

        # 企業別得点推移
        for company in all_companies:
            if not company:
                continue
            trends["companies"][company] = {}
            for year in years:
                score = None
                rank = None
                for item in self.overall.get(year, []):
                    item_company = normalize_company_name(item.get("company", ""))
                    if item_company == company:
                        score = item.get("score")
                        rank = item.get("rank")
                        break
                trends["companies"][company][year] = {
                    "score": score,
                    "rank": rank
                }

        # 年度別統計
        for year in years:
            scores = [item.get("score") for item in self.overall.get(year, []) if item.get("score")]
            if scores:
                trends["average_scores"][year] = round(sum(scores) / len(scores), 2)
                trends["top_score_by_year"][year] = {
                    "score": max(scores),
                    "company": self.overall[year][0].get("company") if self.overall[year] else None
                }

        # 上位企業（最新年度ベース）
        latest_year = max(years)
        if self.overall.get(latest_year):
            trends["top_companies"] = [
                item.get("company") for item in self.overall[latest_year][:10]
            ]

        return trends

    def analyze_item_trends(self) -> Dict[str, Dict]:
        """評価項目別の得点推移を分析

        修正: v6.1
        - 連続記録は「発表回数」を基準にカウント（年度差ではなく）
        - 発表がない年はスキップして連続とカウント
        - 例: 2016〜2025(2018は発表なし)の場合、連続年数は9年

        修正: v7.4
        - 同点1位対応: 同じ得点の企業はすべて1位としてカウント
        """
        if not self.items:
            return {}

        item_trends = {}

        for item_name, year_data in self.items.items():
            if not isinstance(year_data, dict):
                continue

            years = sorted(year_data.keys())
            item_trends[item_name] = {
                "years": years,
                "top_by_year": {},      # 年度別1位
                "consecutive_wins": [],  # 連続1位記録
            }

            # 連続1位計算用（同点1位対応）
            company_current = {}  # 企業ごとの現在の連続状態

            for year in years:
                if not year_data.get(year):
                    continue

                data = year_data[year]
                top = data[0]
                top_score = top.get("score")

                # 年度別1位（表示用、同点含む）
                item_trends[item_name]["top_by_year"][year] = {
                    "company": top.get("company"),
                    "score": top_score
                }

                # v7.4: 同点1位の企業をすべて取得
                top_companies = set()
                for entry in data:
                    company_raw = entry.get("company", "")
                    score = entry.get("score")
                    if score is not None and score == top_score:
                        company = normalize_company_name(company_raw)
                        if company:
                            top_companies.add(company)
                    elif score is not None and score != top_score:
                        break

                # 各1位企業について連続記録を更新
                for company in top_companies:
                    if company in company_current:
                        company_current[company]["count"] += 1
                        company_current[company]["years_list"].append(year)
                    else:
                        company_current[company] = {
                            "start": year,
                            "count": 1,
                            "years_list": [year]
                        }

                # 1位から外れた企業の連続記録を確定
                for company in list(company_current.keys()):
                    if company not in top_companies:
                        streak = company_current.pop(company)
                        item_trends[item_name]["consecutive_wins"].append({
                            "company": company,
                            "start": streak["start"],
                            "end": streak["years_list"][-1] if streak["years_list"] else streak["start"],
                            "years": streak["count"],
                            "years_list": streak["years_list"]
                        })

            # 最後の連続記録を確定
            max_year = max(years) if years else None
            for company, streak in company_current.items():
                item_trends[item_name]["consecutive_wins"].append({
                    "company": company,
                    "start": streak["start"],
                    "end": streak["years_list"][-1] if streak["years_list"] else streak["start"],
                    "years": streak["count"],
                    "years_list": streak["years_list"],
                    "is_current": streak["years_list"][-1] == max_year if streak["years_list"] and max_year else False
                })

        return item_trends

    def analyze_dept_trends(self) -> Dict[str, Dict]:
        """部門別の得点推移を分析

        修正: v6.1
        - 連続記録は「発表回数」を基準にカウント（年度差ではなく）
        - 発表がない年はスキップして連続とカウント
        - 例: 2016〜2025(2018は発表なし)の場合、連続年数は9年

        修正: v7.4
        - 同点1位対応: 同じ得点の企業はすべて1位としてカウント
        """
        if not self.depts:
            return {}

        dept_trends = {}

        for dept_name, year_data in self.depts.items():
            if not isinstance(year_data, dict):
                continue

            years = sorted(year_data.keys())
            dept_trends[dept_name] = {
                "years": years,
                "top_by_year": {},
                "consecutive_wins": [],
            }

            # 連続1位計算用（同点1位対応）
            company_current = {}  # 企業ごとの現在の連続状態

            for year in years:
                if not year_data.get(year):
                    continue

                data = year_data[year]
                top = data[0]
                top_score = top.get("score")

                # 年度別1位（表示用）
                dept_trends[dept_name]["top_by_year"][year] = {
                    "company": top.get("company"),
                    "score": top_score
                }

                # v7.4: 同点1位の企業をすべて取得
                top_companies = set()
                for entry in data:
                    company_raw = entry.get("company", "")
                    score = entry.get("score")
                    if score is not None and score == top_score:
                        company = normalize_company_name(company_raw)
                        if company:
                            top_companies.add(company)
                    elif score is not None and score != top_score:
                        break

                # 各1位企業について連続記録を更新
                for company in top_companies:
                    if company in company_current:
                        company_current[company]["count"] += 1
                        company_current[company]["years_list"].append(year)
                    else:
                        company_current[company] = {
                            "start": year,
                            "count": 1,
                            "years_list": [year]
                        }

                # 1位から外れた企業の連続記録を確定
                for company in list(company_current.keys()):
                    if company not in top_companies:
                        streak = company_current.pop(company)
                        dept_trends[dept_name]["consecutive_wins"].append({
                            "company": company,
                            "start": streak["start"],
                            "end": streak["years_list"][-1] if streak["years_list"] else streak["start"],
                            "years": streak["count"],
                            "years_list": streak["years_list"]
                        })

            # 最後の連続記録を確定
            max_year = max(years) if years else None
            for company, streak in company_current.items():
                dept_trends[dept_name]["consecutive_wins"].append({
                    "company": company,
                    "start": streak["start"],
                    "end": streak["years_list"][-1] if streak["years_list"] else streak["start"],
                    "years": streak["count"],
                    "years_list": streak["years_list"],
                    "is_current": streak["years_list"][-1] == max_year if streak["years_list"] and max_year else False
                })

        return dept_trends


class TopicsAnalyzer:
    """ランキングデータからTOPICSを抽出"""

    def __init__(self, overall_data: Dict, item_data: Dict, ranking_name: str, dept_data: Dict = None):
        """
        Args:
            overall_data: 総合ランキングデータ {年度: [企業データ]}
            item_data: 評価項目別データ {項目名: {年度: [企業データ]}}
            ranking_name: ランキング名
            dept_data: 部門別データ {部門名: {年度: [企業データ]}} (v5.8追加)
        """
        self.overall = overall_data
        self.items = item_data
        self.depts = dept_data or {}
        self.ranking_name = ranking_name

    def analyze(self) -> Dict[str, Any]:
        """
        TOPICS分析を実行

        Returns:
            {
                "recommended": [推奨TOPICS],
                "other": [その他TOPICS],
                "headlines": [見出し案]
            }
        """
        recommended = []
        other = []

        # 1. 連続1位を分析（総合ランキング）
        consecutive = self._analyze_consecutive_wins()
        if consecutive:
            recommended.append(consecutive)

        # 2. 得点差を分析（総合ランキング）
        score_diff = self._analyze_score_difference()
        if score_diff:
            recommended.append(score_diff)

        # 3. 評価項目の独占を分析
        item_dominance = self._analyze_item_dominance()
        if item_dominance:
            recommended.append(item_dominance)

        # 4. 評価項目別の連続1位記録を分析 (v5.8追加)
        item_consecutive = self._analyze_item_consecutive_wins()
        for topic in item_consecutive:
            recommended.append(topic)

        # 5. 部門別の連続1位記録を分析 (v5.8追加)
        dept_consecutive = self._analyze_dept_consecutive_wins()
        for topic in dept_consecutive:
            recommended.append(topic)

        # 6. 部門別の独占状況を分析 (v5.8追加)
        dept_dominance = self._analyze_dept_dominance()
        if dept_dominance:
            recommended.append(dept_dominance)

        # 7. 項目別の特徴を分析
        item_features = self._analyze_item_features()
        other.extend(item_features)

        # 8. 部門別の特徴を分析 (v5.8追加)
        dept_features = self._analyze_dept_features()
        other.extend(dept_features)

        # 9. 順位変動を分析
        rank_changes = self._analyze_rank_changes()
        other.extend(rank_changes)

        # impactでソートして重複を除去
        recommended = sorted(recommended, key=lambda x: x.get("impact", 0), reverse=True)

        # 見出し案を生成
        headlines = self._generate_headlines(recommended)

        return {
            "recommended": recommended,
            "other": other,
            "headlines": headlines
        }

    def _analyze_consecutive_wins(self) -> Dict:
        """連続1位を分析（v7.5: 同点1位対応）"""
        if not self.overall:
            return None

        years = sorted(self.overall.keys(), reverse=True)
        if not years:
            return None

        # 最新の1位（同点1位の企業をすべて取得）
        latest_year = years[0]
        if not self.overall[latest_year]:
            return None

        top_score = self.overall[latest_year][0].get("score")
        top_companies = []
        for entry in self.overall[latest_year]:
            score = entry.get("score")
            if score is not None and score == top_score:
                company = normalize_company_name(entry.get("company", ""))
                if company:
                    top_companies.append(company)
            elif score is not None and score != top_score:
                break

        if not top_companies:
            return None

        # 各1位企業の連続年数をカウント
        best_consecutive = 0
        best_company = top_companies[0]

        for company in top_companies:
            consecutive = 0
            for year in years:
                if not self.overall[year]:
                    continue
                # その年の1位企業（同点含む）を取得
                year_top_score = self.overall[year][0].get("score")
                year_top_companies = set()
                for entry in self.overall[year]:
                    score = entry.get("score")
                    if score is not None and score == year_top_score:
                        c = normalize_company_name(entry.get("company", ""))
                        if c:
                            year_top_companies.add(c)
                    elif score is not None and score != year_top_score:
                        break

                if company in year_top_companies:
                    consecutive += 1
                else:
                    break

            if consecutive > best_consecutive:
                best_consecutive = consecutive
                best_company = company

        if best_consecutive >= 2:
            # 同点1位の場合のタイトル調整
            if len(top_companies) >= 2:
                companies_str = "と".join(top_companies[:2])
                if len(top_companies) > 2:
                    companies_str += f"ら{len(top_companies)}社"
                return {
                    "category": "総合ランキング",
                    "importance": "最重要",
                    "title": f"{companies_str}が同率1位、{best_company}は{best_consecutive}年連続",
                    "evidence": f"{years[best_consecutive-1]}年〜{latest_year}年の総合ランキング1位",
                    "impact": 5
                }
            else:
                return {
                    "category": "総合ランキング",
                    "importance": "最重要",
                    "title": f"{best_company}が{best_consecutive}年連続で総合1位を達成",
                    "evidence": f"{years[best_consecutive-1]}年〜{latest_year}年の総合ランキング1位",
                    "impact": 5
                }
        elif best_consecutive == 1:
            # 前年と比較
            if len(years) >= 2 and self.overall[years[1]]:
                prev_top = normalize_company_name(self.overall[years[1]][0].get("company", ""))
                if prev_top not in top_companies:
                    if len(top_companies) >= 2:
                        companies_str = "と".join(top_companies[:2])
                        return {
                            "category": "総合ランキング",
                            "importance": "重要",
                            "title": f"{companies_str}が同率1位、{prev_top}から首位交代",
                            "evidence": f"{years[1]}年1位の{prev_top}から{latest_year}年は{companies_str}が1位に",
                            "impact": 5
                        }
                    else:
                        return {
                            "category": "総合ランキング",
                            "importance": "重要",
                            "title": f"{best_company}が{prev_top}を抜いて総合1位を獲得",
                            "evidence": f"{years[1]}年1位の{prev_top}から{latest_year}年は{best_company}が1位に",
                            "impact": 5
                        }

        return None

    def _analyze_score_difference(self) -> Dict:
        """得点差を分析（同率順位対応）"""
        if not self.overall:
            return None

        latest_year = max(self.overall.keys())
        data = self.overall[latest_year]

        if len(data) < 2:
            return None

        first = data[0]
        score1 = first.get("score")

        # 0点も有効な値として扱う（Noneのみを除外）
        if score1 is None:
            return None

        # 同率1位のチェック: 同じ得点の企業をすべて収集
        tied_companies = [first["company"]]
        for entry in data[1:]:
            if entry.get("score") == score1:
                tied_companies.append(entry["company"])
            else:
                break  # 得点が異なる企業が出たらループ終了

        # 同率1位が2社以上の場合
        if len(tied_companies) >= 2:
            if len(tied_companies) == 2:
                return {
                    "category": "総合ランキング",  # v5.9: カテゴリ追加
                    "importance": "重要",
                    "title": f"{tied_companies[0]}と{tied_companies[1]}が同率1位",
                    "evidence": f"両社とも{score1}点で並ぶ",
                    "impact": 5
                }
            else:
                # 3社以上の同率1位
                companies_str = "、".join(tied_companies)
                return {
                    "category": "総合ランキング",  # v5.9: カテゴリ追加
                    "importance": "重要",
                    "title": f"{len(tied_companies)}社が同率1位で並ぶ",
                    "evidence": f"{companies_str}（いずれも{score1}点）",
                    "impact": 5
                }

        # 同率1位でない場合、2位との差を分析
        second = data[1]
        score2 = second.get("score")

        if score2 is None:
            return None

        diff = round(score1 - score2, 1)

        if diff >= 2.0:
            return {
                "category": "総合ランキング",  # v5.9: カテゴリ追加
                "importance": "重要",
                "title": f"1位と2位の得点差{diff}点、{first['company']}が大きく引き離す",
                "evidence": f"{first['company']}({score1}点) vs {second['company']}({score2}点)",
                "impact": 4
            }
        elif diff <= 0.5:
            return {
                "category": "総合ランキング",  # v5.9: カテゴリ追加
                "importance": "注目",
                "title": f"1位と2位の得点差わずか{diff}点の僅差",
                "evidence": f"{first['company']}({score1}点) vs {second['company']}({score2}点)",
                "impact": 3
            }

        return None

    def _analyze_item_dominance(self) -> Dict:
        """評価項目の独占状況を分析"""
        if not self.items:
            return None

        # 各社の1位獲得数をカウント
        wins = {}
        actual_items = 0  # 実際にデータがある項目数（空データを除外）

        for item_name, year_data in self.items.items():
            # 新形式（経年データ）の場合は最新年度を使用
            if isinstance(year_data, dict):
                if year_data:
                    latest_year = max(year_data.keys())
                    data = year_data[latest_year]
                else:
                    continue
            else:
                data = year_data

            if data:
                actual_items += 1  # データがある項目のみカウント
                top_company = data[0].get("company", "")
                if top_company:
                    wins[top_company] = wins.get(top_company, 0) + 1

        if not wins or actual_items == 0:
            return None

        # 最多1位獲得企業
        top_winner = max(wins.items(), key=lambda x: x[1])
        company, count = top_winner

        if count >= actual_items * 0.6:  # 60%以上で「独占」（実データ数基準）
            return {
                "category": "評価項目別",  # v5.9: カテゴリ追加
                "importance": "重要",
                "title": f"{company}が{actual_items}項目中{count}項目で1位を独占",
                "evidence": f"評価項目別ランキングで圧倒的な強さ",
                "impact": 4
            }
        elif count >= 3:
            return {
                "category": "評価項目別",  # v5.9: カテゴリ追加
                "importance": "注目",
                "title": f"{company}が{count}項目で1位を獲得",
                "evidence": f"複数の評価項目で高評価",
                "impact": 3
            }

        return None

    def _analyze_item_features(self) -> List[str]:
        """評価項目別の特徴を分析"""
        features = []

        if not self.items:
            return features

        for item_name, year_data in self.items.items():
            # 新形式（経年データ）の場合は最新年度を使用
            if isinstance(year_data, dict):
                if year_data:
                    latest_year = max(year_data.keys())
                    data = year_data[latest_year]
                else:
                    continue
            else:
                data = year_data

            if len(data) >= 2:
                first = data[0]
                second = data[1]

                score1 = first.get("score")
                score2 = second.get("score")

                # 0点も有効な値として扱う（Noneのみを除外）
                if score1 is not None and score2 is not None:
                    diff = round(score1 - score2, 1)

                    if diff >= 3.0:
                        features.append(
                            f"『{item_name}』で{first['company']}が{score1}点、"
                            f"2位と{diff}点差の圧倒的高評価"
                        )

        return features[:3]  # 上位3つまで

    def _analyze_rank_changes(self) -> List[str]:
        """順位変動を分析"""
        changes = []

        if len(self.overall) < 2:
            return changes

        years = sorted(self.overall.keys(), reverse=True)
        latest = self.overall[years[0]]
        previous = self.overall[years[1]]

        if not latest or not previous:
            return changes

        # 前年の順位をマップ化
        prev_ranks = {d.get("company"): d.get("rank") for d in previous}

        for company_data in latest:
            company = company_data.get("company")
            current_rank = company_data.get("rank")
            prev_rank = prev_ranks.get(company)

            if prev_rank and current_rank:
                if prev_rank - current_rank >= 2:
                    changes.append(
                        f"{company}が前年{prev_rank}位→{current_rank}位に躍進"
                    )
                elif current_rank - prev_rank >= 2:
                    changes.append(
                        f"{company}が前年{prev_rank}位→{current_rank}位に後退"
                    )

        return changes[:2]

    def _generate_headlines(self, recommended: List[Dict]) -> List[str]:
        """見出し案を生成"""
        headlines = []

        if not recommended:
            return ["データ不足のため見出し案を生成できませんでした"]

        # 最重要トピックから見出しを生成
        main_topic = recommended[0] if recommended else None

        if main_topic:
            # パターンA: メインのみ
            headlines.append(f"「{main_topic['title'].split('が')[0]}」{main_topic['title'].split('が')[1] if 'が' in main_topic['title'] else main_topic['title']}")

            # パターンB: メイン + サブ
            if len(recommended) >= 2:
                sub = recommended[1]["title"]
                # 簡略化
                sub_short = sub.split("、")[0] if "、" in sub else sub[:30]
                headlines.append(f"{main_topic['title']}　〜{sub_short}〜")

            # パターンC: ランキング名を含む
            headlines.append(f"『{self.ranking_name}』満足度調査　{main_topic['title']}")

        return headlines

    # ========================================
    # v5.8追加: 評価項目別・部門別の連続記録分析
    # ========================================

    def _analyze_item_consecutive_wins(self) -> List[Dict]:
        """評価項目別の連続1位記録を分析（v7.5: 同点1位対応）"""
        topics = []

        if not self.items:
            return topics

        for item_name, year_data in self.items.items():
            if not isinstance(year_data, dict) or not year_data:
                continue

            years = sorted(year_data.keys())
            if len(years) < 2:
                continue

            latest_year = years[-1]

            # 最新年の1位企業（同点含む）を取得
            if not year_data.get(latest_year):
                continue
            latest_data = year_data[latest_year]
            top_score = latest_data[0].get("score")
            latest_top_companies = set()
            for entry in latest_data:
                score = entry.get("score")
                if score is not None and score == top_score:
                    company = normalize_company_name(entry.get("company", ""))
                    if company:
                        latest_top_companies.add(company)
                elif score is not None and score != top_score:
                    break

            # 各1位企業の連続年数をカウント
            for company in latest_top_companies:
                consecutive_count = 0
                streak_start = None
                for year in reversed(years):
                    if not year_data.get(year):
                        continue
                    data = year_data[year]
                    year_top_score = data[0].get("score")
                    year_top_companies = set()
                    for entry in data:
                        score = entry.get("score")
                        if score is not None and score == year_top_score:
                            c = normalize_company_name(entry.get("company", ""))
                            if c:
                                year_top_companies.add(c)
                        elif score is not None and score != year_top_score:
                            break

                    if company in year_top_companies:
                        consecutive_count += 1
                        streak_start = year
                    else:
                        break

                if consecutive_count >= 3:  # 3年以上連続
                    topics.append({
                        "importance": "注目",
                        "title": f"『{item_name}』で{company}が{consecutive_count}年連続1位",
                        "evidence": f"{streak_start}年〜{latest_year}年の評価項目別ランキング",
                        "impact": min(4, 2 + consecutive_count // 2),
                        "category": "評価項目別"
                    })

        # impactが高い順にソートして上位3件まで
        topics = sorted(topics, key=lambda x: x["impact"], reverse=True)[:3]
        return topics

    def _analyze_dept_consecutive_wins(self) -> List[Dict]:
        """部門別の連続1位記録を分析（v7.5: 同点1位対応）"""
        topics = []

        if not self.depts:
            return topics

        for dept_name, year_data in self.depts.items():
            if not isinstance(year_data, dict) or not year_data:
                continue

            years = sorted(year_data.keys())
            if len(years) < 2:
                continue

            latest_year = years[-1]

            # 最新年の1位企業（同点含む）を取得
            if not year_data.get(latest_year):
                continue
            latest_data = year_data[latest_year]
            top_score = latest_data[0].get("score")
            latest_top_companies = set()
            for entry in latest_data:
                score = entry.get("score")
                if score is not None and score == top_score:
                    company = normalize_company_name(entry.get("company", ""))
                    if company:
                        latest_top_companies.add(company)
                elif score is not None and score != top_score:
                    break

            # 各1位企業の連続年数をカウント
            for company in latest_top_companies:
                consecutive_count = 0
                streak_start = None
                for year in reversed(years):
                    if not year_data.get(year):
                        continue
                    data = year_data[year]
                    year_top_score = data[0].get("score")
                    year_top_companies = set()
                    for entry in data:
                        score = entry.get("score")
                        if score is not None and score == year_top_score:
                            c = normalize_company_name(entry.get("company", ""))
                            if c:
                                year_top_companies.add(c)
                        elif score is not None and score != year_top_score:
                            break

                    if company in year_top_companies:
                        consecutive_count += 1
                        streak_start = year
                    else:
                        break

                if consecutive_count >= 3:  # 3年以上連続
                    topics.append({
                        "importance": "注目",
                        "title": f"『{dept_name}』部門で{company}が{consecutive_count}年連続1位",
                        "evidence": f"{streak_start}年〜{latest_year}年の部門別ランキング",
                        "impact": min(4, 2 + consecutive_count // 2),
                        "category": "部門別"
                    })

        topics = sorted(topics, key=lambda x: x["impact"], reverse=True)[:3]
        return topics

    def _analyze_dept_dominance(self) -> Dict:
        """部門別の独占状況を分析"""
        if not self.depts:
            return None

        # 各社の1位獲得数をカウント
        wins = {}
        actual_depts = 0

        for dept_name, year_data in self.depts.items():
            if not isinstance(year_data, dict):
                continue

            if year_data:
                latest_year = max(year_data.keys())
                data = year_data[latest_year]
            else:
                continue

            if data:
                actual_depts += 1
                top_company = data[0].get("company", "")
                if top_company:
                    wins[top_company] = wins.get(top_company, 0) + 1

        if not wins or actual_depts == 0:
            return None

        top_winner = max(wins.items(), key=lambda x: x[1])
        company, count = top_winner

        if count >= actual_depts * 0.6:  # 60%以上で「独占」
            return {
                "importance": "重要",
                "title": f"{company}が{actual_depts}部門中{count}部門で1位を独占",
                "evidence": f"部門別ランキングで圧倒的な強さ",
                "impact": 4,
                "category": "部門別"
            }
        elif count >= 3:
            return {
                "importance": "注目",
                "title": f"{company}が{count}部門で1位を獲得",
                "evidence": f"複数の部門で高評価",
                "impact": 3,
                "category": "部門別"
            }

        return None

    def _analyze_dept_features(self) -> List[str]:
        """部門別の特徴を分析（得点差が大きい部門など）"""
        features = []

        if not self.depts:
            return features

        for dept_name, year_data in self.depts.items():
            if not isinstance(year_data, dict):
                continue

            if year_data:
                latest_year = max(year_data.keys())
                data = year_data[latest_year]
            else:
                continue

            if len(data) >= 2:
                first = data[0]
                second = data[1]

                score1 = first.get("score")
                score2 = second.get("score")

                if score1 is not None and score2 is not None:
                    diff = round(score1 - score2, 1)

                    if diff >= 3.0:
                        features.append(
                            f"『{dept_name}』部門で{first['company']}が{score1}点、"
                            f"2位と{diff}点差の圧倒的高評価"
                        )

        return features[:3]  # 上位3つまで

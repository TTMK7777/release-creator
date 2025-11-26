# -*- coding: utf-8 -*-
"""
TOPICS分析ロジック（ルールベース）
"""

from typing import Dict, List, Any, Optional
from collections import defaultdict


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
        """連続1位記録を計算（年度欠落を考慮）"""
        if not self.overall:
            return []

        years = sorted(self.overall.keys())
        company_streaks = defaultdict(list)  # 企業ごとの連続1位期間

        current_company = None
        streak_start = None
        prev_year = None
        actual_consecutive_years = 0  # 実際の連続年数

        for year in years:
            if not self.overall[year]:
                continue

            top_company = self.overall[year][0].get("company", "")

            # 年度が連続しているかチェック（欠落年度がある場合は連続を切る）
            is_consecutive_year = prev_year is None or year == prev_year + 1

            if top_company == current_company and is_consecutive_year:
                # 連続中
                actual_consecutive_years += 1
            else:
                # 連続が途切れた or 新しい連続開始
                if current_company and streak_start and actual_consecutive_years >= 1:
                    company_streaks[current_company].append({
                        "start": streak_start,
                        "end": prev_year,
                        "years": actual_consecutive_years
                    })
                current_company = top_company
                streak_start = year
                actual_consecutive_years = 1

            prev_year = year

        # 最後の連続記録
        if current_company and streak_start and actual_consecutive_years >= 1:
            company_streaks[current_company].append({
                "start": streak_start,
                "end": prev_year,
                "years": actual_consecutive_years
            })

        # 結果を整形
        results = []
        for company, streaks in company_streaks.items():
            for streak in streaks:
                if streak["years"] >= 1:
                    results.append({
                        "company": company,
                        "start_year": streak["start"],
                        "end_year": streak["end"],
                        "years": streak["years"],
                        "is_current": streak["end"] == max(self.overall.keys())
                    })

        # 連続年数でソート
        results.sort(key=lambda x: (-x["years"], -x["end_year"]))
        return results

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
        """最多1位獲得を計算"""
        win_counts = defaultdict(lambda: {"count": 0, "years": []})

        for year, data in self.overall.items():
            if data:
                top_company = data[0].get("company", "")
                if top_company:
                    win_counts[top_company]["count"] += 1
                    win_counts[top_company]["years"].append(year)

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

    def _calc_first_appearances(self) -> List[Dict]:
        """初登場年を計算"""
        first_year = {}

        for year in sorted(self.overall.keys()):
            for item in self.overall[year]:
                company = item.get("company", "")
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
        """総合ランキングの得点推移を分析"""
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

        # 全企業を収集
        all_companies = set()
        for year_data in self.overall.values():
            for item in year_data:
                all_companies.add(item.get("company", ""))

        # 企業別得点推移
        for company in all_companies:
            if not company:
                continue
            trends["companies"][company] = {}
            for year in years:
                score = None
                rank = None
                for item in self.overall.get(year, []):
                    if item.get("company") == company:
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
        """評価項目別の得点推移を分析"""
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

            # 年度別1位
            current_company = None
            streak_start = None

            for year in years:
                if year_data.get(year):
                    top = year_data[year][0]
                    item_trends[item_name]["top_by_year"][year] = {
                        "company": top.get("company"),
                        "score": top.get("score")
                    }

                    # 連続1位計算
                    top_company = top.get("company")
                    if top_company == current_company:
                        pass
                    else:
                        if current_company and streak_start:
                            prev_year = years[years.index(year) - 1]
                            item_trends[item_name]["consecutive_wins"].append({
                                "company": current_company,
                                "start": streak_start,
                                "end": prev_year,
                                "years": prev_year - streak_start + 1
                            })
                        current_company = top_company
                        streak_start = year

            # 最後の連続
            if current_company and streak_start:
                item_trends[item_name]["consecutive_wins"].append({
                    "company": current_company,
                    "start": streak_start,
                    "end": years[-1],
                    "years": years[-1] - streak_start + 1,
                    "is_current": True
                })

        return item_trends

    def analyze_dept_trends(self) -> Dict[str, Dict]:
        """部門別の得点推移を分析"""
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

            current_company = None
            streak_start = None

            for year in years:
                if year_data.get(year):
                    top = year_data[year][0]
                    dept_trends[dept_name]["top_by_year"][year] = {
                        "company": top.get("company"),
                        "score": top.get("score")
                    }

                    top_company = top.get("company")
                    if top_company == current_company:
                        pass
                    else:
                        if current_company and streak_start:
                            prev_year = years[years.index(year) - 1]
                            dept_trends[dept_name]["consecutive_wins"].append({
                                "company": current_company,
                                "start": streak_start,
                                "end": prev_year,
                                "years": prev_year - streak_start + 1
                            })
                        current_company = top_company
                        streak_start = year

            if current_company and streak_start:
                dept_trends[dept_name]["consecutive_wins"].append({
                    "company": current_company,
                    "start": streak_start,
                    "end": years[-1],
                    "years": years[-1] - streak_start + 1,
                    "is_current": True
                })

        return dept_trends


class TopicsAnalyzer:
    """ランキングデータからTOPICSを抽出"""

    def __init__(self, overall_data: Dict, item_data: Dict, ranking_name: str):
        """
        Args:
            overall_data: 総合ランキングデータ {年度: [企業データ]}
            item_data: 評価項目別データ {項目名: [企業データ]}
            ranking_name: ランキング名
        """
        self.overall = overall_data
        self.items = item_data
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

        # 1. 連続1位を分析
        consecutive = self._analyze_consecutive_wins()
        if consecutive:
            recommended.append(consecutive)

        # 2. 得点差を分析
        score_diff = self._analyze_score_difference()
        if score_diff:
            recommended.append(score_diff)

        # 3. 評価項目の独占を分析
        item_dominance = self._analyze_item_dominance()
        if item_dominance:
            recommended.append(item_dominance)

        # 4. 項目別の特徴を分析
        item_features = self._analyze_item_features()
        other.extend(item_features)

        # 5. 順位変動を分析
        rank_changes = self._analyze_rank_changes()
        other.extend(rank_changes)

        # 見出し案を生成
        headlines = self._generate_headlines(recommended)

        return {
            "recommended": recommended,
            "other": other,
            "headlines": headlines
        }

    def _analyze_consecutive_wins(self) -> Dict:
        """連続1位を分析"""
        if not self.overall:
            return None

        years = sorted(self.overall.keys(), reverse=True)
        if not years:
            return None

        # 最新の1位
        latest_year = years[0]
        if not self.overall[latest_year]:
            return None

        top_company = self.overall[latest_year][0].get("company", "")

        # 連続年数をカウント
        consecutive = 0
        for year in years:
            if self.overall[year] and self.overall[year][0].get("company") == top_company:
                consecutive += 1
            else:
                break

        if consecutive >= 2:
            return {
                "importance": "最重要",
                "title": f"{top_company}が{consecutive}年連続で総合1位を達成",
                "evidence": f"{years[-consecutive+1] if consecutive > 1 else latest_year}年〜{latest_year}年の総合ランキング1位",
                "impact": 5
            }
        elif consecutive == 1:
            # 前年と比較
            if len(years) >= 2:
                prev_top = self.overall[years[1]][0].get("company", "") if self.overall[years[1]] else ""
                if prev_top != top_company:
                    return {
                        "importance": "重要",
                        "title": f"{top_company}が{prev_top}を抜いて総合1位を獲得",
                        "evidence": f"{years[1]}年1位の{prev_top}から{latest_year}年は{top_company}が1位に",
                        "impact": 5
                    }

        return None

    def _analyze_score_difference(self) -> Dict:
        """得点差を分析"""
        if not self.overall:
            return None

        latest_year = max(self.overall.keys())
        data = self.overall[latest_year]

        if len(data) < 2:
            return None

        first = data[0]
        second = data[1]

        score1 = first.get("score")
        score2 = second.get("score")

        # 0点も有効な値として扱う（Noneのみを除外）
        if score1 is None or score2 is None:
            return None

        diff = round(score1 - score2, 1)

        if diff >= 2.0:
            return {
                "importance": "重要",
                "title": f"1位と2位の得点差{diff}点、{first['company']}が大きく引き離す",
                "evidence": f"{first['company']}({score1}点) vs {second['company']}({score2}点)",
                "impact": 4
            }
        elif diff <= 0.5:
            return {
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
                "importance": "重要",
                "title": f"{company}が{actual_items}項目中{count}項目で1位を独占",
                "evidence": f"評価項目別ランキングで圧倒的な強さ",
                "impact": 4
            }
        elif count >= 3:
            return {
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

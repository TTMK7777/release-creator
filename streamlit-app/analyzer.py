# -*- coding: utf-8 -*-
"""
TOPICS分析ロジック（ルールベース）
"""

from typing import Dict, List, Any

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

        score1 = first.get("score", 0)
        score2 = second.get("score", 0)

        if not score1 or not score2:
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
        total_items = len(self.items)

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
                top_company = data[0].get("company", "")
                if top_company:
                    wins[top_company] = wins.get(top_company, 0) + 1

        if not wins:
            return None

        # 最多1位獲得企業
        top_winner = max(wins.items(), key=lambda x: x[1])
        company, count = top_winner

        if count >= total_items * 0.6:  # 60%以上で「独占」
            return {
                "importance": "重要",
                "title": f"{company}が{total_items}項目中{count}項目で1位を独占",
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

                score1 = first.get("score", 0)
                score2 = second.get("score", 0)

                if score1 and score2:
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

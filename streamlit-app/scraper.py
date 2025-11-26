# -*- coding: utf-8 -*-
"""
オリコン顧客満足度サイトのスクレイパー
"""

import requests
from bs4 import BeautifulSoup
import re
from typing import Dict, List, Optional
import time

class OriconScraper:
    """オリコン顧客満足度サイトからランキングデータを取得

    対応サブドメイン:
    - life.oricon.co.jp: 生活系（金融、保険、通信、住宅、生活サービスなど）
    - juken.oricon.co.jp: 教育系（英会話、家庭教師、通信教育、スイミングなど）
    - career.oricon.co.jp: キャリア系（転職エージェント、派遣会社など）
    """

    # サブドメインのマッピング（slug → サブドメイン）
    SUBDOMAIN_MAP = {
        # 教育系 → juken.oricon.co.jp
        "online-english": "juken",
        "kids-english": "juken",
        "tutor": "tutor",  # juken
        "online-study": "juken",
        "kids-swimming": "juken",
        "english-school": "juken",
        # キャリア系 → career.oricon.co.jp
        "_agent": "career",
        "_staffing": "career",
        "_staffing_manufacture": "career",
        # その他は life.oricon.co.jp（デフォルト）
    }

    # 評価項目の日英対応表
    EVALUATION_ITEMS = {
        "procedure": "加入手続き",
        "campaign": "キャンペーン",
        "initial": "初期設定のしやすさ",
        "connection": "通信速度",
        "plan": "料金プラン",
        "lineup": "端末ラインナップ",
        "cost-performance": "利用料金",
        "support": "サポートサービス",
        "option": "付帯サービス",
        # 他のランキング用
        "premium": "保険料",
        "coverage": "補償内容",
        "claim": "保険金・給付金",
        "service": "サービス",
        "app": "アプリ",
        "website": "サイト",
        # 教育系
        "curriculum": "カリキュラム",
        "teacher": "講師",
        "price": "料金",
        "facility": "施設",
        # 住宅系
        "design": "デザイン",
        "quality": "品質",
        "after-service": "アフターサービス",
    }

    def __init__(self, ranking_slug: str, ranking_name: str):
        """
        Args:
            ranking_slug: URL用のランキング名（例: mobile-carrier, _fx, card-loan/nonbank）
                          @type02 などを付与すると、そのセクションのみを抽出
            ranking_name: 表示用のランキング名（例: 携帯キャリア）
        """
        self.ranking_name = ranking_name

        # 調査タイプを分離（例: _fx@type02 → _fx, type02）
        if "@" in ranking_slug:
            ranking_slug, self.survey_type = ranking_slug.split("@", 1)
        else:
            self.survey_type = "type01"  # デフォルトは顧客満足度調査

        self.ranking_slug = ranking_slug

        # サブパスを分離（例: card-loan/nonbank → card-loan, nonbank）
        if "/" in ranking_slug:
            parts = ranking_slug.split("/", 1)
            base_slug = parts[0]
            self.subpath = parts[1]
        else:
            base_slug = ranking_slug
            self.subpath = ""

        # サブドメインを決定（教育系、キャリア系など）
        subdomain = "life"  # デフォルト
        for slug_pattern, domain in self.SUBDOMAIN_MAP.items():
            if base_slug == slug_pattern or base_slug.startswith(slug_pattern):
                subdomain = domain
                break

        # tutorの特殊処理（マッピングに "tutor": "tutor" と誤記があったので修正）
        if base_slug in ["online-english", "kids-english", "tutor", "online-study", "kids-swimming", "english-school"]:
            subdomain = "juken"
        elif base_slug in ["_agent", "_staffing", "_staffing_manufacture"]:
            subdomain = "career"

        self.BASE_URL = f"https://{subdomain}.oricon.co.jp"

        # アンダースコア形式の処理（_fxの場合はrank_fxになる）
        if base_slug.startswith("_"):
            self.url_prefix = f"rank{base_slug}"  # rank_fx
        else:
            self.url_prefix = f"rank-{base_slug}"  # rank-mobile-carrier形式
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })
        # 使用したURLを記録
        self.used_urls = {
            "overall": [],
            "items": [],
            "departments": []
        }

    def get_overall_rankings(self, year_range: tuple = (2020, 2024)) -> Dict[int, List[Dict]]:
        """
        総合ランキングを取得

        Args:
            year_range: (開始年, 終了年) のタプル

        Returns:
            {年度: [{"rank": 1, "company": "...", "score": 69.5}, ...]}
        """
        results = {}
        start_year, end_year = year_range

        for year in range(end_year, start_year - 1, -1):  # 新しい年から古い年へ
            # サブパスがある場合の処理
            subpath_part = f"/{self.subpath}" if self.subpath else ""

            if year == end_year:
                # 最新年度（URLに年度なし）
                url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"
            else:
                # 過去年度
                url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/"

            data = self._fetch_ranking_page(url, self.survey_type)
            if data:
                results[year] = data
                self.used_urls["overall"].append({"year": year, "url": url, "survey_type": self.survey_type, "status": "success"})
                pass  # データ取得成功
            else:
                self.used_urls["overall"].append({"year": year, "url": url, "survey_type": self.survey_type, "status": "not_found"})
                pass  # データなし

            time.sleep(0.3)  # サーバー負荷軽減

        return results

    def get_evaluation_items(self, year_range: tuple = None) -> Dict[str, Dict[int, List[Dict]]]:
        """
        評価項目別ランキングを取得（経年対応）

        Args:
            year_range: (開始年, 終了年) のタプル。Noneの場合は最新年度のみ

        Returns:
            {"加入手続き": {2024: [データ], 2023: [データ], ...}, ...}
        """
        results = {}

        # まず総合ページから評価項目リストを取得
        subpath_part = f"/{self.subpath}" if self.subpath else ""
        main_url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"
        items = self._discover_evaluation_items(main_url)

        if not items:
            return results

        # 年度範囲を決定
        if year_range:
            start_year, end_year = year_range
            years = list(range(end_year, start_year - 1, -1))
        else:
            years = [2024]  # 最新年度のみ

        for item_slug, item_name in items.items():
            results[item_name] = {}

            for year in years:
                if year == max(years):
                    # 最新年度
                    url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/evaluation-item/{item_slug}.html"
                else:
                    # 過去年度
                    url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/evaluation-item/{item_slug}.html"

                data = self._fetch_ranking_page(url, self.survey_type)
                if data:
                    results[item_name][year] = data
                    self.used_urls["items"].append({
                        "name": f"{item_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "success"
                    })
                else:
                    self.used_urls["items"].append({
                        "name": f"{item_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "not_found"
                    })

                time.sleep(0.3)

        return results

    def get_departments(self, year_range: tuple = None) -> Dict[str, Dict[int, List[Dict]]]:
        """
        部門別ランキングを取得（経年対応）

        Args:
            year_range: (開始年, 終了年) のタプル

        Returns:
            {"50代": {2024: [データ], 2023: [データ], ...}, ...}
        """
        results = {}

        # 総合ページから部門リストを取得
        subpath_part = f"/{self.subpath}" if self.subpath else ""
        main_url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"
        departments = self._discover_departments(main_url)

        if not departments:
            return results

        # 年度範囲を決定
        if year_range:
            start_year, end_year = year_range
            years = list(range(end_year, start_year - 1, -1))
        else:
            years = [2024]

        for dept_path, dept_name in departments.items():
            results[dept_name] = {}

            for year in years:
                if year == max(years):
                    # 最新年度
                    url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/{dept_path}"
                else:
                    # 過去年度
                    # パスの形式によって挿入位置を調整
                    # 例: age/50s.html → 2023/age/50s.html
                    url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/{dept_path}"

                data = self._fetch_ranking_page(url, self.survey_type)
                if data:
                    results[dept_name][year] = data
                    self.used_urls["departments"].append({
                        "name": f"{dept_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "success"
                    })
                else:
                    self.used_urls["departments"].append({
                        "name": f"{dept_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "not_found"
                    })

                time.sleep(0.3)

        return results

    def _discover_departments(self, url: str) -> Dict[str, str]:
        """
        ページから部門別リンクを動的に発見

        Returns:
            {"age/50s.html": "50代", "genre/sports.html": "スポーツ", ...}
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            departments = {}

            # 部門別リンクのパターン（評価項目以外）
            # 例: /rank-xxx/age/50s.html, /rank-xxx/genre/sports.html
            dept_patterns = [
                r"/(age|genre|contract|new-contract|device|business|beginner|type|purpose)/",
                r"/[a-z\-]+\.html$"  # evaluation-item以外の.htmlリンク
            ]

            all_links = soup.find_all("a", href=True)

            for link in all_links:
                href = link.get("href", "")

                # 評価項目は除外
                if "evaluation-item" in href:
                    continue

                # 自身のランキングのリンクか確認
                if self.url_prefix not in href:
                    continue

                # 部門別パターンにマッチするか
                for pattern in dept_patterns:
                    if re.search(pattern, href):
                        # パスを抽出
                        match = re.search(rf"/{self.url_prefix}/(?:\d{{4}}/)?(.+)$", href)
                        if match:
                            dept_path = match.group(1)
                            # 評価項目でない、数字のみのパス（年度）でないことを確認
                            if dept_path and not dept_path.isdigit() and "evaluation-item" not in dept_path:
                                # リンクテキストを取得
                                name = link.get_text(strip=True)
                                if name and len(name) < 30:  # 長すぎるテキストは除外
                                    departments[dept_path] = name
                        break

            return departments

        except Exception as e:
            pass  # 部門リスト取得エラー
            return {}

    def _discover_evaluation_items(self, url: str) -> Dict[str, str]:
        """
        ページから評価項目リストを動的に発見

        Returns:
            {"procedure": "加入手続き", ...}
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            items = {}

            # サイドバーやナビゲーションから評価項目リンクを探す
            eval_links = soup.find_all("a", href=re.compile(r"/evaluation-item/"))

            for link in eval_links:
                href = link.get("href", "")
                match = re.search(r"/evaluation-item/([^/]+)\.html", href)
                if match:
                    slug = match.group(1)
                    # 日本語名を取得
                    name = link.get_text(strip=True)
                    if not name:
                        name = self.EVALUATION_ITEMS.get(slug, slug)
                    items[slug] = name

            # 見つからなければデフォルトを使用
            if not items:
                # 携帯キャリアのデフォルト項目
                default_items = [
                    "procedure", "campaign", "initial", "connection",
                    "plan", "lineup", "cost-performance", "support", "option"
                ]
                items = {slug: self.EVALUATION_ITEMS.get(slug, slug) for slug in default_items}

            return items

        except Exception as e:
            pass  # 評価項目リスト取得エラー
            return {}

    def _fetch_ranking_page(self, url: str, survey_type: str = "type01") -> List[Dict]:
        """
        ランキングページからデータを抽出

        Args:
            url: 取得するURL
            survey_type: 調査タイプ（"type01"=顧客満足度調査, "type02"=FP調査など）

        Returns:
            [{"rank": 1, "company": "...", "score": 69.5}, ...]
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            rankings = []
            seen_companies = set()  # 重複チェック用

            # まず特定の調査タイプのセクションを探す（type01-main, type01-topなど）
            target_section = None

            # 優先順位: main > top > side-top
            for suffix in ["-main", "-top", "-side-top", ""]:
                section_id = f"{survey_type}{suffix}"
                target_section = soup.find(id=section_id)
                if target_section:
                    break

            # セクションが見つからない場合は、ページ全体から探すが、
            # type02関連のセクションは除外する
            if not target_section:
                # type02セクションを除外したsoupを作成
                for exclude_type in ["type02-main", "type02-top", "type02-side-top", "type02-side-btm"]:
                    exclude_section = soup.find(id=exclude_type)
                    if exclude_section:
                        exclude_section.decompose()  # DOMから削除
                target_section = soup

            # ランキングボックスを探す（複数のパターンに対応）
            ranking_boxes = target_section.find_all("article", class_=re.compile(r"ranking"))

            if not ranking_boxes:
                # 別のパターンを試す
                ranking_boxes = target_section.find_all("div", class_=re.compile(r"ranking-box"))

            if not ranking_boxes:
                # さらに別のパターン
                ranking_boxes = target_section.find_all("li", class_=re.compile(r"rank"))

            for box in ranking_boxes:
                try:
                    data = self._extract_ranking_data(box)
                    if data:
                        company = data.get("company", "")
                        # 重複企業をスキップ
                        if company and company not in seen_companies:
                            rankings.append(data)
                            seen_companies.add(company)
                except Exception as e:
                    continue

            # 順位でソート、同順位の場合は得点で降順ソート
            rankings.sort(key=lambda x: (x.get("rank", 999), -x.get("score", 0)))

            return rankings

        except Exception as e:
            pass  # ページ取得エラー
            return []

    def _extract_ranking_data(self, element) -> Optional[Dict]:
        """
        HTML要素からランキングデータを抽出

        Returns:
            {"rank": 1, "company": "楽天モバイル", "score": 69.5}
        """
        data = {}

        # 順位を抽出
        rank = None

        # パターン1: imgタグのalt属性
        img = element.find("img", alt=re.compile(r"\d+位"))
        if img:
            match = re.search(r"(\d+)位", img.get("alt", ""))
            if match:
                rank = int(match.group(1))

        # パターン2: クラス名から
        if not rank:
            class_str = " ".join(element.get("class", []))
            match = re.search(r"rank-?(\d+)", class_str)
            if match:
                rank = int(match.group(1))

        # パターン3: テキストから
        if not rank:
            rank_text = element.find(class_=re.compile(r"rank|icon"))
            if rank_text:
                match = re.search(r"(\d+)", rank_text.get_text())
                if match:
                    rank = int(match.group(1))

        if not rank:
            return None

        data["rank"] = rank

        # 企業名を抽出
        company = None

        # パターン1: h3 itemprop="name"
        h3 = element.find("h3", itemprop="name")
        if h3:
            company = h3.get_text(strip=True)

        # パターン2: 単純なh3
        if not company:
            h3 = element.find("h3")
            if h3:
                company = h3.get_text(strip=True)

        # パターン3: company-nameクラス
        if not company:
            name_elem = element.find(class_=re.compile(r"company|name"))
            if name_elem:
                company = name_elem.get_text(strip=True)

        if company:
            # 余分な文字を除去
            company = re.sub(r"\s+", " ", company).strip()
            data["company"] = company

        # 得点を抽出
        score = None

        # パターン1: score-pointクラス
        score_elem = element.find(class_=re.compile(r"score"))
        if score_elem:
            score_text = score_elem.get_text()
            match = re.search(r"(\d+\.?\d*)", score_text)
            if match:
                score = float(match.group(1))

        # パターン2: strongタグ内の数値
        if not score:
            strong = element.find("strong")
            if strong:
                match = re.search(r"(\d+\.?\d*)", strong.get_text())
                if match:
                    score = float(match.group(1))

        if score:
            data["score"] = score

        return data if "company" in data else None

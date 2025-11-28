# -*- coding: utf-8 -*-
"""
オリコン顧客満足度サイトのスクレイパー
v4.3 - 動的年度検出機能追加
- トップページから実際の発表年度を自動検出
- 未発表年度のスキップ機能
- リトライ処理（v4.2より継続）
"""

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import re
from typing import Dict, List, Optional
import time
import logging

logger = logging.getLogger(__name__)

class OriconScraper:
    """オリコン顧客満足度サイトからランキングデータを取得

    対応サブドメイン:
    - life.oricon.co.jp: 生活系（金融、保険、通信、住宅、生活サービスなど）
    - juken.oricon.co.jp: 教育系（英会話、家庭教師、通信教育、スイミングなど）
    - career.oricon.co.jp: キャリア系（転職エージェント、派遣会社など）
    """

    # サブドメインのマッピング（slug → サブドメイン）
    SUBDOMAIN_MAP = {
        # ========================================
        # 教育系 → juken.oricon.co.jp
        # ========================================
        # 英語
        "online-english": "juken",
        "kids-english": "juken",
        "_english": "juken",
        # 塾・受験
        "_college": "juken",
        "college-individual": "juken",
        "college-video": "juken",
        "highschool": "juken",
        "highschool-individual": "juken",
        "_junior": "juken",
        "public-junior": "juken",
        "supplementary-school": "juken",
        "kids-school": "juken",
        # 通信教育・資格
        "online-study": "juken",
        "tutor": "juken",
        "cc": "juken",
        "license": "juken",
        # スポーツ
        "kids-swimming": "juken",
        # ========================================
        # キャリア系 → career.oricon.co.jp
        # ========================================
        # 就活
        "new-graduates-hiring-website": "career",
        "reversed-job-offer": "career",
        # アルバイト
        "arbeit": "career",
        # 転職
        "job-change": "career",
        "job-change_woman": "career",
        "job-change_scout": "career",
        # 転職エージェント
        "_agent": "career",
        "_agent_nurse": "career",
        "_agent_nursing": "career",
        "_agent_hi-and-middle-class": "career",
        # 派遣
        "_staffing": "career",
        "_staffing_manufacture": "career",
        "temp-staff": "career",
        "employment": "career",
        # ========================================
        # その他は life.oricon.co.jp（デフォルト）
        # ========================================
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

        # サブドメインの特殊処理（SUBDOMAIN_MAPで漏れた場合のフォールバック）
        if base_slug in ["online-english", "kids-english", "tutor", "online-study", "kids-swimming", "_english"]:
            subdomain = "juken"
        elif base_slug in ["_agent", "_staffing", "_staffing_manufacture", "job-change"]:
            subdomain = "career"

        self.BASE_URL = f"https://{subdomain}.oricon.co.jp"

        # アンダースコア形式の処理（_fxの場合はrank_fxになる）
        if base_slug.startswith("_"):
            self.url_prefix = f"rank{base_slug}"  # rank_fx
        else:
            self.url_prefix = f"rank-{base_slug}"  # rank-mobile-carrier形式
        # セッション設定（リトライ処理追加）
        self.session = requests.Session()

        # リトライ戦略: 500, 502, 503, 504エラー時に最大3回リトライ
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,  # 1, 2, 4秒と増加
            status_forcelist=[500, 502, 503, 504],
            allowed_methods=["GET"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })
        # 使用したURLを記録
        self.used_urls = {
            "overall": [],
            "items": [],
            "departments": []
        }
        # トップページの実際の年度をキャッシュ
        self._actual_top_year = None

    def _ensure_actual_top_year(self) -> int:
        """
        _actual_top_yearが未設定の場合、トップページから年度を検出して設定する。
        検出できない場合は現在年を返す。
        """
        if self._actual_top_year is None:
            subpath_part = f"/{self.subpath}" if self.subpath else ""
            top_url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"
            detected_year = self._detect_actual_year(top_url)
            if detected_year:
                self._actual_top_year = detected_year
            else:
                # フォールバック: 現在年を使用
                from datetime import datetime
                self._actual_top_year = datetime.now().year
                logger.warning(f"年度検出失敗、現在年を使用: {self._actual_top_year}")
        return self._actual_top_year

    def _detect_actual_year(self, url: str) -> Optional[int]:
        """
        トップページから実際の発表年度を検出

        更新日から年度を推定する。
        優先順位:
        1. 最終更新日（最も信頼性が高い）
        2. タイトルの年度表記
        3. ページ冒頭の年度表記
        4. 過去リンクからの推定（フォールバック）

        ※調査期間は使用しない（更新日が正式な年度基準のため）

        Returns:
            検出された年度（例: 2024）、検出できない場合はNone
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            text = soup.get_text()

            # パターン1: 最終更新日から検出（最も信頼性が高い）
            # 例: 「最終更新日：2025-11-01」「更新日: 2025/11/01」
            update_match = re.search(r'(?:最終)?更新日[：:\s]*(\d{4})[-/]\d{1,2}[-/]\d{1,2}', text)
            if update_match:
                year = int(update_match.group(1))
                logger.info(f"最終更新日から年度検出: {year}年")
                return year

            # パターン2: タイトルから検出（ページ上部に表示されることが多い）
            # 例: 「2025年 オリコン顧客満足度」「2025年オリコン」
            title_match = re.search(r'(\d{4})年\s*オリコン', text)
            if title_match:
                year = int(title_match.group(1))
                logger.info(f"タイトルから年度検出: {year}年")
                return year

            # パターン3: ページ冒頭の年度表記
            # 例: 「2025年 ネット証券」のような表記
            # 最初の数行内で「20XX年」を探す
            lines = text.split('\n')[:30]  # 最初の30行を対象
            for line in lines:
                year_match = re.search(r'^.*?(\d{4})年', line.strip())
                if year_match:
                    year = int(year_match.group(1))
                    if 2000 <= year <= 2030:  # 妥当な年度範囲
                        logger.info(f"ページ冒頭から年度検出: {year}年")
                        return year

            # パターン4: 過去ランキングリンクから推定（フォールバック、信頼性低）
            # 最新の過去年度リンク + 1 = 現在の年度
            # ※ 年度が飛んでいる場合は不正確になるため、最終手段として使用
            past_links = soup.find_all('a', href=re.compile(r'/\d{4}/?$'))
            if past_links:
                years = []
                for link in past_links:
                    href = link.get('href', '')
                    year_match = re.search(r'/(\d{4})/?$', href)
                    if year_match:
                        years.append(int(year_match.group(1)))
                if years:
                    max_past_year = max(years)
                    # 過去年度の最大値 + 1 が現在の年度
                    inferred_year = max_past_year + 1
                    logger.info(f"過去リンクから年度推定（フォールバック）: {inferred_year}年（過去最大: {max_past_year}年）")
                    return inferred_year

            logger.warning(f"年度を検出できませんでした: {url}")
            return None

        except Exception as e:
            logger.error(f"年度検出エラー: {e}")
            return None

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

        # サブパスがある場合の処理
        subpath_part = f"/{self.subpath}" if self.subpath else ""

        # トップページのURLを構築
        top_url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"

        # トップページから実際の発表年度を検出（キャッシュ利用）
        if self._actual_top_year is None:
            self._actual_top_year = self._detect_actual_year(top_url)
            if self._actual_top_year:
                logger.info(f"トップページの実際の年度: {self._actual_top_year}年")
            else:
                # 検出できない場合はend_yearを使用
                self._actual_top_year = end_year
                logger.warning(f"年度検出できず、end_year({end_year})を使用")

        actual_top_year = self._actual_top_year

        for year in range(end_year, start_year - 1, -1):  # 新しい年から古い年へ

            if year == actual_top_year:
                # トップページの年度と一致 → 年度なしURL
                url = top_url
                logger.info(f"{year}年: トップページURL使用 {url}")
            elif year > actual_top_year:
                # まだ発表されていない年度 → スキップ
                logger.info(f"{year}年: 未発表のためスキップ（現在の最新: {actual_top_year}年）")
                self.used_urls["overall"].append({
                    "year": year,
                    "url": "-",
                    "survey_type": self.survey_type,
                    "status": "not_published"
                })
                continue
            else:
                # 過去年度 - サブパスがある場合は /subpath/year/ 形式を優先
                # 例: rank_fitness/24hours/2024/ （正しい形式）
                url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/{year}/"

            data = self._fetch_ranking_page(url, self.survey_type)
            if data:
                results[year] = data
                self.used_urls["overall"].append({"year": year, "url": url, "survey_type": self.survey_type, "status": "success"})
            else:
                # 代替パターン1: /year/subpath/ 形式を試す
                # 例: rank_fitness/2024/24hours/
                if self.subpath:
                    alt_url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/"
                    data = self._fetch_ranking_page(alt_url, self.survey_type)
                    if data:
                        results[year] = data
                        self.used_urls["overall"].append({"year": year, "url": alt_url, "survey_type": self.survey_type, "status": "success"})
                        continue

                # 代替パターン2: 特殊年度パターン（例: 2014-2015）
                special_url = f"{self.BASE_URL}/{self.url_prefix}/{year}-{year+1}{subpath_part}/"
                data = self._fetch_ranking_page(special_url, self.survey_type)
                if data:
                    results[year] = data
                    self.used_urls["overall"].append({"year": f"{year}-{year+1}", "url": special_url, "survey_type": self.survey_type, "status": "success"})
                else:
                    self.used_urls["overall"].append({"year": year, "url": url, "survey_type": self.survey_type, "status": "not_found"})

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
        # _ensure_actual_top_year() で年度を確実に初期化
        actual_top_year = self._ensure_actual_top_year()

        if year_range:
            start_year, end_year = year_range
            years = list(range(end_year, start_year - 1, -1))
        else:
            years = [actual_top_year]  # 検出済みの年度を使用

        for item_slug, item_name in items.items():
            results[item_name] = {}

            for year in years:
                # 未発表年度はスキップ
                if year > actual_top_year:
                    self.used_urls["items"].append({
                        "name": f"{item_name}({year}年)",
                        "url": "-",
                        "survey_type": self.survey_type,
                        "status": "not_published"
                    })
                    continue

                if year == actual_top_year:
                    # トップページの年度と一致 → 年度なしURL
                    url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/evaluation-item/{item_slug}.html"
                else:
                    # 過去年度 - /subpath/year/ 形式を優先
                    url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/{year}/evaluation-item/{item_slug}.html"

                data = self._fetch_ranking_page(url, self.survey_type)
                if data:
                    # ページタイトルから実際の名称を取得
                    page_title = self._extract_page_title(url)
                    results[item_name][year] = data
                    self.used_urls["items"].append({
                        "name": f"{item_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "success",
                        "page_title": page_title,  # ページから取得した実際の名称
                        "item_slug": item_slug,
                        "year": year
                    })
                else:
                    # 代替パターン: /year/subpath/ 形式を試す
                    if self.subpath:
                        alt_url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/evaluation-item/{item_slug}.html"
                        data = self._fetch_ranking_page(alt_url, self.survey_type)
                        if data:
                            page_title = self._extract_page_title(alt_url)
                            results[item_name][year] = data
                            self.used_urls["items"].append({
                                "name": f"{item_name}({year}年)",
                                "url": alt_url,
                                "survey_type": self.survey_type,
                                "status": "success",
                                "page_title": page_title,
                                "item_slug": item_slug,
                                "year": year
                            })
                            continue

                    self.used_urls["items"].append({
                        "name": f"{item_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "not_found",
                        "item_slug": item_slug,
                        "year": year
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
        # _ensure_actual_top_year() で年度を確実に初期化
        actual_top_year = self._ensure_actual_top_year()

        if year_range:
            start_year, end_year = year_range
            years = list(range(end_year, start_year - 1, -1))
        else:
            years = [actual_top_year]  # 検出済みの年度を使用

        for dept_path, dept_name in departments.items():
            results[dept_name] = {}

            for year in years:
                # 未発表年度はスキップ
                if year > actual_top_year:
                    self.used_urls["departments"].append({
                        "name": f"{dept_name}({year}年)",
                        "url": "-",
                        "survey_type": self.survey_type,
                        "status": "not_published"
                    })
                    continue

                if year == actual_top_year:
                    # トップページの年度と一致 → 年度なしURL
                    url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/{dept_path}"
                else:
                    # 過去年度 - /subpath/year/ 形式を優先
                    url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/{year}/{dept_path}"

                data = self._fetch_ranking_page(url, self.survey_type)
                if data:
                    # ページタイトルから実際の名称を取得
                    page_title = self._extract_page_title(url)
                    results[dept_name][year] = data
                    self.used_urls["departments"].append({
                        "name": f"{dept_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "success",
                        "page_title": page_title,  # ページから取得した実際の名称
                        "dept_path": dept_path,
                        "year": year
                    })
                else:
                    # 代替パターン: /year/subpath/ 形式を試す
                    if self.subpath:
                        alt_url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/{dept_path}"
                        data = self._fetch_ranking_page(alt_url, self.survey_type)
                        if data:
                            page_title = self._extract_page_title(alt_url)
                            results[dept_name][year] = data
                            self.used_urls["departments"].append({
                                "name": f"{dept_name}({year}年)",
                                "url": alt_url,
                                "survey_type": self.survey_type,
                                "status": "success",
                                "page_title": page_title,
                                "dept_path": dept_path,
                                "year": year
                            })
                            continue

                    self.used_urls["departments"].append({
                        "name": f"{dept_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "not_found",
                        "dept_path": dept_path,
                        "year": year
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

    def _extract_page_title(self, url: str) -> Optional[str]:
        """
        ページから評価項目名・部門名を抽出

        ページ内のh1, h2, title等から項目名を取得する。
        例: "【2025年】ネット証券の取扱商品 オリコン顧客満足度ランキング" → "取扱商品"
        例: "【2012年】ネット証券の取扱商品量ランキング・比較" → "取扱商品量"

        Returns:
            抽出された項目名、または None
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            # パターン1: h1タグから取得
            h1 = soup.find("h1")
            if h1:
                text = h1.get_text(strip=True)
                extracted = self._extract_item_name_from_title(text)
                if extracted:
                    return extracted

            # パターン2: og:title メタタグから取得
            og_title = soup.find("meta", property="og:title")
            if og_title:
                text = og_title.get("content", "")
                extracted = self._extract_item_name_from_title(text)
                if extracted:
                    return extracted

            # パターン3: titleタグから取得
            title = soup.find("title")
            if title:
                text = title.get_text(strip=True)
                extracted = self._extract_item_name_from_title(text)
                if extracted:
                    return extracted

            return None

        except Exception as e:
            return None

    def _extract_item_name_from_title(self, text: str) -> Optional[str]:
        """
        タイトル文字列から評価項目名・部門名を抽出

        対応パターン:
        - 【2025年】ネット証券の取扱商品 オリコン顧客満足度ランキング → 取扱商品
        - 【2012年】ネット証券の取扱商品量ランキング・比較 → 取扱商品量
        - 2012年 取扱商品量｜ネット証券ランキング → 取扱商品量
        - 【2025年】ネット証券 初心者のランキング → 初心者

        非対応（Noneを返す）:
        - 【最新】ネット証券のランキング・比較 → None（評価項目ページではない）
        """
        if not text:
            return None

        # パターン1: 【年度】XXXのYYY ランキング → YYY を抽出
        # 例: 【2025年】ネット証券の取扱商品 オリコン → 取扱商品
        # 「の」の後に具体的な項目名があり、その後にオリコンorランキングが続く
        match = re.search(r"の(.+?)(?:\s+オリコン|\s+ランキング|ランキング)", text)
        if match:
            item_name = match.group(1).strip()
            # 「満足度」で終わる場合は除去
            item_name = re.sub(r"\s*満足度$", "", item_name)
            # 年度だけの場合はスキップ
            if not re.match(r"^\d{4}年?$", item_name) and item_name:
                # 「ランキング・比較」などの一般的な語句は除外
                if item_name not in ["ランキング", "比較", "ランキング・比較"]:
                    return item_name

        # パターン2: YYYY年 XXX｜ → XXX を抽出
        # 例: 2012年 取扱商品量｜ネット証券ランキング → 取扱商品量
        match = re.search(r"\d{4}年\s+(.+?)(?:｜|\||ランキング)", text)
        if match:
            item_name = match.group(1).strip()
            if item_name and not re.match(r"^\d{4}年?$", item_name):
                # 「ランキング・比較」などの一般的な語句は除外
                if item_name not in ["ランキング", "比較", "ランキング・比較"]:
                    return item_name

        # パターン3: XXX YYYのランキング → YYY（スペース区切り）
        # 例: 【2025年】ネット証券 初心者のランキング → 初心者
        match = re.search(r"\s([^\s]+?)のランキング", text)
        if match:
            item_name = match.group(1).strip()
            if item_name and not re.match(r"^\d{4}年?$", item_name):
                return item_name

        return None

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

        Note:
            順位抽出の優先順位:
            1. icon-rank クラス（総合順位の正しい表示場所）
            2. imgタグのalt属性（ただし td.rank 内は除外）
            3. クラス名から（rank-1, rank01 など）

            評価項目別テーブル内の順位（td.rank 内の img）を
            誤って総合順位として取得しないよう注意が必要。
        """
        data = {}

        # 順位を抽出
        rank = None

        # パターン1（最優先）: icon-rank クラスから総合順位を取得
        # これが正しい総合順位の表示場所（評価項目別テーブル内の順位と混同しない）
        icon_rank = element.find(class_=re.compile(r"icon-rank"))
        if icon_rank:
            rank_text = icon_rank.get_text(strip=True)
            match = re.search(r"(\d+)", rank_text)
            if match:
                rank = int(match.group(1))

        # パターン2: imgタグのalt属性から（ただし td.rank 内は除外）
        # 評価項目別テーブル内の順位を誤取得しないため
        if not rank:
            imgs = element.find_all("img", alt=re.compile(r"\d+位"))
            for img in imgs:
                # 親要素をチェック - td.rank（評価項目別テーブル内）は除外
                parent = img.parent
                if parent and parent.name == "td" and "rank" in parent.get("class", []):
                    continue  # テーブル内の順位はスキップ
                # ranking-score セクション内も除外（評価項目別・ジャンル別テーブル）
                if img.find_parent(class_="ranking-score"):
                    continue
                match = re.search(r"(\d+)位", img.get("alt", ""))
                if match:
                    rank = int(match.group(1))
                    break

        # パターン3: クラス名から（例: rank-1, rank01）
        if not rank:
            class_str = " ".join(element.get("class", []))
            match = re.search(r"rank-?(\d+)", class_str)
            if match:
                rank = int(match.group(1))

        # パターン4: フォールバック - icon クラスを持つ要素のテキストから
        if not rank:
            rank_elem = element.find(class_=re.compile(r"^icon"))
            if rank_elem:
                match = re.search(r"(\d+)", rank_elem.get_text())
                if match:
                    rank = int(match.group(1))

        # 順位の検証（1-100の範囲内であること）
        if not rank or rank < 1 or rank > 100:
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

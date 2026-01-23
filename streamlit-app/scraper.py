# -*- coding: utf-8 -*-
"""
オリコン顧客満足度サイトのスクレイパー
v7.11 - URLバリデーション機能追加
- validate_url()でURL存在確認（404検出）
- 代替URL提案機能（サブドメイン/スラッグ修正候補）
- URL_SLUG_MAPで最新のURL構造に対応

v7.10 - リソース管理改善
- close()メソッドを追加してセッションリソースを解放
- コンテキストマネージャー（with文）対応

v7.9 - SiteStructureAnalyzer統合
- SiteStructureAnalyzerを導入し、1回のリクエストで構造を解析
- 評価項目・部門の検出を効率化（HTTPリクエスト削減）
- analyze_structure()メソッドを追加
- 旧ロジック（_discover_*）はフォールバックとして残す

v7.8 - sort-nav TABLE構造の複数リンク対応
- _extract_departments_from_sort_nav でTD内の全リンクを取得するよう修正
- 派遣会社の業務内容別（type/logistics.html等）が正しく検出されるように
- 1つのTD内に複数リンクがある場合にすべて取得

v7.7 - sort-nav TABLE構造対応
- _extract_departments_from_sort_nav を実際のサイト構造（TABLE）に対応
- 旧SECTION構造はフォールバックとして残す
- 全215ランキングの包括テストで確認済み: TABLE構造 97.2%、SECTION構造 0%

v7.6 - v7.5誤検出修正
- evaluation-item は「評価項目別」であり「部門別」ではないため、除外対象に復元
"""

import requests
from requests.adapters import HTTPAdapter
from requests.exceptions import RequestException
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import re
from typing import Dict, List, Optional
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# v7.9: SiteStructureAnalyzer統合
from site_analyzer import SiteStructureAnalyzer, SiteStructure

logger = logging.getLogger(__name__)

class OriconScraper:
    """オリコン顧客満足度サイトからランキングデータを取得

    対応サブドメイン:
    - life.oricon.co.jp: 生活系（金融、保険、通信、住宅、生活サービスなど）
    - juken.oricon.co.jp: 教育系（英会話、家庭教師、通信教育、スイミングなど）
    - career.oricon.co.jp: キャリア系（転職エージェント、派遣会社など）
    """

    # rankプレフィックスなしのパターン（URLが /xxx/ の形式）
    # 例: medical_insurance → /medical_insurance/ (NOT /rank-medical_insurance/)
    NO_RANK_PREFIX = {
        "medical_insurance",  # 医療保険
    }

    # ========================================
    # 定数定義（v7.0 リファクタリング）
    # ========================================

    # HTTPリクエスト関連の設定
    REQUEST_DELAY_SEC = 0.2  # リクエスト間の遅延（秒）
    REQUEST_TIMEOUT_SEC = 10  # タイムアウト（秒）
    MAX_DEPT_NAME_LENGTH = 30  # 部門名の最大文字数

    # 部門別リンクのパターン（評価項目以外）
    DEPT_PATTERNS = [
        r"/(age|contract|new-contract|device|business|beginner|type|purpose|nisa|ideco|style|sim|sp)(?:/|\.html)",
        r"/(plan)(?:/|\.html)",  # 携帯キャリア: プラン別（一般的なプラン、シニア向けプラン等）v7.0追加
        r"/(specialty|manufacturer)(?:/|\.html)",  # バイク販売店
        r"/(general|publisher|original)(?:/|\.html)",  # 電子コミック/マンガアプリ
        r"/(hokkaido|tohoku|kanto|kinki|tokai|chugoku-shikoku|kyushu-okinawa|koshinetsu-hokuriku|nationwide)(?:/|\.html)",  # 地域別
        r"/(east|west)(?:/|\.html)",  # テーマパーク
        r"/(investment-products|sp-sec|support)(?:/|\.html)",  # ネット証券: 投資商品別、スマホ証券、サポート
        r"/(foreign-stocks|investment-trust)(?:\.html)?",  # ネット証券: 外国株式、投資信託
        r"/(preschooler|grade-schooler)(?:/|\.html)?",  # 子ども英語教室: 幼児、小学生（トップページ用）
        r"/(grade)(?:/|\.html)",  # 子ども英語教室: 低学年、高学年（grade-schooler配下）
        r"/(genre)(?:/|\.html)",  # SVOD: ジャンル別
        # v7.3追加: 引越し会社用パターン
        r"/(family)(?:/|\.html)",  # 引越し会社: 家族構成別（単身/家族）
        r"/(prefecture)(?:/|\.html)",  # 引越し会社: 都道府県別
        r"/(area)(?:/|\.html)",  # 引越し会社: 地域別
        r"/(gender)(?:/|\.html)",  # 引越し会社: 性別
        # v7.4追加: 英会話スクール用パターン
        r"/(level)(?:/|\.html)",  # 英会話スクール: レベル別
        r"/(class)(?:/|\.html)",  # 英会話スクール: クラス別（グループ/個人など）
        # v7.4追加: 保険代理店型パターン
        r"/(agency)(?:/|\.html)",  # 自動車保険・バイク保険: 代理店型
        r"/(category)/([^/]+)\.html",  # カテゴリ別（agent等）
        # v7.5で追加した evaluation-item パターンは v7.6 で削除
        # 理由: evaluation-item は「評価項目別」であり「部門別」ではない
        # v7.12追加: 価格帯別パターン
        r"/(price)(?:/|\.html)",  # ハウスメーカー注文住宅: 価格帯別（2000万円未満、2000-3000万円等）
    ]

    # 除外パターン（部門ページではないもの）
    EXCLUDE_URL_PATTERNS = [
        r"/evaluation-item",  # 評価項目別ページ（部門ではない）- v7.6で復元
        r"/column",           # コラム・解説ページ
        r"/special/",         # 特集・解説ページ
        r"-basic",            # 基本解説ページ
        r"/howto",            # ハウツー
        r"/how_to",           # ハウツー（アンダースコア）
        r"/recommend",        # おすすめ
        r"/compare",          # 比較
        r"/company/",         # 企業詳細ページ（部門ではない）v7.0追加
        r"/education/",       # 教育コラムページ（部門ではない）v7.0追加
        # v7.4追加: 誤検出防止用除外パターン
        r"/school_list/",     # 教室リストページ（部門ではない）
        r"/town/",            # 都道府県・市区町村選択ページ
        r"[?&]pref=",         # 都道府県クエリパラメータ（位置に依存しない）
        r"[?&]area=",         # エリアクエリパラメータ（位置に依存しない）
        r"/search",           # 検索ページ
        r"/ranking-list",     # ランキング一覧ページ
    ]

    # v7.4追加: 無効な部門名（誤検出防止用）
    # v7.9変更: 都道府県名を削除（派遣会社・引越し会社等では有効な部門のため）
    # 都道府県別部門はprefectureパターンで検出される
    INVALID_DEPT_NAMES = [
        # 一般的な無効名
        "ランキング", "一覧", "比較", "おすすめ", "検索結果", "教室一覧",
        # 年度（例: 2024年, 2023年）
        "2020年", "2021年", "2022年", "2023年", "2024年", "2025年", "2026年", "2027年",
    ]

    # 除外するh3見出し（sort-nav内で部門ではないセクション）
    # v7.6変更: 「評価項目別」「評価項目」を除外対象に復元
    # 理由: evaluation-item は「評価項目別」であり「部門別」ではない
    EXCLUDE_HEADINGS = [
        "評価項目別", "評価項目",  # v7.6で復元
        "過去のランキング", "過去ランキング",
        "関連ランキング", "関連する"
    ]

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

    # ========================================
    # v7.11: URL変換マップ（旧スラッグ→新スラッグ）
    # オリコンサイトでURL構造が変更されたものを修正
    # ========================================
    URL_SLUG_MAP = {
        # life.oricon.co.jp の変更
        "moving-company": {"slug": "_move", "domain": "life"},  # 引越し会社
        "travel-reservation-site": {"slug": "travel-website", "domain": "life"},  # 旅行予約サイト
        "fitness-gym": {"slug": "_fitness", "domain": "life"},  # フィットネスジム
        "_hikari": {"slug": "_internet", "domain": "life"},  # 光回線
        "home-wifi": {"slug": "_internet", "domain": "life"},  # ホームWi-Fi→光回線
        # life → career への移動
        "_site": {"slug": "job-change", "domain": "career"},  # 転職サイト
        "_haken": {"slug": "_staffing", "domain": "career"},  # 派遣会社
        # life → juken への移動
        "english-school": {"slug": "_english", "domain": "juken"},  # 英会話スクール
        "programming-school": {"slug": "kids-programming", "domain": "juken"},  # プログラミングスクール
        # juken.oricon.co.jp の変更
        "kids-english-school": {"slug": "kids-english", "domain": "juken"},  # 子ども英語教室
        "swimming-school": {"slug": "kids-swimming", "domain": "juken"},  # スイミングスクール
    }

    # サブパスの変換マップ（旧サブパス→新サブパス）
    SUBPATH_MAP = {
        "infant": "preschooler",  # 幼児 → 未就学児
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

        # URL prefix の決定
        # パターン0: NO_RANK_PREFIX → そのまま（rankなし）
        # パターン1: _fx → rank_fx
        # パターン2: rank_certificate → rank_certificate（そのまま）
        # パターン3: mobile-carrier → rank-mobile-carrier
        if base_slug in self.NO_RANK_PREFIX:
            self.url_prefix = base_slug  # medical_insurance（rankなし）
        elif base_slug.startswith("_"):
            self.url_prefix = f"rank{base_slug}"  # rank_fx
        elif base_slug.startswith("rank_"):
            self.url_prefix = base_slug  # rank_certificate（そのまま）
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
        # スレッドセーフ用ロック（v8.1追加）
        self._url_lock = threading.Lock()
        # トップページの実際の年度をキャッシュ
        self._actual_top_year = None
        # トップページの更新日をキャッシュ (year, month)
        self._update_date = None
        # v7.9: SiteStructureAnalyzerのキャッシュ
        self._site_structure: Optional[SiteStructure] = None
        self._structure_analyzer = SiteStructureAnalyzer()

    def close(self):
        """セッションを閉じてリソースを解放（v7.10追加）"""
        if self.session:
            self.session.close()
            logger.debug("Scraperセッションを閉じました")

    def __enter__(self):
        """コンテキストマネージャー対応（v7.10追加）"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """コンテキストマネージャー終了時にセッションを閉じる（v7.10追加）"""
        self.close()
        return False

    # ========================================
    # v7.11: URLバリデーション機能
    # ========================================

    def validate_url(self, url: str = None) -> dict:
        """
        URLの有効性を確認（v7.11追加）

        Args:
            url: 検証するURL。Noneの場合はトップページURLを使用

        Returns:
            {
                "is_valid": bool,           # URLが有効か（200-299ステータス）
                "status_code": int|None,    # HTTPステータスコード
                "url_tested": str,          # テスト対象URL
                "error_type": str|None,     # エラータイプ（"404", "timeout"等）
                "suggestions": List[dict]   # 代替URL提案
            }
        """
        if url is None:
            subpath_part = f"/{self.subpath}" if self.subpath else ""
            url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"

        result = {
            "is_valid": False,
            "status_code": None,
            "url_tested": url,
            "error_type": None,
            "suggestions": []
        }

        try:
            # HEADリクエストを試行（軽量）
            response = self.session.head(url, timeout=5, allow_redirects=True)

            # 405 Method Not Allowed の場合はGETに切り替え
            if response.status_code == 405:
                response = self.session.get(url, timeout=5)

            result["status_code"] = response.status_code
            result["is_valid"] = 200 <= response.status_code < 300

            if not result["is_valid"]:
                result["error_type"] = self._classify_http_error(response.status_code)
                # 404の場合は代替URLを提案
                if response.status_code == 404:
                    result["suggestions"] = self._suggest_alternative_urls(url)

        except requests.exceptions.Timeout:
            result["error_type"] = "timeout"
            result["suggestions"] = [{"url": url, "reason": "タイムアウト - 再試行してください"}]
        except requests.exceptions.ConnectionError as e:
            if "name resolution" in str(e).lower() or "dns" in str(e).lower():
                result["error_type"] = "dns_error"
            else:
                result["error_type"] = "connection_error"
            result["suggestions"] = self._suggest_alternative_urls(url)
        except Exception as e:
            result["error_type"] = "unknown_error"
            logger.error(f"URLバリデーションエラー: {e}")

        return result

    def _classify_http_error(self, status_code: int) -> str:
        """HTTPステータスコードからエラータイプを分類"""
        if status_code == 404:
            return "404_not_found"
        elif status_code == 403:
            return "403_forbidden"
        elif status_code >= 500:
            return "server_error"
        else:
            return f"http_{status_code}"

    def _suggest_alternative_urls(self, failed_url: str) -> list:
        """
        失敗したURLから代替URLを提案（v7.11追加）

        提案ロジック:
        1. URL_SLUG_MAPに一致するスラッグがあれば変換
        2. サブドメインを変更（life↔juken↔career）
        3. スラッグの命名規則を変換（rank-xxx ↔ rank_xxx）
        """
        suggestions = []

        # 現在のURLからスラッグとサブドメインを抽出
        import urllib.parse
        parsed = urllib.parse.urlparse(failed_url)
        current_domain = parsed.netloc.split('.')[0]  # life, juken, career

        # パスからスラッグを抽出（例: /rank-xxx/subpath/ → xxx）
        path_parts = parsed.path.strip('/').split('/')
        if not path_parts:
            return suggestions

        slug_part = path_parts[0]
        # rank- または rank_ プレフィックスを除去
        if slug_part.startswith('rank-'):
            base_slug = slug_part[5:]
        elif slug_part.startswith('rank_'):
            base_slug = '_' + slug_part[5:]
        else:
            base_slug = slug_part

        # 1. URL_SLUG_MAPでの変換チェック
        if base_slug in self.URL_SLUG_MAP:
            mapping = self.URL_SLUG_MAP[base_slug]
            new_slug = mapping["slug"]
            new_domain = mapping["domain"]

            # 新しいURLプレフィックスを構築
            if new_slug.startswith('_'):
                new_prefix = f"rank{new_slug}"
            else:
                new_prefix = f"rank-{new_slug}"

            # サブパスを取得（あれば）
            subpath_parts = path_parts[1:] if len(path_parts) > 1 else []
            # サブパスの変換（例: infant → preschooler）
            converted_subpath = []
            for sp in subpath_parts:
                if sp in self.SUBPATH_MAP:
                    converted_subpath.append(self.SUBPATH_MAP[sp])
                else:
                    converted_subpath.append(sp)
            subpath_str = '/' + '/'.join(converted_subpath) if converted_subpath else ''

            new_url = f"https://{new_domain}.oricon.co.jp/{new_prefix}{subpath_str}/"
            suggestions.append({
                "url": new_url,
                "reason": f"URL構造変更: {base_slug} → {new_slug} ({new_domain}ドメイン)",
                "confidence": "high"
            })

        # 2. サブドメイン変更の提案
        other_domains = [d for d in ["life", "juken", "career"] if d != current_domain]
        for alt_domain in other_domains:
            alt_url = failed_url.replace(
                f"https://{current_domain}.oricon.co.jp",
                f"https://{alt_domain}.oricon.co.jp"
            )
            suggestions.append({
                "url": alt_url,
                "reason": f"サブドメイン変更: {current_domain} → {alt_domain}",
                "confidence": "medium"
            })

        # 3. スラッグ命名規則の変換（rank- ↔ rank_）
        if 'rank-' in failed_url:
            alt_url = failed_url.replace('rank-', 'rank_')
            suggestions.append({
                "url": alt_url,
                "reason": "URLパターン変更: rank- → rank_",
                "confidence": "medium"
            })
        elif 'rank_' in failed_url:
            alt_url = failed_url.replace('rank_', 'rank-')
            suggestions.append({
                "url": alt_url,
                "reason": "URLパターン変更: rank_ → rank-",
                "confidence": "medium"
            })

        # 重複を除去し、最大5件に制限
        seen_urls = set()
        unique_suggestions = []
        for s in suggestions:
            if s["url"] not in seen_urls and s["url"] != failed_url:
                seen_urls.add(s["url"])
                unique_suggestions.append(s)
        return unique_suggestions[:5]

    def get_corrected_url(self) -> str:
        """
        URL_SLUG_MAPを適用した正しいURLを返す（v7.11追加）

        Returns:
            修正後のトップページURL
        """
        base_slug = self.ranking_slug.split('/')[0] if '/' in self.ranking_slug else self.ranking_slug

        # URL_SLUG_MAPに一致するか確認
        if base_slug in self.URL_SLUG_MAP:
            mapping = self.URL_SLUG_MAP[base_slug]
            new_slug = mapping["slug"]
            new_domain = mapping["domain"]

            # 新しいURLプレフィックスを構築
            if new_slug.startswith('_'):
                new_prefix = f"rank{new_slug}"
            else:
                new_prefix = f"rank-{new_slug}"

            # サブパスの変換
            subpath = self.subpath
            if subpath in self.SUBPATH_MAP:
                subpath = self.SUBPATH_MAP[subpath]

            subpath_part = f"/{subpath}" if subpath else ""
            return f"https://{new_domain}.oricon.co.jp/{new_prefix}{subpath_part}/"

        # 変換不要な場合は通常のURLを返す
        subpath_part = f"/{self.subpath}" if self.subpath else ""
        return f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"

    def analyze_structure(self, auto_correct: bool = True) -> SiteStructure:
        """
        サイト構造を解析（v7.9追加, v7.11強化）

        1回のHTTPリクエストで評価項目・部門・過去年度を一括取得。
        結果はキャッシュされ、2回目以降は即座に返す。

        v7.11追加: auto_correct=TrueでURL自動修正を試行

        Args:
            auto_correct: 404時に代替URLを自動的に試行するか（デフォルトTrue）

        Returns:
            SiteStructure: サイト構造情報
        """
        if self._site_structure is not None:
            return self._site_structure

        subpath_part = f"/{self.subpath}" if self.subpath else ""
        top_url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"

        # v7.11: URLバリデーションを追加
        validation = self.validate_url(top_url)
        if not validation["is_valid"] and auto_correct:
            logger.warning(f"URLアクセス失敗: {top_url} ({validation['error_type']})")

            # 代替URLを試行
            for suggestion in validation["suggestions"]:
                alt_url = suggestion["url"]
                alt_validation = self.validate_url(alt_url)
                if alt_validation["is_valid"]:
                    logger.info(f"代替URL成功: {alt_url} ({suggestion['reason']})")
                    top_url = alt_url
                    break
            else:
                # 全ての代替URLが失敗した場合
                if validation["suggestions"]:
                    logger.warning(f"代替URLの提案: {[s['url'] for s in validation['suggestions'][:3]]}")

        self._site_structure = self._structure_analyzer.analyze(top_url, self.url_prefix)

        # 年度情報を同期
        if self._site_structure.current_year and self._actual_top_year is None:
            self._actual_top_year = self._site_structure.current_year

        return self._site_structure

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
        ※整合性チェック: 更新日年度とタイトル年度が異なる場合はタイトル優先
          （サイトメンテナンスで更新日のみ変更されるケースへの対応）

        Returns:
            検出された年度（例: 2024）、検出できない場合はNone
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            text = soup.get_text()

            # 各パターンから年度を検出（整合性チェック用に変数に保持）
            update_year = None
            title_year = None
            header_year = None
            inferred_year = None

            # パターン1: 最終更新日から検出
            # 例: 「最終更新日：2025-11-01」「更新日: 2025/11/01」
            update_match = re.search(r'(?:最終)?更新日[：:\s]*(\d{4})[-/]\d{1,2}[-/]\d{1,2}', text)
            if update_match:
                update_year = int(update_match.group(1))
                logger.debug(f"最終更新日から年度検出: {update_year}年")

            # パターン2: タイトルから検出（ページ上部に表示されることが多い）
            # 例: 「2025年 オリコン顧客満足度」「2025年オリコン」
            title_match = re.search(r'(\d{4})年\s*オリコン', text)
            if title_match:
                title_year = int(title_match.group(1))
                logger.debug(f"タイトルから年度検出: {title_year}年")

            # パターン3: ページ冒頭の年度表記
            # 例: 「2025年 ネット証券」のような表記
            lines = text.split('\n')[:30]  # 最初の30行を対象
            for line in lines:
                year_match = re.search(r'^.*?(\d{4})年', line.strip())
                if year_match:
                    year = int(year_match.group(1))
                    if 2000 <= year <= 2030:  # 妥当な年度範囲
                        header_year = year
                        logger.debug(f"ページ冒頭から年度検出: {header_year}年")
                        break

            # パターン4: 過去ランキングリンクから推定（フォールバック、信頼性低）
            # ※ 2014-2015形式にも対応
            past_links = soup.find_all('a', href=re.compile(r'/\d{4}(?:-\d{4})?/?$'))
            if past_links:
                years = []
                for link in past_links:
                    href = link.get('href', '')
                    year_match = re.search(r'/(\d{4}(?:-\d{4})?)/?$', href)
                    if year_match:
                        year_str = year_match.group(1)
                        # ハイフン付き年度（例: 2014-2015）は終了年を使用
                        if "-" in year_str:
                            years.append(int(year_str.split("-")[1]))
                        else:
                            years.append(int(year_str))
                if years:
                    max_past_year = max(years)
                    inferred_year = max_past_year + 1
                    logger.debug(f"過去リンクから年度推定: {inferred_year}年（過去最大: {max_past_year}年）")

            # ===== 整合性チェック =====
            # 更新日年度とタイトル年度が両方存在し、かつ異なる場合
            # → サイトメンテナンスで更新日のみ変更された可能性があるため、タイトル年度を優先
            if update_year and title_year and update_year != title_year:
                logger.warning(
                    f"年度不一致検出: 更新日={update_year}年, タイトル={title_year}年, URL={url} "
                    f"→ タイトル年度を採用"
                )
                return title_year

            # 通常ケース: 既存の優先順位で返す
            if update_year:
                logger.info(f"最終更新日から年度検出: {update_year}年")
                return update_year

            if title_year:
                logger.info(f"タイトルから年度検出: {title_year}年")
                return title_year

            if header_year:
                logger.info(f"ページ冒頭から年度検出: {header_year}年")
                return header_year

            if inferred_year:
                logger.info(f"過去リンクから年度推定（フォールバック）: {inferred_year}年")
                return inferred_year

            logger.warning(f"年度を検出できませんでした: {url}")
            return None

        except Exception as e:
            logger.error(f"年度検出エラー: {e}")
            return None

    def get_update_date(self) -> Optional[tuple]:
        """
        トップページから更新日（年, 月）を取得する。

        調査概要の「更新日: 2025/01/06」などから年月を抽出。
        キャッシュを使用し、2回目以降は高速に返す。

        Returns:
            (year, month) のタプル、取得できない場合はNone
        """
        if self._update_date is not None:
            return self._update_date

        try:
            subpath_part = f"/{self.subpath}" if self.subpath else ""
            top_url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/"

            response = self.session.get(top_url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            text = soup.get_text()

            # 更新日パターン: 「最終更新日：2025-01-06」「更新日: 2025/01/06」など
            update_match = re.search(r'(?:最終)?更新日[：:\s]*(\d{4})[-/](\d{1,2})[-/]\d{1,2}', text)
            if update_match:
                year = int(update_match.group(1))
                month = int(update_match.group(2))
                self._update_date = (year, month)
                logger.info(f"更新日を検出: {year}年{month}月")
                return self._update_date

            logger.warning(f"更新日を検出できませんでした: {top_url}")
            return None

        except Exception as e:
            logger.error(f"更新日取得エラー: {e}")
            return None

    def _fetch_year_data(self, year: int, actual_top_year: int, top_url: str, subpath_part: str) -> tuple:
        """
        年度ごとのデータ取得（並列処理用ヘルパー）

        Returns:
            (year, data, url_info) のタプル
        """
        # 並列実行時のサーバー負荷軽減
        time.sleep(0.1)

        if year == actual_top_year:
            url = top_url
            logger.info(f"{year}年: トップページURL使用 {url}")
        elif year > actual_top_year:
            return (year, None, {
                "year": year,
                "url": "-",
                "survey_type": self.survey_type,
                "status": "not_published"
            })
        else:
            url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/{year}/"

        data = self._fetch_ranking_page(url, self.survey_type)
        if data:
            return (year, data, {"year": str(year), "url": url, "survey_type": self.survey_type, "status": "success"})
        
        # 代替パターン1: /year/subpath/ 形式を試す
        if self.subpath:
            alt_url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/"
            data = self._fetch_ranking_page(alt_url, self.survey_type)
            if data:
                return (year, data, {"year": str(year), "url": alt_url, "survey_type": self.survey_type, "status": "success"})

        # 代替パターン2: サブパスなしの親パス（過去年は分類がない場合）
        # 例: credit-card/free-annual/2023/ → credit-card/2023/
        if self.subpath:
            parent_url = f"{self.BASE_URL}/{self.url_prefix}/{year}/"
            data = self._fetch_ranking_page(parent_url, self.survey_type)
            if data:
                logger.info(f"{year}年: サブパスなし親パスにフォールバック {parent_url}")
                return (year, data, {"year": str(year), "url": parent_url, "survey_type": self.survey_type, "status": "success", "fallback": "parent_path"})

        # 代替パターン3: 特殊年度パターン
        special_url = f"{self.BASE_URL}/{self.url_prefix}/{year}-{year+1}{subpath_part}/"
        data = self._fetch_ranking_page(special_url, self.survey_type)
        if data:
            return (year + 1, data, {
                "year": str(year + 1),
                "year_format": f"{year}-{year+1}",
                "url": special_url,
                "survey_type": self.survey_type,
                "status": "success"
            })
        
        return (year, None, {"year": str(year), "url": url, "survey_type": self.survey_type, "status": "not_found"})

    def get_overall_rankings(self, year_range: tuple = (2020, 2024)) -> Dict[int, List[Dict]]:
        """
        総合ランキングを取得（v8.1: 並列処理対応）

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

        # 並列処理で年度ごとのデータを取得（v8.1追加）
        years_to_fetch = list(range(end_year, start_year - 1, -1))
        max_workers = min(5, len(years_to_fetch))  # 最大5並列（サーバー負荷考慮）
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self._fetch_year_data, year, actual_top_year, top_url, subpath_part): year
                for year in years_to_fetch
            }
            
            for future in as_completed(futures):
                year, data, url_info = future.result()
                # スレッドセーフにURL情報を追加
                with self._url_lock:
                    self.used_urls["overall"].append(url_info)

                if data:
                    year_key = str(year) if isinstance(year, int) else year
                    results[year_key] = data

        # 特殊年度パターン（YYYY-YYYY形式）を独立した年度として追加取得
        for year in range(end_year - 1, start_year - 1, -1):
            special_year_str = f"{year}-{year+1}"
            if special_year_str in results:
                continue

            special_url = f"{self.BASE_URL}/{self.url_prefix}/{special_year_str}{subpath_part}/"
            data = self._fetch_ranking_page(special_url, self.survey_type)
            if data:
                results[special_year_str] = data
                self.used_urls["overall"].append({
                    "year": special_year_str,
                    "url": special_url,
                    "survey_type": self.survey_type,
                    "status": "success",
                    "note": "特殊年度形式（独立データ）"
                })
                logger.info(f"特殊年度形式を独立データとして取得: {special_year_str} ({special_url})")

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
            consecutive_not_found = 0  # v7.11: 連続404カウンタ
            MAX_CONSECUTIVE_NOT_FOUND = 3  # 連続3回404で残りの年度をスキップ

            for year in years:
                # v7.11: 連続404が続いたら早期終了
                if consecutive_not_found >= MAX_CONSECUTIVE_NOT_FOUND:
                    logger.debug(f"{item_name}: 連続{MAX_CONSECUTIVE_NOT_FOUND}回404のため残りの年度をスキップ")
                    break

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
                    results[item_name][str(year)] = data  # 文字列で統一
                    self.used_urls["items"].append({
                        "name": f"{item_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "success",
                        "page_title": page_title,  # ページから取得した実際の名称
                        "item_slug": item_slug,
                        "year": str(year)  # 文字列で統一
                    })
                    consecutive_not_found = 0  # v7.11: 成功時はカウンタリセット
                else:
                    # 代替パターン: /year/subpath/ 形式を試す
                    if self.subpath:
                        alt_url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/evaluation-item/{item_slug}.html"
                        data = self._fetch_ranking_page(alt_url, self.survey_type)
                        if data:
                            page_title = self._extract_page_title(alt_url)
                            results[item_name][str(year)] = data  # 文字列で統一
                            self.used_urls["items"].append({
                                "name": f"{item_name}({year}年)",
                                "url": alt_url,
                                "survey_type": self.survey_type,
                                "status": "success",
                                "page_title": page_title,
                                "item_slug": item_slug,
                                "year": str(year)  # 文字列で統一
                            })
                            consecutive_not_found = 0  # v7.11: 成功時はカウンタリセット
                            continue

                    self.used_urls["items"].append({
                        "name": f"{item_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "not_found",
                        "item_slug": item_slug,
                        "year": year
                    })
                    consecutive_not_found += 1  # v7.11: 404時はカウンタ増加

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
            consecutive_not_found = 0  # v7.11: 連続404カウンタ
            MAX_CONSECUTIVE_NOT_FOUND = 3  # 連続3回404で残りの年度をスキップ

            for year in years:
                # v7.11: 連続404が続いたら早期終了（過去データが存在しない部門を効率的に処理）
                if consecutive_not_found >= MAX_CONSECUTIVE_NOT_FOUND:
                    logger.debug(f"{dept_name}: 連続{MAX_CONSECUTIVE_NOT_FOUND}回404のため残りの年度をスキップ")
                    break

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
                    # ページタイトルから実際の名称を取得（部門用の抽出関数を使用）
                    page_title = self._extract_page_title_for_dept(url)
                    results[dept_name][str(year)] = data  # 文字列で統一
                    self.used_urls["departments"].append({
                        "name": f"{dept_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "success",
                        "page_title": page_title,  # ページから取得した実際の名称
                        "dept_path": dept_path,
                        "year": str(year)  # 文字列で統一
                    })
                    consecutive_not_found = 0  # v7.11: 成功時はカウンタリセット
                else:
                    # 代替パターン1: /year/subpath/ 形式を試す
                    found_alt = False
                    if self.subpath:
                        alt_url = f"{self.BASE_URL}/{self.url_prefix}/{year}{subpath_part}/{dept_path}"
                        data = self._fetch_ranking_page(alt_url, self.survey_type)
                        if data:
                            page_title = self._extract_page_title_for_dept(alt_url)
                            results[dept_name][str(year)] = data  # 文字列で統一
                            self.used_urls["departments"].append({
                                "name": f"{dept_name}({year}年)",
                                "url": alt_url,
                                "survey_type": self.survey_type,
                                "status": "success",
                                "page_title": page_title,
                                "dept_path": dept_path,
                                "year": str(year)  # 文字列で統一
                            })
                            consecutive_not_found = 0  # v7.11: 成功時はカウンタリセット
                            found_alt = True

                    # 代替パターン2: YYYY-YYYY 特殊年度形式を試す（2014-2015など）
                    # 一部のランキングでは年度がハイフン付き形式で表現される
                    if not found_alt and year >= 2014 and year <= 2016:
                        special_year_formats = [
                            f"{year}-{year+1}",  # 例: 2014-2015
                            f"{year-1}-{year}",  # 例: 2013-2014（yearが終了年の場合）
                        ]
                        for special_year in special_year_formats:
                            special_url = f"{self.BASE_URL}/{self.url_prefix}{subpath_part}/{special_year}/{dept_path}"
                            data = self._fetch_ranking_page(special_url, self.survey_type)
                            if data:
                                page_title = self._extract_page_title_for_dept(special_url)
                                results[dept_name][str(year)] = data
                                self.used_urls["departments"].append({
                                    "name": f"{dept_name}({special_year}年)",
                                    "url": special_url,
                                    "survey_type": self.survey_type,
                                    "status": "success",
                                    "page_title": page_title,
                                    "dept_path": dept_path,
                                    "year": str(year)
                                })
                                consecutive_not_found = 0
                                found_alt = True
                                logger.info(f"特殊年度形式で取得成功: {special_year} → {dept_name}")
                                break

                    if found_alt:
                        continue

                    self.used_urls["departments"].append({
                        "name": f"{dept_name}({year}年)",
                        "url": url,
                        "survey_type": self.survey_type,
                        "status": "not_found",
                        "dept_path": dept_path,
                        "year": year
                    })
                    consecutive_not_found += 1  # v7.11: 404時はカウンタ増加

                time.sleep(0.3)

        return results

    def _discover_departments(self, url: str) -> Dict[str, str]:
        """
        ページから部門別リンクを動的に発見

        v7.0: ハイブリッドアプローチ
        1. sort-nav クラスから自動検出（優先）
        2. sort-nav がない場合はレガシー dept_patterns を使用

        v7.0リファクタリング:
        - クラス変数 DEPT_PATTERNS, EXCLUDE_URL_PATTERNS を参照
        - レガシーロジックでもアンカーテキストを直接使用（HTTPリクエスト削減）
        - 例外処理を具体化（RequestException）

        Returns:
            {"beginner/": "初心者", "age/50s.html": "50代", ...}
        """
        try:
            response = self.session.get(url, timeout=self.REQUEST_TIMEOUT_SEC)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            # ========================================
            # Phase 1: sort-nav からの自動検出（優先）
            # ========================================
            sort_nav = soup.find(class_="sort-nav")
            if sort_nav:
                departments = self._extract_departments_from_sort_nav(sort_nav, url)
                if departments:
                    logger.info(f"sort-nav から {len(departments)} 件の部門を検出: {url}")
                    return departments
                # sort-nav はあるが部門が見つからない場合はレガシーにフォールバック
                logger.info(f"sort-nav に部門なし、レガシーパターンを試行: {url}")

            # ========================================
            # Phase 2: レガシー dept_patterns による検出（改善版）
            # v7.0改善: アンカーテキストを直接使用（HTTPリクエスト不要）
            # ========================================
            all_links = soup.find_all("a", href=True)

            # サブパスを考慮したベースパターンを構築
            subpath_part = f"/{self.subpath}" if self.subpath else ""
            base_pattern = rf"/{self.url_prefix}{subpath_part}/(?:\d{{4}}/)?(.+?)(?:\?.*)?(?:#.*)?$"

            # 部門情報を直接格納（HTTPリクエスト不要）
            departments = {}

            for link in all_links:
                href = link.get("href", "")

                # 除外パターンにマッチする場合はスキップ（クラス変数を参照）
                if any(re.search(pat, href) for pat in self.EXCLUDE_URL_PATTERNS):
                    continue

                # 自身のランキングのリンクか確認
                if self.url_prefix not in href:
                    continue

                # 部門別パターンにマッチするか（クラス変数を参照）
                for pattern in self.DEPT_PATTERNS:
                    if re.search(pattern, href):
                        # パスを抽出（サブパスを考慮、クエリパラメータとハッシュを除外）
                        match = re.search(base_pattern, href)
                        if match:
                            dept_path = match.group(1)
                            # 数字のみのパス（年度）でない、クエリパラメータを含まないことを確認
                            if dept_path and not dept_path.rstrip('/').isdigit() and '?' not in dept_path:
                                # 重複チェック（既に登録済みの場合はスキップ）
                                if dept_path not in departments:
                                    # アンカーテキストを部門名として使用（HTTPリクエスト不要）
                                    dept_name = link.get_text(strip=True)
                                    # v7.4: バリデーション層追加 - 部門名の妥当性チェック強化
                                    if self._is_valid_dept_name(dept_name):
                                        departments[dept_path] = dept_name
                        break  # 一致したpatternループを抜ける

            if departments:
                logger.info(f"レガシーパターンから {len(departments)} 件の部門を検出: {url}")

            return departments

        except RequestException as e:
            logger.warning(f"部門リスト取得エラー（HTTPエラー）({url}): {e}")
            return {}
        except Exception as e:
            logger.error(f"部門リスト取得エラー（予期せぬエラー）({url}): {e}")
            return {}

    def _extract_departments_from_sort_nav(self, sort_nav, base_url: str) -> Dict[str, str]:
        """
        sort-nav クラスから部門リンクを抽出

        v7.7 修正: 実際のサイト構造（TABLE構造）に対応

        実際のsort-nav構造:
        <div class="sort-nav">
          <table>
            <tr>
              <th>TOP</th>
              <td></td>  ← 総合（除外）
            </tr>
            <tr>
              <th>評価項目別ランキング</th>
              <td><a href=".../evaluation-item/...">口座開設</a></td>  ← 除外
            </tr>
            <tr>
              <th>業態別ランキング</th>
              <td><a href=".../business/#1">FX専業</a></td>  ← 部門
            </tr>
          </table>
        </div>

        Args:
            sort_nav: sort-nav要素のBeautifulSoupオブジェクト
            base_url: ベースURL

        Returns:
            {"business/": "FX専業", ...}
        """
        departments = {}

        # サブパスを考慮
        subpath_part = f"/{self.subpath}" if self.subpath else ""
        base_pattern = rf"/{self.url_prefix}{subpath_part}/(?:\d{{4}}/)?(.+?)(?:\?.*)?(?:#.*)?$"

        # TABLE構造を処理（v7.7: 実際のサイト構造に対応）
        table = sort_nav.find("table")
        if table:
            for tr in table.find_all("tr"):
                th = tr.find("th")
                if not th:
                    continue

                heading_text = th.get_text(strip=True)

                # 除外対象の見出しはスキップ
                # - "TOP" は総合ページ
                # - EXCLUDE_HEADINGSに含まれるもの（評価項目別など）
                if heading_text == "TOP":
                    continue
                if any(exclude in heading_text for exclude in self.EXCLUDE_HEADINGS):
                    continue

                # この行内のTDからリンクを取得
                # v7.8: TD内の全リンクを取得（find_allを使用）
                for td in tr.find_all("td"):
                    # TD内のすべてのリンクを取得（派遣会社の業務内容別など）
                    links = td.find_all("a", href=True)
                    if not links:
                        continue

                    for link in links:
                        href = link.get("href", "")
                        link_text = link.get_text(strip=True)

                        # EXCLUDE_URL_PATTERNSに一致するリンクは除外
                        if any(re.search(pattern, href) for pattern in self.EXCLUDE_URL_PATTERNS):
                            continue

                        # 年度リンク（/2024/や/2014-2015/など）は除外
                        if re.search(r"/\d{4}(?:-\d{4})?/?$", href):
                            continue

                        # 自身のランキングのリンクか確認
                        if self.url_prefix not in href:
                            continue

                        # パスを抽出（#1などのフラグメントを除去）
                        href_clean = href.split('#')[0]
                        match = re.search(base_pattern, href_clean)
                        if match:
                            dept_path = match.group(1)
                            # 数字のみのパス（年度）は除外
                            if dept_path and not dept_path.rstrip('/').isdigit() and '?' not in dept_path:
                                # v7.4: バリデーション層 - 部門名の妥当性チェック
                                if self._is_valid_dept_name(link_text):
                                    departments[dept_path] = link_text

        # フォールバック: 旧SECTION構造（互換性のため残す）
        if not departments:
            sections = sort_nav.find_all("section")
            for section in sections:
                h3 = section.find("h3")
                if not h3:
                    continue

                heading_text = h3.get_text(strip=True)
                if any(exclude in heading_text for exclude in self.EXCLUDE_HEADINGS):
                    continue

                links = section.find_all("a", href=True)
                for link in links:
                    href = link.get("href", "")
                    link_text = link.get_text(strip=True)

                    if any(re.search(pattern, href) for pattern in self.EXCLUDE_URL_PATTERNS):
                        continue
                    # 年度リンク（/2024/や/2014-2015/など）は除外
                    if re.search(r"/\d{4}(?:-\d{4})?/?$", href):
                        continue
                    if self.url_prefix not in href:
                        continue

                    match = re.search(base_pattern, href)
                    if match:
                        dept_path = match.group(1)
                        if dept_path and not dept_path.rstrip('/').isdigit() and '?' not in dept_path:
                            if self._is_valid_dept_name(link_text):
                                departments[dept_path] = link_text

        return departments

    def _is_valid_dept_name(self, dept_name: str) -> bool:
        """
        部門名の妥当性をチェック（v7.4追加: バリデーション層）

        誤検出防止のため、以下を検証:
        1. 空でない（空白文字のみも除外）
        2. 最大文字数以下
        3. 無効な部門名リストに含まれていない（都道府県単体など）

        Args:
            dept_name: 検証する部門名

        Returns:
            True: 有効な部門名
            False: 無効な部門名（除外すべき）
        """
        # 空白文字のみの文字列を除外
        if not dept_name or not dept_name.strip():
            return False

        # 正規化（前後の空白除去）
        dept_name = dept_name.strip()

        if len(dept_name) > self.MAX_DEPT_NAME_LENGTH:
            return False

        # 無効な部門名リストに含まれていないか確認
        # （都道府県名も含まれているため、別途の正規表現チェックは不要）
        if dept_name in self.INVALID_DEPT_NAMES:
            logger.debug(f"無効な部門名を除外: {dept_name}")
            return False

        # 年度パターン（例: 2024年, 2023）を除外
        if re.match(r"^\d{4}年?$", dept_name):
            logger.debug(f"年度パターンを除外: {dept_name}")
            return False

        return True

    def _normalize_dept_url(self, url: str) -> str:
        """
        部門URLを正規化（v7.4追加: URL正規化レイヤー）

        - クエリパラメータを除去
        - ハッシュを除去
        - 複数の連続スラッシュを1つに
        - 末尾スラッシュを統一

        Args:
            url: 正規化前のURL

        Returns:
            正規化されたURL
        """
        if not url:
            return ""

        # クエリパラメータとハッシュを除去
        url = re.sub(r'[?#].*$', '', url)

        # 複数の連続スラッシュを1つに（プロトコル部分を除く）
        url = re.sub(r'(?<!:)/+', '/', url)

        # 末尾スラッシュを統一（.htmlで終わる場合は追加しない）
        if not url.endswith('.html'):
            url = url.rstrip('/') + '/'

        return url

    def _discover_evaluation_items(self, url: str) -> Dict[str, str]:
        """
        ページから評価項目リストを動的に発見

        survey_typeに応じてフィルタリング:
        - type01（顧客満足度）: #1 または ハッシュなしのURL
        - type02（FP評価等）: #2 のURL

        Returns:
            {"procedure": "加入手続き", ...}
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            items = {}

            # 対象のハッシュを決定
            # type01 → #1 または ハッシュなし
            # type02 → #2
            target_hash = "#1" if self.survey_type == "type01" else f"#{self.survey_type[-1]}"

            # サイドバーやナビゲーションから評価項目リンクを探す
            eval_links = soup.find_all("a", href=re.compile(r"/evaluation-item/"))

            for link in eval_links:
                href = link.get("href", "")
                match = re.search(r"/evaluation-item/([^/]+)\.html(#\d)?", href)
                if match:
                    slug = match.group(1)
                    url_hash = match.group(2) if match.group(2) else ""

                    # survey_typeに応じたフィルタリング
                    if self.survey_type == "type01":
                        # type01: ハッシュなし、#1 のみ許可（#2は除外）
                        if url_hash == "#2":
                            continue
                    else:
                        # type02以降: 対応するハッシュのみ許可
                        if url_hash != target_hash:
                            continue

                    # 日本語名を取得
                    name = link.get_text(strip=True)
                    if not name:
                        name = self.EVALUATION_ITEMS.get(slug, slug)
                    items[slug] = name

            # 見つからない場合は空を返す（ランキングごとに異なるため、デフォルト適用は危険）
            if not items:
                logger.info(f"評価項目が検出されませんでした: {url}")

            return items

        except Exception as e:
            logger.warning(f"評価項目リスト取得エラー ({url}): {e}")
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
        タイトル文字列から評価項目名を抽出

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

    def _extract_dept_name_from_title(self, text: str) -> Optional[str]:
        """
        タイトル文字列から部門名を抽出

        対応パターン:
        - 【2025年】初心者向けのネット証券 オリコン... → 初心者
        - 【2025年】初心者におすすめのネット証券 オリコン... → 初心者
        - 【2025年】50代向けの生命保険 オリコン... → 50代
        - 【2025年】ネット証券 NISAのランキング → NISA
        - 【2025年】スポーツ向けのクレジットカード → スポーツ
        - 【2025年】PCユーザーにおすすめのネット証券 → PCユーザー
        - 【2025年】FXの初心者ランキング・比較 → 初心者
        - 【2025年】FXのスキャルピングトレードランキング・比較 → スキャルピングトレード
        - 【2025年】FXのPCランキング・比較 → PC
        - 【アニメ】動画配信サービスのジャンル別ランキング → アニメ (SVOD)
        - NISA（シンプルなタイトル） → NISA

        非対応（Noneを返す）:
        - 【最新】ネット証券のランキング・比較 → None
        """
        if not text:
            return None

        # パターン-1（最優先）: 【ジャンル名】XXXサービスの... → ジャンル名 (SVOD向け)
        # 例: 【アニメ】動画配信サービスのジャンル別ランキング → アニメ
        # 例: 【洋画】動画配信サービスのジャンル別ランキング → 洋画
        match = re.search(r"【([^年】]+?)】(?:動画配信|定額制)", text)
        if match:
            dept_name = match.group(1).strip()
            # 年度（2025年など）でない場合のみ
            if dept_name and not re.match(r"^\d{4}年?$", dept_name) and dept_name not in ["最新"]:
                return dept_name

        # パターン0: 既知のシンプルな部門名（ホワイトリスト方式）
        # 誤検出を防ぐため、明示的にリスト化
        KNOWN_SIMPLE_DEPT_NAMES = ["NISA", "iDeCo", "つみたてNISA", "ジュニアNISA", "新NISA",
                                   "外国株式", "投資信託", "スマホ証券", "初心者", "中長期", "スイングトレード",
                                   "幼児", "小学生", "低学年", "高学年"]  # 子ども英語教室
        simple_text = text.strip()
        if simple_text in KNOWN_SIMPLE_DEPT_NAMES:
            return simple_text
        # 部分一致でもOK（タイトルに含まれていれば抽出）
        for known_name in KNOWN_SIMPLE_DEPT_NAMES:
            if known_name in text:
                return known_name

        # パターン0.6: 【年度】XXXに関する満足度の高い → XXX を抽出（v7.1追加）
        # 例: 【2025年】デイトレードに関する満足度の高いネット証券 → デイトレード
        match = re.search(r"】([^\s【】]+?)に関する満足度の高い", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン0.7: 【年度】XXXの運用におすすめの → XXX を抽出（v7.1追加）
        # 例: 【2025年】外国株式の運用におすすめのネット証券 → 外国株式
        # 例: 【2025年】国内株式の運用におすすめのネット証券 → 国内株式
        match = re.search(r"】([^\s【】]+?)の運用におすすめの", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン0.8: 【年度】XXXのYYYランキング → YYY を抽出（FX向け）
        # 例: 【2025年】FXの初心者ランキング・比較 → 初心者
        # 例: 【2025年】FXのスキャルピングトレードランキング・比較 → スキャルピングトレード
        # 例: 【2025年】FXのPCランキング・比較 → PC
        # ※「満足度」を含む場合は除外（誤マッチ防止）
        # ※「におすすめの」「のおすすめ」「に強い」「を希望」を含む場合は除外（派遣会社等の誤マッチ防止 v7.10）
        if "におすすめの" not in text and "のおすすめ" not in text and "に強い" not in text and "を希望" not in text:
            match = re.search(r"】[^\s【】]+の(.+?)ランキング", text)
            if match:
                dept_name = match.group(1).strip()
                # 「顧客満足度」「満足度」などの一般的な語句は除外
                if dept_name and dept_name not in ["顧客満足度", "オリコン顧客満足度", "満足度"] and "満足度" not in dept_name and len(dept_name) <= 20:
                    return dept_name

        # パターン1: 【年度】XXX向けのYYY → XXX を抽出
        # 例: 【2025年】初心者向けのネット証券 → 初心者
        match = re.search(r"】([^\s【】]+?)向けの", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン2: 【年度】XXXにおすすめのYYY → XXX を抽出
        # 例: 【2025年】初心者におすすめのネット証券 → 初心者
        match = re.search(r"】([^\s【】]+?)におすすめの", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン2.5: 【年度】XXXに強いYYY → XXX を抽出（派遣会社向け v7.10）
        # 例: 【2025年】オフィス・事務系に強い派遣会社 → オフィス・事務系
        match = re.search(r"】([^\s【】]+?)に強い", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 20:
                return dept_name

        # パターン2.55: 【年度】XXXを希望おすすめYYY → XXX を抽出（派遣会社・雇用形態向け v7.10）
        # 例: 【2025年】無期雇用派遣を希望おすすめ派遣会社 → 無期雇用派遣 → 無期雇用
        match = re.search(r"】([^\s【】]+?)を希望", text)
        if match:
            dept_name = match.group(1).strip()
            # "派遣" を末尾から除去
            if dept_name.endswith("派遣"):
                dept_name = dept_name[:-2]
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン2.6: 【年度】XXXのおすすめYYY → XXX を抽出（派遣会社向け v7.10）
        # 例: 【2025年】北海道地方のおすすめ派遣会社 → 北海道地方 → 北海道
        # 例: 【2025年】物流系のおすすめ派遣会社 → 物流系
        # 例: 【2025年】東京都のおすすめ派遣会社 → 東京都
        match = re.search(r"】([^\s【】]+?)のおすすめ", text)
        if match:
            dept_name = match.group(1).strip()
            # "地方" を除去して地域名のみにする
            if dept_name.endswith("地方"):
                dept_name = dept_name[:-2]
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン3: 【年度】XXXに人気のYYY → XXX
        match = re.search(r"】([^\s【】]+?)に人気の", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン4: 【年度】XXXユーザーにおすすめの → XXXユーザー
        # 例: 【2025年】PCユーザーにおすすめのネット証券 → PCユーザー
        match = re.search(r"】([^\s【】]+?ユーザー)におすすめの", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン5: YYYY年 XXX向けの → XXX
        match = re.search(r"\d{4}年[】\s]*([^\s【】]+?)向けの", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン6: YYYY年 XXXにおすすめの → XXX
        match = re.search(r"\d{4}年[】\s]*([^\s【】]+?)におすすめの", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and len(dept_name) <= 15:
                return dept_name

        # パターン7: XXX YYYのランキング → YYY（スペース区切り）
        # 例: 【2025年】ネット証券 NISAのランキング → NISA
        match = re.search(r"\s([^\s]+?)のランキング", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and not re.match(r"^\d{4}年?$", dept_name) and len(dept_name) <= 15:
                return dept_name

        # パターン8: YYYY年 XXX｜ → XXX
        match = re.search(r"\d{4}年\s+(.+?)(?:｜|\||ランキング)", text)
        if match:
            dept_name = match.group(1).strip()
            if dept_name and not re.match(r"^\d{4}年?$", dept_name) and len(dept_name) <= 15:
                return dept_name

        return None

    def _extract_page_title_for_dept(self, url: str) -> Optional[str]:
        """
        部門ページから部門名を抽出

        Args:
            url: ページURL

        Returns:
            部門名（例: "初心者", "50代"）
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            # パターン1: h1タグから取得
            h1 = soup.find("h1")
            if h1:
                text = h1.get_text(strip=True)
                extracted = self._extract_dept_name_from_title(text)
                if extracted:
                    return extracted

            # パターン2: og:title メタタグから取得
            og_title = soup.find("meta", property="og:title")
            if og_title:
                text = og_title.get("content", "")
                extracted = self._extract_dept_name_from_title(text)
                if extracted:
                    return extracted

            # パターン3: titleタグから取得
            title = soup.find("title")
            if title:
                text = title.get_text(strip=True)
                extracted = self._extract_dept_name_from_title(text)
                if extracted:
                    return extracted

            return None

        except Exception as e:
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

            # フォールバック: 古いHTML構造（2014-2015年頃のページ）への対応
            # 古いページは ul.rankin > li > p.rank + p.name 構造を使用
            if not rankings:
                legacy_list = target_section.find("ul", class_="rankin")
                if legacy_list:
                    logger.info(f"古いHTML構造（ul.rankin）を検出: {url}")
                    for li in legacy_list.find_all("li"):
                        try:
                            rank_elem = li.find("p", class_="rank")
                            name_elem = li.find("p", class_="name")

                            if rank_elem and name_elem:
                                rank_text = rank_elem.get_text(strip=True)
                                rank_match = re.search(r"(\d+)", rank_text)
                                rank = int(rank_match.group(1)) if rank_match else None

                                # 企業名はリンクテキストまたは直接テキストから取得
                                name_link = name_elem.find("a")
                                company = name_link.get_text(strip=True) if name_link else name_elem.get_text(strip=True)

                                if rank and company and company not in seen_companies:
                                    rankings.append({
                                        "rank": rank,
                                        "company": company,
                                        "score": None  # 古いページには得点がない場合がある
                                    })
                                    seen_companies.add(company)
                        except Exception as e:
                            continue

            # 順位でソート、同順位の場合は得点で降順ソート
            rankings.sort(key=lambda x: (x.get("rank", 999), -(x.get("score") or 0)))

            return rankings

        except Exception as e:
            logger.debug(f"ページ取得エラー ({url}): {e}")
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

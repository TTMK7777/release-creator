"""
URL解決ロジック（ハイブリッドアプローチ）

マスターデータを優先、URL推測ロジックをフォールバックとして使用します。

使用例:
    resolver = URLResolver(master_data_path="data/master_data.json")
    url = resolver.get_url("2078")  # マスターデータから取得

    # マスターデータにない場合は推測
    url = resolver.get_url("new-ranking-slug")  # URL推測で生成

バージョン: 1.0
作成日: 2026-01-09
"""

import logging
from typing import Dict, Optional, Tuple
from urllib.parse import urlparse
import re

# 既存のインポート（scraper.pyから移行）
from src.data_access.master_data_loader import MasterDataLoader

logger = logging.getLogger(__name__)


class URLResolver:
    """
    ハイブリッドURL解決
    優先度: マスターデータ > URL推測ロジック
    """

    # URL推測用のマッピング（既存のscraper.pyから移行）
    SUBDOMAIN_MAP = {
        # 教育系 → juken.oricon.co.jp
        "online-english": "juken",
        "kids-english": "juken",
        "_english": "juken",
        "_college": "juken",
        "highschool": "juken",
        "_junior": "juken",
        "online-study": "juken",
        "tutor": "juken",
        "cc": "juken",
        "license": "juken",
        "kids-swimming": "juken",
        "kids-programming": "juken",
        "programming-for-kids": "juken",
        "cram-school": "juken",
        "elementary-school-tutoring": "juken",

        # キャリア系 → career.oricon.co.jp
        "new-graduates-hiring-website": "career",
        "reversed-job-offer": "career",
        "arbeit": "career",
        "job-change": "career",
        "_agent": "career",
        "_staffing": "career",
        "dispatch": "career",
        "employment-agency": "career",
        "recruiting": "career",
        "career-change": "career",
        "recruitment-support": "career",
        "senior-employment": "career",
        "freelance-agent": "career",
        "side-job": "career",
        "training": "career",
        "talent-management": "career",
    }

    URL_SLUG_MAP = {
        # life.oricon.co.jp の変更
        "moving-company": {"slug": "_move", "domain": "life"},
        "travel-reservation-site": {"slug": "travel-website", "domain": "life"},
        "fitness-gym": {"slug": "_fitness", "domain": "life"},
        "_hikari": {"slug": "_internet", "domain": "life"},
        "home-wifi": {"slug": "_internet", "domain": "life"},

        # life → career への移動
        "_site": {"slug": "job-change", "domain": "career"},
        "_haken": {"slug": "_staffing", "domain": "career"},

        # life → juken への移動
        "english-school": {"slug": "_english", "domain": "juken"},
        "programming-school": {"slug": "kids-programming", "domain": "juken"},

        # juken.oricon.co.jp の変更
        "kids-english-school": {"slug": "kids-english", "domain": "juken"},
        "swimming-school": {"slug": "kids-swimming", "domain": "juken"},
    }

    def __init__(self, master_data_path: str = "data/master_data.json"):
        """
        Args:
            master_data_path: マスターデータのパス
        """
        self.master_data_loader = MasterDataLoader(master_data_path)

    def get_url(self, identifier: str) -> Tuple[str, str]:
        """
        URLを取得（優先順位付き）

        Args:
            identifier: ランキングID（数字）またはスラッグ

        Returns:
            (URL, モード) のタプル
            モード: "master_data", "inference", "error"
        """
        # 優先度1: マスターデータから取得
        try:
            url = self.master_data_loader.get_ranking_url(identifier)
            logger.debug(f"✅ マスターデータからURL取得: {url}")
            return url, "master_data"

        except KeyError:
            logger.debug(f"⚠️  マスターデータに未登録: {identifier}")

        except Exception as e:
            logger.error(f"❌ マスターデータ読み込みエラー: {e}")

        # 優先度2: URL推測ロジック
        try:
            url = self._infer_url_from_slug(identifier)
            logger.warning(
                f"⚠️  URL推測モードで動作: {url}\n"
                f"   推奨: マスターデータに追加してください（ID: {identifier}）"
            )
            return url, "inference"

        except Exception as e:
            logger.error(f"❌ URL推測失敗: {e}")

        # 優先度3: エラー
        error_message = (
            f"URL取得失敗: {identifier}\n"
            f"  - マスターデータに未登録\n"
            f"  - URL推測ロジックでも生成できません"
        )
        logger.error(error_message)
        raise ValueError(error_message)

    def _infer_url_from_slug(self, slug: str) -> str:
        """
        スラッグからURLを推測（既存ロジック）

        Args:
            slug: ランキングスラッグ

        Returns:
            推測されたURL
        """
        # URL_SLUG_MAPで変換が必要か確認
        base_slug = slug.split('/')[0].split('@')[0]

        if base_slug in self.URL_SLUG_MAP:
            mapping = self.URL_SLUG_MAP[base_slug]
            slug = mapping["slug"]
            subdomain = mapping["domain"]
        else:
            # サブドメイン決定
            subdomain = self._determine_subdomain(base_slug)

        # URLプレフィックス決定
        url_prefix = self._build_url_prefix(base_slug)

        # サブパスの処理
        subpath = ""
        if "/" in slug:
            parts = slug.split("/", 1)
            url_prefix = self._build_url_prefix(parts[0])
            subpath = "/" + parts[1]

        return f"https://{subdomain}.oricon.co.jp/{url_prefix}{subpath}/"

    def _determine_subdomain(self, slug: str) -> str:
        """
        スラッグからサブドメインを決定

        Args:
            slug: ランキングスラッグ

        Returns:
            サブドメイン（life, juken, career）
        """
        base_slug = slug.split('/')[0].split('@')[0]

        for slug_pattern, domain in self.SUBDOMAIN_MAP.items():
            if base_slug == slug_pattern or base_slug.startswith(slug_pattern):
                return domain

        return "life"  # デフォルト

    def _build_url_prefix(self, slug: str) -> str:
        """
        スラッグからURLプレフィックスを構築

        Args:
            slug: ランキングスラッグ

        Returns:
            URLプレフィックス（例: "rank_fx", "rank-mobile-carrier"）
        """
        if slug.startswith("_"):
            return f"rank{slug}"  # _fx → rank_fx
        elif slug.startswith("rank_"):
            return slug  # rank_certificate → rank_certificate
        elif slug.startswith("rank-"):
            return slug  # rank-mobile-carrier → rank-mobile-carrier
        else:
            return f"rank-{slug}"  # mobile-carrier → rank-mobile-carrier

    def get_alternative_urls(self, identifier: str) -> list:
        """
        代替URL候補を取得（404時の復旧用）

        Args:
            identifier: ランキングID（数字）またはスラッグ

        Returns:
            代替URL候補のリスト
        """
        alternatives = []

        # マスターデータから同じスラッグの別URLを検索
        try:
            # identifierが数字の場合はスラッグを取得
            if identifier.isdigit():
                ranking = self.master_data_loader.get_ranking(identifier)
                if ranking:
                    slug = ranking.get("slug")
                    if slug:
                        master_urls = self.master_data_loader.find_by_slug(slug)
                        alternatives.extend(master_urls)
            else:
                # スラッグの場合は直接検索
                master_urls = self.master_data_loader.find_by_slug(identifier)
                alternatives.extend(master_urls)

        except Exception as e:
            logger.debug(f"マスターデータからの代替URL取得失敗: {e}")

        # URL推測ロジックでも代替URLを生成
        try:
            inferred_url = self._infer_url_from_slug(identifier)
            if inferred_url not in alternatives:
                alternatives.append(inferred_url)

        except Exception as e:
            logger.debug(f"URL推測による代替URL生成失敗: {e}")

        return alternatives

    def extract_slug_from_url(self, url: str) -> Optional[str]:
        """
        URLからスラッグを抽出

        Args:
            url: URL

        Returns:
            スラッグ（抽出できない場合はNone）
        """
        # rank- または rank_ の後ろを取得
        match = re.search(r'oricon\.co\.jp/(rank[_-])?([^/]+)', url)

        if match:
            prefix = match.group(1)  # rank- または rank_
            slug = match.group(2)

            # rank_ で始まる場合は _ を付ける
            if prefix and prefix.startswith('rank_'):
                return '_' + slug

            return slug

        return None

    def get_statistics(self) -> Dict:
        """
        統計情報を取得

        Returns:
            マスターデータの統計情報
        """
        return self.master_data_loader.get_statistics()

    def search_rankings(self, keyword: str) -> list:
        """
        ランキング名で検索

        Args:
            keyword: 検索キーワード

        Returns:
            マッチしたランキング情報のリスト
        """
        return self.master_data_loader.search_by_name(keyword)


class URLResolverWithValidation(URLResolver):
    """
    URL解決 + 検証機能付き
    scraper.pyのRankingScraperと統合する際に使用
    """

    def __init__(
        self,
        master_data_path: str = "data/master_data.json",
        session=None
    ):
        """
        Args:
            master_data_path: マスターデータのパス
            session: requests.Sessionインスタンス（URL検証用）
        """
        super().__init__(master_data_path)
        self.session = session

    def get_url_with_validation(self, identifier: str) -> Tuple[str, str]:
        """
        URLを取得し、存在確認も実施

        Args:
            identifier: ランキングID（数字）またはスラッグ

        Returns:
            (URL, モード) のタプル
            モード: "master_data", "inference", "recovered"
        """
        # URL取得
        url, mode = self.get_url(identifier)

        # URL検証（sessionが設定されている場合のみ）
        if self.session:
            validation_result = self._validate_url(url)

            if not validation_result["valid"]:
                logger.warning(f"⚠️  URL検証失敗: {url}")

                # 代替URLで復旧を試行
                alternatives = self.get_alternative_urls(identifier)

                for alternative in alternatives:
                    if alternative == url:
                        continue  # 既に失敗したURLはスキップ

                    logger.info(f"   代替URLを試行: {alternative}")

                    if self._validate_url(alternative)["valid"]:
                        logger.info(f"✅ 代替URL使用: {alternative}")
                        return alternative, "recovered"

                # すべて失敗
                raise ConnectionError(
                    f"URLにアクセスできません: {url}\n"
                    f"代替URLもすべて失敗しました"
                )

        return url, mode

    def _validate_url(self, url: str) -> Dict:
        """
        URLの存在確認

        Args:
            url: 検証するURL

        Returns:
            検証結果の辞書
        """
        try:
            response = self.session.get(url, timeout=5)

            if response.status_code == 200:
                return {"valid": True, "url": url}
            elif response.status_code == 404:
                return {"valid": False, "status": 404}
            else:
                return {"valid": False, "status": response.status_code}

        except Exception as e:
            logger.error(f"URL検証エラー: {e}")
            return {"valid": False, "error": str(e)}


# ユーティリティ関数
def resolve_url(
    identifier: str,
    master_data_path: str = "data/master_data.json"
) -> str:
    """
    URLを解決（シンプルなインターフェース）

    Args:
        identifier: ランキングID（数字）またはスラッグ
        master_data_path: マスターデータのパス

    Returns:
        URL
    """
    resolver = URLResolver(master_data_path)
    url, _ = resolver.get_url(identifier)
    return url

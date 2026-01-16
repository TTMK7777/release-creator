"""
マスターデータローダー

CSDataViewerから生成されたmaster_data.jsonを読み込み、
ランキング情報を提供します。

使用例:
    loader = MasterDataLoader("data/master_data.json")
    ranking = loader.get_ranking("2078")
    url = ranking["url"]

バージョン: 1.0
作成日: 2026-01-09
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Optional
from functools import lru_cache
import hashlib

logger = logging.getLogger(__name__)


class MasterDataLoader:
    """マスターデータを読み込み、ランキング情報を提供"""

    def __init__(self, data_path: str = "data/master_data.json"):
        """
        Args:
            data_path: マスターデータのパス
        """
        self.data_path = Path(data_path)
        self.data = None
        self.rankings_by_id = {}
        self.rankings_by_slug = {}

        # データ読み込み
        self._load_with_fallback()

    def _load_with_fallback(self):
        """フォールバック機能付きデータ読み込み"""
        try:
            # メインファイルを読み込み
            self.data = self._load_json(self.data_path)
            self._build_indexes()

            logger.info(
                f"✅ マスターデータ読み込み成功: {len(self.rankings_by_id)}件"
            )

        except FileNotFoundError:
            logger.error(f"❌ マスターデータが見つかりません: {self.data_path}")
            self._try_load_backup()

        except json.JSONDecodeError as e:
            logger.error(f"❌ JSON解析エラー: {e}")
            self._try_load_backup()

        except Exception as e:
            logger.error(f"❌ 予期しないエラー: {e}")
            self._try_load_backup()

    def _load_json(self, path: Path) -> Dict:
        """JSONファイルを読み込み"""
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # データ検証
        self._validate_data(data)

        return data

    def _validate_data(self, data: Dict):
        """データの妥当性を検証"""
        required_fields = ["version", "source", "total_rankings", "rankings"]

        for field in required_fields:
            if field not in data:
                raise ValueError(f"必須フィールド '{field}' が見つかりません")

        # ランキング数の整合性確認
        declared_count = data["total_rankings"]
        actual_count = len(data["rankings"])

        if declared_count != actual_count:
            logger.warning(
                f"⚠️  ランキング数が不一致: "
                f"宣言={declared_count}件, 実際={actual_count}件"
            )

        # チェックサム検証
        if "checksum" in data:
            self._verify_checksum(data)

    def _verify_checksum(self, data: Dict):
        """チェックサムを検証"""
        expected_checksum = data["checksum"]
        data_without_checksum = {k: v for k, v in data.items() if k != "checksum"}
        json_str = json.dumps(data_without_checksum, sort_keys=True, ensure_ascii=False)
        actual_checksum = hashlib.sha256(json_str.encode()).hexdigest()

        if expected_checksum != actual_checksum:
            logger.warning("⚠️  チェックサムが一致しません（データ改変の可能性）")

    def _build_indexes(self):
        """高速検索用のインデックスを構築"""
        if not self.data or "rankings" not in self.data:
            return

        for ranking in self.data["rankings"]:
            # ID別インデックス
            ranking_id = ranking.get("id")
            if ranking_id:
                self.rankings_by_id[ranking_id] = ranking

            # スラッグ別インデックス
            slug = ranking.get("slug")
            if slug:
                # 複数ランキングが同じスラッグを持つ場合はリスト化
                if slug not in self.rankings_by_slug:
                    self.rankings_by_slug[slug] = []
                self.rankings_by_slug[slug].append(ranking)

        logger.debug(
            f"インデックス構築完了: "
            f"ID={len(self.rankings_by_id)}件, "
            f"スラッグ={len(self.rankings_by_slug)}種類"
        )

    def _try_load_backup(self):
        """バックアップファイルから読み込み"""
        backup_path = Path(str(self.data_path) + ".backup")

        if backup_path.exists():
            logger.warning(f"⚠️  バックアップから復元: {backup_path}")
            try:
                self.data = self._load_json(backup_path)
                self._build_indexes()
                logger.info("✅ バックアップからの復元に成功")
                return
            except Exception as e:
                logger.error(f"❌ バックアップの読み込みに失敗: {e}")

        # バックアップも失敗した場合
        logger.critical("❌ マスターデータが利用できません")
        raise FileNotFoundError(
            f"マスターデータが見つかりません: {self.data_path}\n"
            f"バックアップも利用できません: {backup_path}"
        )

    def get_ranking(self, ranking_id: str) -> Optional[Dict]:
        """
        ランキング情報を取得

        Args:
            ranking_id: ランキングID

        Returns:
            ランキング情報（見つからない場合はNone）
        """
        return self.rankings_by_id.get(ranking_id)

    def get_ranking_url(self, ranking_id: str) -> str:
        """
        ランキングURLを取得

        Args:
            ranking_id: ランキングID

        Returns:
            URL

        Raises:
            KeyError: ランキングIDが見つからない場合
        """
        ranking = self.get_ranking(ranking_id)

        if not ranking:
            raise KeyError(f"ランキングID '{ranking_id}' が見つかりません")

        return ranking["url"]

    def find_by_slug(self, slug: str) -> List[str]:
        """
        スラッグからURLリストを取得（代替URL検索用）

        Args:
            slug: スラッグ（例: "online-english", "_fx"）

        Returns:
            URLのリスト
        """
        rankings = self.rankings_by_slug.get(slug, [])
        return [r["url"] for r in rankings]

    def get_all_rankings(
        self,
        category: str = None,
        subdomain: str = None,
        active_only: bool = True
    ) -> List[Dict]:
        """
        全ランキング（またはフィルタ条件に合致するランキング）を取得

        Args:
            category: カテゴリ名でフィルタ（例: "教育"）
            subdomain: サブドメインでフィルタ（例: "juken"）
            active_only: アクティブなランキングのみ取得

        Returns:
            ランキング情報のリスト
        """
        if not self.data or "rankings" not in self.data:
            return []

        rankings = self.data["rankings"]

        # フィルタリング
        if active_only:
            rankings = [r for r in rankings if r.get("active", True)]

        if category:
            rankings = [r for r in rankings if r.get("category_name") == category]

        if subdomain:
            rankings = [r for r in rankings if r.get("subdomain") == subdomain]

        return rankings

    def get_statistics(self) -> Dict:
        """
        統計情報を取得

        Returns:
            統計情報の辞書
        """
        if not self.data:
            return {}

        return self.data.get("statistics", {})

    def get_version(self) -> str:
        """
        マスターデータのバージョンを取得

        Returns:
            バージョン文字列
        """
        if not self.data:
            return "unknown"

        return self.data.get("version", "unknown")

    def search_by_name(self, keyword: str) -> List[Dict]:
        """
        ランキング名で検索

        Args:
            keyword: 検索キーワード

        Returns:
            マッチしたランキング情報のリスト
        """
        if not self.data or "rankings" not in self.data:
            return []

        keyword_lower = keyword.lower()

        return [
            r for r in self.data["rankings"]
            if keyword_lower in r.get("name", "").lower()
        ]

    def is_valid_ranking_id(self, ranking_id: str) -> bool:
        """
        ランキングIDが存在するか確認

        Args:
            ranking_id: ランキングID

        Returns:
            存在する場合True
        """
        return ranking_id in self.rankings_by_id

    def reload(self):
        """マスターデータを再読み込み"""
        logger.info("マスターデータを再読み込み中...")
        self._load_with_fallback()


# グローバルインスタンス（シングルトンパターン）
_global_loader: Optional[MasterDataLoader] = None


def get_master_data_loader(data_path: str = "data/master_data.json") -> MasterDataLoader:
    """
    グローバルなMasterDataLoaderインスタンスを取得

    Args:
        data_path: マスターデータのパス

    Returns:
        MasterDataLoaderインスタンス
    """
    global _global_loader

    if _global_loader is None:
        _global_loader = MasterDataLoader(data_path)

    return _global_loader


@lru_cache(maxsize=128)
def get_ranking_url_cached(ranking_id: str, data_path: str = "data/master_data.json") -> str:
    """
    ランキングURLを取得（キャッシュ付き）

    Args:
        ranking_id: ランキングID
        data_path: マスターデータのパス

    Returns:
        URL
    """
    loader = get_master_data_loader(data_path)
    return loader.get_ranking_url(ranking_id)

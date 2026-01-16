"""
URLマスター管理モジュール

CSDataViewerから生成したURLマスター（JSON）を一元管理し、
スラッグからURL取得、カテゴリ別フィルタリングなどの機能を提供。

Usage:
    from url_manager import URLManager, get_url_manager

    # シングルトンインスタンス取得
    manager = get_url_manager()

    # URL取得
    url = manager.get_url("_insurance")

    # カテゴリ別に取得
    insurance_rankings = manager.get_rankings_by_category("保険")
"""

import json
from pathlib import Path
from typing import Dict, List, Optional, Any
from functools import lru_cache
from dataclasses import dataclass


@dataclass
class RankingEntry:
    """ランキングエントリ"""
    id: int
    name: str
    url: str
    subdomain: str
    category: str
    slug: str
    is_active: bool = True

    @classmethod
    def from_dict(cls, slug: str, data: Dict[str, Any]) -> "RankingEntry":
        """辞書からRankingEntryを作成"""
        return cls(
            id=data.get("id", 0),
            name=data.get("name", ""),
            url=data.get("url", ""),
            subdomain=data.get("subdomain", ""),
            category=data.get("category", ""),
            slug=slug,
            is_active=data.get("is_active", True)
        )


class URLManager:
    """URLマスター管理クラス"""

    def __init__(self, master_file: Optional[Path] = None):
        """
        Args:
            master_file: URLマスターJSONファイルのパス（省略時はデフォルト）
        """
        if master_file is None:
            master_file = Path(__file__).parent / "data" / "url_master.json"

        self.master_file = master_file
        self._rankings: Dict[str, RankingEntry] = {}
        self._metadata: Dict[str, Any] = {}
        self._load()

    def _load(self) -> None:
        """マスターファイルを読み込み"""
        if not self.master_file.exists():
            raise FileNotFoundError(f"URL master file not found: {self.master_file}")

        with open(self.master_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        self._metadata = data.get("metadata", {})

        for slug, entry_data in data.get("rankings", {}).items():
            self._rankings[slug] = RankingEntry.from_dict(slug, entry_data)

    def reload(self) -> None:
        """マスターファイルを再読み込み"""
        self._rankings.clear()
        self._metadata.clear()
        self._load()

    @property
    def total_count(self) -> int:
        """総件数を取得"""
        return len(self._rankings)

    @property
    def metadata(self) -> Dict[str, Any]:
        """メタデータを取得"""
        return self._metadata.copy()

    def get_url(self, slug: str) -> Optional[str]:
        """
        スラッグからURLを取得

        Args:
            slug: ランキングスラッグ（例: "_insurance", "card-loan/nonbank"）

        Returns:
            URL文字列、見つからない場合はNone
        """
        entry = self._rankings.get(slug)
        if entry and entry.is_active:
            return entry.url
        return None

    def get_entry(self, slug: str) -> Optional[RankingEntry]:
        """
        スラッグからRankingEntryを取得

        Args:
            slug: ランキングスラッグ

        Returns:
            RankingEntry、見つからない場合はNone
        """
        entry = self._rankings.get(slug)
        if entry and entry.is_active:
            return entry
        return None

    def get_name(self, slug: str) -> Optional[str]:
        """
        スラッグからランキング名を取得

        Args:
            slug: ランキングスラッグ

        Returns:
            ランキング名、見つからない場合はNone
        """
        entry = self._rankings.get(slug)
        if entry and entry.is_active:
            return entry.name
        return None

    def get_all_rankings(self) -> List[RankingEntry]:
        """アクティブな全ランキングを取得"""
        return [e for e in self._rankings.values() if e.is_active]

    def get_all_slugs(self) -> List[str]:
        """アクティブな全スラッグを取得"""
        return [e.slug for e in self._rankings.values() if e.is_active]

    def get_rankings_by_category(self, category: str) -> List[RankingEntry]:
        """
        カテゴリ別にランキングを取得

        Args:
            category: カテゴリ名

        Returns:
            該当するRankingEntryのリスト
        """
        return [
            e for e in self._rankings.values()
            if e.category == category and e.is_active
        ]

    def get_rankings_by_subdomain(self, subdomain: str) -> List[RankingEntry]:
        """
        サブドメイン別にランキングを取得

        Args:
            subdomain: サブドメイン（life, juken, career）

        Returns:
            該当するRankingEntryのリスト
        """
        return [
            e for e in self._rankings.values()
            if e.subdomain == subdomain and e.is_active
        ]

    def get_categories(self) -> List[str]:
        """全カテゴリを取得"""
        return sorted(set(e.category for e in self._rankings.values() if e.is_active))

    def get_subdomains(self) -> List[str]:
        """全サブドメインを取得"""
        return sorted(set(e.subdomain for e in self._rankings.values() if e.is_active))

    def search(self, keyword: str) -> List[RankingEntry]:
        """
        キーワードでランキングを検索

        Args:
            keyword: 検索キーワード

        Returns:
            名前にキーワードを含むRankingEntryのリスト
        """
        keyword_lower = keyword.lower()
        return [
            e for e in self._rankings.values()
            if keyword_lower in e.name.lower() and e.is_active
        ]

    def exists(self, slug: str) -> bool:
        """
        スラッグが存在するか確認

        Args:
            slug: ランキングスラッグ

        Returns:
            存在すればTrue
        """
        entry = self._rankings.get(slug)
        return entry is not None and entry.is_active

    def validate(self) -> List[str]:
        """
        マスターデータの検証

        Returns:
            問題のあるスラッグのリスト
        """
        issues = []
        for slug, entry in self._rankings.items():
            if not entry.url or not entry.url.startswith("http"):
                issues.append(f"{slug}: Invalid URL")
            if not entry.name:
                issues.append(f"{slug}: Missing name")
            if not entry.subdomain:
                issues.append(f"{slug}: Missing subdomain")
        return issues

    def to_ranking_options(self) -> Dict[str, str]:
        """
        app.pyのranking_options形式に変換

        Returns:
            {ランキング名: スラッグ} の辞書
        """
        return {
            entry.name: entry.slug
            for entry in self._rankings.values()
            if entry.is_active
        }

    def to_ranking_options_by_category(self) -> Dict[str, Dict[str, str]]:
        """
        カテゴリ別にranking_options形式で返す

        Returns:
            {カテゴリ: {ランキング名: スラッグ}} の辞書
        """
        result = {}
        for entry in self._rankings.values():
            if not entry.is_active:
                continue
            if entry.category not in result:
                result[entry.category] = {}
            result[entry.category][entry.name] = entry.slug
        return result


# シングルトンインスタンス
_manager_instance: Optional[URLManager] = None


@lru_cache(maxsize=1)
def get_url_manager(master_file: Optional[str] = None) -> URLManager:
    """
    URLManagerのシングルトンインスタンスを取得

    Args:
        master_file: マスターファイルパス（省略時はデフォルト）

    Returns:
        URLManagerインスタンス
    """
    global _manager_instance
    if _manager_instance is None:
        path = Path(master_file) if master_file else None
        _manager_instance = URLManager(path)
    return _manager_instance


# 便利関数（直接インポート用）
def get_url(slug: str) -> Optional[str]:
    """スラッグからURLを取得（便利関数）"""
    return get_url_manager().get_url(slug)


def get_ranking_name(slug: str) -> Optional[str]:
    """スラッグからランキング名を取得（便利関数）"""
    return get_url_manager().get_name(slug)


if __name__ == "__main__":
    # テスト実行
    manager = get_url_manager()

    print(f"Total rankings: {manager.total_count}")
    print(f"Categories: {manager.get_categories()}")
    print(f"Subdomains: {manager.get_subdomains()}")

    # URL取得テスト
    test_slugs = ["_insurance", "card-loan/nonbank", "online-english"]
    for slug in test_slugs:
        url = manager.get_url(slug)
        name = manager.get_name(slug)
        print(f"\n{slug}:")
        print(f"  Name: {name}")
        print(f"  URL: {url}")

    # 検索テスト
    results = manager.search("保険")
    print(f"\n'保険' search results: {len(results)} items")
    for r in results[:5]:
        print(f"  - {r.name} ({r.slug})")

    # バリデーション
    issues = manager.validate()
    if issues:
        print(f"\nValidation issues: {len(issues)}")
        for issue in issues[:5]:
            print(f"  - {issue}")
    else:
        print("\nNo validation issues")

# -*- coding: utf-8 -*-
"""
企業名マスタ管理モジュール (v1.0)
プレスリリースの正誤チェック用

機能:
- 正式名称マスタによる企業名検証
- エイリアス（旧社名、略称）の正規化
- 表記ゆれの検出とFuzzy Matching
- 企業名の妥当性チェック
"""

import re
from typing import Dict, List, Optional, Tuple
from difflib import SequenceMatcher


# ========================================
# 正式名称マスタ
# ※ 実際の運用時にはExcel/CSVから読み込むことも可能
# ========================================
COMPANY_MASTER: Dict[str, Dict] = {
    # === 証券会社 ===
    "SBI証券": {"category": "証券", "aliases": ["SBI", "SBI証券株式会社"]},
    "楽天証券": {"category": "証券", "aliases": ["楽天証券株式会社"]},
    "マネックス証券": {"category": "証券", "aliases": ["マネックス"]},
    "松井証券": {"category": "証券", "aliases": ["松井証券株式会社"]},
    "auカブコム証券": {"category": "証券", "aliases": ["カブドットコム証券", "カブコム証券"]},
    "GMOクリック証券": {"category": "証券", "aliases": ["GMOクリック"]},
    "野村證券": {"category": "証券", "aliases": ["野村証券", "野村"]},
    "大和証券": {"category": "証券", "aliases": ["大和"]},
    "SMBC日興証券": {"category": "証券", "aliases": ["日興証券", "SMBC日興"]},

    # === FX ===
    "GMO外貨": {"category": "FX", "aliases": ["外貨ex byGMO", "YJFX!", "外貨ex"]},
    "ヒロセ通商": {"category": "FX", "aliases": ["LION FX", "ヒロセ"]},
    "SBI FXトレード": {"category": "FX", "aliases": ["SBI FX"]},
    "DMM FX": {"category": "FX", "aliases": ["DMM.com証券"]},
    "外為どっとコム": {"category": "FX", "aliases": ["外為ドットコム"]},

    # === 携帯キャリア ===
    "NTTドコモ": {"category": "通信", "aliases": ["ドコモ", "docomo"]},
    "au": {"category": "通信", "aliases": ["KDDI", "エーユー"]},
    "ソフトバンク": {"category": "通信", "aliases": ["SoftBank"]},
    "楽天モバイル": {"category": "通信", "aliases": ["Rakuten Mobile"]},

    # === 格安SIM/MVNO ===
    "UQモバイル": {"category": "MVNO", "aliases": ["UQ mobile", "UQ"]},
    "ワイモバイル": {"category": "MVNO", "aliases": ["Y!mobile", "Ymobile"]},
    "mineo": {"category": "MVNO", "aliases": ["マイネオ"]},
    "IIJmio": {"category": "MVNO", "aliases": ["IIJ"]},
    "OCN モバイル ONE": {"category": "MVNO", "aliases": ["OCNモバイル", "OCN"]},
    "ahamo": {"category": "MVNO", "aliases": ["アハモ"]},
    "povo": {"category": "MVNO", "aliases": ["ポヴォ"]},
    "LINEMO": {"category": "MVNO", "aliases": ["ラインモ"]},

    # === 保険 ===
    "ソニー生命": {"category": "保険", "aliases": ["ソニー生命保険"]},
    "プルデンシャル生命": {"category": "保険", "aliases": ["プルデンシャル"]},
    "アフラック": {"category": "保険", "aliases": ["アフラック生命", "Aflac"]},
    "メットライフ生命": {"category": "保険", "aliases": ["メットライフ"]},
    "オリックス生命": {"category": "保険", "aliases": ["オリックス生命保険"]},

    # === 転職 ===
    "リクルートエージェント": {"category": "転職", "aliases": ["リクルート"]},
    "doda": {"category": "転職", "aliases": ["デューダ", "DODA"]},
    "マイナビエージェント": {"category": "転職", "aliases": ["マイナビ"]},
    "パソナキャリア": {"category": "転職", "aliases": ["パソナ"]},
    "JAC Recruitment": {"category": "転職", "aliases": ["JACリクルートメント", "ＪＡＣリクルートメント", "JAC"]},

    # === 引越し ===
    "サカイ引越センター": {"category": "引越し", "aliases": ["サカイ"]},
    "アート引越センター": {"category": "引越し", "aliases": ["アート"]},
    "アリさんマークの引越社": {"category": "引越し", "aliases": ["アリさん", "引越社"]},
    "日本通運": {"category": "引越し", "aliases": ["日通"]},
    "ハート引越センター": {"category": "引越し", "aliases": ["ハート"]},

    # === 動画配信 ===
    "Netflix": {"category": "動画配信", "aliases": ["ネットフリックス", "ネトフリ"]},
    "Amazon Prime Video": {"category": "動画配信", "aliases": ["Amazonプライムビデオ", "プライムビデオ"]},
    "U-NEXT": {"category": "動画配信", "aliases": ["ユーネクスト"]},
    "Hulu": {"category": "動画配信", "aliases": ["フールー"]},
    "Disney+": {"category": "動画配信", "aliases": ["ディズニープラス", "Disney Plus"]},
    "DMM TV": {"category": "動画配信", "aliases": ["DMMTV"]},

    # === 必要に応じて追加 ===
}


# ========================================
# エイリアス→正式名称の逆引き辞書を自動生成
# ========================================
def _build_alias_lookup() -> Dict[str, str]:
    """エイリアスから正式名称への逆引き辞書を構築"""
    lookup = {}
    for official_name, info in COMPANY_MASTER.items():
        # 正式名称自身も登録
        lookup[official_name] = official_name
        lookup[official_name.lower()] = official_name
        # エイリアスを登録
        for alias in info.get("aliases", []):
            lookup[alias] = official_name
            lookup[alias.lower()] = official_name
    return lookup

ALIAS_LOOKUP = _build_alias_lookup()


# ========================================
# 企業名正規化関数
# ========================================
def normalize_company_name(company: str) -> str:
    """企業名を正規化（正式名称に変換）

    Args:
        company: 入力企業名

    Returns:
        正規化された企業名（マスタにない場合は文字列正規化のみ）
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
    if company in ALIAS_LOOKUP:
        return ALIAS_LOOKUP[company]
    if normalized in ALIAS_LOOKUP:
        return ALIAS_LOOKUP[normalized]
    if normalized.lower() in ALIAS_LOOKUP:
        return ALIAS_LOOKUP[normalized.lower()]

    return normalized


def get_official_name(company: str) -> Optional[str]:
    """企業名に対応する正式名称を取得（マスタに存在する場合のみ）

    Args:
        company: 入力企業名

    Returns:
        正式名称（マスタにない場合はNone）
    """
    normalized = normalize_company_name(company)
    if normalized in COMPANY_MASTER:
        return normalized
    return None


def get_company_category(company: str) -> Optional[str]:
    """企業のカテゴリを取得

    Args:
        company: 企業名

    Returns:
        カテゴリ（マスタにない場合はNone）
    """
    official = get_official_name(company)
    if official and official in COMPANY_MASTER:
        return COMPANY_MASTER[official].get("category")
    return None


# ========================================
# Fuzzy Matching（類似度検索）
# ========================================
def find_similar_companies(
    company: str,
    threshold: float = 0.7,
    max_results: int = 5
) -> List[Tuple[str, float]]:
    """類似企業名を検索（Fuzzy Matching）

    Args:
        company: 検索する企業名
        threshold: 類似度の閾値（0.0-1.0）
        max_results: 最大結果数

    Returns:
        [(企業名, 類似度), ...] のリスト（類似度降順）
    """
    if not company:
        return []

    normalized = normalize_company_name(company)
    results = []

    # 全マスタ企業と比較
    for official_name in COMPANY_MASTER.keys():
        # 正式名称との類似度
        ratio = SequenceMatcher(None, normalized.lower(), official_name.lower()).ratio()
        if ratio >= threshold:
            results.append((official_name, ratio))

        # エイリアスとの類似度もチェック
        for alias in COMPANY_MASTER[official_name].get("aliases", []):
            alias_ratio = SequenceMatcher(None, normalized.lower(), alias.lower()).ratio()
            if alias_ratio >= threshold and alias_ratio > ratio:
                # より高い類似度があれば更新
                results = [(r[0], r[1]) for r in results if r[0] != official_name]
                results.append((official_name, alias_ratio))
                break

    # 類似度降順でソートして返す
    results.sort(key=lambda x: x[1], reverse=True)
    return results[:max_results]


def suggest_correction(company: str) -> Optional[Tuple[str, float]]:
    """企業名の修正候補を提案

    Args:
        company: 検証する企業名

    Returns:
        (修正候補, 類似度) または None
    """
    # 完全一致または正規化で一致する場合はNone
    if get_official_name(company):
        return None

    # 類似企業を検索
    similar = find_similar_companies(company, threshold=0.75, max_results=1)
    if similar:
        return similar[0]
    return None


# ========================================
# 検証関数
# ========================================
def validate_company_name(company: str) -> Dict:
    """企業名の妥当性を検証

    Args:
        company: 検証する企業名

    Returns:
        検証結果の辞書
    """
    result = {
        "input": company,
        "normalized": normalize_company_name(company),
        "is_valid": False,
        "official_name": None,
        "category": None,
        "suggestion": None,
        "warnings": []
    }

    if not company:
        result["warnings"].append("企業名が空です")
        return result

    # 正式名称チェック
    official = get_official_name(company)
    if official:
        result["is_valid"] = True
        result["official_name"] = official
        result["category"] = get_company_category(company)

        # 表記が正式名称と異なる場合は警告
        if company != official:
            result["warnings"].append(f"正式名称は「{official}」です")
    else:
        # 類似候補を検索
        suggestion = suggest_correction(company)
        if suggestion:
            result["suggestion"] = {
                "name": suggestion[0],
                "similarity": suggestion[1]
            }
            result["warnings"].append(
                f"マスタに登録がありません。「{suggestion[0]}」(類似度{suggestion[1]:.0%})ではありませんか？"
            )
        else:
            result["warnings"].append("マスタに登録がない企業名です（新規追加が必要な可能性があります）")

    return result


def batch_validate_companies(companies: List[str]) -> List[Dict]:
    """複数企業名を一括検証

    Args:
        companies: 企業名のリスト

    Returns:
        検証結果のリスト
    """
    return [validate_company_name(c) for c in companies]


# ========================================
# マスタ管理
# ========================================
def add_company(
    official_name: str,
    category: str,
    aliases: Optional[List[str]] = None
) -> bool:
    """企業をマスタに追加（実行時のみ有効、永続化なし）

    Args:
        official_name: 正式名称
        category: カテゴリ
        aliases: エイリアスのリスト

    Returns:
        追加成功フラグ
    """
    global ALIAS_LOOKUP

    if official_name in COMPANY_MASTER:
        return False  # 既に存在

    COMPANY_MASTER[official_name] = {
        "category": category,
        "aliases": aliases or []
    }

    # 逆引き辞書を更新
    ALIAS_LOOKUP = _build_alias_lookup()
    return True


def add_alias(official_name: str, alias: str) -> bool:
    """既存企業にエイリアスを追加（実行時のみ有効、永続化なし）

    Args:
        official_name: 正式名称
        alias: 追加するエイリアス

    Returns:
        追加成功フラグ
    """
    global ALIAS_LOOKUP

    if official_name not in COMPANY_MASTER:
        return False  # 企業が存在しない

    if alias not in COMPANY_MASTER[official_name]["aliases"]:
        COMPANY_MASTER[official_name]["aliases"].append(alias)
        ALIAS_LOOKUP = _build_alias_lookup()

    return True


def get_all_companies() -> List[str]:
    """全企業の正式名称リストを取得"""
    return list(COMPANY_MASTER.keys())


def get_companies_by_category(category: str) -> List[str]:
    """カテゴリで絞り込んだ企業リストを取得"""
    return [
        name for name, info in COMPANY_MASTER.items()
        if info.get("category") == category
    ]


# ========================================
# デバッグ用
# ========================================
if __name__ == "__main__":
    # テスト
    test_names = [
        "SBI証券",
        "SBI",
        "楽天証券株式会社",
        "カブドットコム証券",
        "ドコモ",
        "JACリクルートメント",
        "ＪＡＣリクルートメント",
        "存在しない会社",
        "SBI証権",  # タイポ
    ]

    print("=== 企業名検証テスト ===")
    for name in test_names:
        result = validate_company_name(name)
        print(f"\n入力: {name}")
        print(f"  正規化: {result['normalized']}")
        print(f"  有効: {result['is_valid']}")
        print(f"  正式名称: {result['official_name']}")
        if result['suggestion']:
            print(f"  候補: {result['suggestion']}")
        if result['warnings']:
            print(f"  警告: {result['warnings']}")

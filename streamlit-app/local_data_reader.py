# -*- coding: utf-8 -*-
"""
未公表ローカルランキングデータ読み込みクラス
v1.2 - config.json マッピング対応 + ヘッダー自動検出

共有フォルダに配置された未公表ランキングデータ（CSV / Excel）を読み込み、
既存スクレイパーと同一スキーマで返す。

環境変数:
    LOCAL_DATA_PATH: 共有フォルダのパス（例: \\\\server\\oricon\\local_rankings）
                     未設定の場合は <app_dir>/data/local_rankings/ をフォールバックとして使用

ファイル検索優先順位:
    1. config.json のマッピング（任意のファイル名・フルパスを指定可）
    2. {slug}__{year}.xlsx （Excel命名規則）
    3. {slug}__{year}.csv  （CSV命名規則）

config.json 形式（LOCAL_DATA_PATH 直下に配置）:
    {
        "_fx__2025":              "\\\\server\\oricon\\【資料】FX_ランキング結果2025.xlsx",
        "online-english__2025":   "\\\\server\\oricon\\【資料】英会話_2025.xlsx"
    }

対応列名（英語 / 日本語どちらも可）:
    rank    / 順位
    company / company_name / 企業名 / ランキング対象企業名 / 会社名
    score   / 得点 / スコア / 総合  ← 社内Excelの「総合」列にも対応
"""

import json
import logging
import os
from pathlib import Path
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

# 列名の候補（優先順位順）
_RANK_COLS    = ["rank", "順位", "Rank"]
_COMPANY_COLS = ["company", "company_name", "企業名", "ランキング対象企業名", "会社名"]
_SCORE_COLS   = ["score", "得点", "スコア", "Score", "総合", "合計"]  # 「合計」「総合」= 社内Excel総合スコア列


class LocalDataReader:
    """未公表ローカルランキングデータを読み込むクラス。

    LOCAL_DATA_PATH 環境変数が設定されていればそのパスを参照。
    未設定の場合は <app_dir>/data/local_rankings/ をフォールバックとして使用。

    ファイル解決は config.json マッピング → 命名規則の順で試みる。

    Example:
        reader = LocalDataReader()
        if reader.has_local_data("_fx", 2025):
            df = reader.get_ranking_data("_fx", 2025)
    """

    @property
    def local_dir(self) -> Path:
        """ローカルデータディレクトリのパスを返す。"""
        env_path = os.environ.get("LOCAL_DATA_PATH")
        if env_path:
            return Path(env_path)
        return Path(__file__).parent / "data" / "local_rankings"

    # ------------------------------------------------------------------
    # config.json サポート
    # ------------------------------------------------------------------

    def _load_config(self) -> dict:
        """config.json を読み込む（ファイルがなければ空 dict）。

        config.json 配置先: LOCAL_DATA_PATH / config.json

        Returns:
            dict: {slug__year: filepath, ...}
        """
        config_path = self.local_dir / "config.json"
        if not config_path.exists():
            return {}
        try:
            with open(config_path, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"config.json 読み込みエラー: {config_path} / {e}")
            return {}

    # ------------------------------------------------------------------
    # ファイル検索
    # ------------------------------------------------------------------

    def _find_file(self, slug: str, year: int) -> Optional[Path]:
        """ファイルを検索する。

        優先順位:
        1. config.json マッピング（任意のパス・ファイル名）
        2. {slug}__{year}.xlsx
        3. {slug}__{year}.csv

        Args:
            slug: ランキングスラッグ（例: "_fx", "online-english"）
            year: 年度（例: 2025）

        Returns:
            Optional[Path]: 見つかった場合はファイルパス、なければ None
        """
        key = f"{slug}__{year}"

        # 1. config.json マッピング優先
        config = self._load_config()
        if key in config:
            mapped = Path(config[key])
            if mapped.exists():
                logger.debug(f"config.json マッピング使用: {key} → {mapped}")
                return mapped
            logger.warning(f"config.json のパスが存在しません: {mapped}")

        # 2. 命名規則フォールバック
        base = self.local_dir / key
        for ext in (".xlsx", ".csv"):
            candidate = base.with_suffix(ext)
            if candidate.exists():
                return candidate

        return None

    def has_local_data(self, slug: str, year: int) -> bool:
        """指定のスラッグ・年度のローカルデータが存在するか確認する。"""
        found = self._find_file(slug, year) is not None
        if found:
            logger.debug(f"ローカルデータ検出: slug={slug}, year={year}")
        return found

    # ------------------------------------------------------------------
    # Excel 読み込み（ヘッダー自動検出）
    # ------------------------------------------------------------------

    @staticmethod
    def _detect_skiprows(path: Path, sheet: int = 0) -> int:
        """Excel のヘッダー行を自動検出して、スキップ行数を返す。

        先頭10行を走査し、rank / 順位 / ランキング対象企業名 などの
        キーワードを含む最初の行をヘッダー行とみなす。

        Returns:
            int: skiprows の値（0 = スキップなし）
        """
        rank_keywords = {"rank", "順位", "Rank", "RANK"}
        company_keywords = {"ランキング対象企業名", "企業名", "company", "会社名"}
        trigger_keywords = rank_keywords | company_keywords

        try:
            df_scan = pd.read_excel(
                path, sheet_name=sheet, header=None,
                engine="openpyxl", nrows=10
            )
            for idx, row in df_scan.iterrows():
                for val in row:
                    if str(val).strip() in trigger_keywords:
                        logger.debug(f"ヘッダー行検出: row={idx} ({val!r})")
                        return int(idx)
        except Exception as e:
            logger.warning(f"ヘッダー行検出失敗: {e}")
        return 0

    def _read_raw(self, path: Path) -> Optional[pd.DataFrame]:
        """ファイル拡張子に応じて生の DataFrame を読み込む。"""
        try:
            if path.suffix.lower() == ".xlsx":
                skiprows = self._detect_skiprows(path, sheet=0)
                df = pd.read_excel(
                    path, sheet_name=0,
                    skiprows=skiprows, header=0,
                    engine="openpyxl"
                )
                logger.info(f"Excel 読み込み: {path} (skiprows={skiprows}, {len(df)}行, 列: {list(df.columns[:6])}...)")
            else:
                df = pd.read_csv(path, encoding="utf-8-sig")
                logger.info(f"CSV 読み込み: {path} ({len(df)}行)")
            return df
        except Exception as e:
            logger.error(f"ファイル読み込みエラー: {path} / {e}")
            return None

    # ------------------------------------------------------------------
    # 列名正規化
    # ------------------------------------------------------------------

    @staticmethod
    def _detect_col(df: pd.DataFrame, candidates: list) -> Optional[str]:
        """candidates の中から DataFrame に存在する最初の列名を返す。"""
        for c in candidates:
            if c in df.columns:
                return c
        return None

    def _normalize(self, df: pd.DataFrame, path: Path) -> Optional[pd.DataFrame]:
        """列名を内部フォーマット (rank / company / score) に正規化する。"""
        rank_col    = self._detect_col(df, _RANK_COLS)
        company_col = self._detect_col(df, _COMPANY_COLS)
        score_col   = self._detect_col(df, _SCORE_COLS)

        if rank_col is None:
            logger.error(f"rank 列が見つかりません (file={path}, columns={list(df.columns[:10])})")
            return None
        if company_col is None:
            logger.error(f"company 列が見つかりません (file={path}, columns={list(df.columns[:10])})")
            return None

        col_map = {rank_col: "rank", company_col: "company"}
        if score_col:
            col_map[score_col] = "score"

        result = df[list(col_map.keys())].rename(columns=col_map).copy()

        if "score" not in result.columns:
            result["score"] = None
            logger.warning(f"score 列がありません。None で補完します (file={path})")

        return result

    # ------------------------------------------------------------------
    # 公開 API
    # ------------------------------------------------------------------

    def get_ranking_data(self, slug: str, year: int) -> Optional[pd.DataFrame]:
        """ローカルファイル（Excel/CSV）からランキングデータを読み込む。

        返す DataFrame のスキーマ:
            rank (Int64): 順位
            company (str): 企業名
            score (float|None): 得点（社内Excelの「総合」列）
            _source (str): "local" 固定（下流では drop すること）

        Args:
            slug: ランキングスラッグ
            year: 年度

        Returns:
            Optional[pd.DataFrame]: 正規化済みデータ、取得失敗時は None
        """
        path = self._find_file(slug, year)
        if path is None:
            logger.warning(f"ローカルデータファイルが見つかりません: slug={slug}, year={year}")
            return None

        raw = self._read_raw(path)
        if raw is None:
            return None

        df = self._normalize(raw, path)
        if df is None:
            return None

        # データ型の正規化
        df["rank"]    = pd.to_numeric(df["rank"],    errors="coerce").astype("Int64")
        df["score"]   = pd.to_numeric(df["score"],   errors="coerce")
        df["company"] = df["company"].astype(str)

        # 空行を除去
        df = df.dropna(subset=["rank", "company"])
        df = df[df["company"].str.strip() != ""]

        # UI 表示フラグ（下流では drop すること）
        df["_source"] = "local"

        logger.info(f"ローカルデータ正規化完了: {path} ({len(df)}件, score_col={df.get('score') is not None})")
        return df

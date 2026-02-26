#!/usr/bin/env python3
"""Release Creator ポータブルビルドスクリプト
==========================================
Python embeddable + 全依存ライブラリ + アプリコードを1パッケージに。
Python 未インストールの PC でもそのまま動作するポータブル配布パッケージを生成。

Usage:
    python build/build_portable.py
    python build/build_portable.py --python-version 3.11.9
    python build/build_portable.py --output ./dist

Requirements:
    - Python 3.10+ (ビルド環境用)
    - インターネット接続

Note:
    このスクリプトは開発者のPython環境で実行し、
    配布用パッケージを生成するためのものです。
    エンドユーザーは生成された ReleaseCreator.bat を実行するだけです。
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_PYTHON_VERSION = "3.11.9"

PYTHON_EMBED_URL_TEMPLATE = (
    "https://www.python.org/ftp/python/{version}/"
    "python-{version}-embed-amd64.zip"
)
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"

# アプリケーションファイル（streamlit-app/ からコピー対象）
APP_FILES: list[str] = [
    "app.py",
    "scraper.py",
    "analyzer.py",
    "site_analyzer.py",
    "url_manager.py",
    "release_generator.py",
    "release_tab.py",
    "validator.py",
    "word_generator.py",
    "image_generator.py",
    "company_master.py",
    "requirements.txt",
]

# ディレクトリごとコピー対象
APP_DIRS: list[str] = [
    "src",
]

# git-ignored だが配布に必要なデータディレクトリ（存在する場合のみコピー）
APP_DATA_DIRS: list[str] = [
    "config",
    "data",
]

# コピー除外ディレクトリ
EXCLUDE_DIRS: set[str] = {
    "__pycache__",
    ".git",
    ".pytest_cache",
    ".mypy_cache",
    "test",
    "tests",
    "scripts",
    "docs",
    "build",
    "archive",
    ".venv",
    "venv",
}

# コピー除外拡張子
EXCLUDE_EXTENSIONS: set[str] = {
    ".pyc",
    ".pyo",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _python_tag(version: str) -> str:
    """'3.11.9' -> '311'"""
    major, minor, *_ = version.split(".")
    return f"{major}{minor}"


def _download(url: str, dest: Path, label: str) -> None:
    """ダウンロード + 進捗表示"""

    def _reporthook(block_num: int, block_size: int, total_size: int) -> None:
        if total_size > 0:
            pct = min(100, block_num * block_size * 100 // total_size)
            print(f"\r  ダウンロード中: {pct}%", end="", flush=True)

    print(f"  URL: {url}")
    urllib.request.urlretrieve(url, dest, reporthook=_reporthook)
    print(f"\r  ダウンロード完了: {dest.name}         ")


def _copytree_filtered(src: Path, dst: Path) -> None:
    """__pycache__ 等を除外してコピー"""

    def _ignore(directory: str, contents: list[str]) -> list[str]:
        ignored: list[str] = []
        for item in contents:
            item_path = Path(directory) / item
            if item in EXCLUDE_DIRS:
                ignored.append(item)
            elif item_path.suffix in EXCLUDE_EXTENSIONS:
                ignored.append(item)
        return ignored

    shutil.copytree(src, dst, ignore=_ignore, dirs_exist_ok=True)


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------


class PortableBuilder:
    """ポータブル配布パッケージを生成する"""

    def __init__(
        self,
        project_root: Path,
        output_dir: Path,
        python_version: str,
    ) -> None:
        self.project_root = project_root
        self.output_dir = output_dir
        self.python_version = python_version
        self.tag = _python_tag(python_version)

        # ソースディレクトリ
        self.streamlit_app_dir = project_root / "streamlit-app"

        # 出力先
        self.package_dir = output_dir / "release-creator-portable"
        self.python_dir = self.package_dir / "python"
        self.app_dir = self.package_dir / "app"
        self.templates_dir = project_root / "build" / "templates"

        # Step 間で共有する状態
        self._pth_file: Path | None = None

        # 一時ディレクトリ
        self._tmpdir_obj: tempfile.TemporaryDirectory[str] | None = None
        self.tmp: Path = Path(tempfile.gettempdir())

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def build(self) -> None:
        print("=" * 60)
        print("Release Creator ポータブルビルド")
        print(f"  Python: {self.python_version}")
        print(f"  ソース: {self.streamlit_app_dir}")
        print(f"  出力先: {self.package_dir}")
        print("=" * 60)

        # ソースディレクトリの存在確認
        if not self.streamlit_app_dir.exists():
            raise FileNotFoundError(
                f"streamlit-app/ が見つかりません: {self.streamlit_app_dir}"
            )

        self._tmpdir_obj = tempfile.TemporaryDirectory(prefix="rc_build_")
        self.tmp = Path(self._tmpdir_obj.name)

        try:
            self._step01_clean()
            self._step02_download_python()
            self._step03_fix_pth()
            self._step04_extract_stdlib_zip()
            self._step05_bootstrap_pip()
            self._step06_install_dependencies()
            self._step07_copy_app()
            self._step08_streamlit_config()
            self._step09_place_launchers()
            self._step10_version_file()
            self._step11_verify()
        finally:
            if self._tmpdir_obj:
                self._tmpdir_obj.cleanup()

        print("\n" + "=" * 60)
        print("  ビルド完了!")
        print(f"  パッケージ: {self.package_dir}")
        size_mb = sum(
            f.stat().st_size for f in self.package_dir.rglob("*") if f.is_file()
        ) / (1024 * 1024)
        print(f"  合計サイズ: {size_mb:.0f} MB")
        print("=" * 60)

    # ------------------------------------------------------------------
    # Steps
    # ------------------------------------------------------------------

    def _step01_clean(self) -> None:
        print("\n[1/11] 出力先をクリーンアップ...")
        if self.package_dir.exists():
            print(f"  既存の {self.package_dir.name} を削除")
            shutil.rmtree(self.package_dir)
        self.package_dir.mkdir(parents=True)

    def _step02_download_python(self) -> None:
        print(f"\n[2/11] Python {self.python_version} embeddable をダウンロード...")
        url = PYTHON_EMBED_URL_TEMPLATE.format(version=self.python_version)
        zip_path = self.tmp / f"python-{self.python_version}-embed-amd64.zip"
        _download(url, zip_path, "Python embeddable")

        self.python_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(self.python_dir)
        print(f"  -> {self.python_dir} に展開完了")

    def _step03_fix_pth(self) -> None:
        print(f"\n[3/11] python{self.tag}._pth を修正...")
        pth_file = self.python_dir / f"python{self.tag}._pth"
        if not pth_file.exists():
            # フォールバック: glob で探す
            candidates = list(self.python_dir.glob("python*._pth"))
            if not candidates:
                raise FileNotFoundError(
                    f"python*._pth が見つかりません: {self.python_dir}"
                )
            pth_file = candidates[0]

        content = pth_file.read_text(encoding="utf-8")

        # 1. import site のコメントを解除
        content = content.replace("#import site", "import site")

        # 2. Lib/site-packages パスを追加
        if "Lib\\site-packages" not in content and "Lib/site-packages" not in content:
            content = content.rstrip("\n") + "\nLib\\site-packages\n"

        pth_file.write_text(content, encoding="utf-8")
        self._pth_file = pth_file  # step04 で再利用

        # 3. site-packages ディレクトリを作成
        site_packages = self.python_dir / "Lib" / "site-packages"
        site_packages.mkdir(parents=True, exist_ok=True)

        print(f"  -> {pth_file.name} 修正完了")
        print(f"     import site 有効化 + Lib\\site-packages 追加")

    def _step04_extract_stdlib_zip(self) -> None:
        """python311.zip -> ディレクトリに展開（lib2to3 pickle 問題対策）"""
        print(f"\n[4/11] 標準ライブラリ zip を展開（lib2to3 対策）...")
        stdlib_zip = self.python_dir / f"python{self.tag}.zip"
        if not stdlib_zip.exists():
            candidates = list(self.python_dir.glob("python*.zip"))
            if not candidates:
                print("  -> stdlib zip なし（スキップ）")
                return
            stdlib_zip = candidates[0]

        extract_dir = self.python_dir / stdlib_zip.stem
        extract_dir.mkdir(exist_ok=True)

        with zipfile.ZipFile(stdlib_zip) as zf:
            zf.extractall(extract_dir)

        # zip を削除
        stdlib_zip.unlink()

        # _pth ファイルで zip 参照をディレクトリ参照に置換
        pth_file = self._pth_file or self.python_dir / f"python{self.tag}._pth"
        if pth_file.exists():
            content = pth_file.read_text(encoding="utf-8")
            content = content.replace(stdlib_zip.name, stdlib_zip.stem)
            pth_file.write_text(content, encoding="utf-8")

        print(f"  -> {stdlib_zip.name} -> {extract_dir.name}/ に展開完了")

    def _step05_bootstrap_pip(self) -> None:
        print("\n[5/11] pip をブートストラップ...")
        get_pip_path = self.tmp / "get-pip.py"
        _download(GET_PIP_URL, get_pip_path, "get-pip.py")

        python_exe = self.python_dir / "python.exe"
        result = subprocess.run(
            [str(python_exe), str(get_pip_path), "--no-warn-script-location"],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            print(f"  STDERR: {result.stderr}")
            raise RuntimeError("pip ブートストラップ失敗")

        print("  -> pip インストール完了")

    def _step06_install_dependencies(self) -> None:
        print("\n[6/11] 依存ライブラリをインストール...")
        python_exe = self.python_dir / "python.exe"
        req_file = self.streamlit_app_dir / "requirements.txt"

        if not req_file.exists():
            raise FileNotFoundError(
                f"requirements.txt が見つかりません: {req_file}"
            )

        result = subprocess.run(
            [
                str(python_exe),
                "-m",
                "pip",
                "install",
                "-r",
                str(req_file),
                "--no-warn-script-location",
                "--disable-pip-version-check",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            print(f"  STDOUT (末尾): {result.stdout[-500:]}")
            print(f"  STDERR (末尾): {result.stderr[-500:]}")
            raise RuntimeError("pip install 失敗")

        # インストール済みパッケージ一覧
        list_result = subprocess.run(
            [str(python_exe), "-m", "pip", "list", "--format=columns"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        # ヘッダー行（Package/---）を除外して実パッケージ数をカウント
        pkg_count = sum(
            1
            for line in list_result.stdout.strip().splitlines()
            if line
            and not line.startswith("Package")
            and not line.startswith("---")
        )
        print(f"  -> {pkg_count} パッケージインストール完了")

    def _step07_copy_app(self) -> None:
        print("\n[7/11] アプリケーションファイルをコピー...")
        self.app_dir.mkdir(parents=True, exist_ok=True)
        src_base = self.streamlit_app_dir

        # 個別ファイルをコピー
        for fname in APP_FILES:
            src = src_base / fname
            dst = self.app_dir / fname
            if src.is_file():
                shutil.copy2(src, dst)
                print(f"  -> {fname}")
            else:
                print(f"  [WARN] {fname} が見つかりません（スキップ）")

        # ディレクトリをコピー（必須）
        for dname in APP_DIRS:
            src = src_base / dname
            dst = self.app_dir / dname
            if src.is_dir():
                _copytree_filtered(src, dst)
                file_count = sum(1 for _ in dst.rglob("*") if _.is_file())
                print(f"  -> {dname}/ ({file_count} ファイル)")
            else:
                print(f"  [WARN] {dname}/ が見つかりません（スキップ）")

        # データディレクトリをコピー（存在する場合のみ、git-ignored）
        for dname in APP_DATA_DIRS:
            src = src_base / dname
            dst = self.app_dir / dname
            if src.is_dir():
                _copytree_filtered(src, dst)
                file_count = sum(1 for _ in dst.rglob("*") if _.is_file())
                print(f"  -> {dname}/ ({file_count} ファイル)")
            else:
                print(
                    f"  [注意] {dname}/ が見つかりません（git-ignored の可能性あり）"
                )
                print(
                    f"         配布前に {src_base / dname} を手動配置してください"
                )

    def _step08_streamlit_config(self) -> None:
        print("\n[8/11] Streamlit ローカル設定を配置...")
        config_src = self.templates_dir / "config.toml"
        config_dst = self.app_dir / ".streamlit" / "config.toml"
        config_dst.parent.mkdir(parents=True, exist_ok=True)

        if config_src.exists():
            shutil.copy2(config_src, config_dst)
        else:
            # テンプレートが無い場合はインラインで作成
            config_dst.write_text(
                "[server]\n"
                "headless = true\n"
                'fileWatcherType = "none"\n'
                "\n"
                "[browser]\n"
                "gatherUsageStats = false\n",
                encoding="utf-8",
            )
        print(f"  -> {config_dst}")

    def _step09_place_launchers(self) -> None:
        print("\n[9/11] ランチャーファイルを配置...")
        launcher_files = [
            "ReleaseCreator.bat",
            "start-debug.bat",
            "stop.bat",
            "README.txt",
        ]
        for fname in launcher_files:
            src = self.templates_dir / fname
            dst = self.package_dir / fname
            if src.exists():
                shutil.copy2(src, dst)
                print(f"  -> {fname}")
            else:
                print(f"  [WARN] テンプレート {fname} が見つかりません")

    def _step10_version_file(self) -> None:
        print("\n[10/11] VERSION.txt を作成...")
        version_src = self.templates_dir / "VERSION.txt"
        version_dst = self.package_dir / "VERSION.txt"

        if version_src.exists():
            shutil.copy2(version_src, version_dst)
        else:
            version_dst.write_text("1.0.0\n", encoding="utf-8")

        version = version_dst.read_text(encoding="utf-8").strip()
        print(f"  -> VERSION.txt ({version})")

    def _step11_verify(self) -> None:
        print("\n[11/11] インストールを検証...")
        python_exe = self.python_dir / "python.exe"

        # 主要ライブラリの import テスト（requirements.txt 全パッケージ）
        test_imports = [
            "streamlit",
            "pandas",
            "requests",
            "bs4",
            "openpyxl",
            "xlsxwriter",
            "docx",
            "matplotlib",
            "rapidfuzz",
        ]
        failed: list[str] = []
        for mod in test_imports:
            result = subprocess.run(
                [str(python_exe), "-c", f"import {mod}; print('{mod}: OK')"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            if result.returncode == 0:
                print(f"  [OK] {mod}")
            else:
                err_msg = (
                    result.stderr.strip().splitlines()[-1]
                    if result.stderr.strip()
                    else "unknown"
                )
                print(f"  [NG] {mod}: {err_msg[:80]}")
                failed.append(mod)

        if failed:
            print(f"\n  [WARN] {len(failed)} 個のモジュールで import 失敗:")
            for m in failed:
                print(f"    - {m}")
            print("  -> requirements.txt を確認してください")
        else:
            print(f"  -> 全 {len(test_imports)} モジュール import 成功")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Release Creator ポータブル配布パッケージをビルド",
    )
    parser.add_argument(
        "--python-version",
        default=DEFAULT_PYTHON_VERSION,
        help=f"Python embeddable バージョン (default: {DEFAULT_PYTHON_VERSION})",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="出力ディレクトリ (default: プロジェクトルート/dist)",
    )
    args = parser.parse_args()

    # プロジェクトルート = build/ の親
    project_root = Path(__file__).resolve().parent.parent
    output_dir = Path(args.output) if args.output else project_root / "dist"

    builder = PortableBuilder(
        project_root=project_root,
        output_dir=output_dir,
        python_version=args.python_version,
    )
    builder.build()


if __name__ == "__main__":
    main()

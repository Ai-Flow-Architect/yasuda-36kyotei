"""
スモークテスト：最低限「アプリが起動する」「主要モジュールが import できる」を確認

実行:
    pytest tests/test_smoke.py -v
"""
import os
import sys
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)


def test_entry_point_exists():
    """エントリーポイントファイルが存在する（雛形リポジトリ自体ではスキップ）"""
    if os.path.basename(ROOT) == "_delivery_template":
        pytest.skip("雛形リポジトリのため、エントリーポイントは適用先プロジェクトで検証する")
    candidates = ["main.py", "app.py", "index.html", "server.py"]
    found = [c for c in candidates if os.path.exists(os.path.join(ROOT, c))]
    assert found, f"エントリーポイントが見つかりません（候補: {candidates}）"


def test_env_example_exists():
    """.env.example が存在する（新環境構築のため）"""
    assert os.path.exists(os.path.join(ROOT, ".env.example")), \
        ".env.example が存在しない → 環境変数の引き継ぎができない"


def test_gitignore_excludes_secrets():
    """`.gitignore` が機密情報を除外している"""
    gitignore = os.path.join(ROOT, ".gitignore")
    assert os.path.exists(gitignore), ".gitignore がない"
    with open(gitignore, "r", encoding="utf-8") as f:
        content = f.read()
    required = [".env", "credentials", "secrets"]
    missing = [r for r in required if r not in content]
    assert not missing, f".gitignore に以下が含まれていない: {missing}"


def test_readme_exists():
    """README.md が存在する"""
    assert os.path.exists(os.path.join(ROOT, "README.md")), \
        "README.md がない → 他社説明できない"


@pytest.mark.skip(reason="実装後にスキップ解除してアプリ固有のテストに置換")
def test_main_imports():
    """主要モジュールが import エラーなく読み込める"""
    # 例: import main
    pass

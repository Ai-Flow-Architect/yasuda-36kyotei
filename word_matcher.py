"""
word_matcher.py — 事業所名 × Word ファイル名マッチング
"""
import re
import subprocess
import tempfile
from pathlib import Path


def _normalize(name: str) -> str:
    """会社名を正規化（法人格・記号・スペース除去）"""
    return re.sub(
        r'[\s\u3000　株式会社有限会社合同会社一般社団法人医療法人（）()【】「」・]',
        '', str(name)
    ).lower()


def match_word_files(
    company_name: str,
    word_paths: list[Path],
) -> Path | None:
    """事業所名に最も近い Word ファイルを返す（完全一致優先、部分一致フォールバック）"""
    target = _normalize(company_name)
    if not target:
        return None
    # 完全一致
    for wp in word_paths:
        if _normalize(wp.stem) == target:
            return wp
    # 部分一致
    for wp in word_paths:
        stem = _normalize(wp.stem)
        if target in stem or stem in target:
            return wp
    return None


def build_match_table(
    records: list[dict],
    word_paths: list[Path],
) -> list[dict]:
    """Excelレコード × Word ファイルのマッチング結果を返す"""
    results = []
    for rec in records:
        name = rec.get("事業所名") or ""
        matched = match_word_files(name, word_paths)
        results.append({
            "事業所名": name,
            "送信先メール": rec.get("メールアドレス") or "⚠️ 未設定",
            "Wordファイル": matched.name if matched else "❌ 未マッチ",
            "_matched_path": matched,
            "_record": rec,
        })
    return results


def convert_docx_to_pdf(docx_path: Path, output_dir: Path) -> Path | None:
    """LibreOffice headless で Word → PDF 変換"""
    try:
        result = subprocess.run(
            [
                "libreoffice", "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_dir),
                str(docx_path),
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            return None
        pdf_path = output_dir / (docx_path.stem + ".pdf")
        return pdf_path if pdf_path.exists() else None
    except Exception:
        return None

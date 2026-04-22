"""
word_matcher.py — 事業所名 × Word ファイル名マッチング
命名規則: 0001_36協定書_事業所名.docx（先頭番号_任意テキスト_事業所名）
"""
import re
import subprocess
import tempfile
from pathlib import Path


def _normalize(name: str) -> str:
    """会社名を正規化（法人格・記号・スペース・全角半角除去）"""
    # 全角英数字→半角
    normalized = name.translate(str.maketrans(
        'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏ'
        'ｐｑｒｓｔｕｖｗｘｙｚ'
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯ'
        'ＰＱＲＳＴＵＶＷＸＹＺ'
        '０１２３４５６７８９',
        'abcdefghijklmnopqrstuvwxyz'
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        '0123456789'
    ))
    return re.sub(
        r'[\s　　株式会社有限会社合同会社'
        r'一般社団法人医療法人（）()【】'
        r'「」・]',
        '', str(normalized)
    ).lower()


def _extract_company_from_filename(stem: str) -> str | None:
    """ファイル名ステム（拡張子なし）から事業所名部分を抽出する。
    命名規則: 0001_36協定書_事業所名 → 事業所名 を返す
    アンダースコアで3分割し、3番目を事業所名とみなす。
    """
    parts = stem.split('_', 2)
    if len(parts) >= 3:
        return parts[2]
    return None


def match_word_files(
    company_name: str,
    word_paths: list[Path],
) -> Path | None:
    """事業所名に最も近い Word ファイルを返す。

    優先順位:
    1. ファイル名の事業所名部分（0001_36協定書_事業所名 の3番目）で完全一致
    2. ファイル名の事業所名部分で部分一致
    3. ファイル名全体での従来マッチング（フォールバック）
    """
    target = _normalize(company_name)
    if not target:
        return None

    # 優先: 命名規則ファイル名（3番目セグメント）での完全一致
    for wp in word_paths:
        extracted = _extract_company_from_filename(wp.stem)
        if extracted and _normalize(extracted) == target:
            return wp

    # 優先: 命名規則ファイル名での部分一致
    for wp in word_paths:
        extracted = _extract_company_from_filename(wp.stem)
        if extracted:
            norm_extracted = _normalize(extracted)
            if target in norm_extracted or norm_extracted in target:
                return wp

    # フォールバック: ファイル名全体での従来マッチング
    for wp in word_paths:
        if _normalize(wp.stem) == target:
            return wp
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

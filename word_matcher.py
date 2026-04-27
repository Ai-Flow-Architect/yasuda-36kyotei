"""
word_matcher.py — 事業所名 × Word ファイル名マッチング
命名規則（2種類対応）:
  旧形式: 0001_36協定書_事業所名.docx （先頭が数字）
  新形式: 36協定書_事業所名_様式第9号.docx （先頭が文字列）
"""
import re
import subprocess
import tempfile
import unicodedata
from pathlib import Path


def _normalize(name: str) -> str:
    """会社名を正規化（NFKC正規化 + 法人格・記号・スペース除去）

    unicodedata.normalize('NFKC') により:
    - 全角英数字 → 半角英数字
    - 半角カタカナ → 全角カタカナ
    全角スペース・記号・法人格サフィックスを除去後にlowercaseで統一する。
    """
    normalized = unicodedata.normalize('NFKC', str(name))
    return re.sub(
        r'[\s　株式会社有限会社合同会社'
        r'一般社団法人医療法人（）()【】'
        r'「」・]',
        '', normalized
    ).lower()


def _extract_company_from_filename(stem: str) -> str | None:
    """ファイル名ステム（拡張子なし）から事業所名部分を抽出する。

    旧形式: 0001_36協定書_事業所名 → parts[0]が数字 → parts[2]=事業所名
    新形式: 36協定書_事業所名_様式第9号 → parts[0]が文字列 → parts[1]=事業所名
    """
    parts = stem.split('_', 2)
    if len(parts) < 2:
        return None
    if parts[0].isdigit() and len(parts) >= 3:
        return parts[2]
    return parts[1]


def match_word_files(
    company_name: str,
    word_paths: list[Path],
) -> Path | None:
    """事業所名に最も近い Word ファイルを返す。

    優先順位:
    1. ファイル名から抽出した事業所名部分（新旧形式対応）で完全一致
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


def convert_docx_to_pdf(docx_path: Path, output_dir: Path) -> tuple[Path | None, str]:
    """LibreOffice headless で Word → PDF 変換。

    Returns:
        (pdf_path, error_msg): 成功時は (Path, "")、失敗時は (None, エラーメッセージ)
    """
    try:
        with tempfile.TemporaryDirectory(prefix='lo_profile_') as lo_profile:
            result = subprocess.run(
                [
                    "libreoffice", "--headless", "--norestore",
                    f"--env=UserInstallation=file://{lo_profile}",
                    "--convert-to", "pdf",
                    "--outdir", str(output_dir),
                    str(docx_path),
                ],
                capture_output=True,
                text=True,
                timeout=120,
            )
        if result.returncode != 0:
            err = result.stderr.strip() or f"returncode={result.returncode}"
            return None, f"LibreOffice変換失敗: {err}"
        pdf_path = output_dir / (docx_path.stem + ".pdf")
        if not pdf_path.exists():
            return None, "PDF出力ファイルが見つかりません"
        return pdf_path, ""
    except subprocess.TimeoutExpired:
        return None, "LibreOffice変換タイムアウト（120秒）"
    except Exception as e:
        return None, f"変換エラー: {e}"

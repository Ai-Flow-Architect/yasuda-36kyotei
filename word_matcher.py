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


def _extract_number_from_filename(stem: str) -> str | None:
    """ファイル名先頭の事業所番号を抽出する（例: 0001_36協定書_... → "0001"）。
    先頭セグメントが数字のみの場合のみ抽出する。
    """
    parts = stem.split('_', 1)
    if parts[0].isdigit():
        return parts[0].zfill(4)
    return None


def match_word_files(
    company_name: str,
    word_paths: list[Path],
    office_number: str = "",
) -> Path | None:
    """事業所名（または事業所番号）に最も近い Word ファイルを返す。

    優先順位:
    1. ファイル名先頭の事業所番号と完全一致（最優先・全角半角に依存しない）
    2. ファイル名から抽出した事業所名部分で完全一致
    3. ファイル名の事業所名部分で部分一致
    4. ファイル名全体での従来マッチング（フォールバック）
    """
    # 最優先: 事業所番号によるマッチング（ゼロ埋め正規化して比較）
    if office_number:
        normalized_num = office_number.strip().zfill(4)
        for wp in word_paths:
            file_num = _extract_number_from_filename(wp.stem)
            if file_num and file_num == normalized_num:
                return wp

    target = _normalize(company_name)
    if not target:
        return None

    # 命名規則ファイル名（事業所名セグメント）での完全一致
    for wp in word_paths:
        extracted = _extract_company_from_filename(wp.stem)
        if extracted and _normalize(extracted) == target:
            return wp

    # 命名規則ファイル名での部分一致
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


def match_word_files_multi(
    company_name: str,
    word_paths: list[Path],
    office_number: str = "",
) -> list[Path]:
    """1事業所に紐づく協定書ファイルを「すべて」返す。

    1事業所が様式第9号・様式第9号の2・1年単位変形 など複数の協定書を
    持つケースに対応するため、最も信頼度の高いマッチ階層で該当する
    ファイルを全件返す（バグ①: 複数協定書が添付されない問題の根本対応）。

    優先順位（最初に1件以上ヒットした階層の全件を返す）:
      1. ファイル名先頭の事業所番号と完全一致（最優先・全角半角非依存）
      2. ファイル名から抽出した事業所名部分で完全一致
      3. ファイル名の事業所名部分で部分一致
      4. ファイル名全体での従来マッチング（フォールバック）
    """
    # 1. 事業所番号マッチング（最優先・複数様式は同番号プレフィックスを共有する想定）
    if office_number:
        normalized_num = office_number.strip().zfill(4)
        num_hits = [
            wp for wp in word_paths
            if _extract_number_from_filename(wp.stem) == normalized_num
        ]
        if num_hits:
            return num_hits

    target = _normalize(company_name)
    if not target:
        return []

    # 2. 事業所名セグメント完全一致
    exact = [
        wp for wp in word_paths
        if (ex := _extract_company_from_filename(wp.stem)) and _normalize(ex) == target
    ]
    if exact:
        return exact

    # 3. 事業所名セグメント部分一致
    partial = []
    for wp in word_paths:
        ex = _extract_company_from_filename(wp.stem)
        if ex:
            norm_ex = _normalize(ex)
            if target in norm_ex or norm_ex in target:
                partial.append(wp)
    if partial:
        return partial

    # 4. フォールバック: ファイル名全体マッチング
    whole = [
        wp for wp in word_paths
        if (_normalize(wp.stem) == target
            or target in _normalize(wp.stem)
            or _normalize(wp.stem) in target)
    ]
    return whole


def build_match_table(
    records: list[dict],
    word_paths: list[Path],
) -> list[dict]:
    """Excelレコード × 協定書ファイル（Word/PDF）のマッチング結果を返す"""
    results = []
    for rec in records:
        name = rec.get("事業所名") or ""
        office_number = rec.get("事業所番号") or ""
        matched_paths = match_word_files_multi(name, word_paths, office_number=office_number)
        primary = matched_paths[0] if matched_paths else None
        suffixes = sorted({mp.suffix.upper().lstrip(".") for mp in matched_paths})
        results.append({
            "事業所名": name,
            "送信先メール": rec.get("メールアドレス") or "⚠️ 未設定",
            "協定書ファイル": (
                "／".join(mp.name for mp in matched_paths)
                if matched_paths else "❌ 未マッチ"
            ),
            "件数": len(matched_paths),
            "形式": "／".join(suffixes) if suffixes else "-",
            "_matched_path": primary,        # 後方互換（先頭1件）
            "_matched_paths": matched_paths,  # 全件（バグ①対応）
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

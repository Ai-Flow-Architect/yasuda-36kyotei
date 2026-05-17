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


# 様式サフィックス語（事業所名の後ろに付くため会社名から除去する）。
# 例: 0010_36協定書_丸和_様式第9号 / 36協定書_テスト工業_1年単位
_FORM_SUFFIX_PREFIXES = ('様式第', '様式', '1年単位', '1年変形', '変形', '一年単位', '一年変形')


def _is_form_suffix_segment(seg: str) -> bool:
    """セグメントが様式サフィックス（会社名ではない）かどうか判定する。

    単純な startswith では「変形製作所」「様式工業」「1年単位フーズ」など
    様式語で始まる正当な会社名まで誤って様式サフィックス扱いし、会社名が
    丸ごと消えて添付漏れになる（バグ①と同クラスの逆方向退行）。
    そのため「セグメント全体が様式語」または「様式語の直後が数字/号/の」
    （= 様式第9号 / 様式第9号の2 / 様式9号 等の実在パターン）に限定する。
    """
    for p in _FORM_SUFFIX_PREFIXES:
        if seg == p:
            return True
        if seg.startswith(p):
            rest = seg[len(p):]
            if rest and (rest[0].isdigit() or rest[0] in ('号', 'の')):
                return True
    return False


def _extract_company_from_filename(stem: str) -> str | None:
    """ファイル名ステム（拡張子なし）から事業所名部分を抽出する。

    旧形式: 0001_36協定書_事業所名 → parts[0]が数字 → parts[2]起点が事業所名
    新形式: 36協定書_事業所名_様式第9号 → parts[0]が文字列 → parts[1]起点が事業所名

    会社名の後ろに付く様式サフィックス（様式第9号 / 様式第9号の2 /
    1年単位 / 1年変形 等）はバグ①の誤添付要因になるため除去し、純粋な
    会社名のみを返す（番号prefix欠落＋部分文字列社名の誤マッチ防止）。
    """
    parts = stem.split('_')
    if len(parts) < 2:
        return None
    if parts[0].isdigit() and len(parts) >= 3:
        company_parts = parts[2:]   # 旧形式: 番号_36協定書_<会社名…>
    else:
        company_parts = parts[1:]   # 新形式: 36協定書_<会社名…>
    # 様式サフィックスセグメントを末尾から落とし、純company名を組み立てる
    pure: list[str] = []
    for seg in company_parts:
        if _is_form_suffix_segment(seg):
            break
        pure.append(seg)
    company = '_'.join(pure).strip()
    return company or None


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
    2. 様式サフィックス除去後の事業所名で厳格一致（誤添付防止のため完全一致のみ）
    3. ファイル名全体での完全一致（フォールバック）
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

    # 命名規則ファイル名（様式サフィックス除去後の事業所名セグメント）での
    # 厳格一致。単純な部分文字列一致（例: 丸和 ⊂ 丸和興業）は他事業所の
    # 協定書を誤添付する重大インシデント要因のため採用せず、正規化後の
    # 完全一致のみとする。
    for wp in word_paths:
        extracted = _extract_company_from_filename(wp.stem)
        if extracted and _normalize(extracted) == target:
            return wp

    # フォールバック: ファイル名全体での従来マッチング（完全一致のみ）
    # 番号prefixや様式サフィックス込みの単純部分一致は他事業所の協定書を
    # 誤添付する要因（例: 丸和 ⊂ 丸和興業）のため、ここでは採用しない。
    for wp in word_paths:
        if _normalize(wp.stem) == target:
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
      2. 様式サフィックス除去後の事業所名で厳格一致（誤添付防止・完全一致のみ）
      3. ファイル名全体での完全一致（フォールバック）
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

    # 2. 事業所名セグメント厳格一致（様式サフィックス除去後の完全一致のみ）
    # 単純な部分文字列一致は他事業所の協定書を誤添付する重大インシデント
    # 要因（例: 丸和 ⊂ 丸和興業）のため、正規化後完全一致のみ採用する。
    exact = [
        wp for wp in word_paths
        if (ex := _extract_company_from_filename(wp.stem)) and _normalize(ex) == target
    ]
    if exact:
        return exact

    # 3. フォールバック: ファイル名全体マッチング（完全一致のみ）
    # 番号prefixや様式サフィックス込みの単純部分一致は他事業所の協定書を
    # 誤添付する要因（例: 丸和 ⊂ 丸和興業）のため、ここでは採用しない。
    whole = [wp for wp in word_paths if _normalize(wp.stem) == target]
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

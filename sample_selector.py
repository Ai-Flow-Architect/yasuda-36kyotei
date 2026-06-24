"""
sample_selector.py — 事業所ごとの見本ファイル出し分け

追加開発（2026-06-24）: これまで見本ファイルは「全宛先共通で添付」だったが、
事業所ごとに異なる見本を出し分けたいという要望に対応する。

運用フロー（客様提案の "Excel列に見本名" 方式）:
  1. アプリに見本ファイルを「プール」としてまとめてアップロードする（複数可）
  2. Excel の「見本指定」列に、その事業所へ添付したい見本のファイル名を書く
     （複数指定は 、 , ; / 改行 空白 のいずれでも区切れる）
  3. 各事業所のメール下書きには、指定された見本だけが添付される

マッチングは「完全ファイル名一致 → 拡張子なし一致 → 部分一致」の順で
評価し、誤添付を避けつつ表記ゆれ（全角半角・大文字小文字）を吸収する。
見本ファイル本体は app.py 側でアップロードされるため、本モジュールは
ファイルI/Oを持たず純粋なロジックのみ（テスト容易・低regression）。
"""
import unicodedata
from pathlib import Path

# 見本指定セルの区切り文字（複数の見本を1セルに書けるようにする）
_SPEC_SEPARATORS = ["、", "，", ",", ";", "；", "/", "／", "|", "｜", "\n", "\r", "\t", " ", "　"]

# 部分一致を許容する最小トークン長（短すぎるトークンの誤爆を防ぐ）
_MIN_PARTIAL_LEN = 2


def _normalize(value: str) -> str:
    """NFKC正規化 + 前後空白除去 + 小文字化（全角半角・大小文字の揺れを吸収）。

    None は空文字として扱う（"None" という文字列に化けて誤マッチするのを防ぐ）。
    """
    if value is None:
        return ""
    return unicodedata.normalize("NFKC", str(value)).strip().lower()


def parse_sample_spec(spec: str) -> list[str]:
    """見本指定セルの値を個々のトークン（見本ファイル名候補）へ分解する。

    Args:
        spec: Excelの「見本指定」セルの生値（例 "建設業見本.pdf、製造業見本.pdf"）

    Returns:
        トークン文字列のリスト（空・重複は除去し、元の出現順を保持）
    """
    text = str(spec or "")
    for sep in _SPEC_SEPARATORS:
        text = text.replace(sep, "\x00")
    tokens: list[str] = []
    seen: set[str] = set()
    for raw in text.split("\x00"):
        t = raw.strip()
        if not t:
            continue
        key = _normalize(t)
        if key in seen:
            continue
        seen.add(key)
        tokens.append(t)
    return tokens


# マッチの精度ティア（大きいほど厳密）
_TIER_EXACT = 3    # 完全ファイル名一致（拡張子込み）
_TIER_STEM = 2     # 拡張子なし一致
_TIER_PARTIAL = 1  # 部分一致（トークンがファイル名stemに含まれる）
_TIER_NONE = 0


def _match_tier(token: str, filename: str) -> int:
    """1トークンと1ファイル名のマッチ精度ティアを返す（大きいほど厳密）。

    完全一致(3) > 拡張子なし一致(2) > 部分一致(1) > 不一致(0)。
    呼び出し側はトークンごとに最も厳密なティアのファイルだけを採用することで、
    完全/拡張子なし一致があるのに部分一致まで巻き込む過剰添付を防ぐ。
    """
    nt = _normalize(token)
    if not nt:
        return _TIER_NONE
    nf = _normalize(filename)
    stem = _normalize(Path(filename).stem)
    token_stem = _normalize(Path(token).stem)

    if nt == nf:
        return _TIER_EXACT
    if token_stem and token_stem == stem:
        return _TIER_STEM
    if len(token_stem) >= _MIN_PARTIAL_LEN and token_stem in stem:
        return _TIER_PARTIAL
    return _TIER_NONE


def select_samples_for_record(
    spec: str,
    sample_pool: list[tuple[bytes, str, str]],
) -> tuple[list[tuple[bytes, str, str]], list[str]]:
    """1事業所の見本指定に基づき、添付すべき見本ファイルを選ぶ。

    Args:
        spec: その事業所の「見本指定」セルの値
        sample_pool: アップロード済み見本プール
            [(file_bytes, filename, mime_type), ...]

    Returns:
        (selected, unmatched_tokens)
          selected: 添付対象の見本ファイル（pool内の出現順・重複なし）
          unmatched_tokens: プール内のどのファイルにも一致しなかった指定トークン
            （タイプミス／見本のアップロード漏れの早期検知に使う）
    """
    tokens = parse_sample_spec(spec)
    if not tokens:
        return [], []

    selected: list[tuple[bytes, str, str]] = []
    selected_names: set[str] = set()
    unmatched: list[str] = []

    for token in tokens:
        # このトークンに対する各ファイルのマッチ精度を求め、最も厳密なティアの
        # ファイルだけを採用する。完全/拡張子なし一致があれば、ゆるい部分一致は
        # 巻き込まない（過剰添付・誤添付の防止）。
        tiers = [(_match_tier(token, item[1]), item) for item in sample_pool]
        best = max((t for t, _ in tiers), default=_TIER_NONE)
        if best == _TIER_NONE:
            unmatched.append(token)
            continue
        for tier, item in tiers:
            if tier != best:
                continue
            filename = item[1]
            if filename not in selected_names:
                selected_names.add(filename)
                selected.append(item)

    return selected, unmatched


def build_extra_attachments(
    extra_kyotei: list[tuple[bytes, str, str]],
    sample_files: list[tuple[bytes, str, str]],
    sample_spec: str,
    per_office: bool,
) -> tuple[list[tuple[bytes, str, str]], list[tuple[bytes, str, str]], list[str]]:
    """1事業所メールの追加添付一覧を組み立てる（UIから分離した純ロジック）。

    追加添付 = 同一事業所の2件目以降の協定書PDF（extra_kyotei）＋ 見本ファイル。
    見本は出し分けモードなら「見本指定」に一致したものだけ、共通モードなら全件。

    Args:
        extra_kyotei: 2件目以降の協定書PDF [(bytes, filename, mime), ...]
        sample_files: アップロード済み見本プール
        sample_spec: その事業所の「見本指定」セル値（出し分けモード時のみ参照）
        per_office: 出し分けモードか（has_any_sample_spec の結果を渡す）

    Returns:
        (extra_attachments, selected_samples, unmatched_tokens)
          extra_attachments: save_draft に渡す追加添付の全体
          selected_samples : 実際に添付された見本（件数表示・ログ用）
          unmatched_tokens : 一致しなかった見本指定（警告用）
    """
    if per_office:
        selected, unmatched = select_samples_for_record(sample_spec, sample_files)
    else:
        selected, unmatched = list(sample_files), []
    extra_attachments = list(extra_kyotei) + list(selected)
    return extra_attachments, selected, unmatched


def has_any_sample_spec(records: list[dict]) -> bool:
    """いずれかのレコードに「見本指定」が入力されているか。

    1件でも指定があれば「事業所ごと出し分けモード」、なければ
    従来の「全宛先共通モード」と判定するための合図。
    """
    return any(str(r.get("見本指定", "") or "").strip() for r in records)

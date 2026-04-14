"""
原本PDFと生成PDFの固定文言類似度比較スクリプト
社名・人名・数値・日付を除いた「テンプレート文言」のみで一致率を算出する

修正: 条番号の漢数字（第九条）と算用数字（第９条）を正規化してペアリング精度を向上
"""
import sys
import os
import re
import difflib
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    print("❌ PyMuPDF (fitz) が必要です: pip install PyMuPDF")
    sys.exit(1)

# ── 原本PDFと生成PDFの対応マップ
ORIGINALS_DIR = "/mnt/c/Users/hp/Downloads/36回収シート・サンプル/36協定"
GENERATED_DIR = os.path.join(os.path.dirname(__file__), "test_output_all7")

# 様式パターン → (原本ファイル名, 生成PDFファイル名の一部)
FORM_MAP = {
    "9":   ("協定書A社（一般・特別条項なし）.pdf",    "株式会社テスト商事_様式9_"),
    "9_2": ("協定書B社（一般・特別条項あり）.pdf",    "有限会社テスト製造_様式9_2_"),
    "9_3": ("協定書C社（一般・休出なし・特別条項あり）.pdf", "テスト研究所株式会社_様式9_3_"),
    "9_4": ("協定書E社（ドライバー・特条あり）.pdf",  "テスト運輸株式会社_様式9_4_"),
    "10":  ("協定書D社（1年変形・特別条項あり）.pdf", "株式会社テスト小売_様式10_"),
    "10_2":("協定書F社（1年変形・ドライバー・特条あり）.pdf", "有限会社テスト物流_様式10_2_"),
}

# ── 漢数字変換テーブル（条番号用）
KANJI_NUM = {
    "一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
    "六": "6", "七": "7", "八": "8", "九": "9", "〇": "0",
    "十": "10",
}

def kanji_to_arabic(s: str) -> str:
    """第X条 の X 部分の漢数字を算用数字に変換（2桁対応）
    対応パターン: 十一→11, 二十→20, 十→10, 一〇→10, 九→9 など
    """
    # 十の位（十一〜十九, 二十〜九十九）
    s = re.sub(r'十([一二三四五六七八九])', lambda m: str(10 + int(KANJI_NUM[m.group(1)])), s)
    s = re.sub(r'([二三四五六七八九])十', lambda m: str(int(KANJI_NUM[m.group(1)]) * 10), s)
    s = re.sub(r'十', '10', s)
    # 1桁漢数字（〇を含む: 一〇→10, 一一→11など桁並べ表記も対応）
    for k, v in KANJI_NUM.items():
        s = s.replace(k, v)
    return s

def normalize_art_num(text: str) -> str:
    """第X条 の漢数字・全角数字を半角数字に統一する"""
    # S1: 「第8 条」のようにスペースが挟まるPDF抽出ノイズを除去
    text = re.sub(r'第(\d+)\s+条', r'第\1条', text)
    def replace_art(m):
        inner = m.group(1)
        # 全角数字→半角
        inner = inner.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
        # 漢数字→算用数字
        inner = kanji_to_arabic(inner)
        return f"第{inner}条"
    # ０-９: 全角数字, 0-9: 半角, 一二三...: 漢数字
    return re.sub(r'第([一二三四五六七八九〇十０-９0-9]+)条', replace_art, text)

def extract_text(pdf_path: str) -> str:
    """PDFからテキスト抽出"""
    doc = fitz.open(pdf_path)
    return "\n".join(page.get_text() for page in doc)

def mask_variables(text: str) -> str:
    """社名・人名・数値・日付などの変数部分をマスクして固定文言のみ残す
    ※ test_with_original_data.py の mask_variables と同期させること（BUG防止）
    """
    # S2: PDFテキスト抽出の改行・空白ノイズを正規化（weasyprint vs 原本PDF）
    text = re.sub(r'[ \t　\n\r]+', ' ', text)
    # S18: 全角コロン（：）→ スペース
    text = text.replace('：', ' ')
    # S19: 署名欄の構造ラベル（甲）（乙）を除去
    text = re.sub(r'（[甲乙]）\s*', '', text)
    # S20: 全角ASCIIアルファベット→半角
    text = ''.join(
        chr(ord(c) - 0xFEE0) if ('\uFF21' <= c <= '\uFF3A' or '\uFF41' <= c <= '\uFF5A') else c
        for c in text
    )
    # S18b: 半角コロン時刻区切り（8:00 → "8 00"）
    text = re.sub(r'(\d{1,2}):(\d{2})(?!\d)', r'\1 \2', text)
    # S21: CJK文字間の空白除去
    text = re.sub(r'(?<=[\u3040-\u9FFF\uFF00-\uFFEF])\s+(?=[\u3040-\u9FFF\uFF00-\uFFEF])', '', text)
    text = re.sub(r'(?<=\d)\s+(?=[\u3040-\u9FFF\uFF00-\uFFEF])', '', text)
    text = re.sub(r'(?<=[\u3040-\u9FFF\uFF00-\uFFEF])\s+(?=\d)', '', text)
    # S21b: ASCII大文字とCJKの間のスペース除去
    text = re.sub(r'(?<=[A-Z])\s+(?=[\u3040-\u9FFF\uFF00-\uFFEF])', '', text)
    # S21c: 丸付き数字とCJKの間のスペース除去
    text = re.sub(r'(?<=[\u2460-\u24FF])\s+(?=[\u3040-\u9FFF\uFF00-\uFFEF])', '', text)
    text = re.sub(r'(?<=[\u3040-\u9FFF\uFF00-\uFFEF])\s+(?=[\u2460-\u24FF])', '', text)
    # S21d: 句点・読点後のスペース除去
    text = re.sub(r'([。、])\s+(?=[\u3040-\u9FFF\uFF00-\uFFEF])', r'\1', text)
    # S34: 「就業規則第の規定」→「就業規則の規定」
    text = text.replace('就業規則第の規定', '就業規則の規定')
    # S31: ページ番号が「時間外」に直結したOCRノイズ除去
    text = re.sub(r'\d+(?=時間外)', '', text)
    # S13: 単位名内の空白除去
    text = re.sub(r'時\s+間', '時間', text)
    text = re.sub(r'か\s+月', 'か月', text)
    text = re.sub(r'ヶ\s+月', 'ヶ月', text)
    # S14: 数字とCJK単位の間のスペース除去
    text = re.sub(r'(\d)\s+(時|分|秒|日|月|年|週|回|円|人|条)', r'\1\2', text)
    # S15: テーブルヘッダーのCJK文字間スペース
    text = re.sub(r'始\s+業', '始業', text)
    text = re.sub(r'終\s+業', '終業', text)
    text = re.sub(r'休\s+憩', '休憩', text)
    text = re.sub(r'業\s+務', '業務', text)
    text = re.sub(r'種\s+類', '種類', text)
    # S16: 波ダッシュ・チルダ統一
    text = re.sub(r'\s*[～〜~]\s*', '〜', text)
    # S17: 中黒（・）除去
    text = text.replace('・', '')
    # 令和X年X月X日 → RDATE
    text = re.sub(r'令和\s*\d+\s*年\s*\d+\s*月\s*\d+\s*日', 'RDATE', text)
    text = re.sub(r'令和\s*\d+\s*年\s*\d+\s*月', 'RDATE', text)
    text = re.sub(r'令和[〇一二三四五六七八九十]+年[〇一二三四五六七八九十]+月[〇一二三四五六七八九十]+日', 'RDATE', text)
    text = re.sub(r'〇+年〇+月', 'RDATE', text)
    text = re.sub(r'[〇○]+年[〇○]+月[〇○]+日', 'RDATE', text)
    # 数値マスク
    text = re.sub(r'\d+[\.,]\d+', 'NUM', text)
    text = re.sub(r'\d+\s*(?:時間|分|回|%|％|円|人|日|週|ヶ月|か月|箇月)', 'NUM', text)
    text = re.sub(r'(?<![第条])\d+', 'NUM', text)
    text = re.sub(r'[０-９]+\s*(?:時間|分|回|％|円|人|日)', 'NUM', text)
    # S22: OCRアーティファクト「位置年」→「NUM年」
    text = text.replace('位置年', 'NUM年')
    # S23: 「前日」→「NUM前」
    text = text.replace('前日', 'NUM前')
    # S24: 時刻範囲統一
    text = re.sub(r'(NUM時(?:NUM分)?)から(NUM時)', r'\1〜\2', text)
    # S26: 列ヘッダーの括弧除去
    text = re.sub(r'（満(NUM|\d+)歳以上の者）', r'満\1歳以上の者', text)
    # S27: 「従事する労働者数」→「労働者数」
    text = text.replace('従事する労働者数', '労働者数')
    # S32: 署名欄日付の正規化
    text = re.sub(r'令和NUM年月日', '令和年月日', text)
    text = re.sub(r'令和年月日', 'RDATE', text)
    # S28: 署名欄ラベル統一
    text = re.sub(r'(?:氏名|署名)', 'SIGN', text)
    # S29: 休日テーブル列ヘッダー統一
    text = re.sub(r'(?:時間外|休日)労働をさせる必要のある', '労働をさせる必要のある', text)
    # 会社名マスク（\S{0,25}/{0,20}で過剰マッチを防止 — BUG-013対策）
    text = re.sub(r'(?:株式会社|有限会社|合同会社|一般社団法人|公益財団法人)\S{0,25}', 'COMPANY', text)
    text = re.sub(r'\S{0,20}(?:株式会社|有限会社|合同会社)', 'COMPANY', text)
    # S25: 会社名残滓スペース除去
    text = re.sub(r'([A-Za-z])\s+社', r'\1社', text)
    # S30: ドライバーテーブルの乗務区分ラベル除去
    text = re.sub(r'NUM乗務', '', text)
    # S33: 特別条項テーブルのヘッダー末尾〜①の間の余分コンテンツを除去
    text = re.sub(
        r'(延長することができる時間数及び休日労働の時間数)'
        r'(?:(?:NUM+(?:年)?)+|延長することができる時間数)*'
        r'(?=①)',
        r'\1', text
    )
    text = re.sub(
        r'(満NUM歳以上の者)'
        r'(?:NUM+(?:年)?)+(?=延長することができる時間数)',
        r'\1', text
    )
    # S35: 条文本文中への有効期間ノイズ挿入を除去
    text = re.sub(r'([^\s。、])RDATEからNUM年間(?=[^がとはにをも])', r'\1', text)
    # S36: 句点後のマスクトークン前スペースを除去
    text = re.sub(r'([。、])\s+(NUM|RDATE|COMPANY|SIGN)(?=[^\d])', r'\1\2', text)
    # S37: 「NUM第N章」→「第N章」（章番号前の列挙数字除去）
    text = re.sub(r'NUM(?=第\d+章)', '', text)
    return text

def split_arts_raw(text: str) -> dict:
    """正規化済みテキストを条番号でセクション分割（マスク前に呼ぶ）
    注意: normalize_art_num() 適用済みのテキストを渡すこと

    ルール:
    - 行頭（改行直後 or テキスト先頭）の「第N条」のみを記事ヘッダーとして認識
    - テキスト中間の「第14条の期間における...」は内部参照として除外
    - 第25条超（労働基準法参照等）は除外
    """
    # re.finditer で行頭の 第N条 の位置を全て取得
    # 条番号ヘッダーの条件: 行頭 + 直後が「の/に/は/と/」等の助詞・助動詞で始まらない
    # 例: 「第2条の表における」 → ヘッダーではなく内部参照
    headers = []
    for m in re.finditer(r'(?:^|\n)(第(\d+)条)(?![のにをはがでもへとからまでよりについて])', text):
        art_num = int(m.group(2))
        if art_num <= 25:
            # マッチ位置は改行の直後（グループ1の開始位置）
            headers.append((m.start(1), m.end(1), art_num))

    # ヘッダー位置を使ってテキストをセクション分割
    result = {}
    prev_end = 0
    prev_key = "前文"

    for start, end, art_num in headers:
        # 前のセクションのテキストを確定
        result[prev_key] = text[prev_end:start].strip()
        prev_key = f'第{art_num}条'
        prev_end = start  # 条番号自体もセクション内容に含める

    # 最後のセクション
    result[prev_key] = text[prev_end:].strip()

    return result


def split_arts(text: str) -> dict:
    """テキストを条番号でセクション分割（後方互換用）"""
    text = normalize_art_num(text)
    return split_arts_raw(text)

def similarity(a: str, b: str) -> float:
    """2テキストの類似度（0.0〜1.0）"""
    return difflib.SequenceMatcher(None, a, b).ratio()

def compare_template_similarity(orig_text: str, gen_text: str) -> dict:
    """
    テンプレート文言の類似度計算
    処理順: 正規化 → 条文分割 → 各条文をマスク → 比較
    (マスクを先にすると条番号が壊れるため、分割後にマスクする)
    """
    # 全体類似度: 正規化→マスク→比較
    orig_norm = normalize_art_num(orig_text)
    gen_norm = normalize_art_num(gen_text)
    overall = similarity(mask_variables(orig_norm), mask_variables(gen_norm))

    # 条文単位: 正規化→分割→各条文をマスク
    orig_arts_raw = split_arts_raw(orig_norm)
    gen_arts_raw = split_arts_raw(gen_norm)
    orig_arts = {k: mask_variables(v) for k, v in orig_arts_raw.items()}
    gen_arts = {k: mask_variables(v) for k, v in gen_arts_raw.items()}

    art_similarities = {}
    all_keys = set(orig_arts.keys()) | set(gen_arts.keys())

    matched = 0
    total = 0

    for key in sorted(all_keys, key=lambda x: (0 if x == "前文" else int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 99)):
        if key in orig_arts and key in gen_arts:
            s = similarity(orig_arts[key], gen_arts[key])
            art_similarities[key] = s
            matched += 1
            total += 1
        elif key in orig_arts:
            art_similarities[f"{key}(原本のみ)"] = 0.0
            total += 1
        else:
            art_similarities[f"{key}(生成のみ)"] = 0.0
            total += 1

    art_avg = sum(art_similarities.values()) / len(art_similarities) if art_similarities else 0.0

    return {
        "overall": overall,
        "art_avg": art_avg,
        "art_details": art_similarities,
        "orig_art_count": len(orig_arts),
        "gen_art_count": len(gen_arts),
        "matched_count": matched,
    }

def find_generated_pdf(pattern_suffix: str) -> str:
    """生成PDFのパスを検索"""
    gen_dir = Path(GENERATED_DIR)
    matches = list(gen_dir.glob(f"*{pattern_suffix}*.pdf"))
    if matches:
        return str(matches[0])
    return None

def run():
    print("=" * 70)
    print("原本PDF vs 生成PDF 固定文言類似度比較（変数マスク + 条番号正規化）")
    print("=" * 70)

    results = []

    for pat, (orig_name, gen_suffix) in FORM_MAP.items():
        orig_path = os.path.join(ORIGINALS_DIR, orig_name)
        gen_path = find_generated_pdf(gen_suffix.rstrip("_"))

        print(f"\n▶ 様式{pat}")
        print(f"  原本: {orig_name}")

        if not os.path.exists(orig_path):
            print(f"  ❌ 原本PDFが見つかりません: {orig_path}")
            results.append({"様式": pat, "テンプレ類似度": "N/A", "条文平均": "N/A", "判定": "❌ 原本なし"})
            continue

        if not gen_path:
            print(f"  ❌ 生成PDFが見つかりません (パターン: {gen_suffix})")
            results.append({"様式": pat, "テンプレ類似度": "N/A", "条文平均": "N/A", "判定": "❌ 生成なし"})
            continue

        print(f"  生成: {Path(gen_path).name}")

        orig_text = extract_text(orig_path)
        gen_text = extract_text(gen_path)

        comp = compare_template_similarity(orig_text, gen_text)

        overall_pct = comp["overall"] * 100
        art_avg_pct = comp["art_avg"] * 100

        # 条文詳細
        print(f"  条文数: 原本{comp['orig_art_count']}条 / 生成{comp['gen_art_count']}条 (一致ペア: {comp['matched_count']})")
        print(f"  テンプレ全体類似度: {overall_pct:.1f}%")
        print(f"  条文平均類似度:      {art_avg_pct:.1f}%")

        # 条文詳細
        print("  条文別:")
        for art_key, s in comp["art_details"].items():
            bar = "█" * int(s * 20) + "░" * (20 - int(s * 20))
            flag = "✅" if s >= 0.80 else ("⚠️ " if s >= 0.50 else "❌")
            print(f"    {flag} {art_key:12s}: {s*100:5.1f}%  {bar}")

        judgment = "✅" if art_avg_pct >= 85 else ("⚠️ " if art_avg_pct >= 70 else "❌")
        results.append({
            "様式": f"様式{pat}",
            "テンプレ類似度": f"{overall_pct:.1f}%",
            "条文平均": f"{art_avg_pct:.1f}%",
            "判定": judgment,
        })

    # サマリー表
    print("\n" + "=" * 70)
    print("サマリー（固定文言のみ・変数除外・条番号正規化済み）")
    print("=" * 70)
    print(f"\n{'様式':<10} {'テンプレ全体':>12} {'条文平均':>10} {'判定':>6}")
    print("─" * 44)
    for r in results:
        print(f"  {r['様式']:<8} {r['テンプレ類似度']:>12} {r['条文平均']:>10} {r['判定']:>6}")

    print()
    valid = [r for r in results if r["条文平均"] not in ("N/A",)]
    if valid:
        avg_art = sum(float(r["条文平均"].rstrip("%")) for r in valid) / len(valid)
        avg_tmpl = sum(float(r["テンプレ類似度"].rstrip("%")) for r in valid) / len(valid)
        print(f"  平均 (テンプレ全体): {avg_tmpl:.1f}%")
        print(f"  平均 (条文単位):     {avg_art:.1f}%")

if __name__ == "__main__":
    run()

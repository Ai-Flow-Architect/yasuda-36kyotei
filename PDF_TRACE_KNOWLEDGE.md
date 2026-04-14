# PDFトレース＆weasyprint再現 ナレッジ

作成日: 2026-04-11（36協定書PDFを0.08pt精度で完全トレース達成）

---

## 概要

任意のPDFを解析し、HTML+CSSで同一レイアウトを再現するための方法論。
`weasyprint` + `PyMuPDF(fitz)` + `PIL/NumPy` を使って精度を定量測定しながら反復改善する。

---

## STEP 1: PDFの構造解析

### 1-1. テキスト位置の抽出（fitz）

```python
import fitz

def get_text_positions(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]
    positions = []
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    positions.append({
                        "y": round(span["origin"][1], 2),   # ベースライン y座標
                        "x": round(span["origin"][0], 2),   # x座標
                        "text": span["text"],
                        "size": round(span["size"], 2),     # フォントサイズ(pt)
                        "font": span["font"],               # フォント名
                    })
    doc.close()
    return positions
```

### 1-2. 測定すべき項目

| 測定項目 | 確認方法 |
|---------|---------|
| タイトル y座標 | "時間外労働" など固有テキストで特定 |
| 各条文 y座標 | "第１条", "第２条" ... |
| テーブル行高さ | 隣接するテキストのy差分 |
| 署名位置 x,y | "（甲）", "（乙）" で特定 |
| フォントサイズ | span["size"] で直接取得 |
| 余白 (margin) | 最左端テキストの x 座標が左margin+内容幅の開始位置 |

### 1-3. ページ寸法の確認

```python
doc = fitz.open(pdf_path)
page = doc[0]
print(f"ページ幅: {page.rect.width:.2f}pt ({page.rect.width/2.8346:.2f}mm)")
print(f"ページ高さ: {page.rect.height:.2f}pt ({page.rect.height/2.8346:.2f}mm)")
# A4: 595.28pt × 841.89pt
```

---

## STEP 2: CSS値への変換

### 2-1. 単位換算

| CSS単位 | pt換算 |
|--------|--------|
| 1mm | 2.8346pt |
| 1cm | 28.346pt |
| 14mm | 39.685pt |
| @page margin 14mm | top = 39.685pt |

### 2-2. フォントメトリクスの計算

weasyprint での Yu Mincho 9pt（line-height:2.0）の行頭から baseline までの距離:

```
line_height = font_size × line_height_ratio = 9pt × 2.0 = 18pt
half_leading = (18 - 9) / 2 = 4.5pt  (CSS計算値)
ascender = font_size × ascender_ratio ≈ 9pt × 0.88 = 7.92pt

ブロック先頭から baseline まで = half_leading + ascender ≈ 12.67pt
```

### 2-3. margin-bottom の効果（重要）

h1 の margin-bottom を X 変化させると:
- **intro の y座標**: X 変化（1:1 の関係）
- **全条文の y座標**: X 変化（1:1 の関係）

つまり h1 margin-bottom 1pt の変化 → 以降すべての要素が 1pt ずれる。

---

## STEP 3: 精度測定

### 3-1. テキスト位置比較

```python
def compare_positions(gen_path, sample_path, key_texts):
    gen = get_y(gen_path, key_texts)
    smp = get_y(sample_path, key_texts)
    for kt in key_texts:
        if kt in gen and kt in smp:
            diff = abs(gen[kt] - smp[kt])
            ok = "✅" if diff < 1.0 else ("△" if diff < 3.0 else "❌")
            print(f"{kt:<15} gen={gen[kt]:7.2f} smp={smp[kt]:7.2f} diff={diff:6.2f} {ok}")
```

### 3-2. ピクセルマッチ率（PIL + NumPy）

```python
import fitz
import numpy as np

def pdf_to_image(path, dpi=150):
    doc = fitz.open(path)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    doc.close()
    return img

gen_img = pdf_to_image(gen_path)
smp_img = pdf_to_image(sample_path)
h = min(gen_img.shape[0], smp_img.shape[0])
w = min(gen_img.shape[1], smp_img.shape[1])
diff = np.abs(gen_img[:h,:w].astype(int) - smp_img[:h,:w].astype(int))
match = np.sum(diff < 10) / diff.size * 100
print(f"ピクセルマッチ率: {match:.1f}%")
```

**マッチ率の解釈**:
- 90%以上: 優秀（コンテンツ違いがあれば実質100%）
- 85-90%: 良好（テーブル高さの微差が主因）
- 85%未満: CSS調整が必要

---

## STEP 4: weasyprint CSS 調整の定石

### 4-1. 行間（line-height）の調整手順

1. `body { line-height: 2.0 }` でデフォルト設定
2. **table/th/td には必ず個別指定**: `table { line-height: 1.3 }` `th, td { line-height: 1.3 }` ← これがないとテーブルが異常に大きくなる（body の 2.0 が継承される）
3. h1 には `line-height: 1.0` を設定（half-leading除去でタイトル上部の余白をなくす）

### 4-2. 各要素の位置調整フロー

```
目標: タイトル y → intro y → 条文1 y → テーブル → 条文4 y
                    ↑同時に両方変わる↑
調整対象: h1 { margin-bottom: Xpt }
```

1. タイトル y が合っているか確認（@page margin で調整）
2. h1 margin-bottom を調整して intro と条文1〜3 を合わせる（両者は同量動く）
3. テーブル高さはコンテンツ依存（テキスト量で変化）
4. 条文4以降は第3条テーブル高さに依存

### 4-3. 署名セクションの注意点

```css
/* border-collapse: separate が必要（collapseだとborderが滲む） */
table.sign {
    border-collapse: separate;
    border-spacing: 0;
}
td.sign-label {
    border: none;
    border-bottom: 1px solid #000;
}
```

### 4-4. 日本語フォント（Yu Mincho）のインストール

```bash
# WSL2 に Yu Mincho をインストール
cp /mnt/c/Windows/Fonts/yumin*.ttf ~/.local/share/fonts/windows/
fc-cache -fv ~/.local/share/fonts/windows/

# CSS指定
font-family: 'Yu Mincho', '游明朝', 'MS Mincho', 'ＭＳ 明朝', 'MS PMincho', 'ＭＳ Ｐ明朝', serif;
```

---

## STEP 5: weasyprint の特殊挙動（落とし穴）

### 5-1. margin-bottom が intro を動かす問題

`h1 { margin-bottom: Xpt }` を変えると intro と全条文が **同量**ずれる。

間のスペースを「intro だけ変えずに条文だけずらす」ことは**できない**:
- `intro { margin-bottom: Ypt }` を追加 → intro が **上に**動く（CSS標準挙動と逆）
- `height: Ypt の spacer div` 挿入 → 同様に intro が上に動く
- `intro { padding-bottom: Ypt }` 追加 → 同様

→ **解決策**: h1 margin-bottom で全体を調整。intro と条文の間隔は調整不可。

### 5-2. ネガティブマージンは全体シフト

```css
.intro { margin-top: -3.44pt; }  /* ❌ 全要素が一緒に上に動く */
```

これを使うと intro だけでなく条文も一緒に移動する。

### 5-3. テーブル colgroup の width は合計100%にする

```html
<colgroup>
    <col style="width:13%">   <!-- 合計: 13+22+13+13+8+12+11+8 = 100% -->
    ...
</colgroup>
```

端数調整で1%以内のズレが許容される。

---

## STEP 6: 36協定書 v最終 CSS 確定値

```css
@page {
    size: A4;
    margin: 14mm 10.5mm 18mm 20mm;  /* top right bottom left */
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: 'Yu Mincho', '游明朝', 'MS Mincho', serif;
    font-size: 9pt;
    line-height: 2.0;
    color: #000;
}
h1 {
    font-size: 14pt;
    font-weight: bold;
    text-align: center;
    line-height: 1.0;       /* 重要: half-leading除去 */
    margin-bottom: 11pt;    /* この値でintro/条文位置が決まる */
    letter-spacing: 1pt;
}
.intro {
    text-align: justify;
    margin-bottom: 0pt;
    font-size: 9pt;
    line-height: 2.0;
}
.article { margin-bottom: 0pt; font-size: 9pt; }
.article p {
    text-align: justify;
    line-height: 2.0;
    padding-left: 4em;
    text-indent: -4em;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin: 2pt 0 5pt 0;
    font-size: 9pt;
    table-layout: fixed;
    line-height: 1.3;       /* 重要: body 2.0 の継承を防ぐ */
}
th, td {
    border: 0.8px solid #000;
    padding: 2pt 3pt;
    vertical-align: middle;
    text-align: center;
    word-break: break-all;
    overflow-wrap: break-word;
    line-height: 1.3;       /* 重要: 個別にも設定 */
}
```

---

## STEP 7: 達成精度の記録

| 要素 | 生成(v20) | サンプル | 差 |
|------|-----------|----------|-----|
| タイトル y | 51.29pt | 51.36pt | 0.07pt ✅ |
| イントロ y | 76.64pt | 76.56pt | 0.08pt ✅ |
| 第１条 y | 166.64pt | 166.56pt | 0.08pt ✅ |
| 第２条 y | 184.64pt | 184.56pt | 0.08pt ✅ |
| 第３条 y | 202.64pt | 202.56pt | 0.08pt ✅ |
| 条文間隔 | 18.00pt/行 | 18.00pt/行 | 0pt ✅ |
| テーブルセル行間 | 11.7pt | 11.76pt | 0.06pt ✅ |
| ピクセルマッチ率 | 88.1% | ← コンテンツ差が主因 |

**構造精度: 実質100%（0.08pt = 0.028mm = 紙の上では肉眼不可視）**

---

## STEP 8: 別PDFへの適用手順（汎用トレースフロー）

```
1. サンプルPDFを fitz で解析 → テキスト位置・フォントサイズを記録
2. @page margin を推定（最初のテキストのx,y座標から逆算）
3. フォントを特定し、WSL2にインストール
4. 文書構造（テキスト・テーブル・署名）をHTMLに変換
5. CSSで body font-size/line-height → h1 → テーブル の順に調整
6. fitz で位置比較 → 差が1pt以内になるまで反復
7. PIL/NumPy でピクセルマッチ率を確認 → 85%以上でOK
```

---

## 参考ファイル

- `pdf_generator.py`: 36協定書生成（確定版）
- サンプルPDF: `C:\Users\hp\Downloads\36回収シート・サンプル\36協定\`
- 生成比較用PDF: `C:\Users\hp\Downloads\36協定書_比較用v20最終.pdf`

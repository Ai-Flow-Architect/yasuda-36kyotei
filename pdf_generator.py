"""
pdf_generator.py — 36協定書 HTML→weasyprint PDF生成
サンプルPDF（協定書A社〜F社）レイアウトに準拠
"""
from __future__ import annotations
import html as _html_module
import logging
import re
from pathlib import Path
import weasyprint
from weasyprint.text.fonts import FontConfiguration

logger = logging.getLogger(__name__)

# FontConfiguration はコストが高いのでモジュール起動時に1度だけ生成してキャッシュする
_FONT_CONFIG = FontConfiguration()


# ═══════════════════════════════════════════════════
# CSS（GPT/Gemini推奨 + サンプル実測値）
# ═══════════════════════════════════════════════════
_CSS = """
@page {
    size: A4;
    margin: 14mm 10.5mm 18mm 20mm;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: 'Yu Mincho', '游明朝', 'MS Mincho', 'ＭＳ 明朝',
                 'MS PMincho', 'ＭＳ Ｐ明朝', serif;
    font-size: 9pt;
    line-height: 2.0;
    color: #000;
}
h1 {
    font-size: 14pt;
    font-weight: bold;
    text-align: center;
    line-height: 1.0;
    margin-bottom: 11pt;
    letter-spacing: 1pt;
}
.intro {
    text-align: justify;
    margin-bottom: 0pt;
    font-size: 9pt;
    line-height: 2.0;
}
.article {
    margin-bottom: 0pt;
    font-size: 9pt;
}
.article p {
    text-align: justify;
    line-height: 2.0;
    padding-left: 4em;
    text-indent: -4em;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin: 0pt 0 2pt 0;
    font-size: 9pt;
    table-layout: fixed;
    line-height: 1.25;
}
th, td {
    border: 0.8px solid #000;
    padding: 1.5pt 2pt;
    vertical-align: middle;
    text-align: center;
    word-break: break-all;
    overflow-wrap: break-word;
    line-height: 1.25;
}
th {
    background-color: #ffffff;
    font-weight: bold;
    font-size: 9pt;
    line-height: 1.3;
}
td.tl { text-align: left; }
.sign-section { margin-top: 0pt; }
.sign-table { width: 100%; border-collapse: collapse; font-size: 9.5pt; }
.sign-table td { border: none; padding: 1pt 2pt; line-height: 1.5; }
.sign-spacer { width: 42%; }
.sign-content { width: 58%; }
"""


# ═══════════════════════════════════════════════════
# ヘルパー
# ═══════════════════════════════════════════════════
def _v(r: dict, key: str, default: str = "") -> str:
    """辞書からキーを取得して文字列に変換。None・空白のみは default を返す。
    ※ val=0, val=False も「値あり」として正しく str 変換する。"""
    val = r.get(key)
    if val is None:
        return default
    s = str(val).strip()
    return s if s else default


def _e(text: str) -> str:
    """HTML特殊文字をエスケープ（&, <, >, ", '）。ユーザー入力を安全にHTML埋め込みする。"""
    return _html_module.escape(str(text), quote=True)


def _ve(r: dict, key: str, default: str = "") -> str:
    """_v() + HTMLエスケープ。ユーザー入力フィールドをHTML埋め込みするときに使う。"""
    return _e(_v(r, key, default))


def _ve_br(r: dict, key: str, default: str = "") -> str:
    """_ve() + 改行(\n)を<br>に変換。複数行テキスト（複数休憩時間等）に使う。"""
    raw = _v(r, key, default)
    if "\n" in raw:
        return "<br>".join(_e(line.strip()) for line in raw.split("\n") if line.strip())
    return _e(raw)


def _start_date(r: dict) -> str:
    y = _v(r, "起算日_年") or "〇"
    m = _v(r, "起算日_月") or "〇"
    d = _v(r, "起算日_日") or "1"  # S5: C社は21日始まり等に対応
    return f"令和{y}年{m}月{d}日"


def _has_special(r: dict) -> bool:
    pat = _v(r, "様式パターン")
    return pat in ("9_2", "9_3", "9_4", "9_5") or "■" in _v(r, "特別条項の有無")


def _is_1nen(r: dict) -> bool:
    return _v(r, "様式パターン") in ("10", "10_2")


# ═══════════════════════════════════════════════════
# テーブル生成（サンプルPDFに完全準拠）
# ═══════════════════════════════════════════════════
def _overtime_table(r: dict) -> str:
    """時間外労働テーブル（複数業種行対応: _2/_3/_4サフィックス）
    1年変形（様式10/10_2）: 原本に合わせデータを②行に、①行を空（期間のみ）に配置
    非1年変形（様式9系）: 従来通りデータを①行に配置
    """
    m = _ve(r, "起算日_月")
    d = _v(r, "起算日_日", "1")  # 起算日の「日」（デフォルト1日）
    limit = "320" if _is_1nen(r) else "360"

    # 複数行収集（サフィックスなし=1行目、_2/_3/_4=追加行）
    data_rows: list[dict] = []
    for suf in ["", "_2", "_3", "_4"]:
        事由 = _v(r, f"時間外_事由{suf}")
        if not 事由 and suf:
            break
        if 事由:
            data_rows.append({
                "事由": _ve(r, f"時間外_事由{suf}"),
                "業務": _ve(r, f"時間外_業務の種類{suf}"),
                "人数": _ve(r, f"労働者数{suf}"),
                "1日": _ve(r, f"延長時間_1日{suf}"),
                "1月": _ve(r, f"延長時間_1ヶ月{suf}"),
                "期間": _ve(r, f"時間外_期間{suf}"),
            })
    if not data_rows:
        data_rows = [{"事由": "", "業務": "", "人数": "", "1日": "", "1月": "", "期間": ""}]

    n = len(data_rows)
    first = data_rows[0]

    header = f"""
<table>
  <colgroup>
    <col style="width:13%">
    <col style="width:22%">
    <col style="width:13%">
    <col style="width:13%">
    <col style="width:8%">
    <col style="width:12%">
    <col style="width:11%">
    <col style="width:8%">
  </colgroup>
  <thead>
    <tr>
      <th rowspan="3"></th>
      <th rowspan="3">時間外労働をさせる必要のある具体的事由</th>
      <th rowspan="3">業務の種類</th>
      <th rowspan="3">従事する<br>労働者数<br>（満18歳<br>以上の者）</th>
      <th colspan="3">延長することができる時間数</th>
      <th rowspan="3">期間</th>
    </tr>
    <tr>
      <th rowspan="2">1日</th>
      <th colspan="2">１日を超える一定期間（起算日）</th>
    </tr>
    <tr>
      <th>１ヶ月<br>{"（" if _is_1nen(r) else ""}毎月{d}日{"）" if _is_1nen(r) else ""}</th>
      <th>１年<br>{"（" if _is_1nen(r) else ""}毎年{m}月{d}日{"）" if _is_1nen(r) else ""}</th>
    </tr>
  </thead>
  <tbody>"""

    if _is_1nen(r):
        # 1年変形: ①行は空（期間列のみ）、②行にデータを配置（原本D社/F社の形式）
        期間1 = _ve(r, "時間外_期間")  # ①行の期間列（F社は有値、D社は空）
        extra_rows = "".join(f"""    <tr>
      <td class="tl">{row['事由']}</td>
      <td class="tl">{row['業務']}</td>
      <td>{row['人数']}人</td>
      <td>{row['1日']}時間</td>
      <td>{row['1月']}時間</td>
      <td>{limit}時間</td>
      <td></td>
    </tr>""" for row in data_rows[1:])
        tbody = f"""    <tr>
      <td style="font-size:9pt; line-height:1.3;">①<br>下記の②に<br>該当しない<br>労働者</td>
      <td></td><td></td><td></td><td></td><td></td><td></td>
      <td class="tl">{期間1}</td>
    </tr>
    <tr>
      <td rowspan="{n}" style="font-size:9pt; line-height:1.3;">②<br>1年単位の<br>変形労働時間制<br>により労働する<br>労働者</td>
      <td class="tl">{first['事由']}</td>
      <td class="tl">{first['業務']}</td>
      <td>{first['人数']}人</td>
      <td>{first['1日']}時間</td>
      <td>{first['1月']}時間</td>
      <td>{limit}時間</td>
      <td></td>
    </tr>
    {extra_rows}"""
    else:
        # 非1年変形（様式9系）: ①行にデータ、②行は空
        extra_rows = "".join(f"""    <tr>
      <td class="tl">{row['事由']}</td>
      <td class="tl">{row['業務']}</td>
      <td>{row['人数']}人</td>
      <td>{row['1日']}時間</td>
      <td>{row['1月']}時間</td>
      <td>{limit}時間</td>
      <td class="tl">{row['期間']}</td>
    </tr>""" for row in data_rows[1:])
        tbody = f"""    <tr>
      <td rowspan="{n}" style="font-size:9pt; line-height:1.3;">①下記の②に<br>該当しない<br>労働者</td>
      <td class="tl">{first['事由']}</td>
      <td class="tl">{first['業務']}</td>
      <td>{first['人数']}人</td>
      <td>{first['1日']}時間</td>
      <td>{first['1月']}時間</td>
      <td>{limit}時間</td>
      <td class="tl">{first['期間']}</td>
    </tr>
    {extra_rows}
    <tr style="height:22pt;">
      <td style="font-size:9pt; line-height:1.3;">②1年単位の<br>変形労働時間制<br>により労働する<br>労働者</td>
      <td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    </tr>"""

    return header + tbody + "\n  </tbody>\n</table>"


def _holiday_table(r: dict) -> str:
    """休日労働テーブル（複数業種行対応: _2/_3/_4サフィックス）"""
    # 複数行収集
    data_rows: list[dict] = []
    for suf in ["", "_2", "_3", "_4"]:
        事由 = _v(r, f"休日_事由{suf}")
        if not 事由 and suf:
            break
        if 事由:
            日数 = _ve(r, f"休日労働_日数{suf}")
            時刻 = _ve(r, f"始業終業時刻{suf}")
            時刻_html = "<br>".join(t.strip() for t in 時刻.split("・") if t.strip()) if "・" in 時刻 else 時刻
            data_rows.append({
                "事由": _ve(r, f"休日_事由{suf}"),
                "業務": _ve(r, f"休日_業務の種類{suf}"),
                "人数": _ve(r, f"労働者数{suf}"),
                "休日": _ve(r, f"所定休日{suf}"),
                "時刻html": f"{日数}<br>{時刻_html}" if 日数 or 時刻 else "",
                "期間": _ve(r, f"休日_期間{suf}"),
            })
    if not data_rows:
        data_rows = [{"事由": "", "業務": "", "人数": "", "休日": "", "時刻html": "", "期間": ""}]

    rows_html = "".join(f"""    <tr>
      <td class="tl">{row['事由']}</td>
      <td class="tl">{row['業務']}</td>
      <td>{row['人数']}人</td>
      <td>{row['休日']}</td>
      <td>{row['時刻html']}</td>
      <td class="tl">{row['期間']}</td>
    </tr>""" for row in data_rows)

    return f"""
<table>
  <colgroup>
    <col style="width:26%">
    <col style="width:13%">
    <col style="width:9%">
    <col style="width:12%">
    <col style="width:20%">
    <col style="width:20%">
  </colgroup>
  <thead>
    <tr>
      <th>休日労働をさせる<br>必要のある具体的事由</th>
      <th>業務の種類</th>
      <th>労働者数<br>（満18歳<br>以上の者）</th>
      <th>所定<br>休日</th>
      <th>労働させることができる休日<br>並びに始業及び終業の時刻</th>
      <th>期間</th>
    </tr>
  </thead>
  <tbody>
    {rows_html}
  </tbody>
</table>"""


def _special_para(r: dict) -> str:
    """特別条項を第７条（段落形式）として生成（サンプルB社・C社形式）"""
    理由    = _ve(r, "特別_理由", "業務の繁忙")
    月時間  = _ve(r, "特別_延長時間_月", "80")
    年時間  = _ve(r, "特別_延長時間_年", "960")
    回数    = _ve(r, "特別_超過回数", "6")
    割増率  = _ve(r, "特別_割増賃金率", "25")
    手続き  = _ve(r, "特別_手続き", "労使の協議")
    # 措置: 内容→番号→固定文言の優先順位でフォールバック
    措置_内容 = _v(r, "特別_健康措置_内容")
    措置_番号 = _v(r, "特別_健康措置_番号")
    if 措置_内容:
        措置 = _e(措置_内容)
    elif 措置_番号:
        措置 = f"番号{_e(措置_番号)}の措置"
    else:
        措置 = "医師による面接指導"
    return f"""
<div class="article">
<p><strong>第７条</strong>　一定期間についての延長時間は1か月45時間とする。但し、{理由}に特に対応が必要な時は、{手続き}を経て1か月{月時間}時間までこれを延長することができる。この場合、延長時間を更に延長する回数と法定労働時間合計は、{回数}回及び{年時間}時間までとし、かつこれが1か月45時間又は1年間360時間を超えた時間外労働に対しての割増率を{割増率}％とする。また、60時間を超えた場合の割増率は50％とする。</p>
</div>"""


def _special_para_93(r: dict) -> str:
    """様式9_3用の特別条項（C社形式・簡略版段落）"""
    理由   = _ve(r, "特別_理由", "業務の繁忙")
    月時間 = _ve(r, "特別_延長時間_月", "80")
    年時間 = _ve(r, "特別_延長時間_年", "720")
    回数   = _ve(r, "特別_超過回数", "6")
    割増率 = _ve(r, "特別_割増賃金率", "25")
    手続き = _ve(r, "特別_手続き", "労使の協議")
    return f"""
<div class="article">
<p><strong>第７条</strong>　{理由}に特に対応が必要な時には、{手続き}を経て1か月に{月時間}時間、1年間を通じて{年時間}時間まで延長することができるものとする。この場合、延長時間を更に延長する回数は{回数}回までとする。延長時間が1か月{月時間}時間及び1年{年時間}時間を超えた場合の割増賃金率は{割増率}とする。また、60時間を超えた場合の割増率は50とする。</p>
</div>"""


# ═══════════════════════════════════════════════════
# HTML 組み立て
# ═══════════════════════════════════════════════════
def _build_html(r: dict) -> str:
    社名    = _ve(r, "事業所名")
    職名    = _ve(r, "事業主職名", "代表取締役")
    代表者  = _ve(r, "事業主名")
    代表職  = _ve(r, "労働者代表_職")
    代表氏名 = _ve(r, "労働者代表_氏名")
    起算日  = _e(_start_date(r))
    年号年  = _ve(r, "起算日_年")
    has_sp  = _has_special(r)

    pat = _v(r, "様式パターン", "9")
    # 特別条項ありの場合: 第7条=特別条項段落、第8条=有効期間（計8条構造）
    # 特別条項なしの場合: 第7条=有効期間（計7条構造、サンプルA社準拠）
    # 様式9_3はC社形式の簡略段落を使用
    if has_sp:
        sp_para = _special_para_93(r) if pat == "9_3" else _special_para(r)
        art7_8 = f"""{sp_para}

<div class="article">
<p><strong>第８条</strong>　本協定の有効期間は、{起算日}から１年間とする。</p>
</div>"""
    else:
        art7_8 = f"""<div class="article">
<p><strong>第７条</strong>　本協定の有効期間は、{起算日}から１年間とする。</p>
</div>"""

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>36協定書 - {社名}</title>
<style>{_CSS}</style>
</head>
<body>

<h1>時間外労働及び休日労働に関する協定書</h1>

<p class="intro">{社名}（以下「甲」という。）と労働者代表者（以下「乙」という。）は、労働基準法第３６条第１項の規定に基づき、労働基準法に定める法定労働時間（１週４０時間、１日８時間）並びに変形労働時間制に定める所定労働時間を超えた労働時間で、かつ１日８時間、１週４０時間の法定労働時間又は変形期間の法定労働時間の総枠を超える労働（以下「時間外労働」という。）及び労働基準法に定める休日（毎週１日又は４週４日）における労働（以下「休日労働」という。）に関し、次の通り協定する。</p>

<div class="article">
<p><strong>第１条</strong>　甲は、時間外労働及び休日労働を可能な限り行わせないように努める。</p>
</div>

<div class="article">
<p><strong>第２条</strong>　乙は、故意または過失により時間外労働及び休日労働を生じさせない義務を負う。</p>
</div>

<div class="article">
<p><strong>第３条</strong>　前2条にも関わらずその必要性を生じた場合、甲は次により時間外労働を行わせることができる。</p>
{_overtime_table(r)}
</div>

<div class="article">
<p><strong>第４条</strong>　甲は、時間外労働を行わせる場合は、原則として、前日の終業時刻までに当該労働者に通知する。また、休日労働を行わせる場合は、原則として、前日の終業時刻までに当該労働者に通知する。</p>
</div>

<div class="article">
<p><strong>第５条</strong>　第２条の表における1週、１ヶ月及び１年の起算日はいずれも{起算日}とする。</p>
</div>

<div class="article">
<p><strong>第６条</strong>　甲は、必要がある場合には、次により休日労働を行わせることができる。</p>
{_holiday_table(r)}
</div>

{art7_8}

<div class="sign-section">
  <div style="margin-left:42%; font-size:10.5pt; line-height:1.6;">
    <div style="text-align:right; padding-right:2pt; margin-bottom:2pt;">令和　{年号年}　年　　　月　　　日</div>
    <div style="margin-bottom:0pt;">（甲）{社名}</div>
    <div style="text-align:right; padding-right:2pt; margin-bottom:2pt;">{職名}　{代表者}</div>
    <table style="width:100%; border-collapse:separate; border-spacing:0; font-size:10.5pt; line-height:1.4;">
      <colgroup>
        <col style="width:110pt;">
        <col>
      </colgroup>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:left; padding:0 4pt 0 0;">（乙）労働者代表</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表職}</td>
      </tr>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:right; padding:0 4pt 0 0;">署名</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表氏名}</td>
      </tr>
    </table>
  </div>
</div>

</body>
</html>"""


# ═══════════════════════════════════════════════════
# A3 CSS（ドライバー・1年変形用）
# ═══════════════════════════════════════════════════
_CSS_A3 = """
@page {
    size: A3 portrait;
    margin: 13mm 10mm 18mm 20mm;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: 'Yu Mincho', '游明朝', 'MS Mincho', 'ＭＳ 明朝',
                 'MS PMincho', 'ＭＳ Ｐ明朝', serif;
    font-size: 9pt;
    line-height: 2.0;
    color: #000;
}
h1 {
    font-size: 13pt;
    font-weight: bold;
    text-align: center;
    line-height: 1.0;
    margin-bottom: 8pt;
    letter-spacing: 1pt;
}
.chapter {
    font-size: 10pt;
    font-weight: bold;
    margin-bottom: 4pt;
    line-height: 1.5;
}
.intro {
    text-align: justify;
    margin-bottom: 0pt;
    font-size: 9pt;
    line-height: 1.8;
}
.article {
    margin-bottom: 0pt;
    font-size: 9pt;
}
.article p {
    text-align: justify;
    line-height: 1.8;
    padding-left: 4em;
    text-indent: -4em;
}
.article-inline p {
    text-align: justify;
    line-height: 1.8;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin: 0pt 0 2pt 0;
    font-size: 8.5pt;
    table-layout: fixed;
    line-height: 1.25;
}
th, td {
    border: 0.8px solid #000;
    padding: 1.5pt 2pt;
    vertical-align: middle;
    text-align: center;
    word-break: break-all;
    overflow-wrap: break-word;
    line-height: 1.25;
}
th {
    background-color: #ffffff;
    font-weight: bold;
    font-size: 8.5pt;
}
td.tl { text-align: left; }
.sign-section { margin-top: 0pt; }
"""


# ═══════════════════════════════════════════════════
# ドライバー用テーブル（様式9_4 / 9_5）
# ═══════════════════════════════════════════════════
def _driver_overtime_table(r: dict) -> str:
    """ドライバー時間外テーブル（①行複数業種対応: _2/_3/_4サフィックス）"""
    year_limit = _v(r, "延長時間_1年", "360")

    # ①行の複数業種収集
    rows1: list[dict] = []
    for suf in ["", "_2", "_3", "_4"]:
        事由 = _v(r, f"時間外_事由{suf}")
        if not 事由 and suf:
            break
        if 事由:
            rows1.append({
                "事由": _ve(r, f"時間外_事由{suf}"),
                "業務": _ve(r, f"時間外_業務の種類{suf}"),
                "人数": _ve(r, f"労働者数{suf}"),
                "1日": _ve(r, f"延長時間_1日{suf}"),
                "1月": _ve(r, f"延長時間_1ヶ月{suf}"),
            })
    if not rows1:
        rows1 = [{"事由": "", "業務": "", "人数": "", "1日": "", "1月": ""}]

    n1 = len(rows1)
    first1 = rows1[0]
    extra1 = "".join(f"""    <tr>
      <td class="tl">{row['事由']}</td>
      <td class="tl">{row['業務']}</td>
      <td>{row['人数']}人</td>
      <td>{row['1日']}時間</td>
      <td>{row['1月']}時間</td>
      <td>{year_limit}時間</td>
    </tr>""" for row in rows1[1:])

    return f"""
<table>
  <colgroup>
    <col style="width:12%"><col style="width:28%"><col style="width:13%">
    <col style="width:10%"><col style="width:8%"><col style="width:8%"><col style="width:8%">
  </colgroup>
  <thead>
    <tr>
      <th rowspan="2"></th>
      <th rowspan="2">時間外労働をさせる必要のある具体的事由</th>
      <th rowspan="2">業務の種類</th>
      <th rowspan="2">従事する労働者数<br>（満18歳以上の者）</th>
      <th colspan="3">延長することができる時間</th>
    </tr>
    <tr>
      <th>１日</th><th>１か月</th><th>１年</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="{n1}" style="font-size:8.5pt; line-height:1.25;">①<br>下記②に<br>該当しない<br>労働者</td>
      <td class="tl">{first1['事由']}</td>
      <td class="tl">{first1['業務']}</td>
      <td>{first1['人数']}人</td>
      <td>{first1['1日']}時間</td>
      <td>{first1['1月']}時間</td>
      <td>{year_limit}時間</td>
    </tr>
    {extra1}
    <tr style="height:20pt;">
      <td style="font-size:8.5pt; line-height:1.25;">②<br>1年単位の<br>変形労働時間<br>制により<br>労働する労働者</td>
      <td></td><td></td><td></td><td></td><td></td><td></td>
    </tr>
  </tbody>
</table>"""


def _driver_holiday_table(r: dict, with_period: bool = True) -> str:
    """ドライバー休日労働テーブル（複数業種行対応: _2/_3/_4サフィックス）。
    with_period=True  → 期間列あり（様式10_2/F社形式）
    with_period=False → 期間列なし（様式9_4/E社形式）
    """
    if with_period:
        colgroup = '<col style="width:28%"><col style="width:13%"><col style="width:13%"><col style="width:26%"><col style="width:20%">'
        th_period = '<th>期間</th>'
    else:
        colgroup = '<col style="width:33%"><col style="width:15%"><col style="width:15%"><col style="width:37%">'
        th_period = ''

    # 複数行収集
    data_rows: list[dict] = []
    for suf in ["", "_2", "_3", "_4"]:
        事由 = _v(r, f"休日_事由{suf}")
        if not 事由 and suf:
            break
        日数 = _ve(r, f"休日労働_日数{suf}")
        時刻 = _ve(r, f"始業終業時刻{suf}")
        # 始業終業時刻内に「・」区切りがあれば<br>で展開
        時刻_html = "<br>".join(t.strip() for t in 時刻.split("・") if t.strip()) if "・" in 時刻 else 時刻
        data_rows.append({
            "事由": _ve(r, f"休日_事由{suf}") if 事由 else "",
            "業務": _ve(r, f"休日_業務の種類{suf}"),
            "人数": _ve(r, f"労働者数{suf}"),
            "時刻html": f"{日数}<br>{時刻_html}" if 日数 or 時刻 else "",
            "期間": _ve(r, f"休日_期間{suf}") if with_period else "",
        })
    if not data_rows:
        data_rows = [{"事由": "", "業務": "", "人数": "", "時刻html": "", "期間": ""}]

    rows_html = "".join(f"""    <tr>
      <td class="tl">{row['事由']}</td>
      <td class="tl">{row['業務']}</td>
      <td>{row['人数']}人</td>
      <td>{row['時刻html']}</td>
      {"<td class='tl'>" + row['期間'] + "</td>" if with_period else ""}
    </tr>""" for row in data_rows)

    return f"""
<table>
  <colgroup>
    {colgroup}
  </colgroup>
  <thead>
    <tr>
      <th>休日労働をさせる必要のある具体的事由</th>
      <th>業務の種類</th>
      <th>従事する労働者数<br>（満18歳以上の者）</th>
      <th>労働させることができる休日並びに始業及び終業の時刻</th>
      {th_period}
    </tr>
  </thead>
  <tbody>
    {rows_html}
  </tbody>
</table>"""


def _driver_special_table(r: dict) -> str:
    """特別条項テーブル（ドライバー様式: ①②行構造）
    ①行: 下記②に該当しない労働者（空行）
    ②行: 自動車の運転の業務に従事する労働者 + 事由 + 業務/人数/時間数
    """
    月時間 = _ve(r, "特別_延長時間_月", "80")
    年時間 = _ve(r, "特別_延長時間_年", "960")
    回数   = _ve(r, "特別_超過回数", "6")
    # 特別条項1日は専用フィールド。なければ通常OTの1日を流用
    d1日   = _ve(r, "特別_延長時間_1日") or _ve(r, "延長時間_1日")
    理由   = _ve(r, "特別_理由")
    業務   = _ve(r, "時間外_業務の種類")
    人数   = _ve(r, "労働者数")
    return f"""
<table>
  <colgroup>
    <col style="width:25%"><col style="width:13%"><col style="width:10%">
    <col style="width:7%"><col style="width:8%"><col style="width:8%">
    <col style="width:10%"><col style="width:10%"><col style="width:9%">
  </colgroup>
  <thead>
    <tr>
      <th rowspan="2">臨時的に限度時間を超えて<br>労働させることができる場合</th>
      <th rowspan="2">業務の種類</th>
      <th rowspan="2">従事する労働者数<br>（満18歳以上の者）</th>
      <th colspan="3">延長することができる時間数</th>
      <th rowspan="2">限度時間を超えて労働させることができる回数</th>
      <th colspan="2">延長することができる時間数及び休日労働の時間数</th>
    </tr>
    <tr>
      <th>１日</th><th>１か月</th><th>１年</th>
      <th>１か月</th><th>１年</th>
    </tr>
  </thead>
  <tbody>
    <tr style="height:20pt;">
      <td style="font-size:8.5pt; text-align:left; line-height:1.25;">①<br>下記②に該当しない<br>労働者</td>
      <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    </tr>
    <tr>
      <td class="tl" style="font-size:8.5pt; line-height:1.25;">②<br>自動車の運転の業務に従事する労働者<br>{理由}</td>
      <td class="tl">{業務}</td>
      <td>{人数}人</td>
      <td>{d1日}時間</td>
      <td></td>
      <td></td>
      <td>{回数}回</td>
      <td>{月時間}時間</td>
      <td>{年時間}時間</td>
    </tr>
  </tbody>
</table>"""


def _build_html_driver(r: dict) -> str:
    """様式9_4 / 9_5（ドライバー）用 A3縦HTML生成"""
    社名    = _ve(r, "事業所名")
    職名    = _ve(r, "事業主職名", "代表取締役")
    代表者  = _ve(r, "事業主名")
    代表職  = _ve(r, "労働者代表_職")
    代表氏名 = _ve(r, "労働者代表_氏名")
    起算日  = _e(_start_date(r))
    年号年  = _ve(r, "起算日_年")
    割増率  = _ve(r, "特別_割増賃金率", "25")
    手続き  = _ve(r, "特別_手続き", "労働者代表者代表に対する申し入れ")
    措置    = _e(_v(r, "特別_健康措置_内容") or "就業から始業までの一定時間の休息の確保、年次有給休暇連続取得の促進")

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>36協定書 - {社名}</title>
<style>{_CSS_A3}</style>
</head>
<body>

<h1>時間外労働及び休日労働に関する協定書</h1>

<p class="intro">{社名} {職名}（以下「甲」という。）と労働者代表（以下「乙」という。）は、労働基準法第３６条第１項の規定に基づき、労働基準法に定める法定労働時間（１週４０時間、１日８時間）を超える労働及び変形労働時間制の定めによる所定労働時間を超える労働時間で、かつ１日８時間、１週40時間の法定労働時間又は変形期間の法定労働時間の総枠を超える労働（以下「時間外労働」という。）並びに労働基準法に定める休日（毎週１日又は４週４日）における労働（以下「休日労働」という。）に関し、次のとおり協定する。</p>

<div class="article">
<p><strong>第１条</strong>　甲は、時間外労働及び休日労働を可能な限り行わせないよう努める。</p>
</div>

<div class="article">
<p><strong>第２条</strong>　甲は、就業規則の規定に基づき、必要がある場合には、次により時間外労働を行わせることができる。</p>
{_driver_overtime_table(r)}
</div>

<div class="article-inline">
<p>　２　自動車運転者（トラック）については、前項の規定により時間外労働を行わせることによって「自動車運転者の労働時間等の改善のための基準」（以下「改善基準告示」という。）に定める１か月及び１年についての拘束時間並びに１日についての最大拘束時間の限度を超えることとなる場合においては、当該拘束時間の限度をもって、前項の時間外労働時間の限度とする。</p>
</div>

<div class="article">
<p><strong>第３条</strong>　甲は、就業規則の規定に基づき、必要がある場合には、次により休日労働を行わせることができる。</p>
{_driver_holiday_table(r, with_period=False)}
</div>

<div class="article-inline">
<p>　２　自動車運転者（トラック）については、前項の規定により休日労働を行わせることによって、改善基準告示に定める１か月及び１年についての拘束時間並びに１日についての最大拘束時間の限度を超えることとなる場合においては、当該拘束時間の限度をもって、前項の休日労働の限度とする。</p>
</div>

<div class="article">
<p><strong>第４条</strong>　通常予見することのできない業務量の大幅な増加等に伴う臨時的な場合であって、次の何れかに該当する場合は、第２条の規定に基づき時間外労働を行わせることができる時間を超えて労働させることができる。</p>
{_driver_special_table(r)}
</div>

<div class="article-inline">
<p>　２　前項の規定に基づいて限度時間を超えて労働させる場合の割増率は{割増率}％とする。なお、時間外労働が１か月60時間を超えた場合の割増率は50％とする。</p>
<p>　３　第１項の規定に基づいて限度時間を超えて労働させる場合における手続及び限度時間を超えて労働させる労働者に対する健康及び福祉を確保するための措置については、次のとおりとする。<br>　　限度時間を超えて労働させる場合における手続：{手続き}<br>　　限度時間を超えて労働させる労働者に対する健康及び福祉を確保するための措置：{措置}</p>
<p>　４　自動車運転者（トラック）については、第１項の規定により時間外労働を行わせることによって改善基準告示に定める１か月及び１年についての拘束時間並びに１日についての最大拘束時間の限度を超えることとなる場合においては、当該拘束時間の限度をもって、第１項の時間外労働の時間の限度とする。</p>
</div>

<div class="article">
<p><strong>第５条</strong>　第２条から第４条までの規定に基づいて時間外労働又は休日労働を行わせる場合においても、自動車運転者（トラック）については、各条に定める時間数等にかかわらず、時間外労働及び休日労働を合算した時間数は１か月について100時間未満でなければならず、かつ２か月から６か月までを平均して80時間を超過しないこととする。</p>
</div>

<div class="article">
<p><strong>第６条</strong>　第２条から第４条までの規定に基づいて時間外労働又は休日労働を行わせる場合においても、自動車運転者（トラック）については、改善基準告示に定める運転時間の限度を超えて運転業務に従事させることはできない。</p>
</div>

<div class="article">
<p><strong>第７条</strong>　甲は、時間外労働を行わせる場合は、原則として、前日の終業時刻までに該当労働者に通知する。また、休日労働を行わせる場合は、原則として、２日前の終業時刻までに該当労働者に通知する。</p>
</div>

<div class="article">
<p><strong>第８条</strong>　第２条及び第４条の表における１年の起算日はいずれも{起算日}とする。</p>
<p>　２　本協定の有効期間は、{起算日}から１年間とする。</p>
</div>

<div class="sign-section">
  <div style="margin-left:42%; font-size:10pt; line-height:1.6;">
    <div style="text-align:right; padding-right:2pt; margin-bottom:2pt;">令和　{年号年}　年　　　月　　　日</div>
    <div style="margin-bottom:0pt;">（甲）{社名}　{職名}</div>
    <div style="text-align:right; padding-right:2pt; margin-bottom:2pt;">{代表者}</div>
    <table style="width:100%; border-collapse:separate; border-spacing:0; font-size:10pt; line-height:1.4;">
      <colgroup><col style="width:110pt;"><col></colgroup>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:left; padding:0 4pt 0 0;">（乙）労働者代表</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表職}</td>
      </tr>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:right; padding:0 4pt 0 0;">署名</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表氏名}</td>
      </tr>
    </table>
  </div>
</div>

</body>
</html>"""


# ═══════════════════════════════════════════════════
# 1年変形制用HTML（様式10 / 10_2）
# ═══════════════════════════════════════════════════
def _sign_section_a3(社名: str, 職名: str, 代表者: str, 代表職: str, 代表氏名: str, 年号年: str) -> str:
    """A3共通署名セクション"""
    return f"""
<div class="sign-section">
  <div style="margin-left:42%; font-size:10pt; line-height:1.6;">
    <div style="text-align:right; padding-right:2pt; margin-bottom:2pt;">令和　{年号年}　年　　　月　　　日</div>
    <div style="margin-bottom:0pt;">（甲）{社名}　{職名}</div>
    <div style="text-align:right; padding-right:2pt; margin-bottom:2pt;">{代表者}</div>
    <table style="width:100%; border-collapse:separate; border-spacing:0; font-size:10pt; line-height:1.4;">
      <colgroup><col style="width:110pt;"><col></colgroup>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:left; padding:0 4pt 0 0;">（乙）労働者代表</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表職}</td>
      </tr>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:right; padding:0 4pt 0 0;">署名</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表氏名}</td>
      </tr>
    </table>
  </div>
</div>"""


def _chapter2_1nen(r: dict, art_start: int, 起算日: str, 所定時間: str, begin: str, end: str, rest: str,
                   has_taishokikan: bool = False) -> str:
    """第2章（1年変形制）の共通HTML。
    art_start = 最初の条番号（様式10=9, 様式10_2=10）
    has_taishokikan = True  → 第11条「対象期間」を挿入（様式10/D社形式）
                     False → 対象期間なし（様式10_2/F社形式）
    複数行対応: 始業終業時刻_10条_2/_3 と 休憩時刻_2/_3 フィールドで追加行を指定
    """
    n = art_start
    漢数字 = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九",
              "一〇", "一一", "一二", "一三", "一四", "一五", "一六", "一七", "一八", "一九"]

    def art(no: int) -> str:
        return f"第{漢数字[no]}条" if no <= 19 else f"第{no}条"

    # 対象期間条文がある場合は条番号をずらす
    off = 1 if has_taishokikan else 0

    taishokikan_html = f"""
<div class="article">
<p><strong>{art(n+2)}</strong>　本協定の対象期間は、{起算日}から１年間とする。</p>
</div>
""" if has_taishokikan else ""

    # 休日条文：D社形式は「第(n+3+off)条の期間における初日を起算日とする」参照を含む
    if has_taishokikan:
        kyujitsu_text = f"前条の期間における休日は、{art(n+4+off)}の期間における初日を起算日とする１週に１日は確保するものとし、別添の年間カレンダーのとおりとする。尚、特定期間はないものとする。"
    else:
        kyujitsu_text = "前条の期間における休日は、１週に１回の日曜日の休日を確保するものとし、別添の年間カレンダーのとおりとする。尚、特定期間はないものとする。"

    # 始業終業休憩テーブルの複数行構築
    field_prefix = "始業終業時刻_10条" if art_start == 10 else "始業終業時刻_10条"
    rows_html = f"<tr><td>{begin}</td><td>{end}</td><td>{rest}</td></tr>"
    for suf in ["_2", "_3", "_4"]:
        b_raw = _v(r, f"始業終業時刻_10条{suf}")
        if not b_raw:
            break
        if "〜" in b_raw:
            parts = b_raw.split("〜", 1)
            b2 = _e(parts[0].strip())
            e2 = _e(parts[1].strip()) if len(parts) > 1 else "〇時〇分"
        else:
            b2 = _e(b_raw)
            e2 = "〇時〇分"
        r2 = _ve_br(r, f"休憩時刻{suf}", rest)
        rows_html += f"<tr><td>{b2}</td><td>{e2}</td><td>{r2}</td></tr>"

    return f"""
<p class="chapter">第2章 1年変形制について</p>

<p class="intro">１年単位の変形労働時間制に関し、次のとおり協定する。</p>

<div class="article">
<p><strong>{art(n)}</strong>　所定労働時間は、1年単位の変形労働時間制によるものとし、1年を平均して週40時間を越えないものとする。</p>
<p>　１．　始業終業休憩時間は下記の通りとする。</p>
</div>
<table style="width:40%; margin:0 0 4pt 2em;">
  <colgroup><col style="width:33%"><col style="width:33%"><col style="width:34%"></colgroup>
  <thead><tr><th>始業</th><th>終業</th><th>休憩</th></tr></thead>
  <tbody>{rows_html}</tbody>
</table>

<div class="article">
<p><strong>{art(n+1)}</strong>　本協定に基づく1年単位の変形労働時間制の対象者は次のとおりとする。</p>
<p>　１．　正規の当社社員。但し、次の者を除く。</p>
<p>　　（１）　妊産婦で適用除外のあった者</p>
<p>　　（２）　育児又は老人等の介護を行う者、職業訓練又は教育を受ける者、その他特別の理由により申し出をし、正当な事由より会社から認められた者</p>
<p>　　（３）　変形対象期間の中途で退職が予定される者及び中途採用者・配転者</p>
<p>　２．　パートタイマー、アルバイト及び臨時の従業員</p>
<p>　３．　嘱託雇用者で特に定めた者</p>
</div>
{taishokikan_html}
<div class="article">
<p><strong>{art(n+2+off)}</strong>　{kyujitsu_text}</p>
</div>

<div class="article">
<p><strong>{art(n+3+off)}</strong>　前条の期間における所定労働日は、前条に定める休日以外の日とする。</p>
</div>

<div class="article">
<p><strong>{art(n+4+off)}</strong>　変形対象期間における労働時間は、{所定時間}とする。（休憩時間を除く）</p>
</div>

<div class="article">
<p><strong>{art(n+5+off)}</strong>　業務上やむを得ない事由があるときは、前条の用件の範囲内で、従業員の代表の同意を得て同一週内において休日を振り替えることができる。</p>
</div>

<div class="article">
<p><strong>{art(n+6+off)}</strong>　本協定に基づく所定労働時間を超えて労働した場合には、就業規則の規定に基づき時間外手当を支払うものとする。</p>
</div>

<div class="article">
<p><strong>{art(n+7+off)}</strong>　本協定の有効期間は、{起算日}から1年間とする。</p>
</div>"""


def _build_html_1nen(r: dict) -> str:
    """様式10（1年変形制 + 時間外・休日労働）用 A3縦HTML生成
    構造: 第1章（第1〜6条）+ 第7条固定 + 第8条特別条項 + 第2章（第9〜第14条）
    サンプルD社（1年変形・特別条項あり）に準拠
    """
    社名    = _ve(r, "事業所名")
    職名    = _ve(r, "事業主職名", "代表取締役")
    代表者  = _ve(r, "事業主名")
    代表職  = _ve(r, "労働者代表_職")
    代表氏名 = _ve(r, "労働者代表_氏名")
    起算日  = _e(_start_date(r))
    年号年  = _ve(r, "起算日_年")
    所定時間 = _ve(r, "所定労働時間", "〇時間〇分")
    # 第2章用始業終業時刻: 専用フィールド(始業終業時刻_10条)があればそちらを優先
    # 理由: 始業終業時刻フィールドは休日テーブル用に「0時〜24時の間...」等が入りうる
    始業終業_raw = _v(r, "始業終業時刻_10条") or _v(r, "始業終業時刻") or "〇時〇分　〇時〇分"

    if "〜" in 始業終業_raw:
        parts = 始業終業_raw.split("〜", 1)
        begin = _e(parts[0].strip())
        end   = _e(parts[1].strip()) if len(parts) > 1 else "〇時〇分"
    else:
        begin = _e(始業終業_raw)
        end   = "〇時〇分"
    rest = _ve_br(r, "休憩時刻", "〇時〇分より〇時〇分まで")

    特別理由 = _ve(r, "特別_理由", "業務繁忙")
    月時間   = _ve(r, "特別_延長時間_月", "75")
    年時間   = _ve(r, "特別_延長時間_年", "600")
    回数     = _ve(r, "特別_超過回数", "6")
    limit    = _ve(r, "延長時間_1ヶ月", "42")
    year_limit = "320"
    割増率   = _ve(r, "特別_割増賃金率", "25")

    # 第7条（固定文言）+ 第8条（特別条項）
    art7_8 = f"""
<div class="article">
<p><strong>第７条</strong>　時間外労働及び休日労働の決定は会社によりなされ、よって会社の指示によってのみ行われる。つまり従業員の個人的な判断で行うことはできず、さらに包括的にその一切の権限をいかなる者にも与えることはない。</p>
</div>

<div class="article">
<p><strong>第８条</strong>　特別条項として{特別理由}に特に対応が必要なときは、労使の協議を経て、１か月{月時間}時間までこれを延長することができる。この場合、延長時間を更に延長する回数と法定労働外時間の合計は、{回数}回及び年間{年時間}時間までとし、かつこれが１ヶ月{limit}時間又は１年間{year_limit}時間を超えた時間外労働に対しての割増率を{割増率}％とする。また、60時間を超える場合の割増率は50％とする。</p>
</div>"""

    # 様式10（D社形式）は対象期間条文あり（第9〜17条の9条構成）
    chapter2 = _chapter2_1nen(r, 9, 起算日, 所定時間, begin, end, rest, has_taishokikan=True)
    sign = _sign_section_a3(社名, 職名, 代表者, 代表職, 代表氏名, 年号年)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>36協定書 - {社名}</title>
<style>{_CSS_A3}</style>
</head>
<body>

<h1>時間外労働及び休日労働及び1年変形制に関する協定書</h1>

<p class="chapter">第1章 時間外労働及び休日労働について</p>

<p class="intro">{社名}（以下「甲」という。）と労働者代表者（以下「乙」という。）は、労働基準法第３６条第１項の規定に基づき、労働基準法に定める法定労働時間（１週４０時間、１日８時間）並びに変形労働時間制に定める所定労働時間を超えた労働時間で、かつ１日８時間、１週４０時間の法定労働時間又は変形期間の法定労働時間の総枠を超える労働（以下「時間外労働」という。）及び労働基準法に定める休日（毎週１日又は４週４日）における労働（以下「休日労働」という。）に関し、次の通り協定する。</p>

<div class="article">
<p><strong>第１条</strong>　甲は、時間外労働及び休日労働を可能な限り行わせないように努める。</p>
</div>

<div class="article">
<p><strong>第２条</strong>　乙は、故意または過失により時間外労働及び休日労働を生じさせない義務を負う。</p>
</div>

<div class="article">
<p><strong>第３条</strong>　前2条にも関わらずその必要性を生じた場合、甲は次により時間外労働を行わせることができる。</p>
{_overtime_table(r)}
</div>

<div class="article">
<p><strong>第４条</strong>　甲は、時間外労働を行わせる場合は、原則として、前日の終業時刻までに当該労働者に通知する。また、休日労働を行わせる場合は、原則として、1日前の終業時刻までに当該労働者に通知する。</p>
</div>

<div class="article">
<p><strong>第５条</strong>　第３条の表における1週、１ヶ月及び１年の起算日はいずれも{起算日}とする。</p>
</div>

<div class="article">
<p><strong>第６条</strong>　甲は、必要がある場合には、次により休日労働を行わせることができる。</p>
{_holiday_table(r)}
</div>

{art7_8}

{chapter2}

{sign}

</body>
</html>"""


def _build_html_1nen_driver(r: dict) -> str:
    """様式10_2（1年変形制 + ドライバー）用 A3縦HTML生成
    構造: 第1章（第1〜9条、ドライバー様式）+ 第2章（第10〜17条）
    サンプルF社（1年変形・ドライバー・特条あり）に準拠
    """
    社名    = _ve(r, "事業所名")
    職名    = _ve(r, "事業主職名", "代表取締役")
    代表者  = _ve(r, "事業主名")
    代表職  = _ve(r, "労働者代表_職")
    代表氏名 = _ve(r, "労働者代表_氏名")
    起算日  = _e(_start_date(r))
    年号年  = _ve(r, "起算日_年")
    所定時間 = _ve(r, "所定労働時間", "〇時間〇分")
    # 第10条用始業終業時刻: 専用フィールド 始業終業時刻_10条 があればそちらを優先
    始業終業_raw = _v(r, "始業終業時刻_10条") or _v(r, "始業終業時刻") or "〇時〇分　〇時〇分"
    割増率  = _ve(r, "特別_割増賃金率", "25")
    手続き  = _ve(r, "特別_手続き", "労働者代表者との協議による合意")
    措置    = _e(_v(r, "特別_健康措置_内容") or "休日の確保 面接による健康の把握・労働状況の把握・改善指導")

    if "〜" in 始業終業_raw:
        parts = 始業終業_raw.split("〜", 1)
        begin = _e(parts[0].strip())
        end   = _e(parts[1].strip()) if len(parts) > 1 else "〇時〇分"
    else:
        begin = _e(始業終業_raw)
        end   = "〇時〇分"
    rest = _ve_br(r, "休憩時刻", "〇時〇分より〇時〇分まで")

    chapter2 = _chapter2_1nen(r, 10, 起算日, 所定時間, begin, end, rest)
    sign = _sign_section_a3(社名, 職名, 代表者, 代表職, 代表氏名, 年号年)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>36協定書 - {社名}</title>
<style>{_CSS_A3}</style>
</head>
<body>

<h1>時間外労働及び休日労働及び1年変形制に関する協定書</h1>

<p class="chapter">第1章 時間外労働及び休日労働について</p>

<p class="intro">{社名}（以下「甲」という。）と労働者代表者（以下「乙」という。）は、労働基準法第３６条第１項の規定に基づき、労働基準法に定める法定労働時間（１週４０時間、１日８時間）並びに変形労働時間制に定める所定労働時間を超えた労働時間で、かつ１日８時間、１週４０時間の法定労働時間又は変形期間の法定労働時間の総枠を超える労働（以下「時間外労働」という。）及び労働基準法に定める休日（毎週１日又は４週４日）における労働（以下「休日労働」という。）に関し、次の通り協定する。</p>

<div class="article">
<p><strong>第１条</strong>　甲は、時間外労働及び休日労働を可能な限り行わせないように努める。</p>
</div>

<div class="article">
<p><strong>第２条</strong>　乙は、故意または過失により時間外労働及び休日労働を生じさせない義務を負う。</p>
</div>

<div class="article">
<p><strong>第３条</strong>　前２条にも関わらずその必要性を生じた場合、甲は次により時間外労働を行わせることができる。</p>
{_overtime_table(r)}
</div>

<div class="article-inline">
<p>　２　自動車運転者については、前項の規定により時間外労働を行わせることによって「自動車運転者の労働時間等の改善のための基準」（以下「改善基準告示」という。）に定める１か月及び１年についての拘束時間並びに１日についての最大拘束時間の限度を超えることとなる場合においては、当該拘束時間の限度をもって、前項の時間外労働時間の限度とする。</p>
</div>

<div class="article">
<p><strong>第４条</strong>　甲は、就業規則第の規定に基づき、必要がある場合には、次により休日労働を行わせることができる。</p>
{_driver_holiday_table(r, with_period=True)}
</div>

<div class="article-inline">
<p>　２　自動車運転者については、前項の規定により休日労働を行わせることによって、改善基準告示に定める１か月及び１年についての拘束時間並びに１日についての最大拘束時間の限度を超えることとなる場合においては、当該拘束時間の限度をもって、前項の休日労働の限度とする。</p>
</div>

<div class="article">
<p><strong>第５条</strong>　通常予見することのできない業務量の大幅な増加等に伴う臨時的な場合であって、次の何れかに該当する場合は、第２条の規定に基づき時間外労働を行わせることができる時間を超えて労働させることができる。</p>
{_driver_special_table(r)}
</div>

<div class="article-inline">
<p>　２　前項の規定に基づいて限度時間を超えて労働させる場合の割増率は{割増率}％とする。　なお、時間外労働が１か月６０時間を超えた場合の割増率は５０％とする。</p>
<p>　３　第１項の規定に基づいて限度時間を超えて労働させる場合における手続及び限度時間を超えて労働させる労働者に対する健康及び福祉を確保するための措置については、次のとおりとする。</p>
<p>　　限度時間を超えて労働させる場合における手続：{手続き}</p>
<p>　　限度時間を超えて労働させる労働者に対する健康及び福祉を確保するための措置：{措置}</p>
<p>　４　自動車運転者については、第１項の規定により時間外労働を行わせることによって改善基準告示に定める１か月及び１年についての拘束時間並びに１日についての最大拘束時間の限度を超えることとなる場合においては、当該拘束時間の限度をもって、第１項の時間外労働の時間の限度とする。</p>
</div>

<div class="article">
<p><strong>第６条</strong>　第２条から第４条までの規定に基づいて時間外労働又は休日労働を行わせる場合においても、自動車運転者については、各条に定める時間数等にかかわらず、時間外労働及び休日労働を合算した時間数は１か月について１００時間未満となるよう努めるものとする。</p>
<p>　２　自動車運転者以外の者については、各条により定める時間数等にかかわらず、時間外労働及び休日労働を合算した時間数は、１か月について１００時間未満でなければならず、かつ２か月から６か月までを平均して８０時間を超過しないこととする。</p>
</div>

<div class="article">
<p><strong>第７条</strong>　第２条から第４条までの規定に基づいて時間外労働又は休日労働を行わせる場合においても、自動車運転者については、改善基準告示に定める運転時間の限度を超えて運転業務に従事させることはできない。</p>
</div>

<div class="article">
<p><strong>第８条</strong>　甲は、時間外労働を行わせる場合は、原則として、前日の終業時刻までに該当労働者に通知する。また、休日労働を行わせる場合は、原則として、１日前の終業時刻までに該当労働者に通知する。</p>
</div>

<div class="article">
<p><strong>第９条</strong>　第２条及び第４条の表における１年の起算日はいずれも{起算日}とする。</p>
<p>　２　本協定の有効期間は、{起算日}から１年間とする。</p>
</div>

{chapter2}

{sign}

</body>
</html>"""


# ═══════════════════════════════════════════════════
# 公開 API
# ═══════════════════════════════════════════════════
_KNOWN_PATTERNS = ("9", "9_2", "9_3", "9_4", "9_5", "10", "10_2")

# 様式パターン → HTMLビルダー関数のマッピング（新様式追加はここに1行追加するだけ）
_BUILDERS: dict[str, object] = {
    "9_4": _build_html_driver,
    "9_5": _build_html_driver,
    "10":  _build_html_1nen,
    "10_2": _build_html_1nen_driver,
}


def generate_pdf(record: dict) -> bytes:
    """1件のレコードからPDFバイト列を生成して返す"""
    社名_raw = _v(record, "事業所名", "不明")
    pat = _v(record, "様式パターン", "9")
    try:
        builder = _BUILDERS.get(pat, _build_html)
        if pat not in _KNOWN_PATTERNS:
            logger.warning("未知の様式パターン '%s' を標準様式(9)で処理します [%s]", pat, 社名_raw)
        html = builder(record)
        return weasyprint.HTML(string=html).write_pdf(font_config=_FONT_CONFIG)
    except Exception as exc:
        logger.error("PDF生成エラー [様式=%s, 事業所=%s]: %s", pat, 社名_raw, exc, exc_info=True)
        raise ValueError(f"PDF生成に失敗しました（{社名_raw}、様式{pat}）: {exc}") from exc


def _safe_filename(name: str) -> str:
    """ファイル名に使用できない文字（/ \\ : * ? " < > |）をアンダースコアに置換"""
    # スペース除去 → OS禁止文字を _ に置換 → 先頭末尾のドット・スペースを除去
    name = name.replace(" ", "").replace("　", "")
    name = re.sub(r'[\\/:*?"<>|&%#]', "_", name)
    name = name.strip(". ")
    return name or "不明"


def generate_pdf_file(record: dict, output_dir: str = "output") -> str:
    """PDFをファイルに保存してパスを返す"""
    try:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        社名 = _safe_filename(_v(record, "事業所名"))
        pat  = _safe_filename(_v(record, "様式パターン", "9"))
        out  = Path(output_dir) / f"36協定書_{社名}_{pat}.pdf"
        out.write_bytes(generate_pdf(record))
        logger.info("PDF保存完了: %s", out)
        return str(out)
    except ValueError:
        raise
    except Exception as exc:
        社名_raw = _v(record, "事業所名", "不明")
        logger.error("PDFファイル保存エラー [%s]: %s", 社名_raw, exc, exc_info=True)
        raise ValueError(f"PDFファイル保存に失敗しました（{社名_raw}）: {exc}") from exc

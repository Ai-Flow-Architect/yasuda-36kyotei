"""
36協定自動化ツール - Streamlit Webアプリ（Pattern A）
社労士事務所向け。
Excel（管理情報）+ Word（完成済み協定書）をアップロード
→ 事業所名マッチング → Word→PDF変換 → Yahoo Mail 下書き一括保存
"""
import base64
import io
import os
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))

# 起動時の import エラーを画面に出す（"Oh no. Error running app." の代わりに具体的な原因を表示）
try:
    from excel_reader import read_excel
    from graph_converter import convert_docx_to_pdf_graph, convert_docx_to_pdf_graph_personal
    from mail_drafter import save_draft
    from mail_sender import (
        build_email_body, build_subject,
        FEE_TYPE_STANDARD, FEE_TYPE_ANNUAL_CALENDAR,
    )
    from sample_selector import (
        build_extra_attachments, has_any_sample_spec, select_samples_for_record,
    )
    from word_matcher import build_match_table, convert_docx_to_pdf
except Exception as _import_err:
    st.set_page_config(page_title="36協定自動化ツール", page_icon="📄", layout="centered")
    st.error("⚠️ アプリの起動に失敗しました（モジュール読み込みエラー）")
    st.code(f"{type(_import_err).__name__}: {_import_err}\n\n{traceback.format_exc()}")
    st.stop()

# ============================================================
# ページ設定
# ============================================================
st.set_page_config(
    page_title="36協定自動化ツール",
    page_icon="📄",
    layout="centered",
)

# ============================================================
# カスタムCSS
# ============================================================
st.markdown("""
<style>
    .main-title {
        font-size: 1.8rem;
        font-weight: bold;
        color: #1a1a2e;
        text-align: center;
        padding: 1rem 0 0.3rem 0;
    }
    .sub-title {
        font-size: 0.95rem;
        color: #555;
        text-align: center;
        margin-bottom: 2rem;
    }
    .step-box {
        background: #f8f9ff;
        border-left: 4px solid #4a6cf7;
        padding: 1rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin: 1.2rem 0 0.5rem 0;
    }
    .step-label {
        font-size: 0.75rem;
        font-weight: bold;
        color: #4a6cf7;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .step-title {
        font-size: 1.05rem;
        font-weight: bold;
        color: #1a1a2e;
        margin-top: 0.2rem;
    }
    .result-card {
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin: 0.5rem 0;
    }
    .error-card {
        background: #fef2f2;
        border: 1px solid #fca5a5;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin: 0.5rem 0;
    }
    .footer {
        text-align: center;
        color: #aaa;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #eee;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# パスワード認証
# ============================================================
def check_password() -> bool:
    correct_pw = None
    try:
        correct_pw = st.secrets["password"]
    except Exception:
        correct_pw = os.environ.get("APP_PASSWORD", "")

    if not correct_pw:
        return True

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown('<div class="main-title">📄 36協定自動化ツール</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">社会保険労務士法人あさひ労務管理センター</div>', unsafe_allow_html=True)
    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("#### 🔒 パスワードを入力してください")
        pw = st.text_input("パスワード", type="password", key="pw_input",
                           label_visibility="collapsed", placeholder="パスワードを入力")
        if st.button("ログイン", use_container_width=True, type="primary"):
            if pw == correct_pw:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("パスワードが違います。もう一度お試しください。")
    return False


# ============================================================
# Microsoft Graph API設定をSecretsから取得
# ============================================================
def get_graph_config() -> dict:
    keys = ["ms_tenant_id", "ms_client_id", "ms_client_secret", "ms_user_email", "ms_refresh_token"]
    config = {}
    for k in keys:
        try:
            val = st.secrets[k]
        except Exception:
            val = None
        if val is None:
            val = os.environ.get(k.upper(), "")
        config[k] = str(val) if val else ""
    return config


def pdf_convert(docx_path: Path, output_dir: Path) -> tuple[bytes | None, str]:
    """Graph API優先、未設定ならLibreOfficeにフォールバック。

    認証モード優先順位:
      1. 個人Microsoftアカウント: ms_client_id + ms_refresh_token（ms_client_secretなし）
      2. M365 Business: ms_tenant_id + ms_client_id + ms_client_secret + ms_user_email
      3. LibreOffice（フォールバック）

    Returns:
        (pdf_bytes, error_msg): 成功時は (bytes, "")、失敗時は (None, エラーメッセージ)
    """
    cfg = get_graph_config()

    # 個人アカウントモード（リフレッシュトークン）
    if cfg.get("ms_client_id") and cfg.get("ms_refresh_token"):
        return convert_docx_to_pdf_graph_personal(
            docx_path,
            client_id=cfg["ms_client_id"],
            refresh_token=cfg["ms_refresh_token"],
            user_email=cfg.get("ms_user_email", ""),
        )

    # M365 Businessモード（client_credentials）
    if all(cfg.get(k) for k in ["ms_tenant_id", "ms_client_id", "ms_client_secret", "ms_user_email"]):
        return convert_docx_to_pdf_graph(
            docx_path,
            tenant_id=cfg["ms_tenant_id"],
            client_id=cfg["ms_client_id"],
            client_secret=cfg["ms_client_secret"],
            user_email=cfg["ms_user_email"],
        )

    # フォールバック: LibreOffice
    pdf_path, err = convert_docx_to_pdf(docx_path, output_dir)
    return (pdf_path.read_bytes() if pdf_path else None), err


WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _guess_mime(filename: str) -> str:
    """ファイル名から MIME タイプを推定する（見本ファイル用）"""
    ext = Path(filename).suffix.lower()
    return {
        ".pdf": "application/pdf",
        ".docx": WORD_MIME,
        ".doc": "application/msword",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls": "application/vnd.ms-excel",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
    }.get(ext, "application/octet-stream")


# ============================================================
# Yahoo IMAP設定をSecretsから取得
# ============================================================
def get_imap_config() -> dict:
    keys = ["yahoo_user", "yahoo_password", "差出人名", "差出人所属", "差出人電話", "担当者名"]
    config = {}
    for k in keys:
        try:
            val = st.secrets[k]
        except Exception:
            val = None
        if val is None:
            val = os.environ.get(k.upper(), "")
        config[k] = str(val) if val is not None else ""
    return config


# ============================================================
# メインアプリ
# ============================================================
def main() -> None:
    # ロゴ＋タイトル
    logo_path = Path(__file__).parent / "assets" / "logo.jpg"
    if logo_path.exists():
        logo_b64 = base64.b64encode(logo_path.read_bytes()).decode()
        st.markdown(f"""
        <div style="text-align:center; padding: 1rem 0 0.2rem 0;">
            <img src="data:image/jpeg;base64,{logo_b64}" style="height:64px; width:64px; object-fit:contain; border-radius:50%;">
        </div>""", unsafe_allow_html=True)
    st.markdown('<div class="main-title">36協定自動化ツール</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">Excel＋Wordをアップロードするだけで、PDF変換・メール下書き保存が完了します</div>',
        unsafe_allow_html=True,
    )

    # session_state 初期化
    defaults = {
        "records": [],
        "match_table": [],
        "last_excel_name": "",
        "last_word_names": [],
        "pdf_zip_bytes": None,
        "pdf_data": [],
        "draft_results": [],
        "_word_tmp_dir": "",
        "sample_files": [],  # 見本ファイル: list[tuple[bytes, filename, mime_type]]
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # --------------------------------------------------------
    # STEP 1: Excel アップロード
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 1</div>
        <div class="step-title">📂 Excelファイルをアップロード</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_excel = st.file_uploader(
        "36協定管理Excelファイル（.xlsx）を選択",
        type=["xlsx"],
        label_visibility="collapsed",
    )

    if uploaded_excel is None:
        st.info("👆 まずExcelファイルを選択してください。")
        _show_footer()
        return

    # ファイルが差し替わったらリセット
    if st.session_state.last_excel_name != uploaded_excel.name:
        st.session_state.last_excel_name = uploaded_excel.name
        st.session_state.records = []
        st.session_state.match_table = []
        st.session_state.pdf_zip_bytes = None
        st.session_state.pdf_data = []
        st.session_state.draft_results = []

    # Excel読み取り
    if not st.session_state.records:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded_excel.read())
            tmp_path = tmp.name
        try:
            excel_records, excel_warnings = read_excel(tmp_path)
            st.session_state.records = excel_records
            if excel_warnings:
                with st.expander(f"⚠️ 入力データの警告 {len(excel_warnings)} 件"):
                    for w in excel_warnings:
                        st.warning(w)
        except Exception as e:
            st.markdown(
                f'<div class="error-card">❌ Excelの読み取りに失敗しました。<br><small>{e}</small></div>',
                unsafe_allow_html=True,
            )
            return
        finally:
            os.unlink(tmp_path)

    records = st.session_state.records
    if not records:
        st.warning("Excelにデータが見つかりませんでした。内容を確認してください。")
        return

    # --------------------------------------------------------
    # STEP 2: Excel プレビュー
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 2</div>
        <div class="step-title">📋 読み取り結果の確認</div>
    </div>
    """, unsafe_allow_html=True)

    st.success(f"**{len(records)} 件** のデータを読み取りました。")

    preview_rows = [
        {
            "#": i + 1,
            "事業所名": r.get("事業所名", "（未入力）"),
            "更新月": r.get("更新月", ""),
            "送信先メール": r.get("メールアドレス") or "⚠️ 未設定",
        }
        for i, r in enumerate(records)
    ]
    st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # STEP 3: Word または PDF アップロード → 確認
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 3</div>
        <div class="step-title">📄 協定書ファイルをアップロード（Word または PDF）</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_words = st.file_uploader(
        "完成済み36協定書（.pdf / .docx / .doc）を選択（複数可）",
        type=["pdf", "docx", "doc"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if not uploaded_words:
        st.info("👆 ファイルを選択するとマッチング結果が表示されます。")
        _show_footer()
        return

    # Wordファイルが変わったらリセット
    current_word_names = sorted(f.name for f in uploaded_words)
    if st.session_state.last_word_names != current_word_names:
        st.session_state.last_word_names = current_word_names
        st.session_state.match_table = []
        st.session_state.pdf_zip_bytes = None
        st.session_state.pdf_data = []
        st.session_state.draft_results = []

    # Wordを一時ファイルに保存しマッチング
    if not st.session_state.match_table:
        word_tmp_dir = tempfile.mkdtemp()
        word_paths = []
        for wf in uploaded_words:
            p = Path(word_tmp_dir) / wf.name
            p.write_bytes(wf.read())
            word_paths.append(p)
        st.session_state.match_table = build_match_table(records, word_paths)
        st.session_state._word_tmp_dir = word_tmp_dir

    match_table = st.session_state.match_table

    # マッチング結果プレビュー
    st.markdown("**マッチング結果**")
    preview = [
        {
            "事業所名": row["事業所名"],
            "送信先メール": row["送信先メール"],
            "協定書ファイル": row["協定書ファイル"],
            "件数": row.get("件数", 0),
            "形式": row["形式"],
        }
        for row in match_table
    ]
    st.dataframe(preview, use_container_width=True, hide_index=True)

    unmatched = [r for r in match_table if r["_matched_path"] is None]
    if unmatched:
        st.warning(
            f"⚠️ {len(unmatched)} 件がマッチしていません: "
            + ", ".join(r["事業所名"] for r in unmatched)
        )

    matched_count = len(match_table) - len(unmatched)
    no_email_count = sum(1 for r in match_table if "⚠️" in str(r["送信先メール"]))
    st.info(
        f"マッチ: **{matched_count}/{len(match_table)}** 件 ／ "
        f"メール未設定: **{no_email_count}** 件"
    )

    # PDF準備ボタン（PDF直接アップロードの場合は「変換」ではなく「読み込み」）
    all_paths = [p for r in match_table for p in r.get("_matched_paths", [])]
    pdf_count = sum(1 for p in all_paths if p.suffix.lower() == ".pdf")
    word_count = len(all_paths) - pdf_count
    if not st.session_state.pdf_zip_bytes:
        btn_label = "📄 PDFを準備する"
        if pdf_count > 0 and word_count == 0:
            btn_label = f"📄 PDF {pdf_count}件を読み込む（変換なし）"
        elif pdf_count > 0:
            btn_label = f"📄 PDF {pdf_count}件を読み込む ＋ Word {word_count}件を変換する"
        if st.button(btn_label, type="primary", use_container_width=True):
            _run_pdf_only(match_table)

    # PDF生成済み → ダウンロード + 確認チェック
    if st.session_state.pdf_zip_bytes:
        st.success(f"✅ **{len(st.session_state.pdf_data)} 件** のPDFを準備しました。")
        st.download_button(
            label="📥 PDF ZIP をダウンロードして内容を確認する",
            data=st.session_state.pdf_zip_bytes,
            file_name="36協定書_PDF一括.zip",
            mime="application/zip",
            use_container_width=True,
        )

        # --------------------------------------------------------
        # STEP 4: PDF確認済み → Yahoo下書き一括保存
        # --------------------------------------------------------
        st.markdown("""
        <div class="step-box">
            <div class="step-label">STEP 4</div>
            <div class="step-title">📨 PDFを確認したら、Yahoo メールの下書きを一括保存する</div>
        </div>
        """, unsafe_allow_html=True)

        imap_config = get_imap_config()
        imap_ok = bool(imap_config.get("yahoo_user") and imap_config.get("yahoo_password"))

        if not imap_ok:
            st.error("⚠️ Yahoo Mail の設定が完了していません。（yahoo_user / yahoo_password 未設定）")
            _show_footer()
            return

        st.info(f"📤 差出人アカウント: **{imap_config.get('yahoo_user', '')}**")

        st.info(
            "代行手数料はWordファイル名から自動判定します。\n"
            "「36協定及び1年変形」を含むファイル → 12,000円版、それ以外 → 5,000円版"
        )

        # 見本出し分けモード判定:
        # Excelの「見本指定」列に1件でも入力があれば事業所ごとの出し分け、
        # なければ従来どおり全宛先共通で添付する。
        per_office_samples = has_any_sample_spec(records)

        if per_office_samples:
            st.markdown("**📎 見本ファイル（事業所ごとに出し分けて添付します）**")
            st.info(
                "📌 **事業所ごとの出し分けモード**\n\n"
                "Excelの「見本指定」列に書かれたファイル名と、ここにアップロードした"
                "見本を照合し、各事業所には指定された見本だけを添付します。"
                "出し分けに使う見本は、まとめてここにアップロードしてください。"
            )
            sample_help = (
                "ここにアップロードした見本のうち、Excelの「見本指定」列で"
                "指定されたものだけが各事業所のメールに添付されます。"
            )
        else:
            st.markdown("**📎 見本ファイル（全宛先のメールに共通で添付されます）**")
            sample_help = (
                "ここにアップロードしたファイルは、全ての事業所宛てメール下書きに"
                "共通で添付されます。"
            )

        sample_uploads = st.file_uploader(
            "見本書類（PDF・Word・画像など、複数選択可）",
            type=["pdf", "docx", "doc", "xlsx", "xls", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="sample_uploader",
            label_visibility="collapsed",
            help=sample_help,
        )
        if sample_uploads:
            sample_files: list[tuple[bytes, str, str]] = []
            for up in sample_uploads:
                data = up.read()
                sample_files.append((data, up.name, _guess_mime(up.name)))
            st.session_state.sample_files = sample_files
            for up, (data, _, _) in zip(sample_uploads, sample_files):
                st.caption(f"・{up.name}（{len(data)//1024}KB）")
        else:
            st.session_state.sample_files = []
            st.caption("見本を添付しない場合はそのまま次へ進めます。")

        # 出し分けモードでは「どの事業所にどの見本が付くか」を事前プレビューし、
        # 指定したのにプールに無い見本（タイプミス／アップロード漏れ）を警告する。
        # 見本をまだアップロードしていない段階では全件が未一致になり誤解を招くため、
        # アップロード後にのみプレビューを表示する。
        if per_office_samples and st.session_state.sample_files:
            _show_sample_match_preview(
                st.session_state.pdf_data, st.session_state.sample_files
            )

        st.markdown("---")

        # 締切月入力（メール本文の「〇月15日」を設定）
        first_record = st.session_state.pdf_data[0].get("record", {}) if st.session_state.pdf_data else {}
        try:
            default_締切月 = str(int(first_record.get("更新月", "0") or "0") - 1)
            if default_締切月 == "0":
                default_締切月 = "12"
        except (ValueError, TypeError):
            default_締切月 = ""
        col_month, col_day = st.columns(2)
        with col_month:
            締切月 = st.text_input(
                "締切月（例: 4）",
                value=default_締切月,
                key="締切月_input",
            )
        with col_day:
            締切日 = st.text_input(
                "締切日（例: 15）",
                value="15",
                key="締切日_input",
            )
        st.caption(
            f"→ メール本文に「{締切月 or '〇'}月{締切日 or '〇'}日まで」と記載されます。"
            "日付は自由に変更できます。"
        )
        imap_config["締切月"] = 締切月
        imap_config["締切日"] = 締切日

        confirmed = st.checkbox("PDFの内容を確認しました。Yahoo メールの下書きに保存します。")

        if confirmed:
            if st.button(
                "📨 Yahoo メール下書きを一括保存する",
                type="primary",
                use_container_width=True,
            ):
                _run_draft_only(
                    st.session_state.pdf_data,
                    imap_config,
                    sample_files=st.session_state.sample_files,
                )

        # 下書き保存結果
        if st.session_state.draft_results:
            ok_count = sum(1 for r in st.session_state.draft_results if "成功" in str(r["結果"]))
            fail_count = len(st.session_state.draft_results) - ok_count

            if fail_count == 0:
                st.markdown(
                    f'<div class="result-card">✅ <strong>{ok_count} 件</strong> '
                    f'すべて完了しました。Yahoo Mail の下書きフォルダをご確認ください。</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.warning(f"{ok_count} 件成功 / {fail_count} 件失敗")

            st.dataframe(st.session_state.draft_results, use_container_width=True, hide_index=True)

    _show_footer()


def _dedupe_zip_name(name: str, used_names: set[str]) -> str:
    """ZIP内ファイル名の衝突を避ける。

    同一事業所が同名stemの協定書を複数持つ場合、後勝ち上書きで
    1件サイレント欠落するため、衝突時は連番サフィックス（_2,_3…）を
    付与してからwritestrする。
    """
    if name not in used_names:
        used_names.add(name)
        return name
    stem, dot, ext = name.rpartition(".")
    base = stem if dot else name
    suffix = f".{ext}" if dot else ""
    n = 2
    while f"{base}_{n}{suffix}" in used_names:
        n += 1
    unique = f"{base}_{n}{suffix}"
    used_names.add(unique)
    return unique


def _run_pdf_only(match_table: list[dict]) -> None:
    """Word→PDF変換のみ実行し、ZIPとpdf_dataをsession_stateに保存する"""
    pdf_zip_buf = io.BytesIO()
    pdf_data = []
    convert_errors = []
    used_zip_names: set[str] = set()  # ZIP内ファイル名衝突防止（サイレント欠落対策）
    total = len(match_table)
    progress = st.progress(0, text="PDF変換中...")

    with zipfile.ZipFile(pdf_zip_buf, "w", zipfile.ZIP_DEFLATED) as pdf_zf:
        with tempfile.TemporaryDirectory() as pdf_out_dir:
            for i, row in enumerate(match_table):
                progress.progress((i + 1) / total, text=f"PDF変換中... {i+1}/{total}")
                name = row["事業所名"]
                email_addr = str(row["送信先メール"])
                matched_paths = row.get("_matched_paths") or []
                record: dict = row["_record"]

                if not matched_paths:
                    convert_errors.append(f"{name}: ファイル未マッチ")
                    continue

                # 1事業所に複数の協定書（様式9号＋9号の2＋1年変形等）がある場合、
                # すべてPDF化して同じメール下書きに添付する（バグ①対応）
                office_num = record.get("事業所番号", "")
                converted: list[tuple[bytes, str]] = []  # [(pdf_bytes, pdf_filename), ...]
                for j, matched_path in enumerate(matched_paths):
                    if matched_path.suffix.lower() == ".pdf":
                        pdf_bytes = matched_path.read_bytes()
                        pdf_err = ""
                    else:
                        pdf_bytes, pdf_err = pdf_convert(matched_path, Path(pdf_out_dir))
                    if pdf_bytes is None:
                        convert_errors.append(f"{name}（{matched_path.name}）: {pdf_err}")
                        continue
                    # 複数ファイルはファイル名衝突を避けるため元ファイル名を保持
                    base = matched_path.stem
                    if len(matched_paths) == 1:
                        pdf_filename = (
                            f"{office_num}_36協定書_{name}.pdf"
                            if office_num else f"36協定書_{name}.pdf"
                        )
                    else:
                        pdf_filename = (
                            f"{office_num}_{base}.pdf" if office_num else f"{base}.pdf"
                        )
                    pdf_filename = _dedupe_zip_name(pdf_filename, used_zip_names)
                    pdf_zf.writestr(pdf_filename, pdf_bytes)
                    converted.append((pdf_bytes, pdf_filename))

                if not converted:
                    continue  # 全件変換失敗（個別エラーは記録済み）

                primary_bytes, primary_name = converted[0]
                extra_kyotei = [
                    (b, fn, "application/pdf") for b, fn in converted[1:]
                ]
                pdf_data.append({
                    "事業所名": name,
                    "email_addr": email_addr,
                    "record": record,
                    "pdf_bytes": primary_bytes,
                    "pdf_filename": primary_name,
                    "extra_kyotei": extra_kyotei,
                    "kyotei_count": len(converted),
                    "word_filename": " ".join(mp.name for mp in matched_paths),
                })

    progress.empty()

    if convert_errors:
        with st.expander(f"⚠️ PDF変換エラー {len(convert_errors)} 件", expanded=True):
            for err in convert_errors:
                st.error(err)

    st.session_state.pdf_zip_bytes = pdf_zip_buf.getvalue()
    st.session_state.pdf_data = pdf_data
    st.rerun()


def _run_draft_only(
    pdf_data: list[dict],
    imap_config: dict,
    sample_files: list[tuple[bytes, str, str]] | None = None,
) -> None:
    """保存済みPDFデータをもとにYahoo下書きを一括保存する

    Args:
        sample_files: 全宛先共通で添付する見本ファイル一覧
            [(file_bytes, filename, mime_type), ...]
    """
    sample_files = sample_files or []
    # 見本出し分けモード判定: いずれかのレコードに「見本指定」があれば事業所ごと、
    # なければ従来どおり全宛先共通で添付する。
    per_office = has_any_sample_spec([it.get("record", {}) for it in pdf_data])
    results = []
    total = len(pdf_data)
    progress = st.progress(0, text="下書き保存中...")

    for i, item in enumerate(pdf_data):
        progress.progress((i + 1) / total, text=f"下書き保存中... {i+1}/{total}")
        name = item["事業所名"]
        email_addr = item["email_addr"]

        if "⚠️" in email_addr or not email_addr:
            results.append({"事業所名": name, "宛先": "（未設定）", "結果": "⚠️ メールアドレスなし"})
            continue

        # Wordファイル名から代行手数料タイプを自動判定
        word_filename = item.get("word_filename", "")
        auto_fee = FEE_TYPE_ANNUAL_CALENDAR if "36協定及び1年変形" in word_filename else FEE_TYPE_STANDARD

        subject = build_subject(item["record"])
        body = build_email_body(item["record"], imap_config, fee_type=auto_fee)

        # 追加添付（協定書2件目以降＋見本）を組み立てる。
        # 出し分けモードなら「見本指定」に一致した見本のみ、共通モードなら全見本。
        extra_kyotei = item.get("extra_kyotei", [])
        extra_attachments, this_samples, _ = build_extra_attachments(
            extra_kyotei,
            sample_files,
            item.get("record", {}).get("見本指定", ""),
            per_office,
        )
        res = save_draft(
            to_address=email_addr,
            subject=subject,
            body=body,
            pdf_bytes=item["pdf_bytes"],
            pdf_filename=item["pdf_filename"],
            imap_user=imap_config.get("yahoo_user", ""),
            imap_password=imap_config.get("yahoo_password", ""),
            from_address=imap_config.get("yahoo_user", ""),
            extra_attachments=extra_attachments,
        )
        kyotei_n = item.get("kyotei_count", 1)
        sample_n = len(this_samples)
        note_parts = []
        if kyotei_n > 1:
            note_parts.append(f"協定書{kyotei_n}件")
        if sample_n:
            note_parts.append(f"見本{sample_n}件")
        total_att = kyotei_n + sample_n
        sample_note = f"（添付{total_att}件・" + "・".join(note_parts) + "）" if note_parts else ""
        results.append({"事業所名": name, "宛先": email_addr, "結果": f"{res['status']}{sample_note}"})

    progress.empty()
    st.session_state.draft_results = results


def _show_sample_match_preview(
    pdf_data: list[dict],
    sample_files: list[tuple[bytes, str, str]],
) -> None:
    """出し分けモードで「どの事業所にどの見本が付くか」を事前プレビューする。

    Excelの「見本指定」で指定したのにプールに無い見本（タイプミス／
    アップロード漏れ）はサイレント欠落になるため、警告で可視化する。
    """
    if not pdf_data:
        return
    rows = []
    any_unmatched = False
    for item in pdf_data:
        record = item.get("record", {})
        spec = str(record.get("見本指定", "") or "").strip()
        selected, unmatched = select_samples_for_record(spec, sample_files)
        if unmatched:
            any_unmatched = True
        rows.append({
            "事業所名": item.get("事業所名", ""),
            "見本指定": spec or "（なし）",
            "添付される見本": "／".join(s[1] for s in selected) if selected else "（なし）",
            "未一致の指定": "／".join(unmatched) if unmatched else "",
        })

    st.markdown("**🔍 見本の出し分けプレビュー**")
    st.dataframe(rows, use_container_width=True, hide_index=True)
    if any_unmatched:
        st.warning(
            "⚠️ 「未一致の指定」がある事業所は、その見本がアップロードされていないか"
            "ファイル名が一致していません。見本のアップロードとExcelの「見本指定」の"
            "綴りをご確認ください（一致しない指定は添付されません）。"
        )


def _show_footer():
    st.markdown(
        '<div class="footer">36協定自動化ツール｜社会保険労務士法人あさひ労務管理センター</div>',
        unsafe_allow_html=True,
    )


if check_password():
    main()

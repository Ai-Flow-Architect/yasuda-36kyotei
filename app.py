"""
36協定自動化ツール - Streamlit Webアプリ（Pattern A）
社労士事務所向け。
Excel（管理情報）+ Word（完成済み協定書）をアップロード
→ 事業所名マッチング → Word→PDF変換（飯塚様確認）→ Yahoo Mail 下書き一括保存
"""
import base64
import io
import os
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from excel_reader import read_excel
from graph_converter import convert_docx_to_pdf_graph
from mail_drafter import save_draft
from mail_sender import (
    build_email_body, build_subject,
    FEE_TYPE_STANDARD, FEE_TYPE_ANNUAL_CALENDAR,
)
from word_matcher import build_match_table, convert_docx_to_pdf

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
    keys = ["ms_tenant_id", "ms_client_id", "ms_client_secret", "ms_user_email"]
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


def pdf_convert(docx_path: Path, output_dir: Path) -> bytes | None:
    """Graph API優先、未設定ならLibreOfficeにフォールバック"""
    cfg = get_graph_config()
    if all(cfg.get(k) for k in ["ms_tenant_id", "ms_client_id", "ms_client_secret", "ms_user_email"]):
        return convert_docx_to_pdf_graph(
            docx_path,
            tenant_id=cfg["ms_tenant_id"],
            client_id=cfg["ms_client_id"],
            client_secret=cfg["ms_client_secret"],
            user_email=cfg["ms_user_email"],
        )
    # フォールバック: LibreOffice
    pdf_path = convert_docx_to_pdf(docx_path, output_dir)
    return pdf_path.read_bytes() if pdf_path else None


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
    # STEP 3: Word アップロード → PDF変換 → 飯塚様確認
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 3</div>
        <div class="step-title">📄 Wordをアップロードして、PDFに変換・内容を確認</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_words = st.file_uploader(
        "完成済み36協定書のWordファイル（.docx）を選択（複数可）",
        type=["docx", "doc"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if not uploaded_words:
        st.info("👆 Wordファイルを選択するとマッチング結果が表示されます。")
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
            "Wordファイル": row["Wordファイル"],
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

    # PDF変換ボタン
    if not st.session_state.pdf_zip_bytes:
        if st.button("📄 PDFを一括生成する", type="primary", use_container_width=True):
            _run_pdf_only(match_table)

    # PDF生成済み → ダウンロード + 確認チェック
    if st.session_state.pdf_zip_bytes:
        st.success(f"✅ **{len(st.session_state.pdf_data)} 件** のPDFを生成しました。")
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

        # 締切月入力（メール本文の「〇月15日」を設定）
        first_record = st.session_state.pdf_data[0].get("record", {}) if st.session_state.pdf_data else {}
        try:
            default_締切月 = str(int(first_record.get("更新月", "0") or "0") - 1)
            if default_締切月 == "0":
                default_締切月 = "12"
        except (ValueError, TypeError):
            default_締切月 = ""
        締切月 = st.text_input(
            "締切月（例: 4 と入力すると「4月15日まで」とメールに記載されます）",
            value=default_締切月,
            key="締切月_input",
        )
        imap_config["締切月"] = 締切月

        confirmed = st.checkbox("PDFの内容を確認しました。Yahoo メールの下書きに保存します。")

        if confirmed:
            if st.button(
                "📨 Yahoo メール下書きを一括保存する",
                type="primary",
                use_container_width=True,
            ):
                _run_draft_only(st.session_state.pdf_data, imap_config)

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


def _run_pdf_only(match_table: list[dict]) -> None:
    """Word→PDF変換のみ実行し、ZIPとpdf_dataをsession_stateに保存する"""
    pdf_zip_buf = io.BytesIO()
    pdf_data = []
    total = len(match_table)
    progress = st.progress(0, text="PDF変換中...")

    with zipfile.ZipFile(pdf_zip_buf, "w", zipfile.ZIP_DEFLATED) as pdf_zf:
        with tempfile.TemporaryDirectory() as pdf_out_dir:
            for i, row in enumerate(match_table):
                progress.progress((i + 1) / total, text=f"PDF変換中... {i+1}/{total}")
                name = row["事業所名"]
                email_addr = str(row["送信先メール"])
                matched_path = row["_matched_path"]
                record: dict = row["_record"]

                if matched_path is None:
                    continue

                pdf_bytes = pdf_convert(matched_path, Path(pdf_out_dir))
                if pdf_bytes is None:
                    continue

                pdf_filename = f"36協定書_{name}.pdf"
                pdf_zf.writestr(pdf_filename, pdf_bytes)

                pdf_data.append({
                    "事業所名": name,
                    "email_addr": email_addr,
                    "record": record,
                    "pdf_bytes": pdf_bytes,
                    "pdf_filename": pdf_filename,
                    "word_filename": matched_path.name,
                })

    progress.empty()
    st.session_state.pdf_zip_bytes = pdf_zip_buf.getvalue()
    st.session_state.pdf_data = pdf_data
    st.rerun()


def _run_draft_only(pdf_data: list[dict], imap_config: dict) -> None:
    """保存済みPDFデータをもとにYahoo下書きを一括保存する"""
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
        res = save_draft(
            to_address=email_addr,
            subject=subject,
            body=body,
            pdf_bytes=item["pdf_bytes"],
            pdf_filename=item["pdf_filename"],
            imap_user=imap_config.get("yahoo_user", ""),
            imap_password=imap_config.get("yahoo_password", ""),
            from_address=imap_config.get("yahoo_user", ""),
        )
        results.append({"事業所名": name, "宛先": email_addr, "結果": res["status"]})

    progress.empty()
    st.session_state.draft_results = results


def _show_footer():
    st.markdown(
        '<div class="footer">36協定自動化ツール｜社会保険労務士法人あさひ労務管理センター</div>',
        unsafe_allow_html=True,
    )


if check_password():
    main()

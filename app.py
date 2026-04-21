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
import zipfile
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from excel_reader import read_excel
from mail_drafter import save_draft
from mail_sender import build_email_body, build_subject
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
# Yahoo IMAP設定をSecretsから取得
# ============================================================
def get_imap_config() -> dict:
    keys = ["yahoo_user", "yahoo_password", "差出人名", "差出人所属", "差出人電話"]
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
        "draft_results": [],
        "pdf_zip_bytes": None,
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
        st.session_state.draft_results = []
        st.session_state.pdf_zip_bytes = None

    # Excel読み取り
    if not st.session_state.records:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded_excel.read())
            tmp_path = tmp.name
        try:
            st.session_state.records = read_excel(tmp_path)
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
    # STEP 3: Word アップロード → PDF変換 + メール下書き一括保存
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 3</div>
        <div class="step-title">📄 Wordをアップロードし、メールの下書きとPDFを一括生成</div>
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

    # Wordファイルが変わったらマッチングをリセット
    current_word_names = sorted(f.name for f in uploaded_words)
    if st.session_state.last_word_names != current_word_names:
        st.session_state.last_word_names = current_word_names
        st.session_state.match_table = []
        st.session_state.draft_results = []
        st.session_state.pdf_zip_bytes = None

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

    # IMAP設定確認
    imap_config = get_imap_config()
    imap_ok = bool(imap_config.get("yahoo_user") and imap_config.get("yahoo_password"))

    if not imap_ok:
        st.error(
            "⚠️ Yahoo Mail の設定が完了していません。"
            "（yahoo_user / yahoo_password 未設定）"
        )
        _show_footer()
        return

    st.info(f"📤 差出人アカウント: **{imap_config.get('yahoo_user', '')}**")
    confirmed = st.checkbox("上記の内容を確認しました。PDF変換してメール下書きを保存します。")

    if confirmed:
        if st.button(
            "📨 PDF変換 + メール下書きを一括保存する",
            type="primary",
            use_container_width=True,
        ):
            _run_pdf_and_draft(match_table, imap_config)

    # 実行結果表示
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

        # PDFをZIPでダウンロード
        if st.session_state.pdf_zip_bytes:
            st.download_button(
                label="📥 PDF ZIP をダウンロード",
                data=st.session_state.pdf_zip_bytes,
                file_name="36協定書_PDF一括.zip",
                mime="application/zip",
                use_container_width=True,
            )

    _show_footer()


def _run_pdf_and_draft(match_table: list[dict], imap_config: dict) -> None:
    """Word→PDF変換 + Yahoo下書き一括保存を実行"""
    results = []
    pdf_zip_buf = io.BytesIO()
    total = len(match_table)
    progress = st.progress(0, text="処理中...")

    with zipfile.ZipFile(pdf_zip_buf, "w", zipfile.ZIP_DEFLATED) as pdf_zf:
        with tempfile.TemporaryDirectory() as pdf_out_dir:
            for i, row in enumerate(match_table):
                progress.progress((i + 1) / total, text=f"処理中... {i+1}/{total}")
                name = row["事業所名"]
                email_addr = str(row["送信先メール"])
                matched_path = row["_matched_path"]
                record: dict = row["_record"]

                if "⚠️" in email_addr or not email_addr:
                    results.append({"事業所名": name, "宛先": "（未設定）", "結果": "⚠️ メールアドレスなし"})
                    continue

                if matched_path is None:
                    results.append({"事業所名": name, "宛先": email_addr, "結果": "❌ Wordファイル未マッチ"})
                    continue

                # Word → PDF
                pdf_path = convert_docx_to_pdf(matched_path, Path(pdf_out_dir))
                if pdf_path is None:
                    results.append({"事業所名": name, "宛先": email_addr, "結果": "❌ PDF変換失敗"})
                    continue

                pdf_bytes = pdf_path.read_bytes()
                pdf_filename = f"36協定書_{name}.pdf"
                pdf_zf.writestr(pdf_filename, pdf_bytes)

                # Yahoo下書き保存
                subject = build_subject(record)
                body = build_email_body(record, imap_config)
                res = save_draft(
                    to_address=email_addr,
                    subject=subject,
                    body=body,
                    pdf_bytes=pdf_bytes,
                    pdf_filename=pdf_filename,
                    imap_user=imap_config.get("yahoo_user", ""),
                    imap_password=imap_config.get("yahoo_password", ""),
                    from_address=imap_config.get("yahoo_user", ""),
                )
                results.append({"事業所名": name, "宛先": email_addr, "結果": res["status"]})

    progress.empty()
    st.session_state.draft_results = results
    st.session_state.pdf_zip_bytes = pdf_zip_buf.getvalue()


def _show_footer():
    st.markdown(
        '<div class="footer">36協定自動化ツール｜社会保険労務士法人あさひ労務管理センター</div>',
        unsafe_allow_html=True,
    )


if check_password():
    main()

"""
36協定自動化ツール - Streamlit Webアプリ
社労士事務所向け。Excelアップロード → Word+PDF協定書生成 → Yahoo Mail 下書き一括保存（PDF添付）
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
from word_generator import generate_word, FORM_NAMES
from pdf_generator import generate_pdf, _safe_filename
from mail_drafter import save_draft
from mail_sender import build_email_body, build_subject

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
        # st.secrets が None を返した場合に備えて環境変数をフォールバックとし、
        # 最終的に文字列型を保証する（imaplib.login() に None を渡さない）
        if val is None:
            val = os.environ.get(k.upper(), "")
        config[k] = str(val) if val is not None else ""
    return config


# ============================================================
# ファイル名ヘルパー（generate_all_files / save_all_drafts 共通）
# ============================================================
def _make_pdf_filename(name: str, form_type: str, idx: int) -> str:
    """PDF用一意ファイル名を返す。両関数で同じロジックを保証する。"""
    safe = _safe_filename(name or f"企業{idx}")
    return f"36協定書_{safe}_{form_type}_{idx}.pdf"


# ============================================================
# Word + PDF 生成 → 各ZIPバイナリとファイルbytes辞書を返す
# ============================================================
def generate_all_files(
    records: list[dict],
) -> tuple[bytes, bytes, dict[str, bytes], dict[str, bytes], list[dict]]:
    """Word ZIP・PDF ZIP・個別bytes辞書・結果を返す

    Returns:
        word_zip_bytes, pdf_zip_bytes, word_file_bytes, pdf_file_bytes, results
    """
    results = []
    word_zip_buf = io.BytesIO()
    pdf_zip_buf = io.BytesIO()
    word_file_bytes: dict[str, bytes] = {}
    pdf_file_bytes: dict[str, bytes] = {}

    with tempfile.TemporaryDirectory() as tmpdir, \
         zipfile.ZipFile(word_zip_buf, "w", zipfile.ZIP_DEFLATED) as word_zf, \
         zipfile.ZipFile(pdf_zip_buf, "w", zipfile.ZIP_DEFLATED) as pdf_zf:

        for i, record in enumerate(records):
            # None値ガード: record.get()がNoneを返す場合にf"企業{i+1}"へフォールバック
            name = record.get("事業所名") or f"企業{i+1}"
            form_type = record.get("様式パターン") or "9"
            form_label = FORM_NAMES.get(form_type, form_type)
            error_msg = ""
            # 一意ファイル名用インデックス（同一社名・同一様式の衝突防止）
            idx = i + 1

            # Word生成
            word_ok = False
            try:
                out_path = generate_word(record, output_dir=tmpdir)
                # generate_word が返すパスのステムにインデックスを付与して衝突回避
                original = Path(out_path)
                word_filename = f"{original.stem}_{idx}{original.suffix}"
                new_word_path = original.parent / word_filename
                original.rename(new_word_path)
                with open(new_word_path, "rb") as f:
                    word_bytes = f.read()
                word_file_bytes[word_filename] = word_bytes
                word_zf.write(str(new_word_path), arcname=word_filename)
                word_ok = True
            except Exception as e:
                error_msg = f"Word: {e}"

            # PDF生成
            pdf_ok = False
            try:
                pdf_filename = _make_pdf_filename(name, form_type, idx)
                pdf_bytes = generate_pdf(record)
                pdf_file_bytes[pdf_filename] = pdf_bytes
                pdf_tmp = Path(tmpdir) / pdf_filename
                pdf_tmp.write_bytes(pdf_bytes)
                pdf_zf.write(str(pdf_tmp), arcname=pdf_filename)
                pdf_ok = True
            except Exception as e:
                error_msg += f" PDF: {e}"

            if word_ok and pdf_ok:
                status = "✅ 生成完了"
            elif word_ok or pdf_ok:
                status = "⚠️ 一部失敗"
            else:
                status = "❌ 失敗"

            results.append({
                "事業所名": name,
                "様式": form_label,
                "Word": "✅" if word_ok else "❌",
                "PDF": "✅" if pdf_ok else "❌",
                "エラー": error_msg,
            })

    return (
        word_zip_buf.getvalue(),
        pdf_zip_buf.getvalue(),
        word_file_bytes,
        pdf_file_bytes,
        results,
    )


# ============================================================
# Yahoo Mail 下書き一括保存（PDF添付）
# ============================================================
def save_all_drafts(
    records: list[dict],
    pdf_file_bytes: dict[str, bytes],
    imap_config: dict,
) -> list[dict]:
    results = []
    for i, record in enumerate(records):
        # None値ガード: record.get()がNoneを返す場合に空文字へフォールバック
        name = record.get("事業所名") or ""
        email_addr = record.get("メールアドレス") or ""
        form_type = record.get("様式パターン") or "9"
        idx = i + 1
        # _make_pdf_filename で generate_all_files と完全に同じ命名規則を保証
        pdf_filename = _make_pdf_filename(name, form_type, idx)

        if not email_addr:
            results.append({"事業所名": name, "宛先": "（未設定）", "結果": "⚠️ メールアドレスなし"})
            continue

        pdf_bytes = pdf_file_bytes.get(pdf_filename)
        if not pdf_bytes:
            results.append({"事業所名": name, "宛先": email_addr, "結果": "❌ PDFが見つかりません"})
            continue

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

    return results


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
    st.markdown('<div class="sub-title">Excelをアップロードするだけで、36協定書（Word・PDF）の生成・メール下書き保存が完了します</div>',
                unsafe_allow_html=True)

    # session_state 初期化
    if "word_file_bytes" not in st.session_state:
        st.session_state.word_file_bytes = {}
    if "pdf_file_bytes" not in st.session_state:
        st.session_state.pdf_file_bytes = {}
    if "word_zip_bytes" not in st.session_state:
        st.session_state.word_zip_bytes = None
    if "pdf_zip_bytes" not in st.session_state:
        st.session_state.pdf_zip_bytes = None
    if "gen_results" not in st.session_state:
        st.session_state.gen_results = []
    if "records" not in st.session_state:
        st.session_state.records = []
    if "draft_results" not in st.session_state:
        st.session_state.draft_results = []

    # --------------------------------------------------------
    # STEP 1: Excel アップロード
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 1</div>
        <div class="step-title">📂 Excelファイルをアップロード</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "36協定データの Excel ファイル（.xlsx）を選択してください",
        type=["xlsx"],
        label_visibility="collapsed",
    )

    if uploaded is None:
        st.info("👆 Excelファイルを選択すると、処理が始まります。")
        _show_footer()
        return

    # ファイルが差し替わったらsession_stateをリセット
    if "last_filename" not in st.session_state or st.session_state.last_filename != uploaded.name:
        st.session_state.last_filename = uploaded.name
        st.session_state.word_file_bytes = {}
        st.session_state.pdf_file_bytes = {}
        st.session_state.word_zip_bytes = None
        st.session_state.pdf_zip_bytes = None
        st.session_state.gen_results = []
        st.session_state.records = []
        st.session_state.draft_results = []

    # Excel 読み取り
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    try:
        records = read_excel(tmp_path)
    except Exception as e:
        st.markdown(f'<div class="error-card">❌ Excelの読み取りに失敗しました。<br><small>{e}</small></div>',
                    unsafe_allow_html=True)
        return
    finally:
        os.unlink(tmp_path)

    if not records:
        st.warning("Excelにデータが見つかりませんでした。内容を確認してください。")
        return

    st.session_state.records = records

    # --------------------------------------------------------
    # STEP 2: プレビュー
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 2</div>
        <div class="step-title">📋 読み取り結果の確認</div>
    </div>
    """, unsafe_allow_html=True)

    st.success(f"**{len(records)} 件** のデータを読み取りました。")

    # 主要5列サマリー
    preview_rows = []
    for i, r in enumerate(records):
        form_type = r.get("様式パターン", "9")
        form_label = FORM_NAMES.get(form_type, form_type)
        preview_rows.append({
            "#": i + 1,
            "事業所名": r.get("事業所名", "（未入力）"),
            "事業主名": r.get("事業主名", "（未入力）"),
            "協定書の種類": form_label,
            "特別条項": "あり" if form_type in ("9_2", "9_3", "9_4", "9_5") else "なし",
            "送信先メール": r.get("メールアドレス", "⚠️ 未設定"),
        })

    st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    # 全列展開式
    with st.expander("📊 全データを確認する（クリックで展開）"):
        exclude_keys = {"様式パターン"}
        display_records = [
            {k: v for k, v in r.items() if k not in exclude_keys}
            for r in records
        ]
        st.dataframe(display_records, use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # STEP 3: Word + PDF 生成 & ダウンロード
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 3</div>
        <div class="step-title">📝 Word・PDF協定書を一括生成してダウンロード</div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("⚡ Word + PDF を生成する", type="primary", use_container_width=True):
        with st.spinner("協定書を生成しています（Word + PDF）…"):
            word_zip, pdf_zip, word_bytes, pdf_bytes, gen_results = generate_all_files(records)
        st.session_state.word_zip_bytes = word_zip
        st.session_state.pdf_zip_bytes = pdf_zip
        st.session_state.word_file_bytes = word_bytes
        st.session_state.pdf_file_bytes = pdf_bytes
        st.session_state.gen_results = gen_results
        st.session_state.draft_results = []  # 再生成時は下書き結果をリセット

    if st.session_state.gen_results:
        success_count = sum(1 for r in st.session_state.gen_results if "✅" in r["Word"] and "✅" in r["PDF"])
        fail_count = len(st.session_state.gen_results) - success_count

        if fail_count == 0:
            st.markdown(f'<div class="result-card">✅ <strong>{success_count} 件</strong> すべて生成完了しました（Word・PDF）。</div>',
                        unsafe_allow_html=True)
        else:
            st.warning(f"{success_count} 件成功 / {fail_count} 件失敗")

        st.dataframe(st.session_state.gen_results, use_container_width=True, hide_index=True)

        col_w, col_p = st.columns(2)
        with col_w:
            if st.session_state.word_zip_bytes:
                st.download_button(
                    label="📥 Word ZIP をダウンロード",
                    data=st.session_state.word_zip_bytes,
                    file_name="36協定書_Word一括.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
        with col_p:
            if st.session_state.pdf_zip_bytes:
                st.download_button(
                    label="📥 PDF ZIP をダウンロード",
                    data=st.session_state.pdf_zip_bytes,
                    file_name="36協定書_PDF一括.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

    # --------------------------------------------------------
    # STEP 4: Yahoo Mail 下書き一括保存（PDF生成完了後のみ表示）
    # --------------------------------------------------------
    if not st.session_state.pdf_file_bytes:
        _show_footer()
        return

    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 4</div>
        <div class="step-title">✉️ 協定書（PDF）をメール下書きに一括保存</div>
    </div>
    """, unsafe_allow_html=True)

    # 送信予定テーブル
    send_preview = []
    for r in records:
        name = r.get("事業所名", "")
        email_addr = r.get("メールアドレス", "")
        subject = build_subject(r)
        send_preview.append({
            "事業所名": name,
            "宛先メール": email_addr if email_addr else "⚠️ 未設定",
            "件名": subject,
            "添付": "PDF",
        })

    st.markdown("**下書き保存予定の一覧**")
    st.dataframe(send_preview, use_container_width=True, hide_index=True)

    # IMAP設定確認
    imap_config = get_imap_config()
    imap_ok = bool(imap_config.get("yahoo_user") and imap_config.get("yahoo_password"))

    if not imap_ok:
        st.error("⚠️ Yahoo Mail の設定が完了していません。管理者に連絡してください。（yahoo_user / yahoo_password 未設定）")
        _show_footer()
        return

    st.info(f"📤 差出人アカウント: **{imap_config.get('yahoo_user', '')}**")

    confirmed = st.checkbox("上記の宛先にメール下書きを保存することを確認しました")

    if confirmed:
        if st.button("📨 メール下書きを一括保存する", type="primary", use_container_width=True):
            with st.spinner("Yahoo Mail の下書きフォルダに保存しています…"):
                draft_results = save_all_drafts(
                    st.session_state.records,
                    st.session_state.pdf_file_bytes,
                    imap_config,
                )
            st.session_state.draft_results = draft_results

    if st.session_state.draft_results:
        ok_count = sum(1 for r in st.session_state.draft_results if "成功" in r["結果"])
        fail_count = len(st.session_state.draft_results) - ok_count

        if fail_count == 0:
            st.markdown(f'<div class="result-card">✅ <strong>{ok_count} 件</strong> すべて下書き保存しました。Yahoo Mail の下書きフォルダをご確認ください。</div>',
                        unsafe_allow_html=True)
        else:
            st.warning(f"{ok_count} 件保存成功 / {fail_count} 件失敗")

        st.dataframe(st.session_state.draft_results, use_container_width=True, hide_index=True)

    _show_footer()


def _show_footer():
    st.markdown(
        '<div class="footer">36協定自動化ツール｜社会保険労務士法人あさひ労務管理センター</div>',
        unsafe_allow_html=True,
    )


if check_password():
    main()

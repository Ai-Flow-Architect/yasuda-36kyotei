"""
mail_drafter.py — Yahoo Japan IMAP で36協定書PDFを下書き保存
Yahoo Japan: imap.mail.yahoo.co.jp:993 (SSL)
下書きフォルダ名: "Draft"

【日本語ファイル名対応の方針】
Yahoo Japan Mail webUI は RFC 2231 の filename*= をパースできず "Untitled" になる。
互換性最大化のため以下を全て同時に出力する：
  1. Content-Type の name パラメータ（RFC 2047 B-encoding）
  2. Content-Disposition の filename= （RFC 2047 B-encoding、レガシー互換）
  3. Content-Disposition の filename*= （RFC 2231、モダンクライアント互換）
  4. ASCII フォールバック name="attachment_N.pdf" を別途
"""
import base64
import imaplib
import logging
import re
import time
from datetime import datetime, timezone
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
from urllib.parse import quote

logger = logging.getLogger("yasuda_36kyotei")

IMAP_HOST = "imap.mail.yahoo.co.jp"
IMAP_PORT = 993
DRAFT_FOLDER_CANDIDATES = ["Draft", "Drafts", "下書き", "DRAFT", "DRAFTS"]
MAX_RETRY = 3
RETRY_WAIT_SEC = 2


def _find_draft_folder(m: imaplib.IMAP4_SSL) -> str:
    """IMAPフォルダ一覧から下書きフォルダ名を自動検出する。
    見つからない場合は "Draft" をフォールバックとして返す。
    """
    try:
        typ, folder_list = m.list()
        if typ != "OK":
            return DRAFT_FOLDER_CANDIDATES[0]
        for item in folder_list:
            decoded = item.decode("utf-7") if isinstance(item, bytes) else str(item)
            for candidate in DRAFT_FOLDER_CANDIDATES:
                # フォルダ名はクォートあり・なし両方に対応
                if f'"{candidate}"' in decoded or decoded.endswith(f" {candidate}"):
                    logger.debug(f"[フォルダ検出] 下書き候補: {candidate} (行: {decoded})")
                    return candidate
    except Exception as e:
        logger.warning(f"[フォルダ一覧取得失敗] {e}")
    logger.warning(f"[フォルダ未検出] 候補 {DRAFT_FOLDER_CANDIDATES} が見つからず 'Draft' を使用")
    return DRAFT_FOLDER_CANDIDATES[0]


def _rfc2047_b_encode(text: str) -> str:
    """日本語文字列を RFC 2047 B-encoding する（=?UTF-8?B?...?=）

    Yahoo Japan を含む全主要メールクライアントが理解する最も互換性の高い形式。
    """
    if not text:
        return ""
    # ASCII のみなら encoding 不要
    try:
        text.encode("ascii")
        return text
    except UnicodeEncodeError:
        pass
    b64 = base64.b64encode(text.encode("utf-8")).decode("ascii")
    return f"=?UTF-8?B?{b64}?="


def _ascii_fallback(filename: str, idx: int = 0) -> str:
    """非ASCII文字を含むファイル名から ASCII フォールバック名を生成"""
    # 拡張子を保持
    m = re.match(r"^(.*?)(\.[A-Za-z0-9]+)?$", filename)
    base, ext = (m.group(1) or "file", m.group(2) or ".pdf") if m else ("file", ".pdf")
    # ASCII以外を除去
    ascii_base = re.sub(r"[^A-Za-z0-9_\-]+", "", base)
    if not ascii_base:
        ascii_base = f"attachment_{idx or 1}"
    return f"{ascii_base}{ext}"


def _build_attachment_part(pdf_bytes: bytes, pdf_filename: str, idx: int = 0) -> MIMEBase:
    """互換性最大化PDF添付パートを生成する

    Python email モジュールが自動で挿入する RFC 2231 形式を使わず、
    手動で RFC 2047 B-encoding を含む全形式を直接ヘッダーに書き込む。
    """
    part = MIMEBase("application", "pdf")
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)

    # Python が自動で付けたヘッダーをいったん削除
    if "Content-Type" in part:
        del part["Content-Type"]
    if "Content-Disposition" in part:
        del part["Content-Disposition"]

    # 1. 日本語ファイル名 → RFC 2047 B-encoding
    encoded_name = _rfc2047_b_encode(pdf_filename)
    # 2. ASCII フォールバック名（極端な互換用）
    ascii_name = _ascii_fallback(pdf_filename, idx)
    # 3. RFC 2231 用 URL エンコード
    url_encoded = quote(pdf_filename, safe="")

    # Content-Type: application/pdf; name="=?UTF-8?B?...?="
    part["Content-Type"] = f'application/pdf; name="{encoded_name}"'

    # Content-Disposition:
    #   attachment;
    #   filename="=?UTF-8?B?...?=";   ← RFC 2047 (Yahoo Japan 等レガシー互換)
    #   filename*=UTF-8''<urlencoded>  ← RFC 2231 (モダン互換)
    cd = (
        f'attachment; '
        f'filename="{encoded_name}"; '
        f"filename*=UTF-8''{url_encoded}"
    )
    part["Content-Disposition"] = cd

    logger.debug(f"[添付ヘッダー] {pdf_filename} → encoded={encoded_name}, ascii={ascii_name}")
    return part


def _build_message(
    to_address: str,
    subject: str,
    body: str,
    pdf_bytes: bytes,
    pdf_filename: str,
    from_address: str,
    idx: int = 0,
) -> bytes:
    """MIMEメッセージを組み立てて bytes を返す"""
    msg = MIMEMultipart("mixed")
    msg["From"] = from_address
    msg["To"] = to_address
    # Subject も日本語対応で Header() を使う
    msg["Subject"] = Header(subject, "utf-8").encode()
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="yahoo.co.jp")
    msg["MIME-Version"] = "1.0"

    # 本文（UTF-8）
    text_part = MIMEText(body, "plain", "utf-8")
    msg.attach(text_part)

    # PDF 添付
    pdf_part = _build_attachment_part(pdf_bytes, pdf_filename, idx)
    msg.attach(pdf_part)

    return msg.as_bytes()


def save_draft(
    to_address: str,
    subject: str,
    body: str,
    pdf_bytes: bytes,
    pdf_filename: str,
    imap_user: str,
    imap_password: str,
    from_address: str = "",
    idx: int = 0,
) -> dict:
    """Yahoo IMAP経由でPDF添付メールを下書きフォルダに保存する（リトライ付き）

    Args:
        to_address: 宛先メールアドレス
        subject: 件名
        body: 本文
        pdf_bytes: PDF のバイト列
        pdf_filename: 添付ファイル名（例: "36協定書_株式会社サンプル商事.pdf"）
        imap_user: Yahoo メールアドレス
        imap_password: Yahoo パスワード（IMAP用アプリパスワード）
        from_address: 差出人アドレス（省略時は imap_user）
        idx: 連番（ASCII フォールバック名生成用）

    Returns:
        {"to": str, "subject": str, "status": str, "attempts": int}
    """
    sender = from_address or imap_user
    result = {"to": to_address, "subject": subject, "status": "未保存", "attempts": 0}

    # 入力バリデーション
    if not pdf_bytes or len(pdf_bytes) < 100:
        result["status"] = "失敗: PDFバイト列が不正（空または極小）"
        logger.error(f"[PDF不正] To: {to_address} | size={len(pdf_bytes) if pdf_bytes else 0}")
        return result
    if not pdf_filename:
        result["status"] = "失敗: PDFファイル名が空"
        return result

    # メッセージ組み立て
    try:
        raw_message = _build_message(
            to_address=to_address,
            subject=subject,
            body=body,
            pdf_bytes=pdf_bytes,
            pdf_filename=pdf_filename,
            from_address=sender,
            idx=idx,
        )
    except Exception as e:
        result["status"] = f"失敗: メッセージ組み立てエラー: {e}"
        logger.error(f"[組み立て失敗] {e}")
        return result

    # IMAP APPEND（最大3回リトライ）
    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        result["attempts"] = attempt
        try:
            with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=30) as m:
                m.login(imap_user, imap_password)
                draft_folder = _find_draft_folder(m)
                logger.info(f"[下書きフォルダ] 使用フォルダ: {draft_folder}")
                internal_date = imaplib.Time2Internaldate(datetime.now(timezone.utc))
                typ, data = m.append(
                    draft_folder,
                    r"\Draft",
                    internal_date,
                    raw_message,
                )
                if typ != "OK":
                    raise imaplib.IMAP4.error(f"APPEND returned {typ}: {data} (folder={draft_folder})")
            result["status"] = "下書き保存成功"
            logger.info(f"[下書き保存成功] To: {to_address} | attempts={attempt}")
            return result
        except imaplib.IMAP4.error as e:
            last_err = f"IMAP認証/APPENDエラー: {e}"
            logger.warning(f"[IMAP失敗 {attempt}/{MAX_RETRY}] {e}")
            # 認証エラーはリトライしない
            if "AUTHENTICATIONFAILED" in str(e).upper() or "LOGIN" in str(e).upper():
                break
        except (TimeoutError, ConnectionError, OSError) as e:
            last_err = f"接続エラー: {e}"
            logger.warning(f"[接続失敗 {attempt}/{MAX_RETRY}] {e}")
        except Exception as e:
            last_err = f"想定外エラー: {e}"
            logger.error(f"[想定外 {attempt}/{MAX_RETRY}] {e}")

        if attempt < MAX_RETRY:
            time.sleep(RETRY_WAIT_SEC)

    result["status"] = f"失敗: {last_err}"
    return result

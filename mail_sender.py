"""
メール送信モジュール
36協定書をOutlook/SMTP経由で送信する（定型文1種類）
リトライ機能付き（最大3回、間隔5秒）
"""
import logging
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import Optional
from urllib.parse import quote

# モジュール共通ロガー
logger = logging.getLogger("yasuda_36kyotei")

# SMTP送信リトライ設定
MAX_RETRIES: int = 3
RETRY_INTERVAL_SEC: int = 5

# デフォルト定型文
DEFAULT_TEMPLATE: str = """
{宛先名} 様

お世話になっております。
{事業主名}の36協定届（時間外労働及び休日労働に関する協定書）を送付いたします。

添付ファイルをご確認いただき、内容に問題がなければご署名をお願いいたします。
ご不明な点がございましたら、お気軽にお問い合わせください。

何卒よろしくお願い申し上げます。

──────────────────
{差出人名}
{差出人所属}
TEL: {差出人電話}
──────────────────
""".strip()


def create_email(
    to_address: str,
    subject: str,
    body: str,
    attachment_path: Optional[str],
    from_address: str = "",
) -> MIMEMultipart:
    """メールメッセージを作成する"""
    msg = MIMEMultipart()
    msg["From"] = from_address
    msg["To"] = to_address
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # 添付ファイル
    if attachment_path and Path(attachment_path).exists():
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        filename: str = Path(attachment_path).name
        encoded_filename: str = quote(filename)
        part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded_filename}")
        msg.attach(part)

    return msg


def send_email(
    msg: MIMEMultipart,
    smtp_server: str = "smtp.office365.com",
    smtp_port: int = 587,
    username: str = "",
    password: str = "",
    dry_run: bool = True,
) -> dict[str, str]:
    """メールを送信する（リトライ機能付き: 最大3回、間隔5秒）

    Args:
        msg: 送信するMIMEメッセージ
        smtp_server: SMTPサーバーアドレス
        smtp_port: SMTPポート番号
        username: SMTP認証ユーザー名
        password: SMTP認証パスワード
        dry_run: Trueの場合、実際には送信せずログだけ出力する（デモ用）

    Returns:
        送信結果を含む辞書
    """
    result: dict[str, str] = {
        "to": msg["To"],
        "subject": msg["Subject"],
        "status": "未送信",
    }

    if dry_run:
        result["status"] = "DRY_RUN（送信スキップ）"
        logger.info(f"[DRY_RUN] To: {msg['To']} | Subject: {msg['Subject']}")
        return result

    # リトライループ（最大MAX_RETRIES回）
    last_error: Optional[Exception] = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            logger.debug(f"SMTP送信試行 {attempt}/{MAX_RETRIES}: To={msg['To']}")
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(username, password)
                server.send_message(msg)
            result["status"] = "送信成功"
            logger.info(f"[送信成功] To: {msg['To']}")
            return result
        except Exception as e:
            last_error = e
            logger.warning(
                f"[送信失敗 {attempt}/{MAX_RETRIES}] To: {msg['To']} | Error: {e}"
            )
            if attempt < MAX_RETRIES:
                logger.info(f"  {RETRY_INTERVAL_SEC}秒後にリトライします...")
                time.sleep(RETRY_INTERVAL_SEC)

    # 全リトライ失敗
    result["status"] = f"送信失敗（{MAX_RETRIES}回リトライ後）: {str(last_error)}"
    logger.error(f"[送信失敗（リトライ上限）] To: {msg['To']} | Error: {last_error}")
    return result


def build_email_body(data: dict[str, str], config: dict[str, str]) -> str:
    """定型文にデータを埋め込む"""
    template: str = config.get("メールテンプレート", DEFAULT_TEMPLATE)
    return template.format(
        宛先名=data.get("事業主名", "ご担当者"),
        事業主名=data.get("事業所名", ""),
        差出人名=config.get("差出人名", ""),
        差出人所属=config.get("差出人所属", ""),
        差出人電話=config.get("差出人電話", ""),
    )


def build_subject(data: dict[str, str]) -> str:
    """メール件名を生成"""
    事業所名: str = data.get("事業所名", "")
    return f"【36協定届】{事業所名}様 時間外労働及び休日労働に関する協定書"


if __name__ == "__main__":
    # テスト
    config: dict[str, str] = {
        "差出人名": "安田",
        "差出人所属": "朝日事務所",
        "差出人電話": "03-0000-0000",
    }
    data: dict[str, str] = {"事業主名": "テスト太郎", "事業所名": "テスト株式会社"}

    body: str = build_email_body(data, config)
    subject: str = build_subject(data)
    print(f"件名: {subject}")
    print(f"本文:\n{body}")

"""
メール送信モジュール
36協定書をOutlook/SMTP経由で送信する
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

logger = logging.getLogger("yasuda_36kyotei")

MAX_RETRIES: int = 3
RETRY_INTERVAL_SEC: int = 5

MAIL_SUBJECT: str = "36協定の更新について"

# 標準テンプレート（作成提出代行手数料: 5,000円）
MAIL_TEMPLATE_STANDARD: str = """{宛先会社名}
{宛先担当者名} 様

お世話になっております。
36協定の更新時期が近付いてまいりましたのでご連絡申し上げます。
添付しております協定書をご確認いただき、下記についてご記入いただけますようお願い申し上げます。

①人数
②日付(書類記入日)
③労働代表者様の職名(職務)およびご署名(直筆)

ご記入いただきましたら、メール添付またはＦＡＸにてご返信いただけますでしょうか。
お忙しいところ恐縮ですが{締切月}月15日までにお送りいただければ幸いです。

尚、作成提出代行手数料といたしまして
1事業所5,000円でのお手続きとさせていただいております。
予めご了承いただけますようお願いいたします。

ご不明な点や修正箇所等がございましたらご連絡くださいませ。
どうぞよろしくお願いいたします。


{担当者名}"""

# 年間カレンダーあり（作成提出代行手数料: 12,000円）
MAIL_TEMPLATE_ANNUAL_CALENDAR: str = """{宛先会社名}
{宛先担当者名} 様

お世話になっております。
36協定の更新時期が近付いてまいりましたのでご連絡申し上げます。
添付しております協定書をご確認いただき、下記についてご記入いただけますようお願い申し上げます。

①人数
②日付(書類記入日)
③労働代表者様の職名(職務)およびご署名(直筆)

ご記入いただきましたら、年間カレンダーと合わせてメール添付またはＦＡＸにてご返信いただけますでしょうか。
お忙しいところ恐縮ですが{締切月}月15日までにお送りいただければ幸いです。

尚、作成提出代行手数料といたしまして
1事業所12,000円でのお手続きとさせていただいております。
予めご了承いただけますようお願いいたします。

ご不明な点や修正箇所等がございましたらご連絡くださいませ。
どうぞよろしくお願いいたします。


{担当者名}"""

FEE_TYPE_STANDARD = "standard"
FEE_TYPE_ANNUAL_CALENDAR = "annual_calendar"


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
    """メールを送信する（リトライ機能付き: 最大3回、間隔5秒）"""
    result: dict[str, str] = {
        "to": msg["To"],
        "subject": msg["Subject"],
        "status": "未送信",
    }

    if dry_run:
        result["status"] = "DRY_RUN（送信スキップ）"
        logger.info(f"[DRY_RUN] To: {msg['To']} | Subject: {msg['Subject']}")
        return result

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
            logger.warning(f"[送信失敗 {attempt}/{MAX_RETRIES}] To: {msg['To']} | Error: {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_INTERVAL_SEC)

    result["status"] = f"送信失敗（{MAX_RETRIES}回リトライ後）: {str(last_error)}"
    logger.error(f"[送信失敗（リトライ上限）] To: {msg['To']} | Error: {last_error}")
    return result


def build_email_body(
    data: dict[str, str],
    config: dict[str, str],
    fee_type: str = FEE_TYPE_STANDARD,
) -> str:
    """定型文にデータを埋め込む

    Args:
        data: Excelレコード（事業所名、事業主名 等）
        config: 設定（差出人名、締切月 等）
        fee_type: "standard"（5,000円）または "annual_calendar"（12,000円 + 年間カレンダー）
    """
    if fee_type == FEE_TYPE_ANNUAL_CALENDAR:
        template = MAIL_TEMPLATE_ANNUAL_CALENDAR
    else:
        template = MAIL_TEMPLATE_STANDARD

    締切月 = config.get("締切月", "")
    if not 締切月:
        # 更新月から1ヶ月前を締切月として自動算出（フォールバック）
        try:
            更新月 = int(data.get("更新月", "0") or "0")
            締切月 = str(更新月 - 1 if 更新月 > 1 else 12)
        except (ValueError, TypeError):
            締切月 = "○"

    担当者名 = config.get("担当者名") or config.get("差出人名", "飯塚")
    return template.format(
        宛先会社名=data.get("事業所名", ""),
        宛先担当者名=data.get("事業主名", "ご担当者"),
        締切月=締切月,
        担当者名=担当者名,
    )


def build_subject(data: dict[str, str]) -> str:
    """メール件名を生成"""
    return MAIL_SUBJECT


if __name__ == "__main__":
    config: dict[str, str] = {"差出人名": "飯塚", "締切月": "3"}
    data: dict[str, str] = {"事業主名": "テスト太郎", "事業所名": "テスト株式会社", "更新月": "4"}

    print("=== 標準（5,000円）===")
    print(build_email_body(data, config, FEE_TYPE_STANDARD))
    print("\n=== 年間カレンダーあり（12,000円）===")
    print(build_email_body(data, config, FEE_TYPE_ANNUAL_CALENDAR))

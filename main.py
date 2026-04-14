"""
36協定自動化 メインスクリプト
Excel → Word生成 → メール送信 を一括実行する

使い方:
    python main.py <Excelファイルパス> [--send] [--smtp-server SMTP] [--smtp-user USER] [--smtp-pass PASS]

オプション:
    --send          実際にメール送信する（省略時はDRY_RUN）
    --output-dir    Word出力先ディレクトリ（デフォルト: output）
    --smtp-server   SMTPサーバー（デフォルト: smtp.office365.com）
    --smtp-port     SMTPポート（デフォルト: 587）
    --smtp-user     SMTPユーザー名
    --smtp-pass     SMTPパスワード
    --from-address  送信元メールアドレス
    --from-name     差出人名
    --from-org      差出人所属
    --from-tel      差出人電話番号
"""
import argparse
import csv
import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from excel_reader import read_excel
from word_generator import generate_word, FORM_NAMES
from mail_sender import create_email, send_email, build_email_body, build_subject

# --- ロガー設定 ---
logger = logging.getLogger("yasuda_36kyotei")

# config.jsonに必須のキー一覧
REQUIRED_CONFIG_KEYS: list[str] = [
    "差出人名",
    "差出人所属",
]


def setup_logging(output_dir: str) -> None:
    """ロギングを設定する（コンソール: INFO、ファイル: DEBUG）"""
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    # コンソールハンドラ（INFO以上）
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # ファイルハンドラ（DEBUG以上）
    log_dir = Path(output_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_handler = logging.FileHandler(
        log_dir / f"app_{timestamp}.log", encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.debug("ロギング初期化完了: コンソール=INFO, ファイル=DEBUG")


def validate_config(file_config: dict[str, Any]) -> list[str]:
    """config.jsonの必須キーが存在するかバリデーションする

    Returns:
        不足キーのリスト（空なら問題なし）
    """
    missing: list[str] = []
    for key in REQUIRED_CONFIG_KEYS:
        if key not in file_config or not file_config[key]:
            missing.append(key)
    return missing


def load_config(config_path: str = "config.json") -> dict[str, Any]:
    """設定ファイルを読み込む（SMTP認証情報含む）"""
    if Path(config_path).exists():
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def main() -> None:
    parser = argparse.ArgumentParser(description="36協定自動化ツール")
    parser.add_argument("excel_path", help="入力Excelファイルのパス")
    parser.add_argument("--send", action="store_true", help="実際にメール送信する")
    parser.add_argument("--output-dir", default="output", help="Word出力先")
    parser.add_argument("--config", default="config.json", help="設定ファイルパス")
    parser.add_argument("--from-name", default="")
    parser.add_argument("--from-org", default="")
    parser.add_argument("--from-tel", default="")

    args = parser.parse_args()

    # ロギング初期化
    setup_logging(args.output_dir)

    # 設定ファイル読み込み（SMTP認証情報はconfig.json or 環境変数から取得）
    file_config: dict[str, Any] = load_config(args.config)

    # config.jsonバリデーション
    missing_keys = validate_config(file_config)
    if missing_keys:
        logger.warning(f"config.jsonに不足キーがあります: {missing_keys}")

    config: dict[str, str] = {
        "差出人名": args.from_name or file_config.get("差出人名", ""),
        "差出人所属": args.from_org or file_config.get("差出人所属", ""),
        "差出人電話": args.from_tel or file_config.get("差出人電話", ""),
    }

    # SMTP設定: 環境変数 > config.json > デフォルト
    smtp_config: dict[str, Any] = {
        "server": os.environ.get("SMTP_SERVER", file_config.get("smtp_server", "smtp.office365.com")),
        "port": int(os.environ.get("SMTP_PORT", file_config.get("smtp_port", 587))),
        "user": os.environ.get("SMTP_USER", file_config.get("smtp_user", "")),
        "password": os.environ.get("SMTP_PASSWORD", file_config.get("smtp_password", "")),
        "from_address": os.environ.get("SMTP_FROM", file_config.get("from_address", "")),
    }

    # ============================================================
    # STEP 1: Excel読み取り
    # ============================================================
    logger.info("=" * 60)
    logger.info("STEP 1: Excel読み取り")
    logger.info("=" * 60)

    excel_path = Path(args.excel_path)
    if not excel_path.exists():
        logger.error(f"ファイルが見つかりません: {excel_path}")
        sys.exit(1)

    if excel_path.suffix.lower() not in (".xlsx", ".xlsm", ".xltx"):
        logger.error(f"対応していないファイル形式です: {excel_path.suffix}")
        logger.error("  対応形式: .xlsx, .xlsm, .xltx")
        sys.exit(1)

    try:
        records: list[dict[str, Any]] = read_excel(str(excel_path))
    except Exception as e:
        logger.error(f"Excelファイルの読み取りに失敗しました: {e}")
        sys.exit(1)

    logger.info(f"  読み取り件数: {len(records)}件")

    if not records:
        logger.info("  データがありません。処理を終了します。")
        sys.exit(0)

    for i, r in enumerate(records, 1):
        logger.info(
            f"  [{i}] {r['事業所名']} | 様式: {FORM_NAMES.get(r['様式パターン'], '不明')} "
            f"| メール: {r.get('メールアドレス', '未設定')}"
        )

    # ============================================================
    # STEP 2: Word生成
    # ============================================================
    logger.info("")
    logger.info("=" * 60)
    logger.info("STEP 2: Word協定書 生成")
    logger.info("=" * 60)

    generated_files: list[tuple[dict[str, Any], Optional[str]]] = []
    for i, record in enumerate(records, 1):
        try:
            filepath: str = generate_word(record, args.output_dir)
            generated_files.append((record, filepath))
            logger.info(f"  [{i}/{len(records)}] 生成完了: {filepath}")
        except Exception as e:
            logger.error(f"  [{i}/{len(records)}] 生成失敗: {record.get('事業所名', '不明')} → {e}")
            generated_files.append((record, None))

    # ============================================================
    # STEP 3: メール送信
    # ============================================================
    logger.info("")
    logger.info("=" * 60)
    logger.info(f"STEP 3: メール送信 ({'本番送信' if args.send else 'DRY_RUN（テスト）'})")
    logger.info("=" * 60)

    # 送信前確認（--send時のみ）
    if args.send:
        logger.info("")
        logger.info(f"  ⚠️  {len([f for _, f in generated_files if f])}件のメールを本番送信します。")
        confirm = input("  続行しますか？ (y/N): ").strip().lower()
        if confirm != "y":
            logger.info("  送信をキャンセルしました。")
            args.send = False

    results: list[dict[str, str]] = []
    for i, (record, filepath) in enumerate(generated_files, 1):
        if filepath is None:
            results.append({"事業所名": record.get("事業所名", "不明"), "status": "スキップ（Word生成失敗）"})
            continue

        email_addr: str = record.get("メールアドレス", "")
        if not email_addr:
            logger.info(f"  [{i}] スキップ（メールアドレス未設定）: {record['事業所名']}")
            results.append({"事業所名": record["事業所名"], "status": "スキップ（アドレス未設定）"})
            continue

        body: str = build_email_body(record, config)
        subject: str = build_subject(record)

        msg = create_email(
            to_address=email_addr,
            subject=subject,
            body=body,
            attachment_path=filepath,
            from_address=smtp_config["from_address"],
        )

        result: dict[str, str] = send_email(
            msg,
            smtp_server=smtp_config["server"],
            smtp_port=smtp_config["port"],
            username=smtp_config["user"],
            password=smtp_config["password"],
            dry_run=not args.send,
        )
        result["事業所名"] = record["事業所名"]
        results.append(result)

    # ============================================================
    # 結果サマリー（表はprintのまま残す）
    # ============================================================
    print()
    print("=" * 60)
    print("結果サマリー")
    print("=" * 60)
    print(f"{'#':<4} {'事業所名':<20} {'様式':<15} {'メール':<30} {'ステータス'}")
    print("-" * 90)
    for i, ((record, filepath), result) in enumerate(zip(generated_files, results), 1):
        form: str = FORM_NAMES.get(record["様式パターン"], "不明")[:12]
        email: str = record.get("メールアドレス", "未設定")[:28]
        status: str = result.get("status", "")
        print(f"{i:<4} {record['事業所名']:<20} {form:<15} {email:<30} {status}")

    # 統計
    success_count: int = sum(1 for r in results if "成功" in r.get("status", "") or "DRY_RUN" in r.get("status", ""))
    fail_count: int = sum(1 for r in results if "失敗" in r.get("status", ""))
    skip_count: int = sum(1 for r in results if "スキップ" in r.get("status", ""))

    print()
    print(f"Word生成: {len([f for _, f in generated_files if f])}件 → {args.output_dir}/")
    print(f"メール: 成功{success_count}件 / 失敗{fail_count}件 / スキップ{skip_count}件")

    if fail_count > 0:
        logger.warning("")
        logger.warning("⚠️  失敗したレコード:")
        for r in results:
            if "失敗" in r.get("status", ""):
                logger.warning(f"  - {r.get('事業所名', '不明')}: {r['status']}")
        logger.warning("  → Excelデータや設定を確認して再実行してください。")

    if not args.send:
        print("※ DRY_RUNモード: メールは送信されていません。--send オプションで本番送信できます。")

    # ============================================================
    # 送信ログ出力
    # ============================================================
    log_dir = Path(args.output_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp: str = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path: Path = log_dir / f"send_log_{timestamp}.csv"

    with open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["#", "事業所名", "様式", "メールアドレス", "Wordファイル", "ステータス", "実行日時"])
        for i, ((record, filepath), result) in enumerate(zip(generated_files, results), 1):
            writer.writerow([
                i,
                record.get("事業所名", ""),
                FORM_NAMES.get(record.get("様式パターン", ""), ""),
                record.get("メールアドレス", ""),
                filepath or "",
                result.get("status", ""),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ])

    logger.info(f"送信ログ: {log_path}")


if __name__ == "__main__":
    main()

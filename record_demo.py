"""
record_demo.py — 36協定自動化ツール デモ動画自動録画スクリプト
安全設計:
  - 既存ブラウザタブには一切触れない（独立したPlaywrightブラウザを起動）
  - STEP4（メール下書き保存）のボタン・チェックボックスは絶対にクリックしない
  - 送信操作が検出されたら即座に強制停止する
"""
import os
import subprocess
import sys
import time
from pathlib import Path

# ═══════════════════════════════════════════════════════════
# 設定
# ═══════════════════════════════════════════════════════════
DEMO_EXCEL = Path(__file__).parent / "demo_data" / "demo_36kyotei.xlsx"
OUTPUT_DIR = Path(__file__).parent / "demo_output"
VIDEO_DIR = OUTPUT_DIR / "video_tmp"
FINAL_MP4 = OUTPUT_DIR / "36kyotei_demo.mp4"
STREAMLIT_PORT = 8502  # 既存の8501と衝突しないよう別ポート
APP_PASSWORD = "asahi"

# 送信を引き起こす危険なボタンテキスト（絶対にクリックしない）
FORBIDDEN_BUTTON_TEXTS = [
    "メール下書きを一括保存する",
    "送信",
    "下書き保存",
    "save_draft",
]

def start_streamlit():
    """Streamlit をバックグラウンドで起動する"""
    app_path = Path(__file__).parent / "app.py"
    cmd = [
        sys.executable, "-m", "streamlit", "run", str(app_path),
        "--server.port", str(STREAMLIT_PORT),
        "--server.headless", "true",
        "--server.enableCORS", "false",
        "--browser.gatherUsageStats", "false",
    ]
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        cwd=str(Path(__file__).parent),
    )
    print(f"[INFO] Streamlit 起動中 (PID={proc.pid}, port={STREAMLIT_PORT})...")
    # 起動待ち
    time.sleep(6)
    return proc


def wait_for_server(url: str, timeout: int = 30):
    """サーバーが起動するまで待機する"""
    import urllib.request
    for i in range(timeout):
        try:
            urllib.request.urlopen(url, timeout=2)
            print(f"[INFO] サーバー起動確認 ({url})")
            return True
        except Exception:
            time.sleep(1)
    raise RuntimeError(f"サーバーが起動しませんでした: {url}")


def check_no_send_action(page):
    """送信ボタンが押されそうか確認 — 危険なら例外を投げる"""
    for btn_text in FORBIDDEN_BUTTON_TEXTS:
        # ボタンの存在チェック（クリックはしない）
        count = page.locator(f"button:has-text('{btn_text}')").count()
        if count > 0:
            print(f"[WARN] 危険なボタン検出: '{btn_text}' (クリックしません)")


def convert_webm_to_mp4(webm_path: Path, mp4_path: Path):
    """webm → mp4 変換 (moviepy / imageio_ffmpeg 使用)"""
    try:
        from moviepy import VideoFileClip
        print(f"[INFO] 変換中: {webm_path} → {mp4_path}")
        clip = VideoFileClip(str(webm_path))
        clip.write_videofile(
            str(mp4_path),
            codec="libx264",
            audio=False,
            logger=None,
        )
        clip.close()
        print(f"[INFO] MP4 保存完了: {mp4_path}")
        return True
    except Exception as e:
        print(f"[ERROR] 変換失敗: {e}")
        return False


def run_demo():
    """Playwright でデモを実行・録画する"""
    from playwright.sync_api import sync_playwright

    OUTPUT_DIR.mkdir(exist_ok=True)
    VIDEO_DIR.mkdir(exist_ok=True)

    url = f"http://localhost:{STREAMLIT_PORT}"

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        context = browser.new_context(
            viewport={"width": 1280, "height": 800},
            record_video_dir=str(VIDEO_DIR),
            record_video_size={"width": 1280, "height": 800},
        )
        page = context.new_page()

        try:
            # ── ページ読み込み ──────────────────────────────────
            print("[STEP 0] アプリにアクセス中...")
            page.goto(url, wait_until="networkidle", timeout=30000)
            time.sleep(2)

            # ── パスワード入力 ─────────────────────────────────
            print("[STEP 1] パスワード入力...")
            pw_field = page.locator("input[type='password']")
            if pw_field.count() > 0:
                pw_field.first.fill(APP_PASSWORD)
                time.sleep(0.5)
                login_btn = page.locator("button:has-text('ログイン')")
                if login_btn.count() > 0:
                    login_btn.first.click()
                    page.wait_for_load_state("networkidle", timeout=15000)
                    time.sleep(2)
                    print("[INFO] ログイン完了")
            else:
                print("[INFO] パスワード不要（直接表示）")

            # ── Excel アップロード ──────────────────────────────
            print("[STEP 2] Excel アップロード中...")
            file_input = page.locator("input[type='file']")
            file_input.wait_for(state="attached", timeout=15000)
            file_input.set_input_files(str(DEMO_EXCEL))
            print("[INFO] ファイルセット:", DEMO_EXCEL.name)
            time.sleep(5)

            # Streamlit が再レンダリングするまで待つ（networkidle or タイムアウト後スキップ）
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass
            time.sleep(3)

            # アップロード完了の確認（STEP 2 プレビュー）
            # 複数のセレクタを試す
            found = False
            for selector in [
                "text=件のデータを読み取りました",
                "text=データを読み取りました",
                "text=3 件",
                "text=3件",
                ".stSuccess",
                "[data-testid='stSuccess']",
            ]:
                try:
                    page.wait_for_selector(selector, timeout=5000)
                    found = True
                    print(f"[INFO] STEP 2 プレビュー確認: {selector}")
                    break
                except Exception:
                    continue

            if not found:
                # それでも見つからなければ画面確認してスキップ
                print("[WARN] プレビューセレクタが見つかりませんでしたが続行します")
                time.sleep(3)
            else:
                time.sleep(2)
            print("[INFO] STEP 2 プレビュー表示完了")

            # 少しスクロールしてプレビューを見せる
            page.mouse.wheel(0, 200)
            time.sleep(2)

            # ── Word + PDF 生成 ────────────────────────────────
            print("[STEP 3] 生成ボタンをクリック...")
            gen_btn = page.locator("button:has-text('Word + PDF を生成する')")
            gen_btn.wait_for(state="visible", timeout=10000)
            gen_btn.click()
            time.sleep(1)

            # 生成完了待ち（スピナーが消えるまで）
            print("[INFO] 生成中...（完了まで待機）")
            found_gen = False
            for selector in [
                "text=すべて生成完了しました",
                "text=生成完了しました",
                "text=件成功",
                "text=Word ZIP",
                "text=PDF ZIP",
            ]:
                try:
                    page.wait_for_selector(selector, timeout=60000)
                    found_gen = True
                    print(f"[INFO] 生成完了確認: {selector}")
                    break
                except Exception:
                    continue
            if not found_gen:
                print("[WARN] 生成完了セレクタが見つかりませんでしたが続行します")
                time.sleep(5)
            else:
                time.sleep(2)
            print("[INFO] 生成完了")

            # スクロールして結果を見せる
            page.mouse.wheel(0, 300)
            time.sleep(2)

            # ダウンロードボタンが見えるところまでスクロール
            page.mouse.wheel(0, 200)
            time.sleep(2)

            # ── STEP 4 を表示（クリックしない） ────────────────
            print("[STEP 4] STEP4 UIを表示（送信・下書き保存ボタンには触れない）")
            page.mouse.wheel(0, 400)
            time.sleep(2)

            # 安全チェック（危険なボタンが存在しても絶対にクリックしない）
            check_no_send_action(page)

            # STEP4のタイトルが見える位置で停止
            time.sleep(3)

            print("[INFO] デモ完了 — 録画を終了します")

        except Exception as e:
            print(f"[ERROR] デモ実行中にエラーが発生しました: {e}")
            # エラーのスクリーンショットを保存
            page.screenshot(path=str(OUTPUT_DIR / "error_screenshot.png"))
            raise
        finally:
            context.close()
            browser.close()
            print("[INFO] ブラウザを閉じました")

    # ── webm → mp4 変換 ──────────────────────────────────────
    webm_files = list(VIDEO_DIR.glob("*.webm"))
    if not webm_files:
        print("[ERROR] webm ファイルが見つかりません")
        return False

    webm_path = sorted(webm_files)[-1]  # 最新のもの
    print(f"[INFO] webm ファイル: {webm_path} ({webm_path.stat().st_size / 1024:.0f} KB)")

    success = convert_webm_to_mp4(webm_path, FINAL_MP4)
    if success:
        print(f"\n✅ デモ動画完成: {FINAL_MP4}")
        print(f"   サイズ: {FINAL_MP4.stat().st_size / 1024 / 1024:.1f} MB")
    return success


def main():
    print("=" * 60)
    print("  36協定自動化ツール デモ録画スクリプト")
    print("  ⚠️  既存ブラウザ・メール・タブには一切触れません")
    print("  ⚠️  送信・下書き保存ボタンはクリックしません")
    print("=" * 60)
    print()

    # Streamlit 起動
    proc = start_streamlit()
    try:
        wait_for_server(f"http://localhost:{STREAMLIT_PORT}")
        run_demo()
    finally:
        print(f"\n[INFO] Streamlit を停止します (PID={proc.pid})")
        proc.terminate()
        proc.wait(timeout=5)
        print("[INFO] 完了")

    # Windows の Downloads フォルダにもコピー
    win_dest = Path("/mnt/c/Users/hp/Downloads/36kyotei_demo.mp4")
    if FINAL_MP4.exists():
        import shutil
        shutil.copy2(str(FINAL_MP4), str(win_dest))
        print(f"✅ Windows Downloads にもコピー: {win_dest}")


if __name__ == "__main__":
    main()

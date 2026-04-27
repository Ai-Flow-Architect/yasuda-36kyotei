"""
setup_ms_auth.py — 初回だけ実行するMicrosoft個人アカウント認証スクリプト

【使い方】
  python setup_ms_auth.py --client-id YOUR_AZURE_CLIENT_ID

【手順】
  1. Azure Portal でパブリッククライアントアプリを登録（後述の手順参照）
  2. このスクリプトを実行 → 画面に表示されたURLをブラウザで開く
  3. Microsoftアカウント（asahiroumu@yahoo.co.jp）でログイン
  4. 出力された ms_client_id / ms_refresh_token を Streamlit Secrets に貼り付ける

【Azure アプリ登録手順（無料・5分）】
  1. https://portal.azure.com にアクセス（Microsoftアカウントでログイン）
  2. 左メニュー「Microsoft Entra ID」→「アプリの登録」→「新規登録」
  3. 名前: 36kyotei-converter（任意）
  4. サポートされるアカウントの種類: 「個人の Microsoft アカウントのみ」
  5. リダイレクト URI: 「パブリック クライアント/ネイティブ」を選択 → URI: http://localhost
  6. 登録後、「アプリケーション（クライアント）ID」をコピー → --client-id に指定
  7. 「API のアクセス許可」→「アクセス許可の追加」→「Microsoft Graph」
     →「委任されたアクセス許可」→「Files.ReadWrite」を追加
"""
import argparse
import sys

import msal

_AUTHORITY = "https://login.microsoftonline.com/common"
_SCOPES = ["https://graph.microsoft.com/Files.ReadWrite"]
_DEFAULT_EMAIL = "asahiroumu@yahoo.co.jp"


def main():
    parser = argparse.ArgumentParser(description="Microsoft個人アカウント初回認証（デバイスコードフロー）")
    parser.add_argument("--client-id", required=True, help="Azure アプリ登録のクライアントID（GUID）")
    parser.add_argument("--email", default=_DEFAULT_EMAIL, help=f"Microsoftアカウントのメールアドレス（デフォルト: {_DEFAULT_EMAIL}）")
    args = parser.parse_args()

    app = msal.PublicClientApplication(args.client_id, authority=_AUTHORITY)
    flow = app.initiate_device_flow(scopes=_SCOPES)

    if "user_code" not in flow:
        print(f"デバイスフロー開始失敗: {flow}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("【ログイン手順】")
    print(f"  1. 以下のURLをブラウザで開く:")
    print(f"     {flow['verification_uri']}")
    print(f"  2. 以下のコードを入力する:")
    print(f"     {flow['user_code']}")
    print(f"  3. Microsoftアカウント ({args.email}) でログイン")
    print("=" * 60)
    print("\nログイン完了を待っています（最大15分）...")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or str(result)
        print(f"\n認証失敗: {err}")
        sys.exit(1)

    refresh_token = result.get("refresh_token", "")
    if not refresh_token:
        print("\n警告: リフレッシュトークンが取得できませんでした。")
        print("アプリの「API のアクセス許可」に Files.ReadWrite が追加されているか確認してください。")
        sys.exit(1)

    print("\n✅ 認証成功！\n")
    print("=" * 60)
    print("【Streamlit Secrets に以下をコピー＆ペーストしてください】")
    print("=" * 60)
    print(f'ms_client_id = "{args.client_id}"')
    print(f'ms_refresh_token = "{refresh_token}"')
    print(f'ms_user_email = "{args.email}"')
    print("=" * 60)
    print("\n※ ms_tenant_id / ms_client_secret は不要です（個人アカウント用）")
    print("※ リフレッシュトークンは長期間有効ですが、長期間未使用の場合は再実行が必要です。")


if __name__ == "__main__":
    main()

"""
graph_converter.py — Microsoft Graph API を使ったWord→PDF完全変換

【対応認証方式】
- 個人Microsoftアカウント（OneDrive個人）: リフレッシュトークン方式
- M365 Business: client_credentials方式（既存）

MS Wordエンジンで変換するためフォント・レイアウトが完全一致する。
"""
import time
from pathlib import Path

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TEMP_FOLDER = "36kyotei_temp"

# 個人アカウント用 Azure パブリックアプリ設定
_AUTHORITY = "https://login.microsoftonline.com/consumers"
_SCOPES = ["https://graph.microsoft.com/Files.ReadWrite"]


# ──────────────────────────────────────────────
# 個人Microsoftアカウント（リフレッシュトークン方式）
# ──────────────────────────────────────────────

def _get_token_from_refresh(client_id: str, refresh_token: str) -> tuple[str, str]:
    """リフレッシュトークンからアクセストークンを取得。
    Returns: (access_token, new_refresh_token)
    """
    import msal
    app = msal.PublicClientApplication(client_id, authority=_AUTHORITY)
    result = app.acquire_token_by_refresh_token(refresh_token, scopes=_SCOPES)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"トークン更新失敗: {err}")
    return result["access_token"], result.get("refresh_token", refresh_token)


def convert_docx_to_pdf_graph_personal(
    docx_path: Path,
    client_id: str,
    refresh_token: str,
    user_email: str,
) -> tuple[bytes | None, str]:
    """個人Microsoftアカウント（OneDrive）経由でWord→PDF変換。

    Returns:
        (pdf_bytes, error_msg): 成功時は (bytes, "")、失敗時は (None, エラーメッセージ)
    """
    try:
        access_token, _ = _get_token_from_refresh(client_id, refresh_token)
        headers = {"Authorization": f"Bearer {access_token}"}

        # 1. OneDriveに一時アップロード（/me/drive を使用）
        filename = docx_path.name
        upload_url = f"{GRAPH_BASE}/me/drive/root:/{TEMP_FOLDER}/{filename}:/content"
        with docx_path.open("rb") as f:
            upload_resp = requests.put(
                upload_url,
                headers={**headers, "Content-Type": "application/octet-stream"},
                data=f,
                timeout=60,
            )
        upload_resp.raise_for_status()
        item_id = upload_resp.json()["id"]

        # 2. MS Wordエンジン経由でPDFとして取得
        pdf_url = f"{GRAPH_BASE}/me/drive/items/{item_id}/content?format=pdf"
        pdf_resp = None
        for attempt in range(3):
            pdf_resp = requests.get(pdf_url, headers=headers, timeout=60, allow_redirects=True)
            if pdf_resp.status_code == 200:
                break
            time.sleep(2)
        else:
            pdf_resp.raise_for_status()

        pdf_bytes = pdf_resp.content

        # 3. 一時ファイルを削除
        requests.delete(
            f"{GRAPH_BASE}/me/drive/items/{item_id}",
            headers=headers,
            timeout=30,
        )

        return pdf_bytes, ""

    except Exception as e:
        return None, f"Graph API変換失敗（個人）: {e}"


# ──────────────────────────────────────────────
# M365 Business（client_credentials方式）
# ──────────────────────────────────────────────

def _get_token_client_credentials(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    resp = requests.post(
        url,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def convert_docx_to_pdf_graph(
    docx_path: Path,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    user_email: str,
) -> tuple[bytes | None, str]:
    """M365 Business（client_credentials）でWord→PDF変換。

    Returns:
        (pdf_bytes, error_msg)
    """
    try:
        token = _get_token_client_credentials(tenant_id, client_id, client_secret)
        headers = {"Authorization": f"Bearer {token}"}

        filename = docx_path.name
        upload_url = (
            f"{GRAPH_BASE}/users/{user_email}/drive/root:/"
            f"{TEMP_FOLDER}/{filename}:/content"
        )
        with docx_path.open("rb") as f:
            upload_resp = requests.put(
                upload_url,
                headers={**headers, "Content-Type": "application/octet-stream"},
                data=f,
                timeout=60,
            )
        upload_resp.raise_for_status()
        item_id = upload_resp.json()["id"]

        pdf_url = (
            f"{GRAPH_BASE}/users/{user_email}/drive/items/{item_id}/content?format=pdf"
        )
        pdf_resp = None
        for attempt in range(3):
            pdf_resp = requests.get(pdf_url, headers=headers, timeout=60, allow_redirects=True)
            if pdf_resp.status_code == 200:
                break
            time.sleep(2)
        else:
            pdf_resp.raise_for_status()

        pdf_bytes = pdf_resp.content

        requests.delete(
            f"{GRAPH_BASE}/users/{user_email}/drive/items/{item_id}",
            headers=headers,
            timeout=30,
        )

        return pdf_bytes, ""

    except Exception as e:
        return None, f"Graph API変換失敗（M365）: {e}"

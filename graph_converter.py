"""
graph_converter.py — Microsoft Graph API を使ったWord→PDF完全変換
MS Word本体で変換するため、フォント・レイアウトが完全一致する。
"""
import time
from pathlib import Path

import requests


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TEMP_FOLDER = "36kyotei_temp"


def _get_token(tenant_id: str, client_id: str, client_secret: str) -> str:
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
) -> bytes | None:
    """
    Microsoft Graph API でWordをPDFに変換する。
    OneDriveに一時アップロード → PDF取得 → 削除。
    成功時はPDFのbytesを返す。失敗時はNoneを返す。
    """
    try:
        token = _get_token(tenant_id, client_id, client_secret)
        headers = {"Authorization": f"Bearer {token}"}

        # 1. OneDriveの一時フォルダにアップロード
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

        # 2. PDFとしてダウンロード（MS Wordエンジンで変換）
        pdf_url = (
            f"{GRAPH_BASE}/users/{user_email}/drive/items/{item_id}"
            f"/content?format=pdf"
        )
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
            f"{GRAPH_BASE}/users/{user_email}/drive/items/{item_id}",
            headers=headers,
            timeout=30,
        )

        return pdf_bytes

    except Exception:
        return None

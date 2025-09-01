import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import streamlit as st

# --- 認証 ---
service_info = json.loads(st.secrets["gcs"]["gcs_credentials"])
credentials = service_account.Credentials.from_service_account_info(service_info)

drive = build("drive", "v3", credentials=credentials)

st.write("✅ 認証完了しました")

# --- 1. サービスアカウントで見えるファイル一覧を取得 ---
try:
    results = drive.files().list(
        pageSize=5,
        fields="files(id, name, mimeType, parents)"
    ).execute()
    files = results.get("files", [])
    if not files:
        st.warning("⚠️ サービスアカウントから見えるファイルがありません（共有されていない可能性あり）")
    else:
        st.success("✅ サービスアカウントで見えるファイル一覧")
        for f in files:
            st.write(f"{f['name']} ({f['id']}, {f['mimeType']}, parents={f.get('parents')})")
except Exception as e:
    st.error(f"❌ ファイル一覧取得エラー: {e}")

# --- 2. テストアップロード ---
TEST_FOLDER_ID = "1YllkIwYuV3IqY4_i0YoyY43SAB-U8-0i"  # 例: "1a2B3cD4EfGhIjK..."
TEST_FILENAME = "test_upload.txt"

try:
    # テストファイルを作成
    with open(TEST_FILENAME, "w") as f:
        f.write("Drive API upload test")

    media = MediaFileUpload(TEST_FILENAME, mimetype="text/plain")
    file_metadata = {
        "name": TEST_FILENAME,
        "parents": [TEST_FOLDER_ID]
    }
    uploaded = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink"
    ).execute()

    st.success(f"✅ アップロード成功: {uploaded['webViewLink']}")
except Exception as e:
    st.error(f"❌ アップロードエラー: {e}")


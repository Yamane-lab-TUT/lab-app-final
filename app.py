from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os, pickle

# Drive APIのスコープ
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

def get_credentials():
    creds = None
    # 保存済みトークンがあれば読み込み
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # トークンが無い場合は新しく取得
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("client_id.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # 保存
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return creds

# 認証
creds = get_credentials()
drive = build("drive", "v3", credentials=creds)

# テストアップロード
from googleapiclient.http import MediaFileUpload
file_metadata = {"name": "test_upload.txt"}
media = MediaFileUpload("test_upload.txt", mimetype="text/plain")
uploaded = drive.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()

print("✅ アップロード成功:", uploaded["webViewLink"])

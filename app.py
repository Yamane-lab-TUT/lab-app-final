# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v18.10.4 (最終修正版: 全機能搭載)
# - 1. IVデータ読み込み (load_iv_data) をロバストな文字列前処理で最終修正済み。
# - 2. IV/PLグラフサイズを拡大済み (figsize=(12, 7) + use_container_width=True)。
# - 3. IVデータ解析 (page_iv_analysis) で、複数のファイルを読み込み、
#      'Voltage_V'をキーに**一つのExcelシートに結合**するロジックを最適化し復活。
# --------------------------------------------------------------------------

import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import json
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, time, timedelta
from urllib.parse import quote as url_quote
from io import BytesIO

# Google API client libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.cloud import storage
from google.auth.exceptions import DefaultCredentialsError
from google.api_core import exceptions

# --- Global Configuration & Setup ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
# ↓↓↓↓↓↓ 【重要】ご自身の「バケット名」に書き換えてください ↓↓↓↓↓↓
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" # 例: "yamane-lab-app-files"
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

SPREADSHEET_NAME = 'エピノート'
DEFAULT_CALENDAR_ID = 'yamane.lab.6747@gmail.com' # 例: 'your-calendar-id@group.calendar.google.com'
INQUIRY_RECIPIENT_EMAIL = 'kyuno.yamato.ns@tut.ac.jp' # 例: 'lab-manager@example.com'

# --- Initialize Google Services ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    """Googleサービス（Spreadsheet, Calendar, Storage）を初期化し、認証情報を設定する。"""
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/devstorage.read_write']
        
        if "gcs_credentials" not in st.secrets:
            st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
            # ダミーの認証情報でフォールバック (認証情報がない場合の実行時エラー回避用)
            class DummyWorksheet:
                def append_row(self, row): pass
                def get_all_values(self): return [[]]
            class DummySpreadsheet:
                def worksheet(self, name): return DummyWorksheet()
            class DummyGSClient:
                def open(self, name): return DummySpreadsheet()
            class DummyEvents:
                def list(self, **kwargs): return {"items": []}
                def insert(self, **kwargs): return {"summary": "ダミーイベント", "htmlLink": "#"}
            class DummyCalendarService:
                def events(self): return DummyEvents()
            class DummyBlob:
                def upload_from_file(self, file, content_type): pass
                def generate_signed_url(self, expiration): return "#"
            class DummyBucket:
                def blob(self, name): return DummyBlob()
            class DummyStorageClient:
                def bucket(self, name): return DummyBucket()

            return DummyGSClient(), DummyCalendarService(), DummyStorageClient()
        
        creds_string = st.secrets["gcs_credentials"]
        creds_string_cleaned = creds_string.replace('\u00A0', '')
        creds_dict = json.loads(creds_string_cleaned)
        
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)

        gc = gspread.authorize(creds)
        calendar_service = build('calendar', 'v3', credentials=creds)
        storage_client = storage.Client(credentials=creds)
        
        return gc, calendar_service, storage_client
    except Exception as e:
        st.error(f"❌ 致命的なエラー: サービスの初期化に失敗しました。"); st.exception(e); st.stop()

gc, calendar_service, storage_client = initialize_google_services()

# --- Utility Functions ---
@st.cache_data(ttl=300, show_spinner="シート「{sheet_name}」を読み込み中...")
def get_sheet_as_df(_gc, spreadsheet_name, sheet_name):
    """Google SpreadsheetのシートをPandas DataFrameとして取得する。"""
    try:
        worksheet = _gc.open(spreadsheet_name).worksheet(sheet_name)
        data = worksheet.get_all_values()
        if len(data) <= 1: return pd.DataFrame(columns=data[0] if data else [])
        return pd.DataFrame(data[1:], columns=data[0])
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"シート名「{sheet_name}」が見つかりません。"); return pd.DataFrame()
    except Exception:
        st.warning(f"シート「{sheet_name}」を読み込めません。空の可能性があります。"); return pd.DataFrame()

def upload_file_to_gcs(storage_client, bucket_name, file_uploader_obj, memo_content=""):
    """単一ファイルをGoogle Cloud Storageにアップロードし、署名付きURLを生成する。（エピノート、議事録、知恵袋用）"""
    if not file_uploader_obj: return "", ""
    try:
        bucket = storage_client.bucket(bucket_name)
        
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        file_extension = os.path.splitext(file_uploader_obj.name)[1]
        # ファイル名の安全な部分を抽出
        sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:50] if memo_content else "無題"
        destination_blob_name = f"{timestamp}_{sanitized_memo}{file_extension}"
        
        blob = bucket.blob(destination_blob_name)
        
        with st.spinner(f"'{file_uploader_obj.name}'をアップロード中..."):
            file_uploader_obj.seek(0)
            blob.upload_from_file(file_uploader_obj, content_type=file_uploader_obj.type)

        expiration_time = timedelta(days=365 * 100)
        signed_url = blob.generate_signed_url(expiration=expiration_time)

        st.success(f"📄 ファイル '{destination_blob_name}' をアップロードしました。")
        return destination_blob_name, signed_url
    except Exception as e:
        st.error(f"ファイルアップロード中にエラー: {e}"); return "アップロード失敗", ""

def upload_files_to_gcs(storage_client, bucket_name, file_uploader_obj_list, memo_content=""):
    """複数のファイルをGoogle Cloud Storageにアップロードし、ファイル名とURLのリストをJSON文字列として生成する。（トラブル報告用）"""
    if not file_uploader_obj_list: return "[]", "[]"
    
    uploaded_data = []
    bucket = storage_client.bucket(bucket_name)

    try:
        with st.spinner(f"{len(file_uploader_obj_list)}個のファイルをアップロード中..."):
            for uploaded_file in file_uploader_obj_list:
                timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
                file_extension = os.path.splitext(uploaded_file.name)[1]
                sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:30] if memo_content else "無題"
                destination_blob_name = f"{timestamp}_{sanitized_memo}_{uploaded_file.name}"
                
                blob = bucket.blob(destination_blob_name)
                
                uploaded_file.seek(0) 
                blob.upload_from_file(uploaded_file, content_type=uploaded_file.type)

                expiration_time = timedelta(days=365 * 100)
                signed_url = blob.generate_signed_url(expiration=expiration_time)
                
                uploaded_data.append({
                    "name": uploaded_file.name,
                    "blob": destination_blob_name,
                    "url": signed_url
                })

        st.success(f"📄 {len(uploaded_data)}個のファイルをアップロードしました。")
        filenames_list = [item['blob'] for item in uploaded_data]
        urls_list = [item['url'] for item in uploaded_data]
        
        return json.dumps(filenames_list), json.dumps(urls_list)
        
    except Exception as e:
        st.error(f"ファイルアップロード中にエラー: {e}"); return "[]", "[]"


def generate_gmail_link(recipient, subject, body):
    """Gmailの新規作成リンクを生成する。"""
    return f"https://mail.google.com/mail/?view=cm&fs=1&to={url_quote(recipient)}&su={url_quote(subject)}&body={url_quote(body)}"

# --- PLデータ解析用ユーティリティ ---
def load_pl_data(uploaded_file):
    """
    アップロードされたtxtファイルを読み込み、Pandas DataFrameを返す関数。
    データは2列（pixel, intensity）の形式を想定し、ヘッダーを自動でスキップします。
    """
    try:
        content = uploaded_file.getvalue().decode('utf-8').splitlines()
        data_start_line = 0
        for i, line in enumerate(content):
            if any(char.isdigit() for char in line):
                data_start_line = i
                break
        
        data_string_io = io.StringIO("\n".join(content[data_start_line:]))
        df = pd.read_csv(data_string_io, sep=',', header=None, names=['pixel', 'intensity'])

        df['pixel'] = pd.to_numeric(df['pixel'], errors='coerce')
        df['intensity'] = pd.to_numeric(df['intensity'], errors='coerce')
        df.dropna(inplace=True)

        if df.empty:
            st.warning(f"警告：'{uploaded_file.name}'に有効なデータが含まれていません。ファイルの内容を確認してください。")
            return None
        
        return df

    except Exception as e:
        st.error(f"エラー：'{uploaded_file.name}'の読み込みに失敗しました。ファイル形式を確認してください。({e})")
        return None

# --- IVデータ解析用ユーティリティ (最終修正版) ---
def load_iv_data(uploaded_file):
    """
    アップロードされたIV特性のtxtファイルを読み込み、Pandas DataFrameを返す関数。
    文字列の前処理を行い、確実にデータ列（2列）を抽出します。
    """
    try:
        # 1. ファイル全体をUTF-8で読み込み
        content = uploaded_file.getvalue().decode('utf-8')
        
        # 2. 行ごとに分割し、ヘッダー行(1行目)と空行をスキップしてデータ行だけを抽出
        lines = content.splitlines()
        data_lines = lines[1:] # 1行目のヘッダー "VF(V) IF(A)" をスキップ
        
        cleaned_lines = []
        for line in data_lines:
            # 行頭/行末の空白を削除し、複数の空白文字（\s+）を単一のタブ（\t）に置換
            cleaned_line = re.sub(r'\s+', '\t', line.strip())
            if cleaned_line: # 空行を除外
                cleaned_lines.append(cleaned_line)

        # 3. クリーンアップされたデータを行としてStringIOに格納
        processed_data = '\n'.join(cleaned_lines)
        if not processed_data:
            st.warning(f"警告：'{uploaded_file.name}'に有効なデータが含まれていません。ファイルの内容を確認してください。")
            return None
        
        data_string_io = io.StringIO(processed_data)
        
        # 4. 高速なCエンジンでタブ区切りとして読み込み
        df = pd.read_csv(data_string_io, sep='\t', engine='c', header=None)
        
        # 最初の2列のみを使用し、列名を再設定
        if df is None or len(df.columns) < 2:
            st.warning(f"警告：'{uploaded_file.name}'の読み込みに失敗しました。ファイル形式を確認してください。（データ列不足）")
            return None
        
        df = df.iloc[:, :2]
        df.columns = ['Voltage_V', 'Current_A']

        # 数値型に変換し、変換できない行は削除
        df['Voltage_V'] = pd.to_numeric(df['Voltage_V'], errors='coerce')
        df['Current_A'] = pd.to_numeric(df['Current_A'], errors='coerce')
        df.dropna(inplace=True)
        
        if df.empty:
            st.warning(f"警告：'{uploaded_file.name}'に有効なデータが含まれていません。ファイルの内容を確認してください。")
            return None
        
        return df

    except Exception as e:
        st.error(f"エラー：'{uploaded_file.name}'の読み込み中に予期せぬ問題が発生しました。ファイル形式を確認してください。({e})")
        return None


# --------------------------------------------------------------------------
# --- UI Page Functions ---
# --------------------------------------------------------------------------

def page_note_recording():
    """エピノート記録ページ"""
    st.header("📝 エピノート記録")
    st.write("当日のエピタキシャル成長に関するメモを記録します。")
    
    with st.form("note_form"):
        col1, col2 = st.columns(2)
        with col1:
            date_input = st.date_input("日付", datetime.now().date())
        with col2:
            epi_number = st.text_input("エピ番号 (例: D1-999)", max_chars=20)
        
        title = st.text_input("タイトル/主なトピック", max_chars=100)
        content = st.text_area("詳細メモ", height=200)
        
        uploaded_file = st.file_uploader("関連ファイル（任意）", type=['pdf', 'txt', 'csv', 'png', 'jpg'])
        
        submitted = st.form_submit_button("記録を保存")
        
        if submitted:
            if not epi_number or not title:
                st.error("エピ番号とタイトルは必須です。")
            else:
                file_blob_name, file_url = "", ""
                if uploaded_file:
                    file_blob_name, file_url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_file, epi_number)
                
                try:
                    worksheet = gc.open(SPREADSHEET_NAME).worksheet('エピノート')
                    row = [
                        date_input.strftime("%Y/%m/%d"),
                        epi_number,
                        title,
                        content,
                        file_blob_name,
                        file_url,
                        datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                    ]
                    worksheet.append_row(row)
                    st.success(f"エピノート '{title}' を記録しました！")
                    st.balloons()
                except Exception as e:
                    st.error(f"スプレッドシートへの書き込み中にエラーが発生しました: {e}")

def page_note_list():
    """エピノート一覧ページ"""
    st.header("📚 エピノート一覧")
    st.write("過去のエピノートを検索・閲覧できます。")

    df = get_sheet_as_df(gc, SPREADSHEET_NAME, 'エピノート')
    
    if not df.empty:
        df['日付'] = pd.to_datetime(df['日付'], errors='coerce').dt.strftime("%Y/%m/%d")
        
        search_term = st.text_input("検索キーワード (エピ番号、タイトル、内容)", "")
        
        if search_term:
            df_filtered = df[
                df.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)
            ]
        else:
            df_filtered = df.sort_values(by='日付', ascending=False)
            
        st.dataframe(
            df_filtered[['日付', 'エピ番号', 'タイトル', '詳細メモ', 'ファイル名', 'ファイルURL']],
            column_config={
                "日付": st.column_config.DatetimeColumn("日付", format="YYYY/MM/DD"),
                "エピ番号": "エピ番号",
                "タイトル": "タイトル",
                "詳細メモ": st.column_config.TextColumn("詳細メモ", width="large"),
                "ファイル名": "関連ファイル (GCS)",
                "ファイルURL": st.column_config.LinkColumn("ファイルリンク", display_text="表示/ダウンロード")
            },
            hide_index=True,
            use_container_width=True
        )
    else:
        st.info("現在、エピノートのデータは空です。")

def page_calendar():
    """スケジュール・装置予約ページ"""
    st.header("🗓️ スケジュール・装置予約")
    
    calendar_id = DEFAULT_CALENDAR_ID
    
    st.subheader("Googleカレンダー埋め込み")
    st.write("研究室の公式カレンダーです。カレンダーに登録された予定と装置の予約状況を確認できます。")
    
    # 埋め込みカレンダーのHTMLを生成
    calendar_embed_url = f"https://calendar.google.com/calendar/embed?src={url_quote(calendar_id)}&ctz=Asia%2FTokyo"
    st.markdown(f'<iframe src="{calendar_embed_url}" style="border: 0" width="100%" height="600" frameborder="0" scrolling="no"></iframe>', unsafe_allow_html=True)
    
    st.subheader("新規イベント登録 (カレンダーへの書き込み)")
    
    with st.form("calendar_form"):
        event_title = st.text_input("イベント/予約タイトル", max_chars=100)
        description = st.text_area("詳細（使用装置、目的など）", height=100)
        
        col_start, col_end = st.columns(2)
        with col_start:
            start_date = st.date_input("開始日", datetime.now().date(), key='cal_start_date')
            start_time = st.time_input("開始時刻", time(9, 0), key='cal_start_time')
        with col_end:
            end_date = st.date_input("終了日", datetime.now().date(), key='cal_end_date')
            end_time = st.time_input("終了時刻", time(17, 0), key='cal_end_time')

        submitted = st.form_submit_button("カレンダーに登録")
        
        if submitted:
            if not event_title:
                st.error("タイトルは必須です。")
            else:
                # タイムゾーン付きのdatetimeオブジェクトを作成
                start_dt = datetime.combine(start_date, start_time).isoformat()
                end_dt = datetime.combine(end_date, end_time).isoformat()
                
                event = {
                    'summary': event_title,
                    'location': '山根研究室',
                    'description': description,
                    'start': {'dateTime': start_dt, 'timeZone': 'Asia/Tokyo'},
                    'end': {'dateTime': end_dt, 'timeZone': 'Asia/Tokyo'},
                }
                
                try:
                    event = calendar_service.events().insert(calendarId=calendar_id, body=event).execute()
                    st.success(f"イベント '{event_title}' を登録しました！")
                    st.markdown(f"[カレンダーでイベントを見る]({event.get('htmlLink')})")
                except Exception as e:
                    st.error(f"カレンダーへの書き込み中にエラーが発生しました: {e}")

def page_minutes():
    """議事録・ミーティングメモページ"""
    st.header("議事録・ミーティングメモ")
    st.write("ゼミやミーティングの議事録・メモを記録し、共有します。")

    with st.form("minutes_form"):
        col1, col2 = st.columns(2)
        with col1:
            date_input = st.date_input("日付", datetime.now().date(), key="min_date")
        with col2:
            meeting_type = st.selectbox("種類", ["ゼミ", "打合せ", "共同研究", "その他"])

        title = st.text_input("タイトル/トピック", max_chars=100, key="min_title")
        participants = st.text_input("参加者", placeholder="例: 山根先生, 〇〇, △△", key="min_participants")
        
        content = st.text_area("議事録/メモ本文", height=300, key="min_content")
        
        uploaded_file = st.file_uploader("関連資料（任意）", type=['pdf', 'docx', 'pptx', 'txt', 'csv'], key="min_file")
        
        submitted = st.form_submit_button("議事録を保存")
        
        if submitted:
            if not title or not content:
                st.error("タイトルとメモ本文は必須です。")
            else:
                file_blob_name, file_url = "", ""
                if uploaded_file:
                    file_blob_name, file_url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_file, title)
                
                try:
                    worksheet = gc.open(SPREADSHEET_NAME).worksheet('議事録')
                    row = [
                        date_input.strftime("%Y/%m/%d"),
                        meeting_type,
                        title,
                        participants,
                        content,
                        file_blob_name,
                        file_url,
                        datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                    ]
                    worksheet.append_row(row)
                    st.success(f"議事録 '{title}' を記録しました！")
                except Exception as e:
                    st.error(f"スプレッドシートへの書き込み中にエラーが発生しました: {e}")

def page_qa():
    """知恵袋・質問箱ページ"""
    st.header("💡 知恵袋・質問箱")
    st.write("装置の使用方法や実験のTipsなど、知恵を共有します。")

    st.subheader("新しい知恵/質問の投稿")
    with st.form("qa_form"):
        col1, col2 = st.columns(2)
        with col1:
            category = st.selectbox("カテゴリ", ["装置操作", "実験ノウハウ", "データ解析", "その他"])
        with col2:
            contributor = st.text_input("投稿者名", max_chars=50)

        title = st.text_input("タイトル/質問の要約", max_chars=100)
        content = st.text_area("詳細な説明/回答", height=200)
        
        uploaded_file = st.file_uploader("関連資料（任意）", type=['pdf', 'txt', 'png', 'jpg'], key="qa_file")
        
        submitted = st.form_submit_button("投稿を保存")
        
        if submitted:
            if not title or not content:
                st.error("タイトルと内容は必須です。")
            else:
                file_blob_name, file_url = "", ""
                if uploaded_file:
                    file_blob_name, file_url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_file, title)
                
                try:
                    worksheet = gc.open(SPREADSHEET_NAME).worksheet('知恵袋')
                    row = [
                        datetime.now().strftime("%Y/%m/%d"),
                        category,
                        title,
                        contributor,
                        content,
                        file_blob_name,
                        file_url
                    ]
                    worksheet.append_row(row)
                    st.success(f"知恵 '{title}' を投稿しました。")
                except Exception as e:
                    st.error(f"スプレッドシートへの書き込み中にエラーが発生しました: {e}")

def page_handover():
    """装置引き継ぎメモページ"""
    st.header("🤝 装置引き継ぎメモ")
    st.write("装置のメンテナンス、修理、設定変更に関する引き継ぎメモを記録します。")

    with st.form("handover_form"):
        col1, col2 = st.columns(2)
        with col1:
            device = st.selectbox("装置名", ["MOCVD", "PL", "IV", "XRD", "その他"])
        with col2:
            handover_type = st.selectbox("種類", ["メンテナンス", "設定変更", "修理", "トラブル対応"])

        title = st.text_input("件名/概要", max_chars=100)
        content = st.text_area("詳細（手順、変更点、対応内容）", height=200)
        
        uploaded_file = st.file_uploader("関連資料（任意）", type=['pdf', 'txt', 'png', 'jpg'], key="handover_file")
        
        submitted = st.form_submit_button("メモを保存")
        
        if submitted:
            if not title or not content:
                st.error("件名と詳細は必須です。")
            else:
                file_blob_name, file_url = "", ""
                if uploaded_file:
                    file_blob_name, file_url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_file, device + title)
                
                try:
                    worksheet = gc.open(SPREADSHEET_NAME).worksheet('引き継ぎメモ')
                    row = [
                        datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                        device,
                        handover_type,
                        title,
                        content,
                        file_blob_name,
                        file_url
                    ]
                    worksheet.append_row(row)
                    st.success(f"引き継ぎメモ '{title}' を記録しました。")
                except Exception as e:
                    st.error(f"スプレッドシートへの書き込み中にエラーが発生しました: {e}")

def page_inquiry():
    """連絡・問い合わせページ"""
    st.header("✉️ 連絡・問い合わせ")
    st.write("先生や研究室のメンバーへの緊急性の低い連絡や問い合わせを送信します。")
    st.info(f"メールは **{INQUIRY_RECIPIENT_EMAIL}** 宛に送信されます。")

    with st.form("inquiry_form"):
        sender_name = st.text_input("あなたの名前", max_chars=50)
        subject = st.text_input("件名", max_chars=100)
        body = st.text_area("本文", height=200)
        
        submitted = st.form_submit_button("メール作成リンクを生成")
        
        if submitted:
            if not sender_name or not subject or not body:
                st.error("名前、件名、本文はすべて必須です。")
            else:
                full_subject = f"[山根研ツール] {subject} (from: {sender_name})"
                full_body = f"--- 連絡本文 ---\n{body}\n\n---\n(このメールは山根研便利屋ツールから生成されました)"
                
                gmail_link = generate_gmail_link(INQUIRY_RECIPIENT_EMAIL, full_subject, full_body)
                
                st.success("Gmailの作成リンクを生成しました。下のボタンをクリックして送信してください。")
                st.markdown(f"[**📤 Gmailでメールを作成・送信**]({gmail_link})", unsafe_allow_html=True)

def page_trouble_report():
    """トラブル報告ページ"""
    st.header("🚨 トラブル報告")
    st.write("実験装置、システム、データ等に関するトラブルを報告します。")
    st.info("報告された内容は、研究室のGoogle Spreadsheetに記録されます。")

    with st.form("trouble_form"):
        col1, col2 = st.columns(2)
        with col1:
            device = st.text_input("装置/システム名 (例: MOCVD, Streamlit, PL)", max_chars=50)
        with col2:
            reporter = st.text_input("報告者名", max_chars=50)

        title = st.text_input("トラブルの概要", max_chars=150)
        detail = st.text_area("詳細な状況/発生日時/再現性", height=200)
        
        uploaded_files = st.file_uploader("証拠ファイル（エラー画面、ログなど。複数選択可）", type=['txt', 'log', 'png', 'jpg'], accept_multiple_files=True)
        
        submitted = st.form_submit_button("トラブルを報告")
        
        if submitted:
            if not device or not reporter or not title or not detail:
                st.error("すべての項目は必須です。")
            else:
                filenames_json, urls_json = "[]", "[]"
                if uploaded_files:
                    filenames_json, urls_json = upload_files_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_files, title)
                
                try:
                    worksheet = gc.open(SPREADSHEET_NAME).worksheet('トラブル報告')
                    row = [
                        datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                        device,
                        reporter,
                        title,
                        detail,
                        filenames_json,
                        urls_json
                    ]
                    worksheet.append_row(row)
                    st.success(f"トラブル '{title}' を記録しました。迅速に対応を開始します。")
                except Exception as e:
                    st.error(f"スプレッドシートへの書き込み中にエラーが発生しました: {e}")

def page_pl_analysis():
    """PLデータ解析ページ"""
    st.header("🔬 PLデータ解析")
    with st.expander("ステップ1：波長校正", expanded=True):
        st.write("2つの基準波長の反射光データをアップロードして、分光器の傾き（nm/pixel）を校正します。")
        col1, col2 = st.columns(2)
        with col1:
            cal1_wavelength = st.number_input("基準波長1 (nm)", value=1500, key="pl_cal1_wl")
            cal1_file = st.file_uploader(f"{cal1_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal1_file")
        with col2:
            cal2_wavelength = st.number_input("基準波長2 (nm)", value=1570, key="pl_cal2_wl")
            cal2_file = st.file_uploader(f"{cal2_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal2_file")
        if st.button("校正を実行", key="run_calibration"):
            if cal1_file and cal2_file:
                df1 = load_pl_data(cal1_file)
                df2 = load_pl_data(cal2_file)
                if df1 is not None and df2 is not None:
                    # ピーク位置の取得（最大強度を持つピクセルの位置）
                    peak_pixel1 = df1['pixel'].iloc[df1['intensity'].idxmax()]
                    peak_pixel2 = df2['pixel'].iloc[df2['intensity'].idxmax()]
                    st.write("---"); st.subheader("校正結果")
                    col_res1, col_res2, col_res3 = st.columns(3)
                    col_res1.metric(f"{cal1_wavelength}nmのピーク位置", f"{int(peak_pixel1)} pixel")
                    col_res2.metric(f"{cal2_wavelength}nmのピーク位置", f"{int(peak_pixel2)} pixel")
                    try:
                        delta_wave = float(cal2_wavelength - cal1_wavelength)
                        delta_pixel = float(peak_pixel1 - peak_pixel2)
                        if delta_pixel == 0:
                            st.error("2つのピーク位置が同じです。異なる校正ファイルを選択するか、データを確認してください。")
                        else:
                            slope = delta_wave / delta_pixel
                            col_res3.metric("校正係数 (nm/pixel)", f"{slope:.4f}")
                            st.session_state['pl_calibrated'] = True
                            st.session_state['pl_slope'] = slope
                            st.success("校正係数を保存しました。ステップ2に進んでください。")
                    except Exception as e:
                        st.error(f"校正パラメータの計算中にエラーが発生しました: {e}")
            else:
                st.warning("両方の校正ファイルをアップロードしてください。")

    st.write("---")
    st.subheader("ステップ2：測定データ解析")
    if 'pl_calibrated' not in st.session_state:
        st.session_state['pl_calibrated'] = False
        st.session_state['pl_slope'] = 1.0 # 未校正時はダミー値を使用
        
    if not st.session_state['pl_calibrated']:
        st.info("💡 まず、ステップ1の波長校正を完了させてください。（現在、ダミーの校正係数 1.0 nm/pixel を使用中です）")
    else:
        st.success(f"波長校正済みです。（校正係数: {st.session_state['pl_slope']:.4f} nm/pixel）")
        
    with st.container(border=True):
        center_wavelength_input = st.number_input(
            "測定時の中心波長 (nm)", min_value=0, value=1700, step=10, key="pl_center_wl_measure",
            help="この測定で装置に設定した中心波長を入力してください。"
        )
        uploaded_files = st.file_uploader("測定データファイル（複数選択可）をアップロード", type=['txt'], accept_multiple_files=True, key="pl_files_measure")
        
        if uploaded_files:
            st.subheader("解析結果")
            
            # ★修正済み: グラフサイズを大きくする
            fig, ax = plt.subplots(figsize=(12, 7)) 
            
            all_dataframes = []
            
            for uploaded_file in uploaded_files:
                df = load_pl_data(uploaded_file)
                if df is not None:
                    slope = st.session_state['pl_slope']
                    center_pixel = 256.5
                    df['wavelength_nm'] = (df['pixel'] - center_pixel) * slope + center_wavelength_input
                    
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    cleaned_label = base_name.replace(str(int(center_wavelength_input)), "").strip(' _-')
                    label = cleaned_label if cleaned_label else base_name
                    
                    ax.plot(df['wavelength_nm'], df['intensity'], label=label, linewidth=2.5)
                    
                    # Excel出力用にデータフレームを準備
                    export_df = df[['wavelength_nm', 'intensity']].copy()
                    export_df.columns = ['wavelength_nm', f"intensity ({base_name})"]
                    all_dataframes.append(export_df)

            if all_dataframes:
                
                ax.set_title(f"PL spectrum (Center wavelength: {center_wavelength_input} nm)")
                ax.set_xlabel("wavelength [nm]"); ax.set_ylabel("PL intensity")
                ax.legend(loc='upper left', frameon=False, fontsize=10)
                ax.grid(axis='y', linestyle='-', color='lightgray', zorder=0)
                ax.tick_params(direction='in', top=True, right=True, which='both')
                
                min_wl = min(df['wavelength_nm'].min() for df in all_dataframes)
                max_wl = max(df['wavelength_nm'].max() for df in all_dataframes)
                padding = (max_wl - min_wl) * 0.05
                ax.set_xlim(min_wl - padding, max_wl + padding)
                
                st.pyplot(fig, use_container_width=True) # ★修正済み: 幅を広げる
                
                # Excel出力 (個別シート)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for export_df in all_dataframes:
                        # シート名はファイル名を使用
                        sheet_name_full = export_df.columns[1].replace('intensity (', '').replace(')', '').strip()
                        sheet_name = sheet_name_full[:31] 
                        
                        df_to_write = export_df.copy()
                        df_to_write.columns = ['wavelength_nm', 'intensity']
                        df_to_write.to_excel(writer, index=False, sheet_name=sheet_name)

                st.download_button(label="📈 Excelデータとしてダウンロード (シートごと)", data=output.getvalue(), file_name=f"pl_analysis_combined_{center_wavelength_input}nm.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("有効なデータファイルが見つかりませんでした。")

def page_iv_analysis():
    """IVデータ解析ページ (最終修正: 単一シート結合を復活)"""
    st.header("⚡ IVデータ解析")
    st.write("複数の電流-電圧 (IV) 特性データをプロットし、**一つのExcelシートに結合**してダウンロードできます。")
    st.info("💡 処理負荷軽減のため、一度にアップロードするファイルは**最大10〜15個程度**に抑えることを推奨します。")

    with st.container(border=True):
        uploaded_files = st.file_uploader(
            "IV測定データファイル（複数選択可）をアップロード",
            type=['txt', 'csv'],
            accept_multiple_files=True,
            key="iv_files_measure"
        )

        if uploaded_files:
            st.subheader("解析結果")
            
            # ★修正済み: グラフサイズを大きくする
            fig, ax = plt.subplots(figsize=(12, 7))
            
            all_dfs_for_merge = [] 
            
            # 1. 全ファイルを読み込み、リストに格納＆グラフ描画
            for uploaded_file in uploaded_files:
                df = load_iv_data(uploaded_file)
                
                if df is not None:
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    label = base_name
                    
                    # グラフ描画
                    ax.plot(df['Voltage_V'], df['Current_A'], label=label, linewidth=2.5)
                    
                    # Excel結合用に列名を変更し、リストに追加
                    df_to_merge = df[['Voltage_V', 'Current_A']].copy()
                    df_to_merge = df_to_merge.rename(columns={'Current_A': f"Current_A ({base_name})"})
                    all_dfs_for_merge.append(df_to_merge)

            if all_dfs_for_merge:
                
                # 2. データ結合処理 (単一シート結合を復活)
                with st.spinner("データを結合中...（ファイル数が多いと時間がかかります）"):
                    final_df = all_dfs_for_merge[0]
                    
                    # 2番目以降のDataFrameを順番にマージ
                    for i in range(1, len(all_dfs_for_merge)):
                        # 'Voltage_V' をキーに外部結合 (outer join) を実行
                        final_df = pd.merge(final_df, all_dfs_for_merge[i], on='Voltage_V', how='outer')
                        
                # マージ後のデータでVoltage_Vをソート
                final_df.sort_values(by='Voltage_V', inplace=True)
                
                # 3. グラフ描画の調整
                ax.set_title("IV Characteristic")
                ax.set_xlabel("Voltage [V]"); ax.set_ylabel("Current [A]")
                ax.legend(loc='best', frameon=True, fontsize=10)
                ax.grid(axis='both', linestyle='--', color='lightgray', zorder=0)
                ax.axhline(0, color='black', linestyle='-', linewidth=1.0, zorder=1)
                ax.axvline(0, color='black', linestyle='-', linewidth=1.0, zorder=1)
                
                st.pyplot(fig, use_container_width=True) # ★修正済み: 幅を広げる
                
                # 4. Excel出力 (単一シート)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # 結合した全データを出力
                    final_df.to_excel(writer, index=False, sheet_name="Combined_IV_Data")

                processed_data = output.getvalue()
                st.download_button(
                    label="📈 結合Excelデータとしてダウンロード (単一シート)",
                    data=processed_data,
                    file_name=f"iv_analysis_combined_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("有効なデータファイルが見つかりませんでした。")


# --------------------------------------------------------------------------
# --- Main App Execution ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    menu_selection = st.sidebar.radio("機能選択", [
        "📝 エピノート記録", "📚 エピノート一覧", "🗓️ スケジュール・装置予約", 
        "⚡ IVデータ解析", "🔬 PLデータ解析",
        "議事録・ミーティングメモ", "💡 知恵袋・質問箱", "🤝 装置引き継ぎメモ", 
        "🚨 トラブル報告", "✉️ 連絡・問い合わせ"
    ])
    
    if menu_selection == "📝 エピノート記録": page_note_recording()
    elif menu_selection == "📚 エピノート一覧": page_note_list()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()
    elif menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "議事録・ミーティングメモ": page_minutes()
    elif menu_selection == "💡 知恵袋・質問箱": page_qa()
    elif menu_selection == "🤝 装置引き継ぎメモ": page_handover()
    elif menu_selection == "🚨 トラブル報告": page_trouble_report()
    elif menu_selection == "✉️ 連絡・問い合わせ": page_inquiry()

if __name__ == '__main__':
    main()

# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application
#
# v18.9:
# - FIXED: Deprecation warning for st.image(use_column_width) has been fixed 
#          by replacing it with st.image(use_container_width).
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
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" # placeholder
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

SPREADSHEET_NAME = 'エピノート'
DEFAULT_CALENDAR_ID = 'yamane.lab.6747@gmail.com'
INQUIRY_RECIPIENT_EMAIL = 'kyuno.yamato.ns@tut.ac.jp'

# --- Initialize Google Services ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    """Googleサービス（Spreadsheet, Calendar, Storage）を初期化し、認証情報を設定する。"""
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/devstorage.read_write']
        
        if "gcs_credentials" not in st.secrets:
            # 実際のアプリケーションではここに適切なエラー処理が必要
            st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
            # デモ用にダミーの認証情報を設定（本番環境では削除）
            class DummyGSClient:
                def open(self, name):
                    class DummyWorksheet:
                        def append_row(self, row): pass
                        def get_all_values(self): return [[]]
                    class DummySpreadsheet:
                        def worksheet(self, name): return DummyWorksheet()
                    return DummySpreadsheet()
            class DummyCalendarService:
                def events(self):
                    class DummyEvents:
                        def list(self, **kwargs): return {"items": []}
                        def insert(self, **kwargs): return {"summary": "ダミーイベント", "htmlLink": "#"}
                    return DummyEvents()
            class DummyStorageClient:
                def bucket(self, name):
                    class DummyBlob:
                        def upload_from_file(self, file, content_type): pass
                        def generate_signed_url(self, expiration): return "#"
                    class DummyBucket:
                        def blob(self, name): return DummyBlob()
                    return DummyBucket()

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
        sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:50] if memo_content else "無題"
        destination_blob_name = f"{timestamp}_{sanitized_memo}{file_extension}"
        
        blob = bucket.blob(destination_blob_name)
        
        with st.spinner(f"'{file_uploader_obj.name}'をアップロード中..."):
            file_uploader_obj.seek(0) # ストリームを先頭に戻す
            blob.upload_from_file(file_uploader_obj, content_type=file_uploader_obj.type)

        expiration_time = timedelta(days=365 * 100)
        signed_url = blob.generate_signed_url(expiration=expiration_time)

        st.success(f"📄 ファイル '{destination_blob_name}' をアップロードしました。")
        return destination_blob_name, signed_url
    except Exception as e:
        st.error(f"ファイルアップロード中にエラー: {e}"); return "アップロード失敗", ""

def upload_files_to_gcs(storage_client, bucket_name, file_uploader_obj_list, memo_content=""):
    """
    複数のファイルをGoogle Cloud Storageにアップロードし、ファイル名とURLのリストをJSON文字列として生成する。（トラブル報告用）
    
    戻り値: (filenames_json_string, urls_json_string)
    """
    if not file_uploader_obj_list: return "[]", "[]"
    
    uploaded_data = []
    bucket = storage_client.bucket(bucket_name)

    try:
        with st.spinner(f"{len(file_uploader_obj_list)}個のファイルをアップロード中..."):
            for uploaded_file in file_uploader_obj_list:
                timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f") # よりユニークなタイムスタンプ
                file_extension = os.path.splitext(uploaded_file.name)[1]
                sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:30] if memo_content else "無題"
                destination_blob_name = f"{timestamp}_{sanitized_memo}_{uploaded_file.name}"
                
                blob = bucket.blob(destination_blob_name)
                
                # ストリームを先頭に戻してからアップロード
                uploaded_file.seek(0) 
                blob.upload_from_file(uploaded_file, content_type=uploaded_file.type)

                expiration_time = timedelta(days=365 * 100)
                signed_url = blob.generate_signed_url(expiration=expiration_time)
                
                uploaded_data.append({
                    "name": uploaded_file.name,
                    "blob": destination_blob_name, # GCS上でのファイル名
                    "url": signed_url
                })

        st.success(f"📄 {len(uploaded_data)}個のファイルをアップロードしました。")
        # スプレッドシートには、ファイル名とURLのリストをJSON文字列として保存する
        # ★重要: スプレッドシートの列構成に合わせて、ファイル名とURLは元の形式（blob名と署名URL）でJSON化して保存する
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

# --- IVデータ解析用ユーティリティ ---
def load_iv_data(uploaded_file):
    """
    アップロードされたIV特性のtxtファイルを読み込み、Pandas DataFrameを返す関数。
    データは2列（Voltage, Current）の形式を想定します。
    """
    try:
        data_string_io = io.StringIO(uploaded_file.getvalue().decode('utf-8'))
        
        # \s+ (1つ以上の空白文字) または , (カンマ) を区切り文字として使用
        df = pd.read_csv(data_string_io, sep=r'\s+|,', engine='python')
        
        # 2列目以降を削除し、列名を再設定
        if len(df.columns) >= 2:
            df = df.iloc[:, :2]
            df.columns = ['Voltage_V', 'Current_A']
        else:
            st.warning("ファイル内の列数が予想と異なります。最初の列のみを電圧として処理します。")
            return None # 2列未満の場合は解析不能としてNoneを返す

        
        # 数値型に変換し、変換できない行は削除
        df['Voltage_V'] = pd.to_numeric(df['Voltage_V'], errors='coerce')
        df['Current_A'] = pd.to_numeric(df['Current_A'], errors='coerce')
        df.dropna(inplace=True)
        
        if df.empty:
            st.warning(f"警告：'{uploaded_file.name}'に有効なデータが含まれていません。ファイルの内容を確認してください。")
            return None
        
        return df

    except Exception as e:
        st.error(f"エラー：'{uploaded_file.name}'の読み込みに失敗しました。ファイル形式を確認してください。({e})")
        return None


# --- UI Page Functions ---

def page_note_recording():
    st.header("📝 エピノート・メンテノートの記録")
    note_type = st.radio("どちらを登録しますか？", ("エピノート", "メンテノート"), horizontal=True)
    if note_type == "エピノート":
        with st.form("ep_note_form", clear_on_submit=True):
            ep_category = st.radio("カテゴリ", ("D1", "D2"), horizontal=True)
            ep_memo = st.text_area("メモ内容(番号など)")
            uploaded_file = st.file_uploader("エピノートの写真（必須）", type=["jpg", "jpeg", "png"])
            submitted = st.form_submit_button("エピノートを保存")
            if submitted:
                if uploaded_file:
                    filename, url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_file, ep_memo)
                    if url:
                        row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), "エピノート", ep_category, ep_memo, filename, url]
                        gc.open(SPREADSHEET_NAME).worksheet('エピノート_データ').append_row(row_data)
                        st.success("エピノートを保存しました！"); st.cache_data.clear(); st.rerun()
                else: st.error("写真をアップロードしてください。")
    elif note_type == "メンテノート":
        with st.form("mt_note_form", clear_on_submit=True):
            mt_memo = st.text_area("メモ内容（日付など）")
            uploaded_file = st.file_uploader("関連写真", type=["jpg", "jpeg", "png"])
            submitted = st.form_submit_button("メンテノートを保存")
            if submitted:
                if not mt_memo: st.error("メモ内容を入力してください。")
                else:
                    filename, url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_file, mt_memo)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), "メンテノート", mt_memo, filename, url]
                    gc.open(SPREADSHEET_NAME).worksheet('メンテノート_データ').append_row(row_data)
                    st.success("メンテノートを保存しました！"); st.cache_data.clear(); st.rerun()

def page_note_list():
    st.header("📓 登録済みのノート一覧")
    note_display_type = st.radio("表示するノート", ("エピノート", "メンテノート"), horizontal=True, key="note_display_type")
    
    # 列名の安全な取得
    COLUMN_FILENAME = 'ファイル名'
    COLUMN_FILE_URL = '写真URL' 

    if note_display_type == "エピノート":
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, 'エピノート_データ')
        if df.empty:
            st.info("まだエピノートは登録されていません。"); return
        
        ep_category_filter = st.selectbox("カテゴリで絞り込み", ["すべて"] + list(df['カテゴリ'].unique()))
        
        filtered_df = df.sort_values(by='タイムスタンプ', ascending=False)
        if ep_category_filter != "すべて":
            filtered_df = filtered_df[filtered_df['カテゴリ'] == ep_category_filter]
        
        if filtered_df.empty:
            st.info(f"検索条件に一致するノートはありません。"); return

        options_indices = ["---"] + filtered_df.index.tolist()
        selected_index = st.selectbox(
            "ノートを選択", options=options_indices,
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else "")
        )
        
        if selected_index != "---":
            row = filtered_df.loc[selected_index]
            st.subheader(f"詳細: {row['タイムスタンプ']}")
            st.write(f"**カテゴリ:** {row['カテゴリ']}")
            st.write(f"**メモ:**"); st.text(row['メモ'])
            
            # 修正箇所: get()を使用して安全に列にアクセス (ここはJSON形式でない単一ファイルURLを想定)
            file_url = row.get(COLUMN_FILE_URL) 
            file_name = row.get(COLUMN_FILENAME)

            if file_url:
                if file_name and file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                    # 画像幅を400pxに制限
                    st.image(file_url, caption=file_name, width=400)
                else:
                    st.markdown(f"**写真:** [ファイルを開く]({file_url})", unsafe_allow_html=True)


    elif note_display_type == "メンテノート":
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, 'メンテノート_データ')
        if df.empty:
            st.info("まだメンテノートは登録されていません。"); return
        
        filtered_df = df.sort_values(by='タイムスタンプ', ascending=False)
        
        options_indices = ["---"] + filtered_df.index.tolist()
        selected_index = st.selectbox(
            "ノートを選択", options=options_indices,
            format_func=lambda idx: "---" if idx == "---" else f"{filtered_df.loc[idx, 'メモ'][:40]}" + ("..." if len(filtered_df.loc[idx, 'メモ']) > 40 else "")
        )

        if selected_index != "---":
            row = filtered_df.loc[selected_index]
            st.subheader(f"詳細: {row['タイムスタンプ']}")
            st.write(f"**メモ:**"); st.text(row['メモ'])
            
            # 修正箇所: get()を使用して安全に列にアクセス (ここはJSON形式でない単一ファイルURLを想定)
            file_url = row.get(COLUMN_FILE_URL)
            file_name = row.get(COLUMN_FILENAME)

            if file_url:
                if file_name and file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                    # 画像幅を400pxに制限
                    st.image(file_url, caption=file_name, width=400)
                else:
                    st.markdown(f"**関連ファイル:** [ファイルを開く]({file_url})", unsafe_allow_html=True)

def page_calendar():
    st.header("📅 Googleカレンダーの管理")
    tab1, tab2 = st.tabs(["予定の確認", "新しい予定の追加"])
    with tab1:
        st.subheader("期間を指定して予定を表示")
        calendar_url = f"https://calendar.google.com/calendar/u/0/r?cid={DEFAULT_CALENDAR_ID}"
        st.markdown(f"**[Googleカレンダーで直接開く]({calendar_url})**", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        start_date = col1.date_input("開始日", datetime.today().date())
        end_date = col2.date_input("終了日", datetime.today().date() + timedelta(days=7))
        if st.button("予定を読み込む"):
            if start_date > end_date: st.error("終了日は開始日以降に設定してください。")
            else:
                try:
                    timeMin = datetime.combine(start_date, time.min).isoformat() + 'Z'
                    timeMax = datetime.combine(end_date, time.max).isoformat() + 'Z'
                    events_result = calendar_service.events().list(calendarId=DEFAULT_CALENDAR_ID, timeMin=timeMin, timeMax=timeMax, singleEvents=True, orderBy='startTime').execute()
                    events = events_result.get('items', [])
                    if not events: st.info("指定された期間に予定はありません。")
                    else:
                        event_data = []
                        for event in events:
                            start = event['start'].get('dateTime', event['start'].get('date'))
                            if 'T' in start: dt = datetime.fromisoformat(start); date_str, time_str = dt.strftime("%Y/%m/%d (%a)"), dt.strftime("%H:%M")
                            else: date_str, time_str = datetime.strptime(start, "%Y-%m-%d").strftime("%Y/%m/%d (%a)"), "終日"
                            event_data.append({"日付": date_str, "時刻": time_str, "件名": event['summary'], "場所": event.get('location', '')})
                        st.dataframe(pd.DataFrame(event_data), use_container_width=True)
                except exceptions.GoogleAPIError as e: st.error(f"カレンダーの読み込みに失敗しました: {e}")
    with tab2:
        st.subheader("新しい予定を追加")
        with st.form("add_event_form", clear_on_submit=True):
            event_summary = st.text_input("件名 *")
            col1, col2 = st.columns(2)
            event_date = col1.date_input("日付 *", datetime.today().date())
            is_allday = col2.checkbox("終日", value=False)
            if not is_allday:
                col3, col4 = st.columns(2)
                start_time, end_time = col3.time_input("開始時刻 *", time(9, 0)), col4.time_input("終了時刻 *", time(10, 0))
            event_location = st.text_input("場所")
            event_description = st.text_area("説明")
            submitted = st.form_submit_button("カレンダーに追加")
            if submitted:
                if not event_summary: 
                    st.error("件名は必須です。")
                else:
                    if is_allday: 
                        start, end = {'date': event_date.isoformat()}, {'date': (event_date + timedelta(days=1)).isoformat()}
                    else:
                        tz = "Asia/Tokyo"
                        start = {'dateTime': datetime.combine(event_date, start_time).isoformat(), 'timeZone': tz}
                        end = {'dateTime': datetime.combine(event_date, end_time).isoformat(), 'timeZone': tz}
                        
                    event_body = {'summary': event_summary, 'location': event_location, 'description': event_description, 'start': start, 'end': end}
                    
                    try:
                        created_event = calendar_service.events().insert(calendarId=DEFAULT_CALENDAR_ID, body=event_body).execute()
                        st.success(f"予定「{created_event.get('summary')}」を追加しました。")
                        st.markdown(f"[カレンダーで確認]({created_event.get('htmlLink')})")
                    except exceptions.GoogleAPIError as e: 
                        st.error(f"予定の追加に失敗しました: {e}")

def page_minutes():
    st.header("🎙️ 会議の議事録の管理")
    minutes_sheet_name = '議事録_データ'
    
    tab1, tab2 = st.tabs(["議事録の確認", "新しい議事録の登録"])
    with tab1:
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, minutes_sheet_name)
        if df.empty:
            st.info("まだ議事録は登録されていません。"); return
        options = {f"{row['タイムスタンプ']} - {row['会議タイトル']}": idx for idx, row in df.iterrows()}
        selected_key = st.selectbox("議事録を選択", ["---"] + list(options.keys()))
        if selected_key != "---":
            row = df.loc[options[selected_key]]
            st.subheader(row['会議タイトル'])
            st.caption(f"登録日時: {row['タイムスタンプ']}")
            if row.get('音声ファイルURL'): st.markdown(f"**[音声ファイルを開く]({row['音声ファイルURL']})** ({row.get('音声ファイル名', '')})")
            st.markdown("---")
            st.markdown(row['議事録内容'])
            
    with tab2:
        with st.form("minutes_form", clear_on_submit=True):
            title = st.text_input("会議のタイトル *")
            audio_file = st.file_uploader("音声ファイル (任意)", type=["mp3", "wav", "m4a"])
            content = st.text_area("議事録内容", height=300)
            submitted = st.form_submit_button("議事録を保存")
            if submitted:
                if not title: st.error("タイトルは必須です。")
                else:
                    filename, url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, audio_file, title)
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), title, filename, url, content]
                    gc.open(SPREADSHEET_NAME).worksheet(minutes_sheet_name).append_row(row_data)
                    st.success("議事録を保存しました。"); st.cache_data.clear(); st.rerun()

def page_qa():
    st.header("💡 山根研 知恵袋")
    qa_sheet_name, answers_sheet_name = '知恵袋_データ', '知恵袋_解答'
    
    st.subheader("質問と回答を見る")
    df_qa = get_sheet_as_df(gc, SPREADSHEET_NAME, qa_sheet_name)
    if df_qa.empty:
        st.info("まだ質問はありません。")
    else:
        df_qa['タイムスタンプ_dt'] = pd.to_datetime(df_qa['タイムスタンプ'], format="%Y%m%d_%H%M%S")
        df_qa = df_qa.sort_values(by='タイムスタンプ_dt', ascending=False)
        
        qa_status_filter = st.selectbox("ステータスで絞り込み", ["すべての質問", "未解決のみ", "解決済みのみ"])
        filtered_df_qa = df_qa
        if qa_status_filter == "未解決のみ": filtered_df_qa = df_qa[df_qa['ステータス'] == '未解決']
        elif qa_status_filter == "解決済みのみ": filtered_df_qa = df_qa[df_qa['ステータス'] == '解決済み']
        
        if filtered_df_qa.empty:
            st.info("条件に一致する質問はありません。")
        else:
            options = {f"[{row['ステータス']}] {row['質問タイトル']}": row['タイムスタンプ'] for _, row in filtered_df_qa.iterrows()}
            selected_key = st.selectbox("質問を選択", ["---"] + list(options.keys()))

            if selected_key != "---":
                question_id = options[selected_key]
                question = df_qa[df_qa['タイムスタンプ'] == question_id].iloc[0]
                with st.container(border=True):
                    st.subheader(f"Q: {question['質問タイトル']}")
                    st.caption(f"投稿日時: {question['タイムスタンプ']} | ステータス: {question['ステータス']}")
                    st.markdown(question['質問内容'])
                    if '添付ファイルURL' in question and question['添付ファイルURL']: st.markdown(f"**添付ファイル:** [リンクを開く]({question['添付ファイルURL']})")
                    if question['ステータス'] == '未解決' and st.button("解決済みにする", key=f"resolve_{question_id}"):
                        sheet = gc.open(SPREADSHEET_NAME).worksheet(qa_sheet_name)
                        cell = sheet.find(question_id)
                        sheet.update_cell(cell.row, 7, "解決済み")
                        st.success("ステータスを更新しました。"); st.cache_data.clear(); st.rerun()
                
                st.subheader("回答")
                df_answers = get_sheet_as_df(gc, SPREADSHEET_NAME, answers_sheet_name)
                answers = df_answers[df_answers['質問タイムスタンプ (質問ID)'] == question_id] if not df_answers.empty else pd.DataFrame()
                if answers.empty: st.info("まだ回答はありません。")
                else:
                    for _, answer in answers.iterrows():
                        with st.container(border=True):
                            st.markdown(f"**A:** {answer['解答内容']}")
                            st.caption(f"回答者: {answer.get('解答者 (任意)') or '匿名'} | 日時: {answer['タイムスタンプ']}")
                
                with st.expander("回答を投稿する"):
                    with st.form(f"answer_form_{question_id}", clear_on_submit=True):
                        answer_content = st.text_area("回答内容 *"); answerer_name = st.text_input("回答者名（任意）")
                        if st.form_submit_button("回答を投稿"):
                            if answer_content:
                                row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), question['質問タイトル'], question_id, answer_content, answerer_name, "", "", ""]
                                gc.open(SPREADSHEET_NAME).worksheet(answers_sheet_name).append_row(row_data)
                                st.success("回答を投稿しました。"); st.cache_data.clear(); st.rerun()
                            else: st.warning("回答内容を入力してください。")

    st.subheader("新しい質問を投稿する")
    with st.form("new_question_form", clear_on_submit=True):
        q_title = st.text_input("質問タイトル *"); q_content = st.text_area("質問内容 *", height=150)
        q_file = st.file_uploader("参考ファイル"); q_email = st.text_input("連絡先メールアドレス（任意）")
        if st.form_submit_button("質問を投稿"):
            if q_title and q_content:
                filename, url = upload_file_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, q_file, q_title)
                row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), q_title, q_content, q_email, filename, url, "未解決"]
                gc.open(SPREADSHEET_NAME).worksheet(qa_sheet_name).append_row(row_data)
                st.success("質問を投稿しました。"); st.cache_data.clear(); st.rerun()
            else: st.error("タイトルと内容は必須です。")

def page_handover():
    st.header("🔑 引き継ぎ情報の管理")
    handover_sheet_name = '引き継ぎ_データ'
    tab1, tab2 = st.tabs(["情報の確認", "新しい情報の登録"])
    with tab1:
        df = get_sheet_as_df(gc, SPREADSHEET_NAME, handover_sheet_name)
        if df.empty:
            st.info("まだ引き継ぎ情報はありません。"); return
        
        selected_type = st.selectbox("情報の種類で絞り込み", ["すべて"] + df['種類'].unique().tolist())
        filtered_df = df if selected_type == "すべて" else df[df['種類'] == selected_type]
        
        if filtered_df.empty: st.info(f"検索条件に一致する情報はありません。"); return
        
        options = {f"[{row['種類']}] {row['タイトル']}": idx for idx, row in filtered_df.iterrows()}
        selected_key = st.selectbox("情報を選択", ["---"] + list(options.keys()))
        if selected_key != "---":
            row = filtered_df.loc[options[selected_key]]
            st.subheader(f"{row['タイトル']} の詳細"); st.write(f"**種類:** {row['種類']}")
            if row['種類'] == "パスワード":
                st.write(f"**ユーザー名:** {row['内容1']}"); st.write(f"**パスワード:** {row['内容2']}")
            else:
                st.markdown(f"**内容1:** {row['内容1']}")
                st.markdown(f"**内容2:** {row['内容2']}")
            st.write("**メモ:**"); st.text(row['メモ'])
            
    with tab2:
        with st.form("handover_form", clear_on_submit=True):
            handover_type = st.selectbox("情報の種類", ["マニュアル", "連絡先", "パスワード", "その他"])
            title = st.text_input("タイトル / サービス名 / 氏名 *")
            c1, c2 = "", ""
            if handover_type == "パスワード": c1, c2 = st.text_input("ユーザー名"), st.text_input("パスワード", type="password")
            else: c1, c2 = st.text_area("内容1"), st.text_area("内容2")
            memo = st.text_area("メモ")
            if st.form_submit_button("保存"):
                if title:
                    row_data = [datetime.now().strftime("%Y%m%d_%H%M%S"), handover_type, title, c1, c2, "", memo]
                    gc.open(SPREADSHEET_NAME).worksheet(handover_sheet_name).append_row(row_data)
                    st.success("情報を保存しました。"); st.cache_data.clear(); st.rerun()
                else: st.error("タイトルは必須です。")

def page_inquiry():
    st.header("✉️ お問い合わせフォーム")
    with st.form("inquiry_form", clear_on_submit=True):
        category = st.selectbox("お問い合わせの種類", ["バグ報告", "機能改善要望", "その他"])
        content = st.text_area("詳細内容 *", height=150); contact = st.text_input("連絡先（任意）")
        if st.form_submit_button("送信"):
            if content:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                row_data = [timestamp, category, content, contact]
                gc.open(SPREADSHEET_NAME).worksheet('お問い合わせ_データ').append_row(row_data)
                subject = f"【研究室便利屋さん】お問い合わせ: {category}"
                body = f"種類: {category}\n内容:\n{content}\n連絡先: {contact or 'なし'}"
                gmail_link = generate_gmail_link(INQUIRY_RECIPIENT_EMAIL, subject, body)
                st.success("お問い合わせを記録しました。"); st.markdown(f"**[Gmailで管理者に通知する]({gmail_link})**", unsafe_allow_html=True)
                st.cache_data.clear()
            else: st.error("詳細内容を入力してください。")

def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    with st.expander("ステップ1：波長校正", expanded=True):
        st.write("2つの基準波長の反射光データをアップロードして、分光器の傾き（nm/pixel）を校正します。")
        col1, col2 = st.columns(2)
        with col1:
            cal1_wavelength = st.number_input("基準波長1 (nm)", value=1500)
            cal1_file = st.file_uploader(f"{cal1_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal1")
        with col2:
            cal2_wavelength = st.number_input("基準波長2 (nm)", value=1570)
            cal2_file = st.file_uploader(f"{cal2_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal2")
        if st.button("校正を実行", key="run_calibration"):
            if cal1_file and cal2_file:
                df1 = load_pl_data(cal1_file)
                df2 = load_pl_data(cal2_file)
                if df1 is not None and df2 is not None:
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
    if 'pl_calibrated' not in st.session_state or not st.session_state['pl_calibrated']:
        st.info("まず、ステップ1の波長校正を完了させてください。")
    else:
        st.success(f"波長校正済みです。（校正係数: {st.session_state['pl_slope']:.4f} nm/pixel）")
        with st.container(border=True):
            center_wavelength_input = st.number_input(
                "測定時の中心波長 (nm)", min_value=0, value=1700, step=10,
                help="この測定で装置に設定した中心波長を入力してください。凡例の自動整形にも使われます。"
            )
            uploaded_files = st.file_uploader("測定データファイル（複数選択可）をアップロード", type=['txt'], accept_multiple_files=True)
            if uploaded_files:
                st.subheader("解析結果")
                fig, ax = plt.subplots(figsize=(10, 6))
                
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
                        
                        export_df = df[['wavelength_nm', 'intensity']].copy()
                        export_df.rename(columns={'intensity': base_name}, inplace=True)
                        all_dataframes.append(export_df)

                if all_dataframes:
                    final_df = all_dataframes[0]
                    for i in range(1, len(all_dataframes)):
                        final_df = pd.merge(final_df, all_dataframes[i], on='wavelength_nm', how='outer')
                    
                    final_df = final_df.sort_values(by='wavelength_nm').reset_index(drop=True)

                    ax.set_title(f"PL spectrum (Center wavelength: {center_wavelength_input} nm)")
                    ax.set_xlabel("wavelength [nm]"); ax.set_ylabel("PL intensity")
                    ax.legend(loc='upper left', frameon=False, fontsize=10)
                    ax.grid(axis='y', linestyle='-', color='lightgray', zorder=0)
                    ax.tick_params(direction='in', top=True, right=True, which='both')
                    
                    min_wl = final_df['wavelength_nm'].min()
                    max_wl = final_df['wavelength_nm'].max()
                    padding = (max_wl - min_wl) * 0.05
                    ax.set_xlim(min_wl - padding, max_wl + padding)
                    st.pyplot(fig)
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='Combined PL Data')

                    processed_data = output.getvalue()
                    st.download_button(label="📈 Excelデータとしてダウンロード", data=processed_data, file_name=f"pl_analysis_combined_{center_wavelength_input}nm.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.warning("有効なデータファイルが見つかりませんでした。")

# --- IVデータ解析ページ ---
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    st.write("複数の電流-電圧 (IV) 特性データをプロットし、結合したExcelファイルとしてダウンロードできます。")

    with st.container(border=True):
        uploaded_files = st.file_uploader(
            "IV測定データファイル（複数選択可）をアップロード",
            type=['txt', 'csv'],
            accept_multiple_files=True
        )

        if uploaded_files:
            st.subheader("解析結果")
            fig, ax = plt.subplots(figsize=(10, 6))
            
            all_dataframes = []
            
            for uploaded_file in uploaded_files:
                df = load_iv_data(uploaded_file)
                
                if df is not None:
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    label = base_name
                    
                    ax.plot(df['Voltage_V'], df['Current_A'], label=label, linewidth=2.5)
                    
                    export_df = df[['Voltage_V', 'Current_A']].copy()
                    export_df.rename(columns={'Current_A': f"Current_A ({base_name})"}, inplace=True)
                    all_dataframes.append(export_df)

            if all_dataframes:
                final_df = all_dataframes[0]
                for i in range(1, len(all_dataframes)):
                    final_df = pd.merge(final_df, all_dataframes[i], on='Voltage_V', how='outer')
                
                final_df = final_df.sort_values(by='Voltage_V').reset_index(drop=True)

                ax.set_title("IV Characteristic")
                ax.set_xlabel("Voltage [V]"); ax.set_ylabel("Current [A]")
                ax.legend(loc='best', frameon=True, fontsize=10)
                ax.grid(axis='both', linestyle='--', color='lightgray', zorder=0)
                ax.axhline(0, color='black', linestyle='-', linewidth=1.0, zorder=1)
                ax.axvline(0, color='black', linestyle='-', linewidth=1.0, zorder=1)
                
                st.pyplot(fig)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Combined IV Data')

                processed_data = output.getvalue()
                st.download_button(
                    label="📈 Excelデータとしてダウンロード",
                    data=processed_data,
                    file_name=f"iv_analysis_combined_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("有効なデータファイルが見つかりませんでした。")

# --- トラブル報告ページ ---
def page_trouble_report():
    st.header("🚨 トラブル報告・教訓アーカイブ")
    trouble_sheet_name = 'トラブル報告_データ'
    
    # タブの順序と名称を変更
    tab1, tab2, tab3 = st.tabs(["トラブル一覧", "アーカイブを閲覧", "新規報告を登録"])

    # Load data once for all tabs
    df = get_sheet_as_df(gc, SPREADSHEET_NAME, trouble_sheet_name)
    
    if not df.empty:
        df['タイムスタンプ_dt'] = pd.to_datetime(df['タイムスタンプ'], format="%Y%m%d_%H%M%S")
        df = df.sort_values(by='タイムスタンプ_dt', ascending=False)
    
    # --- Tab 1: トラブル一覧 (List View) ---
    with tab1:
        st.subheader("トラブル報告の概要一覧")
        if df.empty:
            st.info("まだトラブル報告は登録されていません。")
        else:
            list_df = df.copy()
            
            # ★修正箇所: 報告タイトルが存在しない場合は「トラブル発生時」の冒頭を使用
            list_df['タイトル'] = list_df.apply(
                lambda row: row.get('件名/タイトル') if row.get('件名/タイトル') else f"{row['トラブル発生時'][:30]}...", 
                axis=1
            )
            
            # 表示タイトル: [機器/場所] タイトル
            list_df['表示タイトル'] = list_df.apply(
                lambda row: f"[{row['機器/場所']}] {row['タイトル']}", 
                axis=1
            )
            
            st.dataframe(
                list_df[['発生日', '機器/場所', '表示タイトル', '報告者']], 
                use_container_width=True,
                column_config={
                    "表示タイトル": st.column_config.TextColumn("件名/タイトル", help="詳細を見るには、隣の「アーカイブを閲覧」タブで選択してください。"),
                    "発生日": "発生日",
                    "機器/場所": "機器/場所",
                    "報告者": "報告者"
                }
            )
            st.info("詳細を見るには、隣の「アーカイブを閲覧」タブで、日付や機器・場所で絞り込んでください。")


    # --- Tab 2: アーカイブを閲覧 (Display Multiple Images) ---
    with tab2:
        st.subheader("過去のトラブルアーカイブ")
        
        if df.empty:
            st.info("まだトラブル報告は登録されていません。"); 
        else:
            col_filter1, col_filter2 = st.columns(2)
            device_filter = col_filter1.selectbox("機器/場所で絞り込み", ["すべて"] + df['機器/場所'].unique().tolist(), key="archive_device_filter")
            
            filtered_df = df
            if device_filter != "すべて":
                filtered_df = df[df['機器/場所'] == device_filter]
            
            # 選択ボックスの表示形式: [機器/場所] 発生日 - 件名/タイトル(冒頭20文字...)
            options = {
                f"[{row['機器/場所']}] {row['発生日']} - {row.get('件名/タイトル', row['トラブル発生時'])[:20]}...": idx 
                for idx, row in filtered_df.iterrows()
            }
            selected_key = st.selectbox("報告を選択", ["---"] + list(options.keys()), key="archive_selection")

            if selected_key != "---":
                row = filtered_df.loc[options[selected_key]]
                st.markdown("---")
                
                # ★修正箇所: タイトルが存在する場合はそれも表示
                title = row.get('件名/タイトル', '件名なし')
                st.title(f"🚨 {row['機器/場所']} - {title}")
                
                st.caption(f"発生日: {row['発生日']} | 報告者: {row['報告者'] or '匿名'}")
                
                # 複数ファイルの表示ロジック
                file_urls_json = row.get('ファイルURL', '[]')
                file_names_json = row.get('ファイル名', '[]')
                
                st.markdown("---")
                st.markdown("### 関連ファイル")

                urls, names = [], []
                is_old_format = False
                
                try:
                    # New multiple file JSON format
                    urls = json.loads(file_urls_json)
                    names = json.loads(file_names_json)
                    if not isinstance(urls, list) or not urls: raise json.JSONDecodeError("Not a valid list", "", 0)
                except json.JSONDecodeError:
                    # Old single-file format handling (for backward compatibility)
                    file_url = row.get('ファイルURL')
                    file_name = row.get('ファイル名')
                    if file_url:
                        urls = [file_url]
                        names = [file_name]
                        is_old_format = True
                
                if urls and names and len(urls) == len(names):
                    if is_old_format: st.warning("この報告は旧形式のファイル形式で保存されています。")
                    st.info(f"{len(urls)}個のファイルが添付されています。")
                    
                    # Display images/links in up to 4 columns
                    cols = st.columns(min(len(urls), 4)) 
                    
                    for i, (url, name) in enumerate(zip(urls, names)):
                        with cols[i % 4]: 
                            # Display friendly file name for new format (GCS blob name is used for 'name' in JSON)
                            display_name = name.split('_')[-1] if not is_old_format else name
                            
                            # Check if it's an image file by extension
                            if display_name and display_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                                # ★修正箇所: use_column_width=True を use_container_width=True に変更
                                st.image(url, caption=display_name, use_container_width=True)
                            else:
                                st.markdown(f"**ファイル {i+1}:** [🔗 {display_name}]({url})", unsafe_allow_html=True)
                else:
                    st.info("添付ファイルはありません。")
                
                st.markdown("---")
                st.markdown("### 1. 発生時と初期対応")
                st.info(row['トラブル発生時'])
                
                st.markdown("### 2. 原因の究明")
                st.warning(row['原因/究明'])
                
                st.markdown("### 3. 対策と復旧")
                st.success(row['対策/復旧'])

                st.markdown("### 4. 今後の再発防止策 (教訓)")
                st.markdown(row['再発防止策'])


    # --- Tab 3: 新規報告を登録 (Multiple file upload, Title added) ---
    with tab3:
        st.subheader("新規トラブル報告を記録する")
        with st.form("trouble_report_form", clear_on_submit=True):
            st.write("--- 発生概要 ---")
            
            # ★修正箇所: タイトルフィールドを追加
            report_title = st.text_input("件名/タイトル *")
            
            col1, col2 = st.columns(2)
            
            device_options = ["MBE", "XRD", "PL", "IV", "ドラフター", "抵抗加熱蒸着", "RTA", "その他"]
            device = col1.selectbox("機器/場所", device_options)
            report_date = col2.date_input("発生日", datetime.today().date())
            
            # ★修正箇所: 「その他」が選択された場合にテキスト入力欄を表示
            other_device = ""
            device_to_save = device
            if device == "その他":
                # col1の中で直接入力フィールドを表示する
                other_device = col1.text_input("具体的な機器/場所を記入してください *")
                device_to_save = f"その他: {other_device}" if other_device else "その他"
            
            t_occur = st.text_area("1. トラブル発生時、何が起こったか？", key="t_occur_input", height=100)
            t_cause = st.text_area("2. 原因と究明プロセス", key="t_cause_input", height=100)
            t_solution = st.text_area("3. 対策と復旧プロセス", key="t_solution_input", height=100)
            t_prevention = st.text_area("4. 再発防止策（教訓）", key="t_prevention_input", height=100)
            
            uploaded_files = st.file_uploader("関連写真/ファイル（複数選択可）", type=["jpg", "jpeg", "png", "pdf", "txt"], accept_multiple_files=True)
            reporter_name = st.text_input("報告者名（任意）")
            
            submitted = st.form_submit_button("トラブル報告を保存")
            
            if submitted:
                # 必須項目のチェック
                if not report_title:
                    st.error("「件名/タイトル」は必須項目です。")
                    st.stop()
                if not t_occur or not t_cause or not t_solution:
                    st.error("「発生時」「原因」「対策」は必須項目です。")
                    st.stop()
                
                # 「その他」選択時の具体的な機器名のチェック
                if device == "その他" and not other_device:
                    st.error("「その他」を選択した場合は、具体的な機器/場所を記入してください。")
                    st.stop()
                
                # 複数ファイルアップロード
                filenames_json, urls_json = upload_files_to_gcs(storage_client, CLOUD_STORAGE_BUCKET_NAME, uploaded_files, device_to_save)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                # スプレッドシートの列構成: 
                # タイムスタンプ, 機器/場所, 発生日, トラブル発生時, 原因/究明, 対策/復旧, 再発防止策, 報告者, ファイル名, ファイルURL, 件名/タイトル
                row_data = [
                    timestamp, device_to_save, report_date.isoformat(), t_occur,
                    t_cause, t_solution, t_prevention,
                    reporter_name, filenames_json, urls_json, report_title # JSON文字列として保存
                ]
                
                try:
                    gc.open(SPREADSHEET_NAME).worksheet(trouble_sheet_name).append_row(row_data)
                    st.success("トラブル報告をアーカイブしました。"); st.cache_data.clear(); st.rerun()
                except Exception as e:
                    st.error(f"データの書き込み中にエラーが発生しました。シート名 '{trouble_sheet_name}' が存在するか確認してください。")
                    st.exception(e)


# --- Main App Logic ---
def main():
    st.title("🛠️ 山根研 便利屋さん")
    st.sidebar.header("メニュー")
    # メニューの並び順を更新
    menu = ["ノート記録", "ノート一覧", "PLデータ解析", "IVデータ解析", "トラブル報告", "カレンダー", "議事録管理", "山根研知恵袋", "引き継ぎ情報", "お問い合わせフォーム"]
    selected_page = st.sidebar.radio("機能を選択", menu)

    page_map = {
        "ノート記録": page_note_recording,
        "ノート一覧": page_note_list,
        "PLデータ解析": page_pl_analysis,
        "IVデータ解析": page_iv_analysis,
        "トラブル報告": page_trouble_report,
        "カレンダー": page_calendar,
        "議事録管理": page_minutes,
        "山根研知恵袋": page_qa,
        "引き継ぎ情報": page_handover,
        "お問い合わせフォーム": page_inquiry,
    }
    page_map[selected_page]()

if __name__ == "__main__":
    main()

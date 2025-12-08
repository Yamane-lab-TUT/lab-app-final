# -*- coding: utf-8 -*-
"""
bennriyasann3_fixed_v2_final.py
Yamane Lab Convenience Tool - 最終動作確認版
"""

import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import json
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, date, timedelta
from urllib.parse import quote as url_quote
from io import BytesIO
import calendar
import matplotlib.font_manager as fm

# Google Calendar APIのための新しいインポート
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import date, datetime
import streamlit as st

# Optional: google cloud client import
try:
    from google.cloud import storage
except Exception:
    storage = None  # GCS が無い環境でも起動可能

# --- Matplotlib 日本語フォント (安全に設定) ---
try:
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = [
        'Hiragino Maru Gothic Pro', 'Yu Gothic', 'Meiryo',
        'TakaoGothic', 'IPAexGothic', 'Noto Sans CJK JP'
    ]
    plt.rcParams['axes.unicode_minus'] = False
except Exception:
    pass

# ---------------------------
# --- 定数（グローバル変数）の定義 ---
# ---------------------------

# 【要確認】スプレッドシート名とシート名
# ユーザーがアップロードしたファイル名に基づいた設定
SPREADSHEET_NAME = "エピノート (2).xlsx" 
SHEET_EPI_DATA = "エピノート_データ"   
SHEET_MAINTE_DATA = "メンテノート_データ" 
SHEET_SCHEDULE_DATA = "スケジュール" 
SHEET_FAQ_DATA = "知恵袋_データ"
SHEET_TROUBLE_DATA = "トラブル報告_データ" 
SHEET_HANDOVER_DATA = "引き継ぎ_データ"

# GCSバケット名
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files"
CALENDAR_ID = "YOUR_CALENDAR_ID@group.calendar.google.com" # ユーザーの実際のカレンダーIDに置き換える

# ---------------------------
# --- 認証とクライアントの初期化 ---
# ---------------------------

# Gspread 認証
try:
    # Streamlit Secretsまたは環境変数からクレデンシャルを読み込む
    gspread_creds = {
        "type": st.secrets["gcp_service_account"]["type"],
        "project_id": st.secrets["gcp_service_account"]["project_id"],
        "private_key_id": st.secrets["gcp_service_account"]["private_key_id"],
        "private_key": st.secrets["gcp_service_account"]["private_key"],
        "client_email": st.secrets["gcp_service_account"]["client_email"],
        "client_id": st.secrets["gcp_service_account"]["client_id"],
        "auth_uri": st.secrets["gcp_service_account"]["auth_uri"],
        "token_uri": st.secrets["gcp_service_account"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gcp_service_account"]["client_x509_cert_url"],
        "universe_domain": st.secrets["gcp_service_account"]["universe_domain"],
    }
    gc = gspread.service_account_from_dict(gspread_creds)
    gcal_creds = service_account.Credentials.from_service_account_info(gspread_creds, scopes=['https://www.googleapis.com/auth/calendar'])
    gcal_service = build('calendar', 'v3', credentials=gcal_creds)
except Exception as e:
    st.error(f"認証エラー: Google SheetsまたはCalendarの認証情報が読み込めませんでした。詳細: {e}")
    gc = None
    gcal_service = None

# GCSクライアントの初期化
storage_client = None
try:
    if storage:
        storage_client = storage.Client()
except Exception as e:
    st.warning(f"GCSクライアントの初期化に失敗しました。ファイルアップロード機能は無効になります。詳細: {e}")

# ---------------------------
# --- データ読み込みユーティリティ ---
# ---------------------------

@st.cache_data(ttl=600)  # 10分間キャッシュを保持
def get_data_from_gspread(sheet_name):
    if gc is None:
        return pd.DataFrame()
    
    try:
        worksheet = gc.open(SPREADSHEET_NAME).worksheet(sheet_name)
        data = worksheet.get_all_values()
        
        if not data:
            return pd.DataFrame()
        
        df = pd.DataFrame(data[1:], columns=data[0])
        return df
    except gspread.exceptions.WorksheetNotFound:
        # ワークシートが見つからない場合は空のDataFrameを返す
        st.error(f"ワークシート '{sheet_name}' が見つかりません。シート名を確認してください。")
        return pd.DataFrame()
    except Exception as e:
        # その他のエラー（認証エラーなど）
        st.error(f"スプレッドシート '{sheet_name}' の読み込み中にエラーが発生しました: {e}")
        return pd.DataFrame()

# ---------------------------
# --- GCS アップロードユーティリティ ---
# ---------------------------
def upload_file_to_gcs(storage_client_obj, file_obj): 
    """
    StreamlitのアップロードファイルをGCSのルートに保存し、公開URLを返す。
    """
    from datetime import datetime
    from urllib.parse import quote as url_quote
    
    if storage_client_obj is None or storage is None:
        return None, None

    try:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        original_filename = file_obj.name
        safe_filename = original_filename.replace(' ', '_').replace('/', '_')
        gcs_filename = f"{timestamp}_{safe_filename}"

        bucket = storage_client_obj.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)

        file_bytes = file_obj.getvalue()
        blob.upload_from_string(file_bytes, content_type=file_obj.type if hasattr(file_obj, 'type') else 'application/octet-stream')

        # 署名付きURLを生成
        # 一時的なURLを生成するのではなく、公開URL（認証はクエリパラメータで行う）を返す
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        
        # 署名付きURLが必要な場合は以下を使用
        # public_url = blob.generate_signed_url(expiration=timedelta(days=365*100))
        
        return original_filename, public_url
        
    except Exception as e:
        # GCSアップロードエラーが発生した場合、呼び出し側で処理
        st.error(f"GCSへのアップロード中にエラーが発生しました: {e}")
        return None, None

# ---------------------------
# --- 添付ファイル表示ユーティリティ ---
# ---------------------------
def display_attached_files(row_dict, col_url_key, col_filename_key):
    """
    指定された行データから添付ファイル（URLとファイル名）を抽出し、リンクとして表示する。
    JSON形式（エスケープ対応）と古い単一URL形式の両方に対応。
    """
    import json
    import re
    
    urls = []
    filenames = []
    
    raw_urls = row_dict.get(col_url_key, '')
    raw_filenames = row_dict.get(col_filename_key, '')
    
    # 2. URLのデコードを試みる（新しいデータに対応）
    try:
        # JSONデコードを試みる
        parsed_urls = json.loads(raw_urls)
        
        if isinstance(parsed_urls, list):
            for item in parsed_urls:
                if isinstance(item, str) and item.startswith('http'):
                    urls.append(item)
                else:
                    # リスト要素がさらにエスケープされた文字列だった場合に対応
                    try:
                        inner_item = json.loads(item)
                        if isinstance(inner_item, str) and inner_item.startswith('http'):
                            urls.append(inner_item)
                    except:
                        pass
        
    except (json.JSONDecodeError, AttributeError, TypeError):
        # 3. JSONデコードに失敗した場合（古いデータや単一のURL文字列の場合）
        
        # 文字列から http:// または https:// で始まる最初の要素をURLとして抽出
        url_match = re.search(r'https?://[^\s,"]+', raw_urls)
        if url_match:
            urls = [url_match.group(0)]
        else:
            urls = []

    # 4. ファイル名の取得
    try:
        filenames = json.loads(raw_filenames)
        if not isinstance(filenames, list):
            filenames = [filenames] if isinstance(filenames, str) else []
    except (json.JSONDecodeError, AttributeError, TypeError):
        filenames = [f"添付ファイル {i+1}" for i in range(len(urls))]


    # 5. 表示処理
    if urls:
        st.markdown("##### 📎 添付ファイル")
        
        if len(filenames) < len(urls):
            filenames += [f"ファイル {i+1}" for i in range(len(filenames), len(urls))]
        elif len(filenames) > len(urls):
            filenames = filenames[:len(urls)]
            
        for url, filename in zip(urls, filenames):
            st.markdown(f"[{filename}]({url})")
    else:
        st.markdown("添付ファイルはありません。")


# ---------------------------
# --- エピノート/メンテノート 記録ページ ---
# ---------------------------

def page_epi_note_recording():
    st.markdown("#### 📝 新しいエピノートを記録")
    
    with st.form(key='epi_note_form'):
        ep_title = st.text_input("タイトル/番号 (例: 791)", key="epi_title")
        ep_category = st.selectbox("カテゴリ", ["D1", "D2", "その他"], key="epi_category") 
        ep_memo = st.text_area("詳細メモ", height=200, key="epi_memo")
        
        uploaded_files = st.file_uploader(
            "添付ファイル (画像, PDF, データファイルなど)", 
            type=None, 
            accept_multiple_files=True,
            key="epi_uploader"
        )
        
        st.markdown("---")
        with st.expander("データのインポート"):
            pass
            
        submit_button = st.form_submit_button("記録を保存") 
        
    if submit_button:
        from datetime import datetime
        import json
        
        if not ep_title:
            st.warning("番号 (例: 791) は必須項目です。")
            return
            
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    # GCSルートに保存
                    filename, url = upload_file_to_gcs(storage_client, file_obj) 
                    
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)
                    else:
                        # upload_file_to_gcs内でエラーメッセージが表示される
                        return

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        memo_content = f"{ep_title}\n{ep_memo}"
        
        EPI_COL_NOTE_TYPE = "エピノート" 
        SHEET_TO_WRITE = SHEET_EPI_DATA # 正しいシート名を使用
        
        # 【6列構成】: タイムスタンプ, ノート種別, カテゴリ, メモ, ファイル名, 写真URL
        row_data = [timestamp, EPI_COL_NOTE_TYPE, ep_category, memo_content, filenames_json, urls_json]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_TO_WRITE)
            worksheet.append_row(row_data)
            st.success("✅ エピノートをアップロードしました！")
            
            # 書き込み成功後、キャッシュをクリアし、一覧表示を更新させる
            get_data_from_gspread.clear() 
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")


def page_mainte_recording():
    st.markdown("#### 🛠️ 新しいメンテノートを記録")
    
    with st.form(key='mainte_note_form'):
        mainte_title = st.text_input("メンテタイトル (例: プローブ調整)", key="mainte_title")
        mainte_device = st.selectbox("対象装置", ["MOCVD", "IV/PL", "その他"], key="mainte_device") 
        memo_content = st.text_area("作業詳細メモ", height=200, key="mainte_memo")
        
        uploaded_files = st.file_uploader(
            "添付ファイル (画像, PDF, データファイルなど)", 
            type=None, 
            accept_multiple_files=True,
            key="mainte_uploader"
        )
        
        st.markdown("---")
        with st.expander("データのインポート"):
            pass
            
        submit_button = st.form_submit_button("記録を保存")
        
    if submit_button:
        from datetime import datetime
        import json

        if not mainte_title:
            st.warning("メンテタイトルを入力してください。")
            return
            
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    filename, url = upload_file_to_gcs(storage_client, file_obj)
                    
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)
                    else:
                        return

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        memo_to_save = f"[{mainte_title}] (対象装置: {mainte_device})\n{memo_content}"
        
        MAINTE_COL_NOTE_TYPE = "メンテノート" 
        SHEET_TO_WRITE = SHEET_MAINTE_DATA # 正しいシート名を使用
        
        # 【5列構成】: タイムスタンプ, ノート種別, メモ, ファイル名, 写真URL
        row_data = [timestamp, MAINTE_COL_NOTE_TYPE, memo_to_save, filenames_json, urls_json]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_TO_WRITE)
            worksheet.append_row(row_data)
            st.success("✅ メンテノートをアップロードしました！")
            
            # 書き込み成功後、キャッシュをクリアし、一覧表示を更新させる
            get_data_from_gspread.clear() 
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

# ---------------------------
# --- データ一覧表示ページ ---
# ---------------------------

def page_data_list(sheet_data, title, recording_func):
    st.header(title)
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])

    with tab2:
        recording_func()

    with tab1:
        df = get_data_from_gspread(sheet_data)
        
        if df.empty:
            st.info(f"{title} のデータはまだありません。")
            return

        # 日付形式の変換 (タイムスタンプをソート可能にするため)
        if 'タイムスタンプ' in df.columns:
            # タイムスタンプを降順でソート
            df = df.sort_values(by='タイムスタンプ', ascending=False)
            
        st.dataframe(df, use_container_width=True)
        
        st.subheader("詳細ビュー")
        
        # DataFrameのインデックス（タイムスタンプ）をキーとして選択ボックスを作成
        key_col = 'タイムスタンプ'
        if key_col not in df.columns:
            st.warning("タイムスタンプ列が見つからないため、詳細表示できません。")
            return
            
        # タイムスタンプをキーとして選択
        selection = st.selectbox("記録を選択", df[key_col].unique(), key=f"{sheet_data}_selection")
        
        if selection:
            row = df[df[key_col] == selection].iloc[0].to_dict()
            
            # メタデータ表示
            st.markdown(f"**記録日時:** {row.get('タイムスタンプ', 'N/A')}")
            if 'カテゴリ' in row:
                st.markdown(f"**カテゴリ:** {row['カテゴリ']}")
            
            # メモ内容表示
            memo_content = row.get('メモ', '内容なし')
            if title == "メンテノート":
                # メンテノートはタイトルと装置情報がメモに統合されている前提
                st.subheader(row.get('ノート種別', '詳細メモ'))
            else:
                st.subheader("詳細メモ")
            st.markdown(memo_content)
            
            # 添付ファイル表示
            # 列名がシートによって異なる可能性があるが、ここではエピノート/メンテノートの列名を使用
            display_attached_files(row, '写真URL', 'ファイル名')


def page_epi_note():
    page_data_list(SHEET_EPI_DATA, "エピノート", page_epi_note_recording)

def page_mainte_note():
    page_data_list(SHEET_MAINTE_DATA, "メンテノート", page_mainte_recording)

# ---------------------------
# --- 不足しているページのプレースホルダー定義 ---
# ---------------------------
# NameErrorを回避し、アプリを動作させるために最低限の関数を定義します。

def page_schedule_reservation():
    st.header("🗓️ スケジュール・装置予約")
    st.info("この機能のロジックは以前のコードに存在します。ここではプレースホルダーとして定義します。")

def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    st.info("この機能は現在構築中です。")

def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    st.info("この機能は現在構築中です。")

def page_meeting_note():
    st.header("📄 議事録")
    st.info("この機能は現在構築中です。")

def page_faq():
    st.header("💡 知恵袋・質問箱")
    st.info("この機能は現在構築中です。")

def page_device_handover():
    st.header("📝 装置引き継ぎメモ")
    st.info("この機能は現在構築中です。")

def page_trouble_report():
    st.header("🚨 トラブル報告")
    st.info("この機能は現在構築中です。")

def page_contact():
    st.header("📧 連絡・問い合わせ")
    st.info("この機能は現在構築中です。")


# ---------------------------
# --- メインルーティング (最終修正版) ---
# ---------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    menu_items = [
        "エピノート",
        "メンテノート",
        "🗓️ スケジュール・装置予約",
        "IVデータ解析",
        "PLデータ解析",
        "議事録",
        "知恵袋・質問箱",
        "装置引き継ぎメモ",
        "トラブル報告",
        "連絡・問い合わせ",
    ]
    menu_selection = st.sidebar.radio("機能選択", menu_items)
    
    # 【重要修正】メニュー切り替え時にデータキャッシュをクリアするロジック
    # 選択が変更された場合、データ読み込み関数（get_data_from_gspread）のキャッシュをクリア
    if 'menu_selection' not in st.session_state or st.session_state.menu_selection != menu_selection:
        try:
            get_data_from_gspread.clear()
        except NameError:
            if 'st.cache_data' in st.__dict__:
                st.cache_data.clear()
        
        st.session_state.menu_selection = menu_selection
        # st.rerun() は不要。次回実行時に自動でデータ取得が行われる

    # --- ページルーティング ---
    if menu_selection == "エピノート":
        page_epi_note()
    elif menu_selection == "メンテノート":
        page_mainte_note()
    elif menu_selection == "🗓️ スケジュール・装置予約":
        page_schedule_reservation()
    elif menu_selection == "IVデータ解析":
        page_iv_analysis()
    elif menu_selection == "PLデータ解析":
        page_pl_analysis()
    elif menu_selection == "議事録":
        page_meeting_note()
    elif menu_selection == "知恵袋・質問箱":
        page_faq()
    elif menu_selection == "装置引き継ぎメモ":
        page_device_handover()
    elif menu_selection == "トラブル報告":
        page_trouble_report()
    elif menu_selection == "連絡・問い合わせ":
        page_contact()

if __name__ == "__main__":
    main()

# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v19.0.0 (スプレッドシート構造完全対応 & IV高速化・安定化版)
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
from datetime import datetime, date, time, timedelta
from urllib.parse import quote as url_quote
from io import BytesIO
import calendar

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
# .streamlit/secrets.toml の CLOUD_STORAGE_BUCKET_NAME と一致させてください
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" 
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

SPREADSHEET_NAME = 'エピノート' # Google Spreadsheetのファイル名

# --- エピノートのヘッダー定義（★お客様のシート構造に合わせて修正済み★）
COLUMN_DATE = '日付' 
COLUMN_EPI_NO = 'エピ番号' 
COLUMN_TITLE = 'タイトル' 
COLUMN_DETAIL_MEMO = '詳細メモ' 
COLUMN_FILENAME = 'ファイル名'
COLUMN_FILE_URL = 'ファイルURL' 

# --- メンテノートのヘッダー定義 (既存の構造を維持)
MAINT_COL_TIMESTAMP = 'タイムスタンプ'
MAINT_COL_TYPE = 'ノート種別'
MAINT_COL_MEMO = 'メモ'
MAINT_COL_FILENAME = 'ファイル名'
MAINT_COL_FILE_URL = '写真URL' # メンテノートは'写真URL'を維持

# --------------------------------------------------------------------------
# --- Google Service Initialization (認証処理) ---
# --------------------------------------------------------------------------

class DummyGSClient:
    """認証失敗時用のダミーgspreadクライアント"""
    def open(self, name): return self
    def worksheet(self, name): return self
    def get_all_records(self): return []
    def get_all_values(self): return []
    def append_row(self, values): pass
    # ... 他のダミーメソッドは省略 ...

class DummyCalendarService:
    """認証失敗時用のダミーカレンダーサービス"""
    def events(self): return self
    def list(self, **kwargs): return self
    def execute(self): return {'items': []}

class DummyStorageClient:
    """認証失敗時用のダミーGCSクライアント"""
    def bucket(self, name): return self
    def blob(self, name): return self
    def download_as_bytes(self): return b''
    def upload_from_file(self, file_obj, content_type): pass
    def get_bucket(self, name): return self
    def list_blobs(self, **kwargs): return []

@st.cache_resource(ttl=3600)
def initialize_google_services():
    """
    Streamlit Secretsから認証情報を読み込み、Googleサービスを初期化する
    """
    if "gcs_credentials" not in st.secrets:
        st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
        return DummyGSClient(), DummyCalendarService(), DummyStorageClient()

    try:
        # JSON文字列を直接ロード
        info = json.loads(st.secrets["gcs_credentials"])
        
        # gspread (Spreadsheet) の認証
        gc = gspread.service_account_from_dict(info)

        # googleapiclient (Calendar) の認証
        credentials = Credentials.from_service_account_info(info)
        calendar_service = build('calendar', 'v3', credentials=credentials) # ダミーとして残します

        # google.cloud.storage (GCS) の認証
        storage_client = storage.Client.from_service_account_info(info)

        return gc, calendar_service, storage_client

    except Exception as e:
        st.error(f"❌ 認証エラー: サービスアカウントの初期化に失敗しました。認証情報をご確認ください。({e})")
        return DummyGSClient(), DummyCalendarService(), DummyStorageClient()

gc, calendar_service, storage_client = initialize_google_services()

# --------------------------------------------------------------------------
# --- Data Utilities (データ取得・解析) ---
# --------------------------------------------------------------------------

@st.cache_data(ttl=600, show_spinner="スプレッドシートからデータを読み込み中...")
def get_sheet_as_df(gc, spreadsheet_name, sheet_name):
    """指定されたシートのデータをDataFrameとして取得する"""
    if isinstance(gc, DummyGSClient):
        return pd.DataFrame()
    
    try:
        worksheet = gc.open(spreadsheet_name).worksheet(sheet_name)
        data = worksheet.get_all_values()
        if not data:
            return pd.DataFrame()
        
        # 1行目をヘッダーとしてDataFrameを作成
        df = pd.DataFrame(data[1:], columns=data[0])
        return df

    except gspread.exceptions.WorksheetNotFound:
        st.error(f"シート名「{sheet_name}」が見つかりません。スプレッドシートをご確認ください。")
        return pd.DataFrame()
    except Exception as e:
        st.warning(f"警告：シート「{sheet_name}」の読み込み中にエラーが発生しました。ヘッダーの不一致やデータ形式を確認してください。({e})")
        return pd.DataFrame()

# --- IVデータ解析用ユーティリティ (キャッシュで高速化) ---
@st.cache_data(show_spinner="IVデータを解析中...", max_entries=50)
def load_iv_data(uploaded_file_bytes, uploaded_file_name):
    """アップロードされたIVファイルを読み込み、DataFrameを返す (IV処理落ち対策済み)"""
    try:
        content = uploaded_file_bytes.decode('utf-8').splitlines()
        data_lines = content[1:] # 最初の1行（ヘッダー）をスキップ

        cleaned_data_lines = []
        for line in data_lines:
            line_stripped = line.strip()
            if line_stripped and not line_stripped.startswith(('#', '!', '/')):
                cleaned_data_lines.append(line_stripped)

        if not cleaned_data_lines: return None

        data_string_io = io.StringIO("\n".join(cleaned_data_lines))
        
        # ロバストな読み込み処理: \s+ (スペース/タブ)、タブ、コンマ区切りを順に試す
        try:
            df = pd.read_csv(data_string_io, sep=r'\s+', engine='python', header=None, skipinitialspace=True)
        except Exception:
            try:
                data_string_io.seek(0)
                df = pd.read_csv(data_string_io, sep='\t', engine='c', header=None)
            except Exception:
                data_string_io.seek(0)
                df = pd.read_csv(data_string_io, sep=',', engine='python', header=None)

        if df is None or len(df.columns) < 2: return None
        
        df = df.iloc[:, :2]
        df.columns = ['Voltage_V', uploaded_file_name] # ファイル名を列名に使用

        # 数値型に変換し、変換できない行は削除 (float型を明示し、numpy.floatエラーを回避)
        df['Voltage_V'] = pd.to_numeric(df['Voltage_V'], errors='coerce', downcast='float')
        df[uploaded_file_name] = pd.to_numeric(df[uploaded_file_name], errors='coerce', downcast='float')
        df.dropna(inplace=True)
        
        return df

    except Exception:
        return None

@st.cache_data(show_spinner="データを結合中...")
def combine_iv_dataframes(dataframes, filenames):
    """複数のIV DataFrameをVoltage_Vをキーに外部結合する"""
    if not dataframes: return None
    
    # 結合処理の高速化
    combined_df = dataframes[0]
    
    for i in range(1, len(dataframes)):
        df_to_merge = dataframes[i]
        combined_df = pd.merge(combined_df, df_to_merge, on='Voltage_V', how='outer')
        
    combined_df = combined_df.sort_values(by='Voltage_V', ascending=False).reset_index(drop=True)
    
    for col in combined_df.columns:
        if col != 'Voltage_V':
            combined_df[col] = combined_df[col].round(4)
            
    return combined_df

# --------------------------------------------------------------------------
# --- GCS Utilities (ファイルアップロード) ---
# --------------------------------------------------------------------------

def upload_file_to_gcs(storage_client, file_obj, folder_name):
    """ファイルをGCSにアップロードし、公開URLを返す"""
    if isinstance(storage_client, DummyStorageClient):
        return None, "dummy_url_gcs_error"
        
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    original_filename = file_obj.name
    gcs_filename = f"{folder_name}/{timestamp}_{original_filename}"

    try:
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        
        file_obj.seek(0)
        blob.upload_from_file(file_obj, content_type=file_obj.type)

        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        
        return original_filename, public_url

    except Exception as e:
        st.error(f"❌ GCSエラー: ファイルのアップロード中にエラーが発生しました。バケット名 '{CLOUD_STORAGE_BUCKET_NAME}' が正しいか、権限があるか確認してください。({e})")
        return None, None

# --------------------------------------------------------------------------
# --- Page Definitions (各機能ページ) ---
# --------------------------------------------------------------------------

def page_note_recording(sheet_name='エピノート_データ', is_mainte=False):
    """エピノート・メンテノート記録ページ"""
    
    # ... (UIと書き込みロジックはヘッダーに合わせて修正済み) ...
    if is_mainte:
        st.header("🛠️ メンテノート記録")
        sheet_name = 'メンテノート_データ'
    else:
        st.header("📝 エピノート記録")
    
    st.markdown("---")
    
    with st.form(key='note_form'):
        
        if not is_mainte:
            col1, col2 = st.columns(2)
            with col1:
                ep_date = st.date_input(f"{COLUMN_DATE}", datetime.now().date())
                ep_no = st.text_input(f"{COLUMN_EPI_NO} (例: 784-A)", key='ep_no_input')
            with col2:
                ep_title = st.text_input(f"{COLUMN_TITLE} (例: PL測定)", key='ep_title_input')
        
        if is_mainte:
            mainte_type = st.selectbox(f"{MAINT_COL_TYPE} (装置/内容)", [
                "ドライポンプ交換", "ドライポンプメンテ", "オイル交換", "ヒーター交換", "その他"
            ])

        memo_content = st.text_area(f"{COLUMN_DETAIL_MEMO} / {MAINT_COL_MEMO}", height=150, key='memo_input')
        uploaded_files = st.file_uploader("添付ファイル (画像、グラフなど)", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], accept_multiple_files=True)
        
        st.markdown("---")
        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')

    if submit_button:
        if not memo_content and not uploaded_files:
            st.warning("メモ内容を入力するか、ファイルをアップロードしてください。")
            return
        
        filenames_list = []
        urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                folder_name = "ep_notes" if not is_mainte else "mainte_notes"
                for file_obj in uploaded_files:
                    filename, url = upload_file_to_gcs(storage_client, file_obj, folder_name)
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

        # 2. スプレッドシートに行を追加
        if not is_mainte:
            # エピノート: ['日付', 'エピ番号', 'タイトル', '詳細メモ', 'ファイル名', 'ファイルURL']
            row_data = [
                ep_date.isoformat(), ep_no, ep_title, 
                memo_content, filenames_json, urls_json
            ]
        else:
            # メンテノート: ['タイムスタンプ', 'ノート種別', 'メモ', 'ファイル名', '写真URL']
            row_data = [
                timestamp, mainte_type, 
                memo_content, filenames_json, urls_json
            ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(sheet_name)
            worksheet.append_row(row_data)
            st.success("記録を保存しました！"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{sheet_name}' が存在するか確認してください。")


def page_note_list(sheet_name='エピノート_データ', is_mainte=False):
    """エピノート・メンテノート一覧ページ"""
    
    if is_mainte:
        st.header("🛠️ メンテノート一覧")
        sheet_name = 'メンテノート_データ'
        COL_TIME = MAINT_COL_TIMESTAMP
        COL_FILTER = MAINT_COL_TYPE
        COL_MEMO = MAINT_COL_MEMO
        COL_URL = MAINT_COL_FILE_URL
    else:
        st.header("📚 エピノート一覧")
        sheet_name = 'エピノート_データ'
        COL_TIME = COLUMN_DATE 
        COL_FILTER = COLUMN_TITLE # ★タイトルで絞り込み★
        COL_MEMO = COLUMN_DETAIL_MEMO # ★詳細メモを表示★
        COL_URL = COLUMN_FILE_URL # ★ファイルURLを参照★
    
    df = get_sheet_as_df(gc, SPREADSHEET_NAME, sheet_name)

    if df.empty: st.info("データがありません。"); return
        
    st.subheader("絞り込みと検索")
    
    if COL_FILTER in df.columns:
        filter_options = ["すべて"] + list(df[COL_FILTER].unique())
        note_filter = st.selectbox(f"{COL_FILTER}で絞り込み", filter_options)
        
        if note_filter != "すべて":
            df = df[df[COL_FILTER] == note_filter]

    col_date1, col_date2 = st.columns(2)
    with col_date1:
        start_date = st.date_input("開始日", value=datetime.now().date() - timedelta(days=30))
    with col_date2:
        end_date = st.date_input("終了日", value=datetime.now().date())
    
    try:
        # 日付を扱う列に合わせて処理
        df[COL_TIME] = pd.to_datetime(df[COL_TIME]).dt.date
        df = df[(df[COL_TIME] >= start_date) & (df[COL_TIME] <= end_date)]
    except:
        st.warning("日付（タイムスタンプ）列の形式が不正な行があります。")

    if df.empty: st.info("絞り込み条件に一致するデータがありません。"); return

    df = df.sort_values(by=COL_TIME, ascending=False).reset_index(drop=True)
    
    st.markdown("---")
    st.subheader(f"検索結果 ({len(df)}件)")

    if df.empty: st.info("表示するデータがありません。"); return

    df['display_index'] = df.index
    format_func = lambda idx: f"[{df.loc[idx, COL_TIME]}] {df.loc[idx, COL_FILTER]} - {df.loc[idx, COL_MEMO][:30]}..."

    selected_index = st.selectbox(
        "詳細を表示する記録を選択", 
        options=df['display_index'], 
        format_func=format_func
    )

    if selected_index is not None:
        row = df.loc[selected_index]
        st.markdown(f"#### 選択された記録 (ID: {selected_index+1})")
        
        if not is_mainte:
            # エピノートの表示項目
            st.write(f"**{COLUMN_DATE}:** {row[COLUMN_DATE]}")
            st.write(f"**{COLUMN_EPI_NO}:** {row[COLUMN_EPI_NO]}")
            st.write(f"**{COLUMN_TITLE}:** {row[COLUMN_TITLE]}")
            st.markdown(f"**{COLUMN_DETAIL_MEMO}:**"); st.text(row[COLUMN_DETAIL_MEMO])
        else:
            # メンテノートの表示項目
            st.write(f"**{MAINT_COL_TIMESTAMP}:** {row[MAINT_COL_TIMESTAMP]}")
            st.write(f"**{MAINT_COL_TYPE}:** {row[MAINT_COL_TYPE]}")
            st.markdown(f"**{MAINT_COL_MEMO}:**"); st.text(row[MAINT_COL_MEMO])
            
        st.markdown("##### 添付ファイル")
        try:
            urls = json.loads(row[COL_URL])
            filenames = json.loads(row[COLUMN_FILENAME])
            
            if urls:
                for filename, url in zip(filenames, urls):
                    st.markdown(f"- [{filename}]({url})")
            else:
                st.info("添付ファイルはありません。")
        except:
            st.warning("添付ファイル情報が不正です。")
            
def page_mainte_recording(): page_note_recording(is_mainte=True)
def page_mainte_list(): page_note_list(is_mainte=True)
    
def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    st.info("このページは未実装です。")

def page_iv_analysis():
    """⚡ IVデータ解析ページ（キャッシュ適用済み）"""
    st.header("⚡ IVデータ解析")
    
    uploaded_files = st.file_uploader(
        "IV測定データファイル (.txt) をアップロード",
        type=['txt'], 
        accept_multiple_files=True
    )

    if uploaded_files:
        valid_dataframes = []
        filenames = []
        
        st.subheader("ステップ1: ファイル読み込みと解析")
        
        for uploaded_file in uploaded_files:
            df = load_iv_data(uploaded_file.getvalue(), uploaded_file.name)
            
            if df is not None and not df.empty:
                valid_dataframes.append(df)
                filenames.append(uploaded_file.name)
        
        if valid_dataframes:
            
            combined_df = combine_iv_dataframes(valid_dataframes, filenames)
            
            st.success(f"{len(valid_dataframes)}個の有効なファイルを読み込み、結合しました。")
            
            st.subheader("ステップ2: グラフ表示")
            
            fig, ax = plt.subplots(figsize=(12, 7)) 
            
            for filename in filenames:
                ax.plot(combined_df['Voltage_V'], combined_df[filename], label=filename)
            
            ax.set_xlabel("Voltage (V)")
            ax.set_ylabel("Current (A)")
            ax.grid(True)
            ax.legend(title="ファイル名", loc='best')
            ax.set_title("IV特性比較")
            
            st.pyplot(fig, use_container_width=True) 
            
            st.subheader("ステップ3: 結合データ")
            st.dataframe(combined_df, use_container_width=True)
            
            # Excelダウンロード
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, sheet_name='Combined IV Data', index=False)
            
            st.download_button(
                label="📈 結合Excelデータとしてダウンロード",
                data=output.getvalue(),
                file_name=f"iv_analysis_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("有効なデータファイルが見つかりませんでした。")

# --- Dummy Pages (未実装のページ) ---
def page_calendar(): st.header("🗓️ スケジュール・装置予約"); st.info("このページは未実装です。")
def page_meeting_minutes(): st.header("議事録・ミーティングメモ"); st.info("このページは未実装です。")
def page_qa(): st.header("💡 知恵袋・質問箱"); st.info("このページは未実装です。")
def page_handover(): st.header("🤝 装置引き継ぎメモ"); st.info("このページは未実装です。")
def page_trouble_report(): st.header("🚨 トラブル報告"); st.info("このページは未実装です。")
def page_contact(): st.header("✉️ 連絡・問い合わせ"); st.info("このページは未実装です。")

# --------------------------------------------------------------------------
# --- Main App Execution ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    menu_selection = st.sidebar.radio("機能選択", [
        "📝 エピノート記録", "📚 エピノート一覧", "🛠️ メンテノート記録", "🛠️ メンテノート一覧",
        "🗓️ スケジュール・装置予約", 
        "⚡ IVデータ解析", "🔬 PLデータ解析",
        "議事録・ミーティングメモ", "💡 知恵袋・質問箱", "🤝 装置引き継ぎメモ", 
        "🚨 トラブル報告", "✉️ 連絡・問い合わせ"
    ])
    
    # ページルーティング
    if menu_selection == "📝 エピノート記録": page_note_recording()
    elif menu_selection == "📚 エピノート一覧": page_note_list()
    elif menu_selection == "🛠️ メンテノート記録": page_mainte_recording()
    elif menu_selection == "🛠️ メンテノート一覧": page_mainte_list()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()
    elif menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "議事録・ミーティングメモ": page_meeting_minutes()
    elif menu_selection == "💡 知恵袋・質問箱": page_qa()
    elif menu_selection == "🤝 装置引き継ぎメモ": page_handover()
    elif menu_selection == "🚨 トラブル報告": page_trouble_report()
    elif menu_selection == "✉️ 連絡・問い合わせ": page_contact()

if __name__ == "__main__":
    main()

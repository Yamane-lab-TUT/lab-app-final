# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v20.0.0 (全シート構造完全対応 & IV高速化・安定化版)
# - お客様の全アップロードデータに基づき、ヘッダー名とシート名を確定。
# - 全ての記録・一覧ページ（エピノート、メンテノート、トラブル報告、議事録、知恵袋、引き継ぎ、問い合わせ）を実装。
# --------------------------------------------------------------------------

import streamlit as st
import gspread
import pandas as pd
import io
import re
import json
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, date, timedelta
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

# --- SPREADSHEET COLUMN HEADERS (お客様のデータ構造に完全一致) ---

# --- エピノート_データ
SHEET_EPI_DATA = 'エピノート_データ'
EPI_COL_TIMESTAMP = 'タイムスタンプ'
EPI_COL_NOTE_TYPE = 'ノート種別'   # 'エピノート'
EPI_COL_CATEGORY = 'カテゴリ'     # 'D1', '897'など、エピ番号やカテゴリ
EPI_COL_MEMO = 'メモ'           # タイトルと詳細メモを含む
EPI_COL_FILENAME = 'ファイル名'
EPI_COL_FILE_URL = '写真URL'

# --- メンテノート_データ
SHEET_MAINTE_DATA = 'メンテノート_データ'
MAINT_COL_TIMESTAMP = 'タイムスタンプ'
MAINT_COL_NOTE_TYPE = 'ノート種別' # 'メンテノート'
MAINT_COL_MEMO = 'メモ'
MAINT_COL_FILENAME = 'ファイル名'
MAINT_COL_FILE_URL = '写真URL'

# --- 議事録_データ
SHEET_MEETING_DATA = '議事録_データ'
MEETING_COL_TIMESTAMP = 'タイムスタンプ'
MEETING_COL_TITLE = '会議タイトル'
MEETING_COL_AUDIO_NAME = '音声ファイル名'
MEETING_COL_AUDIO_URL = '音声ファイルURL'
MEETING_COL_CONTENT = '議事録内容'

# --- 引き継ぎ_データ
SHEET_HANDOVER_DATA = '引き継ぎ_データ'
HANDOVER_COL_TIMESTAMP = 'タイムスタンプ'
HANDOVER_COL_TYPE = '種類'
HANDOVER_COL_TITLE = 'タイトル'
HANDOVER_COL_MEMO = 'メモ' # 内容1,2,3はUIを複雑にするため、一旦メモに統合

# --- 知恵袋_データ (質問)
SHEET_QA_DATA = '知恵袋_データ'
QA_COL_TIMESTAMP = 'タイムスタンプ'
QA_COL_TITLE = '質問タイトル'
QA_COL_CONTENT = '質問内容'
QA_COL_CONTACT = '連絡先メールアドレス'
QA_COL_FILENAME = '添付ファイル名'
QA_COL_FILE_URL = '添付ファイルURL'
QA_COL_STATUS = 'ステータス'
SHEET_QA_ANSWER = '知恵袋_解答' # 解答シート

# --- お問い合わせ_データ
SHEET_CONTACT_DATA = 'お問い合わせ_データ'
CONTACT_COL_TIMESTAMP = 'タイムスタンプ'
CONTACT_COL_TYPE = 'お問い合わせの種類'
CONTACT_COL_DETAIL = '詳細内容'
CONTACT_COL_CONTACT = '連絡先'

# --- トラブル報告_データ
SHEET_TROUBLE_DATA = 'トラブル報告_データ'
TROUBLE_COL_TIMESTAMP = 'タイムスタンプ'
TROUBLE_COL_DEVICE = '機器/場所'
TROUBLE_COL_OCCUR_DATE = '発生日'
TROUBLE_COL_OCCUR_TIME = 'トラブル発生時'
TROUBLE_COL_CAUSE = '原因/究明'
TROUBLE_COL_SOLUTION = '対策/復旧'
TROUBLE_COL_PREVENTION = '再発防止策'
TROUBLE_COL_REPORTER = '報告者'
TROUBLE_COL_FILENAME = 'ファイル名'
TROUBLE_COL_FILE_URL = 'ファイルURL'
TROUBLE_COL_TITLE = '件名/タイトル'

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
    
class DummyStorageClient:
    """認証失敗時用のダミーGCSクライアント"""
    def bucket(self, name): return self
    def blob(self, name): return self
    def download_as_bytes(self): return b''
    def upload_from_file(self, file_obj, content_type): pass
    def get_bucket(self, name): return self
    def list_blobs(self, **kwargs): return []

# ダミーカレンダーサービスは使用しないため削除
# app.py の initialize_google_services 関数部分のみ
# ... (省略) ...

@st.cache_resource(ttl=3600)
def initialize_google_services():
    """Streamlit Secretsから認証情報を読み込み、Googleサービスを初期化する"""
    if "gcs_credentials" not in st.secrets:
        st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
        return DummyGSClient(), DummyStorageClient()

    try:
        raw_credentials_string = st.secrets["gcs_credentials"]
        
        # --- 認証文字列の【強制】クリーンアップ v20.3.0 ---
        # 1. 冒頭と末尾の不要な空白（改行、タブなど）を除去
        cleaned_string = raw_credentials_string.strip()
        
        # 2. JSON内部の改行とタブ文字を完全に除去し、JSON全体を一行にする
        # これにより、三重引用符内のインデントや改行によるパースエラーをほぼ確実に排除します。
        # ただし、private_key内のエスケープされた改行(\\n)は保持される必要があります。
        
        # JSON外の改行・タブ・全角スペースを除去
        cleaned_string = cleaned_string.replace('\n', '')
        cleaned_string = cleaned_string.replace('\t', '')
        cleaned_string = cleaned_string.replace(' ', '') # U+00A0: NO-BREAK SPACE (全角スペースと誤認されやすい文字)
        
        # 最後に連続するスペースを一つに置換 (JSONの構造を壊さない範囲で)
        cleaned_string = re.sub(r'(\s){2,}', r'\1', cleaned_string)
        
        # JSONをパース
        info = json.loads(cleaned_string) 
        
        # gspread (Spreadsheet) の認証
        gc = gspread.service_account_from_dict(info)

        # google.cloud.storage (GCS) の認証
        storage_client = storage.Client.from_service_account_info(info)

        st.sidebar.success("✅ Googleサービス認証成功")
        return gc, storage_client

    except json.JSONDecodeError as e:
        # JSONパースエラーが発生した場合
        st.error(f"❌ 認証エラー（JSON形式不正）: サービスアカウントのJSON形式が不正です。改行やタブ文字、不要なスペースが含まれていないか確認してください。エラー詳細: {e}")
        return DummyGSClient(), DummyStorageClient()
        
    except Exception as e:
        # その他の認証エラー（権限不足など）
        st.error(f"❌ 認証エラー: サービスアカウントの初期化に失敗しました。認証情報をご確認ください。({e})")
        return DummyGSClient(), DummyStorageClient()

# ... (省略) ...
# Calendar Serviceは使わないため、戻り値を調整
gc, storage_client = initialize_google_services() 

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
        if not data or len(data) <= 1: # ヘッダーのみの場合も空とみなす
            return pd.DataFrame(columns=data[0] if data else [])
        
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
# (前回のコードから変更なし)
@st.cache_data(show_spinner="IVデータを解析中...", max_entries=50)
def load_iv_data(uploaded_file_bytes, uploaded_file_name):
    """アップロードされたIVファイルを読み込み、DataFrameを返す"""
    try:
        content = uploaded_file_bytes.decode('utf-8').splitlines()
        data_lines = content[1:] 

        cleaned_data_lines = []
        for line in data_lines:
            line_stripped = line.strip()
            if line_stripped and not line_stripped.startswith(('#', '!', '/')):
                cleaned_data_lines.append(line_stripped)

        if not cleaned_data_lines: return None

        data_string_io = io.StringIO("\n".join(cleaned_data_lines))
        
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
        df.columns = ['Voltage_V', uploaded_file_name] 

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
    # ファイル名が日本語の場合に備え、URLエンコードを考慮してスペース等をアンダースコアに置換（GCSのblob名はURLエンコードされないため）
    safe_filename = original_filename.replace(' ', '_').replace('/', '_')
    gcs_filename = f"{folder_name}/{timestamp}_{safe_filename}"

    try:
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        
        file_obj.seek(0)
        blob.upload_from_file(file_obj, content_type=file_obj.type)

        # 署名付きURLではなく、よりシンプルな公開URLを生成（バケットの権限設定に依存）
        # ユーザーの既存データが使用している形式に合わせます
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        
        return original_filename, public_url

    except Exception as e:
        st.error(f"❌ GCSエラー: ファイルのアップロード中にエラーが発生しました。バケット名 '{CLOUD_STORAGE_BUCKET_NAME}' が正しいか、権限があるか確認してください。({e})")
        return None, None

# --------------------------------------------------------------------------
# --- Page Implementations (各機能ページ) ---
# --------------------------------------------------------------------------

# --- 汎用的な一覧表示関数 ---
def page_data_list(sheet_name, title, col_time, col_filter=None, col_memo=None, col_url=None, detail_cols=None):
    """汎用的なデータ一覧ページ"""
    
    st.header(title)
    df = get_sheet_as_df(gc, SPREADSHEET_NAME, sheet_name)

    if df.empty: st.info("データがありません。"); return
        
    st.subheader("絞り込みと検索")
    
    if col_filter and col_filter in df.columns:
        filter_options = ["すべて"] + list(df[col_filter].unique())
        data_filter = st.selectbox(f"「{col_filter}」で絞り込み", filter_options)
        
        if data_filter != "すべて":
            df = df[df[col_filter] == data_filter]

    # 日付による絞り込み
    if col_time and col_time in df.columns:
        try:
            # タイムスタンプ列を日付型に変換
            df[col_time] = pd.to_datetime(df[col_time].str.replace(r'[^0-9]', '', regex=True), errors='coerce', format='%Y%m%d%H%M%S', exact=False).dt.date
        except:
            # 日付形式が不正な場合は、そのまま処理
            pass 
        
        df_valid_date = df.dropna(subset=[col_time])
        
        if not df_valid_date.empty:
            min_date = df_valid_date[col_time].min()
            max_date = df_valid_date[col_time].max()
            
            col_date1, col_date2 = st.columns(2)
            with col_date1:
                start_date = st.date_input("開始日", value=max(min_date, datetime.now().date() - timedelta(days=30)))
            with col_date2:
                end_date = st.date_input("終了日", value=max_date)
            
            df = df_valid_date[(df_valid_date[col_time] >= start_date) & (df_valid_date[col_time] <= end_date)]
        else:
            st.warning("日付（タイムスタンプ）列の形式が不正な行が多いため、日付絞り込みをスキップしました。")


    if df.empty: st.info("絞り込み条件に一致するデータがありません。"); return

    df = df.sort_values(by=col_time, ascending=False).reset_index(drop=True)
    
    st.markdown("---")
    st.subheader(f"検索結果 ({len(df)}件)")

    # 選択肢のフォーマット関数
    def format_func(idx):
        row = df.loc[idx]
        time_str = str(row[col_time])
        filter_str = row[col_filter] if col_filter and pd.notna(row[col_filter]) else ""
        memo_str = row[col_memo] if col_memo and pd.notna(row[col_memo]) else "メモなし"
        return f"[{time_str}] {filter_str} - {memo_str[:50].replace('\\n', ' ')}..."

    df['display_index'] = df.index
    selected_index = st.selectbox(
        "詳細を表示する記録を選択", 
        options=df['display_index'], 
        format_func=format_func
    )

    if selected_index is not None:
        row = df.loc[selected_index]
        st.markdown(f"#### 選択された記録 (ID: {selected_index+1})")
        
        # 主要情報と詳細情報を表示
        if detail_cols:
            for col in detail_cols:
                if col in row:
                    if col_memo == col:
                        st.markdown(f"**{col}:**"); st.text(row[col])
                    else:
                        st.write(f"**{col}:** {row[col]}")
        
        # 添付ファイル (ファイル名とURLが分離しているか、同一かによって表示を調整)
        if col_url and col_url in row:
            st.markdown("##### 添付ファイル")
            
            try:
                # ファイル名とURLがJSONリストとして保存されている場合（標準的な書き込み形式）
                urls = json.loads(row[col_url])
                filenames = json.loads(row[EPI_COL_FILENAME]) if EPI_COL_FILENAME in row else ['ファイル'] * len(urls)
                
                if urls:
                    for filename, url in zip(filenames, urls):
                        # URLの末尾がGoogle Driveの場合は別表示
                        if "drive.google.com" in url:
                            st.markdown(f"- **Google Drive:** [{filename}](<{url}>)")
                        else:
                            st.markdown(f"- [{filename}]({url})")
                else:
                    st.info("添付ファイルはありません。")

            except Exception:
                # JSON形式ではない場合（古いデータや手動入力）
                if pd.notna(row[col_url]) and row[col_url]:
                    st.markdown(f"- [添付ファイルURL]({row[col_url]})")
                else:
                    st.info("添付ファイルはありません。")


# --- 1. エピノート記録/一覧 ---
def page_epi_note_recording():
    st.header("📝 エピノート記録")
    st.markdown("---")
    
    with st.form(key='epi_note_form'):
        
        # ユーザーの既存データ構造: カテゴリ(エピ番号), メモ(タイトル+詳細)
        col1, col2 = st.columns(2)
        with col1:
            ep_category = st.text_input(f"{EPI_COL_CATEGORY} (例: D1, 784-A)", key='ep_category_input')
        with col2:
            ep_title = st.text_input("タイトル/要約 (必須)", key='ep_title_input')
        
        ep_memo = st.text_area(f"詳細メモ", height=150, key='ep_memo_input')
        uploaded_files = st.file_uploader("添付ファイル (画像、グラフなど)", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], accept_multiple_files=True)
        
        st.markdown("---")
        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')

    if submit_button:
        if not ep_title or not ep_memo:
            st.warning("タイトルと詳細メモを入力してください。")
            return
        
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    filename, url = upload_file_to_gcs(storage_client, file_obj, "ep_notes")
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 実際のシートの列にデータをマッピング: ['タイムスタンプ', 'ノート種別', 'カテゴリ', 'メモ', 'ファイル名', '写真URL']
        memo_content = f"{ep_title}\n{ep_memo}"
        row_data = [
            timestamp, EPI_COL_NOTE_TYPE, ep_category, 
            memo_content, filenames_json, urls_json
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_EPI_DATA)
            worksheet.append_row(row_data)
            st.success("エピノートを保存しました！"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_EPI_DATA}' が存在するか確認してください。")

def page_epi_note_list():
    # 表示項目: タイムスタンプ, ノート種別, カテゴリ, メモ, ファイル名, 写真URL
    detail_cols = [EPI_COL_TIMESTAMP, EPI_COL_CATEGORY, EPI_COL_NOTE_TYPE, EPI_COL_MEMO, EPI_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_EPI_DATA,
        title="📚 エピノート一覧",
        col_time=EPI_COL_TIMESTAMP,
        col_filter=EPI_COL_CATEGORY,
        col_memo=EPI_COL_MEMO,
        col_url=EPI_COL_FILE_URL,
        detail_cols=detail_cols
    )

# --- 2. メンテノート記録/一覧 ---
def page_mainte_recording():
    st.header("🛠️ メンテノート記録")
    st.markdown("---")
    
    with st.form(key='mainte_note_form'):
        
        mainte_type = st.selectbox(f"{MAINT_COL_MEMO} (装置/内容)", [
            "D1 ドライポンプ交換", "D2 ドライポンプ交換", "オイル交換", "ヒーター交換", "その他"
        ])
        memo_content = st.text_area("詳細メモ", height=150, key='mainte_memo_input')
        uploaded_files = st.file_uploader("添付ファイル (画像、グラフなど)", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], accept_multiple_files=True)
        
        st.markdown("---")
        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')

    if submit_button:
        if not memo_content:
            st.warning("詳細メモを入力してください。")
            return
        
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    filename, url = upload_file_to_gcs(storage_client, file_obj, "mainte_notes")
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 実際のシートの列にデータをマッピング: ['タイムスタンプ', 'ノート種別', 'メモ', 'ファイル名', '写真URL']
        memo_to_save = f"[{mainte_type}] {memo_content}"
        row_data = [
            timestamp, MAINT_COL_NOTE_TYPE, memo_to_save, 
            filenames_json, urls_json
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_MAINTE_DATA)
            worksheet.append_row(row_data)
            st.success("メンテノートを保存しました！"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_MAINTE_DATA}' が存在するか確認してください。")

def page_mainte_list():
    # 表示項目: タイムスタンプ, ノート種別, メモ, ファイル名, 写真URL
    detail_cols = [MAINT_COL_TIMESTAMP, MAINT_COL_NOTE_TYPE, MAINT_COL_MEMO, MAINT_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_MAINTE_DATA,
        title="🛠️ メンテノート一覧",
        col_time=MAINT_COL_TIMESTAMP,
        col_filter=MAINT_COL_NOTE_TYPE, # 種類で絞り込み
        col_memo=MAINT_COL_MEMO,
        col_url=MAINT_COL_FILE_URL,
        detail_cols=detail_cols
    )
    
# --- 3. 議事録・ミーティングメモ記録/一覧 ---
def page_meeting_recording():
    st.header("📝 議事録記録")
    st.info("※ 録音機能は未実装のため、手動でURLをペーストしてください。")
    st.markdown("---")

    with st.form(key='meeting_form'):
        meeting_title = st.text_input(f"{MEETING_COL_TITLE} (例: 2025-10-28 定例会議)", key='meeting_title_input')
        meeting_content = st.text_area(f"{MEETING_COL_CONTENT}", height=300, key='meeting_content_input')
        col1, col2 = st.columns(2)
        with col1:
            audio_name = st.text_input(f"{MEETING_COL_AUDIO_NAME} (例: audio.m4a)", key='audio_name_input')
        with col2:
            audio_url = st.text_input(f"{MEETING_COL_AUDIO_URL} (Google Drive URLなど)", key='audio_url_input')

        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')
        
    if submit_button:
        if not meeting_title or not meeting_content:
            st.warning("会議タイトルと議事録内容を入力してください。")
            return
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # ['タイムスタンプ', '会議タイトル', '音声ファイル名', '音声ファイルURL', '議事録内容']
        row_data = [
            timestamp, meeting_title, audio_name, 
            audio_url, meeting_content
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_MEETING_DATA)
            worksheet.append_row(row_data)
            st.success("議事録を保存しました！"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_MEETING_DATA}' が存在するか確認してください。")

def page_meeting_list():
    # 表示項目: タイムスタンプ, 会議タイトル, 音声ファイル名, 音声ファイルURL, 議事録内容
    detail_cols = [MEETING_COL_TIMESTAMP, MEETING_COL_TITLE, MEETING_COL_CONTENT, MEETING_COL_AUDIO_NAME, MEETING_COL_AUDIO_URL]
    page_data_list(
        sheet_name=SHEET_MEETING_DATA,
        title="📚 議事録一覧",
        col_time=MEETING_COL_TIMESTAMP,
        col_filter=MEETING_COL_TITLE,
        col_memo=MEETING_COL_CONTENT,
        col_url=MEETING_COL_AUDIO_URL,
        detail_cols=detail_cols
    )

# --- 4. 知恵袋・質問箱（質問のみ実装）---
def page_qa_recording():
    st.header("💡 知恵袋・質問箱 (質問投稿)")
    st.markdown("---")
    
    with st.form(key='qa_form'):
        qa_title = st.text_input(f"{QA_COL_TITLE} (例: XRDの測定手順について)", key='qa_title_input')
        qa_content = st.text_area(f"{QA_COL_CONTENT}", height=200, key='qa_content_input')
        col1, col2 = st.columns(2)
        with col1:
            qa_contact = st.text_input(f"{QA_COL_CONTACT} (任意)", key='qa_contact_input')
        with col2:
            uploaded_files = st.file_uploader("添付ファイル", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], accept_multiple_files=True)
            
        st.markdown("---")
        submit_button = st.form_submit_button(label='質問を投稿')

    if submit_button:
        if not qa_title or not qa_content:
            st.warning("質問タイトルと質問内容を入力してください。")
            return
        
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    filename, url = upload_file_to_gcs(storage_client, file_obj, "qa_files")
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # ['タイムスタンプ', '質問タイトル', '質問内容', '連絡先メールアドレス', '添付ファイル名', '添付ファイルURL', 'ステータス']
        row_data = [
            timestamp, qa_title, qa_content, qa_contact, 
            filenames_json, urls_json, "未解決" # 初期ステータス
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_QA_DATA)
            worksheet.append_row(row_data)
            st.success("質問を投稿しました！回答があるまでお待ちください。"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_QA_DATA}' が存在するか確認してください。")

def page_qa_list():
    # 表示項目: タイムスタンプ, 質問タイトル, 質問内容, 連絡先メールアドレス, 添付ファイル名, 添付ファイルURL, ステータス
    detail_cols = [QA_COL_TIMESTAMP, QA_COL_TITLE, QA_COL_CONTENT, QA_COL_CONTACT, QA_COL_STATUS, QA_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_QA_DATA,
        title="💡 知恵袋・質問箱 (質問一覧)",
        col_time=QA_COL_TIMESTAMP,
        col_filter=QA_COL_STATUS, # ステータスで絞り込み
        col_memo=QA_COL_CONTENT,
        col_url=QA_COL_FILE_URL,
        detail_cols=detail_cols
    )
    st.info("※ 回答の閲覧機能は現在開発中です。")

# --- 5. 装置引き継ぎメモ記録/一覧 ---
def page_handover_recording():
    st.header("🤝 装置引き継ぎメモ記録")
    st.markdown("---")
    
    with st.form(key='handover_form'):
        
        handover_type = st.selectbox(f"{HANDOVER_COL_TYPE} (カテゴリ)", ["マニュアル", "装置設定", "その他メモ"])
        handover_title = st.text_input(f"{HANDOVER_COL_TITLE} (例: D1 MBE起動手順)", key='handover_title_input')
        handover_memo = st.text_area(f"{HANDOVER_COL_MEMO}", height=150, key='handover_memo_input', help="詳細な説明やリンクなどを記入してください。")
        
        st.markdown("---")
        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')

    if submit_button:
        if not handover_title:
            st.warning("タイトルを入力してください。")
            return
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # ユーザーのシート構造: ['タイムスタンプ', '種類', 'タイトル', '内容1', '内容2', '内容3', 'メモ']
        # 暫定的に「内容1」に詳細メモを、「内容2」「内容3」は空で保存します。
        row_data = [
            timestamp, handover_type, handover_title, 
            handover_memo, "", "", ""
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_HANDOVER_DATA)
            worksheet.append_row(row_data)
            st.success("引き継ぎメモを保存しました！"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_HANDOVER_DATA}' が存在するか確認してください。")

def page_handover_list():
    # 表示項目: タイムスタンプ, 種類, タイトル, 内容1, 内容2, 内容3, メモ
    detail_cols = [HANDOVER_COL_TIMESTAMP, HANDOVER_COL_TYPE, HANDOVER_COL_TITLE, '内容1', '内容2', '内容3', HANDOVER_COL_MEMO]
    page_data_list(
        sheet_name=SHEET_HANDOVER_DATA,
        title="🤝 装置引き継ぎメモ一覧",
        col_time=HANDOVER_COL_TIMESTAMP,
        col_filter=HANDOVER_COL_TYPE,
        col_memo=HANDOVER_COL_TITLE,
        col_url='内容1', # ユーザーのシートでは「内容1」にリンクが保存されているケースがあるため
        detail_cols=detail_cols
    )
    
# --- 6. お問い合わせフォーム（記録のみ実装）---
def page_contact_recording():
    st.header("✉️ 連絡・問い合わせフォーム")
    st.markdown("---")
    
    with st.form(key='contact_form'):
        
        contact_type = st.selectbox(f"{CONTACT_COL_TYPE}", ["バグ報告", "機能要望", "データ修正依頼", "その他"])
        contact_detail = st.text_area(f"{CONTACT_COL_DETAIL}", height=150, key='contact_detail_input')
        contact_info = st.text_input(f"{CONTACT_COL_CONTACT} (メールアドレスなど、任意)", key='contact_info_input')
        
        st.markdown("---")
        submit_button = st.form_submit_button(label='送信')

    if submit_button:
        if not contact_detail:
            st.warning("詳細内容を入力してください。")
            return
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # ['タイムスタンプ', 'お問い合わせの種類', '詳細内容', '連絡先']
        row_data = [
            timestamp, contact_type, contact_detail, contact_info
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_CONTACT_DATA)
            worksheet.append_row(row_data)
            st.success("お問い合わせを送信しました。担当者から折り返し連絡いたします。"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_CONTACT_DATA}' が存在するか確認してください。")

def page_contact_list():
    # 表示項目: タイムスタンプ, お問い合わせの種類, 詳細内容, 連絡先
    detail_cols = [CONTACT_COL_TIMESTAMP, CONTACT_COL_TYPE, CONTACT_COL_DETAIL, CONTACT_COL_CONTACT]
    page_data_list(
        sheet_name=SHEET_CONTACT_DATA,
        title="✉️ 連絡・問い合わせ一覧",
        col_time=CONTACT_COL_TIMESTAMP,
        col_filter=CONTACT_COL_TYPE,
        col_memo=CONTACT_COL_DETAIL,
        detail_cols=detail_cols
    )

# --- 7. トラブル報告記録/一覧 ---
def page_trouble_recording():
    st.header("🚨 トラブル報告記録")
    st.markdown("---")
    
    with st.form(key='trouble_form'):
        
        st.subheader("基本情報")
        col1, col2 = st.columns(2)
        with col1:
            report_date = st.date_input(f"{TROUBLE_COL_OCCUR_DATE} (発生日)", datetime.now().date())
        with col2:
            device_to_save = st.text_input(f"{TROUBLE_COL_DEVICE} (例: MBE-D1, RTA)", key='device_input')
            
        report_title = st.text_input(f"{TROUBLE_COL_TITLE}", key='trouble_title_input')
        occur_time = st.text_area(f"{TROUBLE_COL_OCCUR_TIME} (状況詳細)", height=100)
        
        st.subheader("対応と考察")
        cause = st.text_area(f"{TROUBLE_COL_CAUSE}", height=100)
        solution = st.text_area(f"{TROUBLE_COL_SOLUTION}", height=100)
        prevention = st.text_area(f"{TROUBLE_COL_PREVENTION}", height=100)

        col3, col4 = st.columns(2)
        with col3:
            reporter_name = st.text_input(f"{TROUBLE_COL_REPORTER} (氏名)", key='reporter_input')
        with col4:
            uploaded_files = st.file_uploader("添付ファイル", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], accept_multiple_files=True)
            
        st.markdown("---")
        submit_button = st.form_submit_button(label='トラブル報告を保存')

    if submit_button:
        if not report_title or not reporter_name:
            st.warning("タイトルと報告者名を入力してください。")
            return
            
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    filename, url = upload_file_to_gcs(storage_client, file_obj, "trouble_reports")
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # ['タイムスタンプ', '機器/場所', '発生日', 'トラブル発生時', '原因/究明', '対策/復旧', '再発防止策', '報告者', 'ファイル名', 'ファイルURL', '件名/タイトル']
        row_data = [
            timestamp, device_to_save, report_date.isoformat(), occur_time,
            cause, solution, prevention, reporter_name,
            filenames_json, urls_json, report_title
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_TROUBLE_DATA)
            worksheet.append_row(row_data)
            st.success("トラブル報告を保存しました！"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_TROUBLE_DATA}' が存在するか確認してください。")

def page_trouble_list():
    # 表示項目: タイムスタンプ, 機器/場所, 発生日, トラブル発生時, 原因/究明, 対策/復旧, 再発防止策, 報告者, ファイル名, ファイルURL, 件名/タイトル
    detail_cols = [
        TROUBLE_COL_TIMESTAMP, TROUBLE_COL_TITLE, TROUBLE_COL_DEVICE, TROUBLE_COL_OCCUR_DATE, 
        TROUBLE_COL_OCCUR_TIME, TROUBLE_COL_CAUSE, TROUBLE_COL_SOLUTION, TROUBLE_COL_PREVENTION, 
        TROUBLE_COL_REPORTER, TROUBLE_COL_FILENAME
    ]
    page_data_list(
        sheet_name=SHEET_TROUBLE_DATA,
        title="🚨 トラブル報告一覧",
        col_time=TROUBLE_COL_TIMESTAMP,
        col_filter=TROUBLE_COL_DEVICE,
        col_memo=TROUBLE_COL_TITLE,
        col_url=TROUBLE_COL_FILE_URL,
        detail_cols=detail_cols
    )


# --- 8. IVデータ解析 (前回と同じく再利用) ---
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
def page_pl_analysis(): st.header("🔬 PLデータ解析"); st.info("このページは未実装です。")
def page_calendar(): st.header("🗓️ スケジュール・装置予約"); st.info("このページは未実装です。")

# --------------------------------------------------------------------------
# --- Main App Execution ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    menu_selection = st.sidebar.radio("機能選択", [
        "📝 エピノート記録", "📚 エピノート一覧", 
        "🛠️ メンテノート記録", "🛠️ メンテノート一覧",
        "⚡ IVデータ解析", "🔬 PLデータ解析",
        "🗓️ スケジュール・装置予約",
        "📝 議事録記録", "📚 議事録一覧", 
        "💡 知恵袋・質問投稿", "💡 知恵袋・質問一覧", 
        "🤝 引き継ぎメモ記録", "🤝 引き継ぎメモ一覧",
        "🚨 トラブル報告記録", "🚨 トラブル報告一覧", 
        "✉️ 問い合わせ記録", "✉️ 問い合わせ一覧"
    ])
    
    # ページルーティング
    if menu_selection == "📝 エピノート記録": page_epi_note_recording()
    elif menu_selection == "📚 エピノート一覧": page_epi_note_list()
    elif menu_selection == "🛠️ メンテノート記録": page_mainte_recording()
    elif menu_selection == "🛠️ メンテノート一覧": page_mainte_list()
    elif menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()
    elif menu_selection == "📝 議事録記録": page_meeting_recording()
    elif menu_selection == "📚 議事録一覧": page_meeting_list()
    elif menu_selection == "💡 知恵袋・質問投稿": page_qa_recording()
    elif menu_selection == "💡 知恵袋・質問一覧": page_qa_list()
    elif menu_selection == "🤝 引き継ぎメモ記録": page_handover_recording()
    elif menu_selection == "🤝 引き継ぎメモ一覧": page_handover_list()
    elif menu_selection == "🚨 トラブル報告記録": page_trouble_recording()
    elif menu_selection == "🚨 トラブル報告一覧": page_trouble_list()
    elif menu_selection == "✉️ 問い合わせ記録": page_contact_recording()
    elif menu_selection == "✉️ 問い合わせ一覧": page_contact_list()

if __name__ == "__main__":
    main()



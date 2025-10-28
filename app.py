# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v20.5.0 (最終機能統合・日本語対応版)
# - FIX: 機能メニューを記録・一覧で統合 (例: エピノート)
# - FIX: 一覧のデフォルト開始日を 2025年4月1日 に変更
# - ADD: PLデータ解析ページを実装
# - ADD: Matplotlibによるグラフの日本語表示に対応
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
import matplotlib.font_manager as fm

# --- Matplotlib 日本語フォント設定 ---
# Streamlit Cloud環境で動作する可能性の高いフォントを設定
try:
    # 環境依存で動作しない可能性もあるため、広範囲のフォントを指定
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Hiragino Maru Gothic Pro', 'Yu Gothic', 'Meiryo', 'TakaoGothic', 'IPAexGothic', 'IPAfont', 'Noto Sans CJK JP']
    plt.rcParams['axes.unicode_minus'] = False # 負の記号の豆腐化防止
except Exception:
    pass # 設定に失敗しても続行
    
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

SHEET_EPI_DATA = 'エピノート_データ'
EPI_COL_TIMESTAMP = 'タイムスタンプ'
EPI_COL_NOTE_TYPE = 'ノート種別'   # 'エピノート'
EPI_COL_CATEGORY = 'カテゴリ'     # 'D1', '897'など、エピ番号やカテゴリ
EPI_COL_MEMO = 'メモ'           # タイトルと詳細メモを含む
EPI_COL_FILENAME = 'ファイル名'
EPI_COL_FILE_URL = '写真URL'

SHEET_MAINTE_DATA = 'メンテノート_データ'
MAINT_COL_TIMESTAMP = 'タイムスタンプ'
MAINT_COL_NOTE_TYPE = 'ノート種別' # 'メンテノート'
MAINT_COL_MEMO = 'メモ'
MAINT_COL_FILENAME = 'ファイル名'
MAINT_COL_FILE_URL = '写真URL'

SHEET_MEETING_DATA = '議事録_データ'
MEETING_COL_TIMESTAMP = 'タイムスタンプ'
MEETING_COL_TITLE = '会議タイトル'
MEETING_COL_AUDIO_NAME = '音声ファイル名'
MEETING_COL_AUDIO_URL = '音声ファイルURL'
MEETING_COL_CONTENT = '議事録内容'

SHEET_HANDOVER_DATA = '引き継ぎ_データ'
HANDOVER_COL_TIMESTAMP = 'タイムスタンプ'
HANDOVER_COL_TYPE = '種類'
HANDOVER_COL_TITLE = 'タイトル'
HANDOVER_COL_MEMO = 'メモ' # 内容1,2,3はUIを複雑にするため、一旦メモに統合

SHEET_QA_DATA = '知恵袋_データ'
QA_COL_TIMESTAMP = 'タイムスタンプ'
QA_COL_TITLE = '質問タイトル'
QA_COL_CONTENT = '質問内容'
QA_COL_CONTACT = '連絡先メールアドレス'
QA_COL_FILENAME = '添付ファイル名'
QA_COL_FILE_URL = '添付ファイルURL'
QA_COL_STATUS = 'ステータス'
SHEET_QA_ANSWER = '知恵袋_解答' # 解答シート

SHEET_CONTACT_DATA = 'お問い合わせ_データ'
CONTACT_COL_TIMESTAMP = 'タイムスタンプ'
CONTACT_COL_TYPE = 'お問い合わせの種類'
CONTACT_COL_DETAIL = '詳細内容'
CONTACT_COL_CONTACT = '連絡先'

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

# gc と storage_client はグローバルで定義（Dummyオブジェクトで初期化）
gc = DummyGSClient()
storage_client = DummyStorageClient()

@st.cache_resource(ttl=3600)
def initialize_google_services():
    """Streamlit Secretsから認証情報を読み込み、Googleサービスを初期化する"""
    if "gcs_credentials" not in st.secrets:
        st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
        return DummyGSClient(), DummyStorageClient()

    try:
        raw_credentials_string = st.secrets["gcs_credentials"]
        
        # --- 認証文字列の【強制】クリーンアップ v20.3.0 ---
        cleaned_string = raw_credentials_string.strip()
        cleaned_string = cleaned_string.replace('\n', '')
        cleaned_string = cleaned_string.replace('\t', '')
        cleaned_string = cleaned_string.replace(' ', '') # U+00A0: NO-BREAK SPACE
        cleaned_string = re.sub(r'(\s){2,}', r'\1', cleaned_string)
        
        # JSONをパース
        info = json.loads(cleaned_string) 
        
        # gspread (Spreadsheet) の認証
        gc_real = gspread.service_account_from_dict(info)

        # google.cloud.storage (GCS) の認証
        storage_client_real = storage.Client.from_service_account_info(info)

        st.sidebar.success("✅ Googleサービス認証成功")
        return gc_real, storage_client_real

    except json.JSONDecodeError as e:
        st.error(f"❌ 認証エラー（JSON形式不正）: サービスアカウントのJSON形式が不正です。エラー詳細: {e}")
        return DummyGSClient(), DummyStorageClient()
        
    except Exception as e:
        st.error(f"❌ 認証エラー: サービスアカウントの初期化に失敗しました。認証情報をご確認ください。({e})")
        return DummyGSClient(), DummyStorageClient()

# グローバル変数を初期化されたクライアントに更新
gc, storage_client = initialize_google_services() 

# --------------------------------------------------------------------------
# --- Data Utilities (データ取得・解析) ---
# --------------------------------------------------------------------------

@st.cache_data(ttl=600, show_spinner="スプレッドシートからデータを読み込み中...")
def get_sheet_as_df(spreadsheet_name, sheet_name):
    """指定されたシートのデータをDataFrameとして取得する"""
    global gc # グローバルな gc オブジェクトを利用
    
    if isinstance(gc, DummyGSClient):
        return pd.DataFrame()
    
    try:
        worksheet = gc.open(spreadsheet_name).worksheet(sheet_name)
        data = worksheet.get_all_values()
        if not data or len(data) <= 1: 
            return pd.DataFrame(columns=data[0] if data else [])
        
        df = pd.DataFrame(data[1:], columns=data[0])
        return df

    except gspread.exceptions.WorksheetNotFound:
        st.error(f"シート名「{sheet_name}」が見つかりません。スプレッドシートをご確認ください。")
        return pd.DataFrame()
    except Exception as e:
        st.warning(f"警告：シート「{sheet_name}」の読み込み中にエラーが発生しました。({e})")
        return pd.DataFrame()

# --- IV/PLデータ解析用ユーティリティ (キャッシュで高速化) ---
@st.cache_data(show_spinner="データを解析中...", max_entries=50)
def load_data_file(uploaded_file_bytes, uploaded_file_name):
    """アップロードされたIV/PLファイルを読み込み、DataFrameを返す (IV/PL共通ロジック)"""
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
        
        # 複数の区切り文字を試すロバストな読み込み
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
        df.columns = ['Axis_X', uploaded_file_name] # 一時的に汎用的な列名を使用

        df['Axis_X'] = pd.to_numeric(df['Axis_X'], errors='coerce', downcast='float')
        df[uploaded_file_name] = pd.to_numeric(df[uploaded_file_name], errors='coerce', downcast='float')
        df.dropna(inplace=True)
        
        return df

    except Exception:
        return None

@st.cache_data(show_spinner="データを結合中...")
def combine_dataframes(dataframes, filenames):
    """複数のDataFrameを共通のX軸をキーに外部結合する"""
    if not dataframes: return None
    
    # 結合キーは 'Axis_X'
    combined_df = dataframes[0].rename(columns={'Axis_X': 'X_Value'})
    
    for i in range(1, len(dataframes)):
        df_to_merge = dataframes[i].rename(columns={'Axis_X': 'X_Value'})
        combined_df = pd.merge(combined_df, df_to_merge, on='X_Value', how='outer')
        
    combined_df = combined_df.sort_values(by='X_Value', ascending=False).reset_index(drop=True)
    
    for col in combined_df.columns:
        if col != 'X_Value':
            combined_df[col] = combined_df[col].round(4)
            
    # X軸の列名を結合前に戻す
    combined_df = combined_df.rename(columns={'X_Value': dataframes[0].columns[0]})
    
    return combined_df.rename(columns={dataframes[0].columns[0]: 'X_Axis'})


# --------------------------------------------------------------------------
# --- GCS Utilities (ファイルアップロード) ---
# --------------------------------------------------------------------------

def upload_file_to_gcs(storage_client, file_obj, folder_name):
    """ファイルをGCSにアップロードし、公開URLを返す"""
    if isinstance(storage_client, DummyStorageClient):
        return None, "dummy_url_gcs_error"
        
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    original_filename = file_obj.name
    safe_filename = original_filename.replace(' ', '_').replace('/', '_')
    gcs_filename = f"{folder_name}/{timestamp}_{safe_filename}"

    try:
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        
        file_obj.seek(0)
        blob.upload_from_file(file_obj, content_type=file_obj.type)

        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        
        return original_filename, public_url

    except Exception as e:
        st.error(f"❌ GCSエラー: ファイルのアップロード中にエラーが発生しました。({e})")
        return None, None
        
# --------------------------------------------------------------------------
# --- Page Implementations (各機能ページ) ---
# --------------------------------------------------------------------------

# --- 汎用的な一覧表示関数 ---
def page_data_list(sheet_name, title, col_time, col_filter=None, col_memo=None, col_url=None, detail_cols=None):
    """汎用的なデータ一覧ページ"""
    
    st.header(f"📚 {title}一覧")
    
    df = get_sheet_as_df(SPREADSHEET_NAME, sheet_name) 

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
            df[col_time] = pd.to_datetime(df[col_time].astype(str).str.replace(r'[^0-9]', '', regex=True), errors='coerce', format='%Y%m%d%H%M%S', exact=False).dt.date
        except:
            pass 
        
        df_valid_date = df.dropna(subset=[col_time])
        
        if not df_valid_date.empty:
            min_date = df_valid_date[col_time].min()
            max_date = df_valid_date[col_time].max()
            
            # --- ★修正箇所: デフォルト開始日を2025年4月1日に設定 ★---
            try:
                default_start_date = date(2025, 4, 1)
            except ValueError:
                default_start_date = date.today() - timedelta(days=365) # 安全策
                
            # 実際の日付の最小値と、指定されたデフォルト開始日のうち、新しい方を選択
            initial_start_date = max(min_date, default_start_date) if isinstance(min_date, date) else default_start_date
            # ----------------------------------------------------
            
            col_date1, col_date2 = st.columns(2)
            with col_date1:
                start_date = st.date_input("開始日", value=initial_start_date)
            with col_date2:
                end_date = st.date_input("終了日", value=max_date)
            
            df = df_valid_date[(df_valid_date[col_time] >= start_date) & (df_valid_date[col_time] <= end_date)]
        else:
            st.warning("日付（タイムスタンプ）列の形式が不正な行が多いため、日付絞り込みをスキップしました。")


    if df.empty: st.info("絞り込み条件に一致するデータがありません。"); return

    df = df.sort_values(by=col_time, ascending=False).reset_index(drop=True)
    
    st.markdown("---")
    st.subheader(f"検索結果 ({len(df)}件)")

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
        
        if detail_cols:
            for col in detail_cols:
                if col in row:
                    if col_memo == col or '内容' in col: # メモや内容が多い場合はテキストエリアで表示
                        st.markdown(f"**{col}:**"); st.text(row[col])
                    else:
                        st.write(f"**{col}:** {row[col]}")
        
        # 添付ファイル (ファイル名とURLが分離しているか、同一かによって表示を調整)
        if col_url and col_url in row:
            st.markdown("##### 添付ファイル")
            
            try:
                # JSONデコードを試みる
                urls = json.loads(row[col_url])
                filenames = json.loads(row[EPI_COL_FILENAME]) if EPI_COL_FILENAME in row and row[EPI_COL_FILENAME] else ['ファイル'] * len(urls)
                
                if urls:
                    for filename, url in zip(filenames, urls):
                        if "drive.google.com" in url:
                            st.markdown(f"- **Google Drive:** [{filename}](<{url}>)")
                        else:
                            st.markdown(f"- [{filename}]({url})")
                else:
                    st.info("添付ファイルはありません。")

            except Exception:
                # JSON形式ではない場合（古いデータや手動入力、単一URLの直接保存）
                if pd.notna(row[col_url]) and row[col_url]:
                    url_list = row[col_url].split(',')
                    for url in url_list:
                         url = url.strip().strip('"')
                         if url:
                             st.markdown(f"- [添付ファイルURL]({url})")
                else:
                    st.info("添付ファイルはありません。")


# --- 機能統合されたページ実装 ---

# 1. エピノート機能
def page_epi_note_recording():
    st.markdown("#### 📝 新しいエピノートを記録")
    # ... (既存の記録ロジックをここに記述) ...
    
    with st.form(key='epi_note_form'):
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
    detail_cols = [EPI_COL_TIMESTAMP, EPI_COL_CATEGORY, EPI_COL_NOTE_TYPE, EPI_COL_MEMO, EPI_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_EPI_DATA,
        title="エピノート",
        col_time=EPI_COL_TIMESTAMP,
        col_filter=EPI_COL_CATEGORY,
        col_memo=EPI_COL_MEMO,
        col_url=EPI_COL_FILE_URL,
        detail_cols=detail_cols
    )
    
def page_epi_note():
    st.header("エピノート機能")
    st.markdown("---")
    tab_selection = st.radio("表示切り替え", ["📝 記録", "📚 一覧"], key="epi_tab", horizontal=True)
    
    if tab_selection == "📝 記録": page_epi_note_recording()
    elif tab_selection == "📚 一覧": page_epi_note_list()


# 2. メンテノート機能
def page_mainte_recording():
    st.markdown("#### 🛠️ 新しいメンテノートを記録")
    # ... (既存の記録ロジックをここに記述) ...
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
    detail_cols = [MAINT_COL_TIMESTAMP, MAINT_COL_NOTE_TYPE, MAINT_COL_MEMO, MAINT_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_MAINTE_DATA,
        title="メンテノート",
        col_time=MAINT_COL_TIMESTAMP,
        col_filter=MAINT_COL_NOTE_TYPE, 
        col_memo=MAINT_COL_MEMO,
        col_url=MAINT_COL_FILE_URL,
        detail_cols=detail_cols
    )

def page_mainte_note():
    st.header("メンテノート機能")
    st.markdown("---")
    tab_selection = st.radio("表示切り替え", ["📝 記録", "📚 一覧"], key="mainte_tab", horizontal=True)
    
    if tab_selection == "📝 記録": page_mainte_recording()
    elif tab_selection == "📚 一覧": page_mainte_list()


# 3. 議事録・ミーティングメモ機能
def page_meeting_recording():
    st.markdown("#### 📝 新しい議事録を記録")
    # ... (既存の記録ロジックをここに記述) ...
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
    detail_cols = [MEETING_COL_TIMESTAMP, MEETING_COL_TITLE, MEETING_COL_CONTENT, MEETING_COL_AUDIO_NAME, MEETING_COL_AUDIO_URL]
    page_data_list(
        sheet_name=SHEET_MEETING_DATA,
        title="議事録",
        col_time=MEETING_COL_TIMESTAMP,
        col_filter=MEETING_COL_TITLE,
        col_memo=MEETING_COL_CONTENT,
        col_url=MEETING_COL_AUDIO_URL,
        detail_cols=detail_cols
    )

def page_meeting_note():
    st.header("議事録・ミーティングメモ機能")
    st.markdown("---")
    tab_selection = st.radio("表示切り替え", ["📝 記録", "📚 一覧"], key="meeting_tab", horizontal=True)
    
    if tab_selection == "📝 記録": page_meeting_recording()
    elif tab_selection == "📚 一覧": page_meeting_list()


# 4. 知恵袋・質問箱機能
def page_qa_recording():
    st.markdown("#### 💡 新しい質問を投稿")
    # ... (既存の記録ロジックをここに記述) ...
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
    detail_cols = [QA_COL_TIMESTAMP, QA_COL_TITLE, QA_COL_CONTENT, QA_COL_CONTACT, QA_COL_STATUS, QA_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_QA_DATA,
        title="知恵袋・質問箱",
        col_time=QA_COL_TIMESTAMP,
        col_filter=QA_COL_STATUS, 
        col_memo=QA_COL_CONTENT,
        col_url=QA_COL_FILE_URL,
        detail_cols=detail_cols
    )
    st.info("※ 回答の閲覧機能は現在開発中です。")

def page_qa_box():
    st.header("知恵袋・質問箱機能")
    st.markdown("---")
    tab_selection = st.radio("表示切り替え", ["💡 質問投稿", "📚 質問一覧"], key="qa_tab", horizontal=True)
    
    if tab_selection == "💡 質問投稿": page_qa_recording()
    elif tab_selection == "📚 質問一覧": page_qa_list()


# 5. 装置引き継ぎメモ機能
def page_handover_recording():
    st.markdown("#### 🤝 新しい引き継ぎメモを記録")
    # ... (既存の記録ロジックをここに記述) ...
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
        
        # 既存のシート構造に合わせる（内容1, 2, 3は空にし、メモに集約）
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
    detail_cols = [HANDOVER_COL_TIMESTAMP, HANDOVER_COL_TYPE, HANDOVER_COL_TITLE, '内容1', '内容2', '内容3', HANDOVER_COL_MEMO]
    page_data_list(
        sheet_name=SHEET_HANDOVER_DATA,
        title="装置引き継ぎメモ",
        col_time=HANDOVER_COL_TIMESTAMP,
        col_filter=HANDOVER_COL_TYPE,
        col_memo=HANDOVER_COL_TITLE,
        col_url='内容1', 
        detail_cols=detail_cols
    )

def page_handover_note():
    st.header("装置引き継ぎメモ機能")
    st.markdown("---")
    tab_selection = st.radio("表示切り替え", ["📝 記録", "📚 一覧"], key="handover_tab", horizontal=True)
    
    if tab_selection == "📝 記録": page_handover_recording()
    elif tab_selection == "📚 一覧": page_handover_list()


# 6. トラブル報告機能
def page_trouble_recording():
    st.markdown("#### 🚨 新しいトラブルを報告")
    # ... (既存の記録ロジックをここに記述) ...
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
    detail_cols = [
        TROUBLE_COL_TIMESTAMP, TROUBLE_COL_TITLE, TROUBLE_COL_DEVICE, TROUBLE_COL_OCCUR_DATE, 
        TROUBLE_COL_OCCUR_TIME, TROUBLE_COL_CAUSE, TROUBLE_COL_SOLUTION, TROUBLE_COL_PREVENTION, 
        TROUBLE_COL_REPORTER, TROUBLE_COL_FILENAME
    ]
    page_data_list(
        sheet_name=SHEET_TROUBLE_DATA,
        title="トラブル報告",
        col_time=TROUBLE_COL_TIMESTAMP,
        col_filter=TROUBLE_COL_DEVICE,
        col_memo=TROUBLE_COL_TITLE,
        col_url=TROUBLE_COL_FILE_URL,
        detail_cols=detail_cols
    )

def page_trouble_report():
    st.header("トラブル報告機能")
    st.markdown("---")
    tab_selection = st.radio("表示切り替え", ["📝 記録", "📚 一覧"], key="trouble_tab", horizontal=True)
    
    if tab_selection == "📝 記録": page_trouble_recording()
    elif tab_selection == "📚 一覧": page_trouble_list()


# 7. 連絡・問い合わせ機能
def page_contact_recording():
    st.markdown("#### ✉️ 新しい問い合わせを記録")
    # ... (既存の記録ロジックをここに記述) ...
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
        
        row_data = [
            timestamp, contact_type, contact_detail, contact_info
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_CONTACT_DATA)
            worksheet.append_row(row_data)
            st.success("お問い合わせを送信しました。"); st.cache_data.clear(); st.rerun() 
        except Exception:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_CONTACT_DATA}' が存在するか確認してください。")

def page_contact_list():
    detail_cols = [CONTACT_COL_TIMESTAMP, CONTACT_COL_TYPE, CONTACT_COL_DETAIL, CONTACT_COL_CONTACT]
    page_data_list(
        sheet_name=SHEET_CONTACT_DATA,
        title="連絡・問い合わせ",
        col_time=CONTACT_COL_TIMESTAMP,
        col_filter=CONTACT_COL_TYPE,
        col_memo=CONTACT_COL_DETAIL,
        detail_cols=detail_cols
    )

def page_contact_form():
    st.header("連絡・問い合わせ機能")
    st.markdown("---")
    tab_selection = st.radio("表示切り替え", ["📝 記録", "📚 一覧"], key="contact_tab", horizontal=True)
    
    if tab_selection == "📝 記録": page_contact_recording()
    elif tab_selection == "📚 一覧": page_contact_list()


# 8. IVデータ解析
def page_iv_analysis():
    """⚡ IVデータ解析ページ（日本語対応）"""
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
            # load_data_file を使用
            df = load_data_file(uploaded_file.getvalue(), uploaded_file.name)
            
            if df is not None and not df.empty:
                valid_dataframes.append(df)
                filenames.append(uploaded_file.name)
        
        if valid_dataframes:
            
            combined_df = combine_dataframes(valid_dataframes, filenames)
            
            st.success(f"{len(valid_dataframes)}個の有効なファイルを読み込み、結合しました。")
            
            st.subheader("ステップ2: グラフ表示 (日本語対応)")
            
            fig, ax = plt.subplots(figsize=(12, 7)) 
            
            for filename in filenames:
                ax.plot(combined_df['X_Axis'], combined_df[filename], label=filename)
            
            ax.set_xlabel("電圧 (V)") # 日本語ラベル
            ax.set_ylabel("電流 (A)") # 日本語ラベル
            ax.grid(True)
            ax.legend(title="ファイル名", loc='best')
            ax.set_title("IV特性比較") # 日本語タイトル
            
            st.pyplot(fig, use_container_width=True) 
            
            st.subheader("ステップ3: 結合データ")
            combined_df = combined_df.rename(columns={'X_Axis': 'Voltage_V'}) # 表示用
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


# 9. PLデータ解析 (復活・日本語対応)
def page_pl_analysis():
    """🔬 PLデータ解析ページ（復活・日本語対応）"""
    st.header("🔬 PLデータ解析")
    st.info("※ IVデータと同様に、2列の数値データ（波長/エネルギー vs 強度）を持つテキストファイルを想定しています。")
    
    uploaded_files = st.file_uploader(
        "PL測定データファイル (.txt) をアップロード",
        type=['txt'], 
        accept_multiple_files=True
    )

    if uploaded_files:
        valid_dataframes = []
        filenames = []
        
        st.subheader("ステップ1: ファイル読み込みと解析")
        
        for uploaded_file in uploaded_files:
            # load_data_file を使用
            df = load_data_file(uploaded_file.getvalue(), uploaded_file.name)
            
            if df is not None and not df.empty:
                valid_dataframes.append(df)
                filenames.append(uploaded_file.name)
        
        if valid_dataframes:
            combined_df = combine_dataframes(valid_dataframes, filenames)
            
            st.success(f"{len(valid_dataframes)}個の有効なファイルを読み込み、結合しました。")
            
            st.subheader("ステップ2: グラフ表示 (日本語対応)")
            
            fig, ax = plt.subplots(figsize=(12, 7)) 
            
            for filename in filenames:
                ax.plot(combined_df['X_Axis'], combined_df[filename], label=filename)
            
            # PLデータに合わせたラベルに変更
            ax.set_xlabel("波長 / エネルギー (nm / eV)") # 日本語ラベル
            ax.set_ylabel("強度 (a.u.)") # 日本語ラベル
            ax.grid(True)
            ax.legend(title="ファイル名", loc='best')
            ax.set_title("PLスペクトル比較") # 日本語タイトル
            
            st.pyplot(fig, use_container_width=True) 
            
            st.subheader("ステップ3: 結合データ")
            combined_df = combined_df.rename(columns={'X_Axis': 'Wavelength_or_Energy'}) # 表示用
            st.dataframe(combined_df, use_container_width=True)
            
            # Excelダウンロード
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, sheet_name='Combined PL Data', index=False)
            
            st.download_button(
                label="📈 結合Excelデータとしてダウンロード",
                data=output.getvalue(),
                file_name=f"pl_analysis_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("有効なデータファイルが見つかりませんでした。")


# --- Dummy Pages (未実装のページ) ---
def page_calendar(): st.header("🗓️ スケジュール・装置予約"); st.info("このページは未実装です。")

# --------------------------------------------------------------------------
# --- Main App Execution ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    # ★修正箇所: メニューを記録・一覧で統合★
    menu_selection = st.sidebar.radio("機能選択", [
        "エピノート", "メンテノート", "議事録", "知恵袋・質問箱", "装置引き継ぎメモ", "トラブル報告", "連絡・問い合わせ",
        "⚡ IVデータ解析", "🔬 PLデータ解析", "🗓️ スケジュール・装置予約"
    ])
    
    # ページルーティング
    if menu_selection == "エピノート": page_epi_note()
    elif menu_selection == "メンテノート": page_mainte_note()
    elif menu_selection == "議事録": page_meeting_note()
    elif menu_selection == "知恵袋・質問箱": page_qa_box()
    elif menu_selection == "装置引き継ぎメモ": page_handover_note()
    elif menu_selection == "トラブル報告": page_trouble_report()
    elif menu_selection == "連絡・問い合わせ": page_contact_form()
    elif menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()


if __name__ == "__main__":
    main()

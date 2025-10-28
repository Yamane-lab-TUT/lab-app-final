# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v20.6.1 (PL波長校正対応版)
# - NEW: load_pl_data(uploaded_file) 関数を追加し、データカラム名を 'pixel', 'intensity' に固定。
# - CHG: page_pl_analysis() をユーザー提供の波長校正ロジックに置き換え、校正係数をセッションステートで保持。
# - FIX: 全てのリストでデフォルト開始日を2025/4/1に設定 (v20.6.0から変更なし)。
# --------------------------------------------------------------------------
# [FIXED BY GEMINI] IVデータ解析ロジックを安定版に置き換え (UnboundLocalError, 10ファイル制限対応)
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
from datetime import datetime, date, timedelta
from urllib.parse import quote as url_quote
from io import BytesIO
import calendar
import matplotlib.font_manager as fm

# GCSクライアントのインポート 
try:
    from google.cloud import storage
except ImportError:
    st.error("❌ 警告: `google-cloud-storage` ライブラリが見つかりません。")
    pass
    
# --- Matplotlib 日本語フォント設定 ---
try:
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Hiragino Maru Gothic Pro', 'Yu Gothic', 'Meiryo', 'TakaoGothic', 'IPAexGothic', 'IPAfont', 'Noto Sans CJK JP']
    plt.rcParams['axes.unicode_minus'] = False
except Exception:
    pass
    
# --- Global Configuration & Setup ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
# ↓↓↓↓↓↓ 【重要】ご自身の「バケット名」に書き換えてください ↓↓↓↓↓↓
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" 
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
MAX_COMBINED_FILES = 10 # 結合データを作成するファイルの最大数 [NEW]

SPREADSHEET_NAME = 'エピノート' # Google Spreadsheetのファイル名

# --- SPREADSHEET COLUMN HEADERS (お客様のデータ構造に完全一致) ---

SHEET_EPI_DATA = 'エピノート_データ'
EPI_COL_TIMESTAMP = 'タイムスタンプ'
EPI_COL_NOTE_TYPE = 'ノート種別'
EPI_COL_CATEGORY = 'カテゴリ'
EPI_COL_MEMO = 'メモ' # タイトルと詳細メモを含む
EPI_COL_FILENAME = 'ファイル名'
EPI_COL_FILE_URL = '写真URL'

SHEET_MAINTE_DATA = 'メンテノート_データ'
MAINT_COL_TIMESTAMP = 'タイムスタンプ'
MAINT_COL_NOTE_TYPE = 'ノート種別'
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
    
    if 'storage' not in globals():
        st.error("❌ 致命的なエラー: `google.cloud.storage` のインポートに失敗しました。Streamlitの環境依存と思われます。")
        return DummyGSClient(), DummyStorageClient()
        
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
    global gc
    
    if isinstance(gc, DummyGSClient):
        st.warning("⚠️ 認証エラーのためダミーデータを返します。")
        return pd.DataFrame()
    
    try:
        worksheet = gc.open(spreadsheet_name).worksheet(sheet_name)
        data = worksheet.get_all_values()
        
        if not data or len(data) <= 1: 
            # データがない場合、ヘッダー行だけを基に空のDataFrameを作成
            return pd.DataFrame(columns=data[0] if data else [])
        
        df = pd.DataFrame(data[1:], columns=data[0])
        return df

    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ シート名「{sheet_name}」が見つかりません。スプレッドシートをご確認ください。")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ シート「{sheet_name}」の読み込み中にエラーが発生しました。({e})")
        return pd.DataFrame()

# --- IV/PLデータ解析用コアユーティリティ (PL解析のために維持) ---
def _load_two_column_data_core(uploaded_file_bytes, column_names):
    """IV/PLデータファイルから2列のデータを読み込み、指定されたカラム名を付けてDataFrameを返す"""
    try:
        # ロバストな読み込みロジック 
        # (utf-8デコード済みで渡される前提だったが、load_pl_dataがgetvalue()を渡すため修正なし)
        content = uploaded_file_bytes.decode('utf-8').splitlines()
        data_lines = content[1:] # 1行目をヘッダーとしてスキップ

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
        df.columns = column_names # 指定されたカラム名を使用

        df[column_names[0]] = pd.to_numeric(df[column_names[0]], errors='coerce', downcast='float')
        df[column_names[1]] = pd.to_numeric(df[column_names[1]], errors='coerce', downcast='float')
        df.dropna(inplace=True)
        
        return df

    except Exception:
        return None

# --- IVデータ解析用 (安定版に置き換え) ---
@st.cache_data(show_spinner="IVデータを解析中...", max_entries=50)
def load_iv_data(uploaded_file):
    """アップロードされたIVデータファイル（TXT/CSV）をロバストに読み込む関数。"""
    
    file_name = uploaded_file.name
    
    # ファイルをバイナリとして読み込み、文字列にデコード（UTF-8, Shift-JISを試行）
    try:
        data_string = uploaded_file.getvalue().decode('utf-8')
    except UnicodeDecodeError:
        try:
            data_string = uploaded_file.getvalue().decode('shift_jis')
        except:
            return None, file_name

    try:
        data_io = io.StringIO(data_string)
        
        # skiprows=1で最初のヘッダー行をスキップし、タブ/スペース区切りで読み込む
        df = pd.read_csv(data_io, sep=r'\s+', skiprows=1, header=None, names=['VF(V)', 'IF(A)'])
        
        # データ型を数値に変換（エラーがある行は無視）
        df['VF(V)'] = pd.to_numeric(df['VF(V)'], errors='coerce')
        df['IF(A)'] = pd.to_numeric(df['IF(A)'], errors='coerce')
        df.dropna(inplace=True)

        return df, file_name

    except Exception:
        return None, file_name


# --- PLデータ解析用 (元のコードを維持) ---
@st.cache_data(show_spinner="PLデータを解析中...", max_entries=50)
def load_pl_data(uploaded_file):
    """PLファイル (pixel vs intensity) を読み込み、DataFrame (pixel, intensity) を返す"""
    df = _load_two_column_data_core(uploaded_file.getvalue(), ['pixel', 'intensity'])
    # load_pl_dataは、uploaded_fileオブジェクトを直接受け取るため、getvalue()を使用
    if df is not None and not df.empty:
        return df[['pixel', 'intensity']]
    return None

# (元のコードにあった combine_dataframes は、新しい IV ロジックで不要になったため削除)


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
        # Google Cloud Storageの公開URL形式
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        return original_filename, public_url
    except Exception as e:
        st.error(f"❌ GCSエラー: ファイルのアップロード中にエラーが発生しました。({e})")
        return None, None

# --------------------------------------------------------------------------
# --- Page Implementations (各機能ページ) ---
# --------------------------------------------------------------------------

# --- 汎用的な一覧表示関数 (元のコードを維持) ---
def page_data_list(sheet_name, title, col_time, col_filter=None, col_memo=None, col_url=None, detail_cols=None):
    """汎用的なデータ一覧ページ (R2, R3, R1対応)"""
    st.header(f"📚 {title}一覧")
    df = get_sheet_as_df(SPREADSHEET_NAME, sheet_name)
    if df.empty:
        st.info("データがありません。")
        return

    # ... (元の絞り込み・検索ロジックを維持) ...
    st.subheader("絞り込みと検索")
    
    # カテゴリ・ステータスによる絞り込み
    if col_filter and col_filter in df.columns: 
        # 空白データを 'なし' として扱う
        df[col_filter] = df[col_filter].fillna('なし')
        filter_options = ["すべて"] + sorted(list(df[col_filter].unique()))
        data_filter = st.selectbox(f"「{col_filter}」で絞り込み", filter_options)
        if data_filter != "すべて":
            df = df[df[col_filter] == data_filter]

    # 日付による絞り込み (R2: 開始日を2025/4/1に固定)
    if col_time and col_time in df.columns:
        # タイムスタンプ列のクリーンアップと日付型への変換
        try:
            # タイムスタンプ形式 ('YYYYMMDD_HHMMSS' または 'YYYYMMDDHHMMSS') から日付部分のみを取得
            df['date_only'] = pd.to_datetime(
                df[col_time].astype(str).str.replace(r'[^0-9]', '', regex=True).str[:8],
                errors='coerce', format='%Y%m%d'
            ).dt.date
        except:
            st.warning("⚠️ タイムスタンプ列の形式が不正です。日付による絞り込みをスキップしました。")
            df['date_only'] = pd.NaT # 日付フィルタを無効化
        
        df_valid_date = df.dropna(subset=['date_only'])
        
        if not df_valid_date.empty:
            min_date = df_valid_date['date_only'].min()
            max_date = df_valid_date['date_only'].max()
            
            # R2: デフォルト開始日を2025年4月1日に設定
            try:
                default_start_date = date(2025, 4, 1)
                if default_start_date < min_date:
                    default_start_date = min_date
            except ValueError:
                default_start_date = min_date

            date_range = st.date_input(
                "日付範囲で絞り込み", 
                value=(default_start_date, max_date), 
                min_value=min_date, 
                max_value=max_date
            )
            
            if len(date_range) == 2:
                start_date, end_date = date_range
                df = df[ (df['date_only'] >= start_date) & (df['date_only'] <= end_date) ]
            elif len(date_range) == 1:
                start_date = date_range[0]
                df = df[ df['date_only'] >= start_date ]

    # キーワード検索
    search_query = st.text_input("キーワード検索 (メモ/タイトルなど)", value="")
    if search_query:
        df_search = pd.DataFrame()
        cols_to_search = [c for c in df.columns if c in [col_memo, HANDOVER_COL_TITLE, QA_COL_TITLE, QA_COL_CONTENT]]
        
        for col in cols_to_search:
            # 検索対象列が文字列型であることを確認
            if pd.api.types.is_object_dtype(df[col]):
                df_search = pd.concat([df_search, df[df[col].astype(str).str.contains(search_query, case=False, na=False)]]).drop_duplicates()
        
        df = df_search.sort_values(by=col_time, ascending=False)
    else:
        df = df.sort_values(by=col_time, ascending=False)

    st.subheader(f"検索結果 ({len(df)}件)")

    # 最終的な表示 (詳細列の表示設定)
    display_cols = [col_time]
    if col_filter: display_cols.append(col_filter)
    if detail_cols: display_cols.extend(detail_cols)

    # DataFrame表示
    st.dataframe(df[display_cols].reset_index(drop=True), use_container_width=True)


def page_epi_note():
    # ... (元のコードを維持) ...
    st.header("📝 エピノート記録")
    st.markdown("成長や実験の記録を入力し、指定のGoogle SpreadSheetにアーカイブします。")
    
    # ... (元のロジックを維持) ...
    NOTE_TYPE = 'エピノート'
    FOLDER_NAME = 'epi_files'
    
    col1, col2 = st.columns(2)
    with col1:
        category = st.selectbox("カテゴリ (装置/テーマ)", ["D1", "D2", "MBE", "RTA", "ALD", "その他"], key='epi_category')
    with col2:
        file_attachments = st.file_uploader("写真/ファイル添付", type=['jpg', 'png', 'pdf'], accept_multiple_files=True, key='epi_attachments')

    memo = st.text_area("メモ (記録内容)", height=300, key='epi_memo')
    
    if st.button("記録を保存 (エピノート)", key='save_epi_note_button'):
        if not memo:
            st.error("メモ（記録内容）は必須です。")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        filenames_json = json.dumps([f.name for f in file_attachments])
        urls_list = []
        
        # ファイルアップロード処理
        with st.spinner("ファイルをCloud Storageにアップロード中..."):
            for uploaded_file in file_attachments:
                original_filename, public_url = upload_file_to_gcs(storage_client, uploaded_file, FOLDER_NAME)
                if public_url:
                    urls_list.append(public_url)

        urls_json = json.dumps(urls_list)
        
        row_data = [
            timestamp, NOTE_TYPE, category, memo, filenames_json, urls_json # JSON文字列として保存
        ]
        
        # スプレッドシートへの書き込み
        try:
            gc.open(SPREADSHEET_NAME).worksheet(SHEET_EPI_DATA).append_row(row_data)
            st.success("エピノートをアーカイブしました。")
            # キャッシュをクリアして再読み込み
            st.cache_data.clear(); st.rerun()
        except Exception as e:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_EPI_DATA}' が存在するか確認してください。")
            st.exception(e)

def page_mainte_note():
    # ... (元のコードを維持) ...
    st.header("🛠️ メンテノート記録")
    st.markdown("装置のメンテナンス記録を入力し、指定のGoogle SpreadSheetにアーカイブします。")
    
    # ... (元のロジックを維持) ...
    NOTE_TYPE = 'メンテノート'
    FOLDER_NAME = 'mainte_files'
    
    memo = st.text_area("メモ (記録内容/メンテナンス実施日と内容)", height=200, key='mainte_memo')
    file_attachments = st.file_uploader("写真/ファイル添付", type=['jpg', 'png', 'pdf'], accept_multiple_files=True, key='mainte_attachments')

    if st.button("記録を保存 (メンテノート)", key='save_mainte_note_button'):
        if not memo:
            st.error("メモ（記録内容）は必須です。")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        filenames_json = json.dumps([f.name for f in file_attachments])
        urls_list = []
        
        # ファイルアップロード処理
        with st.spinner("ファイルをCloud Storageにアップロード中..."):
            for uploaded_file in file_attachments:
                original_filename, public_url = upload_file_to_gcs(storage_client, uploaded_file, FOLDER_NAME)
                if public_url:
                    urls_list.append(public_url)

        urls_json = json.dumps(urls_list)
        
        row_data = [
            timestamp, NOTE_TYPE, memo, filenames_json, urls_json # JSON文字列として保存
        ]
        
        # スプレッドシートへの書き込み
        try:
            gc.open(SPREADSHEET_NAME).worksheet(SHEET_MAINTE_DATA).append_row(row_data)
            st.success("メンテノートをアーカイブしました。")
            st.cache_data.clear(); st.rerun()
        except Exception as e:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_MAINTE_DATA}' が存在するか確認してください。")
            st.exception(e)


def page_epi_note_list():
    page_data_list(SHEET_EPI_DATA, "エピノート", EPI_COL_TIMESTAMP, col_filter=EPI_COL_CATEGORY, detail_cols=[EPI_COL_MEMO, EPI_COL_FILENAME, EPI_COL_FILE_URL])

def page_mainte_note_list():
    page_data_list(SHEET_MAINTE_DATA, "メンテノート", MAINT_COL_TIMESTAMP, detail_cols=[MAINT_COL_MEMO, MAINT_COL_FILENAME, MAINT_COL_FILE_URL])

def page_meeting_note():
    # ... (元のコードを維持) ...
    st.header("📋 議事録管理")
    st.markdown("議事録データをアップロードし、Google SpreadSheetに記録します。")
    
    # ... (元のロジックを維持) ...
    MEETING_FOLDER_NAME = 'meeting_audio'
    
    meeting_title = st.text_input("会議タイトル/日付", key='meeting_title')
    audio_file = st.file_uploader("会議の音声ファイル (.m4a, .mp3など)", type=['m4a', 'mp3', 'wav'], key='audio_file')
    content = st.text_area("議事録内容 (または文字起こしテキスト)", height=300, key='meeting_content')
    
    if st.button("議事録をアーカイブ", key='archive_meeting_button'):
        if not meeting_title or not content:
            st.error("会議タイトルと議事録内容は必須です。")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        audio_filename = ""
        audio_url = ""
        
        if audio_file:
            with st.spinner("音声ファイルをCloud Storageにアップロード中..."):
                audio_filename, audio_url = upload_file_to_gcs(storage_client, audio_file, MEETING_FOLDER_NAME)

        row_data = [
            timestamp, meeting_title, audio_filename, audio_url, content
        ]
        
        # スプレッドシートへの書き込み
        try:
            gc.open(SPREADSHEET_NAME).worksheet(SHEET_MEETING_DATA).append_row(row_data)
            st.success("議事録をアーカイブしました。")
            st.cache_data.clear(); st.rerun()
        except Exception as e:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_MEETING_DATA}' が存在するか確認してください。")
            st.exception(e)

def page_meeting_note_list():
    page_data_list(SHEET_MEETING_DATA, "議事録", MEETING_COL_TIMESTAMP, detail_cols=[MEETING_COL_TITLE, MEETING_COL_AUDIO_NAME, MEETING_COL_AUDIO_URL, MEETING_COL_CONTENT])


# --------------------------------------------------------------------------
# --- Page Implementations: IVデータ解析 (安定動作版) ---
# --------------------------------------------------------------------------
# **[REPLACED] Stable page_iv_analysis (fixes UnboundLocalError and 10-file limit)**
def page_iv_analysis():
    st.header("⚡ IV Data Analysis (IVデータ解析)")
    st.markdown(f"IVデータファイルを選択し、グラフ描画とデータのエクスポートを行います。**ファイル数が{MAX_COMBINED_FILES}個以下の場合、結合データも作成します。**")

    uploaded_files = st.file_uploader(
        "IVデータファイル（.txt または .csv）を選択してください",
        type=['txt', 'csv'],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.subheader("📊 IV Characteristic Plot")
        
        # グラフサイズを大きく
        fig, ax = plt.subplots(figsize=(12, 7))
        
        all_data_for_export = [] # 各ファイルのDFとファイル名を格納
        
        # 1. データの読み込みとグラフ描画
        for uploaded_file in uploaded_files:
            # 新しい安定版のロード関数を使用
            df, file_name = load_iv_data(uploaded_file) 
            
            if df is not None and not df.empty:
                voltage_col = 'VF(V)'
                current_col = 'IF(A)'
                
                # グラフにプロット
                ax.plot(df[voltage_col], df[current_col], label=file_name)
                
                # エクスポート用に[Voltage_V, Current_A_filename]のDFを作成
                df_export = df.rename(columns={voltage_col: 'Voltage_V', current_col: f'Current_A_{file_name}'})
                all_data_for_export.append({'name': file_name, 'df': df_export})

        
        # グラフ設定 (文字化け対策: すべて英語)
        ax.set_title('IV Characteristic Plot', fontsize=16)
        ax.set_xlabel('Voltage (V)', fontsize=14)
        ax.set_ylabel('Current (A)', fontsize=14)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend(title='File Name', loc='best')
        ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))
        
        st.pyplot(fig, use_container_width=True)
        plt.close(fig) # メモリ解放

        # ------------------------------------------------------------------
        # 2. Excelエクスポート (条件分岐ロジック)
        # ------------------------------------------------------------------
        if all_data_for_export:
            st.subheader("📝 データのエクスポート")
            
            output = BytesIO()
            file_count = len(all_data_for_export)
            
            # 10個以下の場合は結合フラグをTrueに
            SHOULD_COMBINE = file_count <= MAX_COMBINED_FILES
            
            if SHOULD_COMBINE:
                st.info(f"✅ ファイル数が{file_count}個のため、個別シートに加えて**結合データシート**を作成します。")
            else:
                st.warning(f"⚠️ ファイル数が{file_count}個と多いため、クラッシュ防止のため**個別データシートのみ**を作成します。（結合シートはスキップされます）")
            
            with st.spinner("データをExcelに書き込んでいます..."):
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    
                    # --- (A) 各ファイルを別シートに出力 (共通処理) ---
                    for data_item in all_data_for_export:
                        file_name = data_item['name']
                        df_export = data_item['df']
                        
                        sheet_name = file_name.replace('.txt', '').replace('.csv', '')
                        # Excelのシート名制限(31文字)
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:28] 
                        
                        df_export.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 個別DFのメモリを直後に解放
                        del df_export

                    # --- (B) 結合データを出力 (10個以下の場合のみ) ---
                    if SHOULD_COMBINE:
                        
                        # 最初のデータフレームを基準に結合を開始
                        start_df = all_data_for_export[0]['df']
                        combined_df = start_df.copy() 
                        
                        # 2つ目以降のデータフレームを 'Voltage_V' をキーに結合
                        for i in range(1, len(all_data_for_export)):
                            item = all_data_for_export[i]
                            df_current = item['df']
                            # 'Voltage_V'をキーに、2つ目の列（電流データ）のみを結合
                            combined_df = pd.merge(combined_df, df_current[['Voltage_V', df_current.columns[1]]], on='Voltage_V', how='outer')
                        
                        # 電圧順にソート
                        combined_df.sort_values(by='Voltage_V', inplace=True)
                        
                        # 結合DFのプレビュー
                        st.dataframe(combined_df.head())
                        
                        # 結合DFを最終シートに出力
                        combined_df.to_excel(writer, sheet_name='__COMBINED_DATA__', index=False)
                        
                        # 処理落ち対策: 結合DFのメモリを直後に解放
                        del combined_df
                        
            
            processed_data = output.getvalue()
            
            download_label = "📈 結合/個別データを含むExcelファイルとしてダウンロード" if SHOULD_COMBINE else "📁 全データを個別シートに保存してダウンロード"
            
            st.download_button(
                label=download_label,
                data=processed_data,
                file_name=f"iv_analysis_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        else:
            st.warning("有効なデータファイルが見つかりませんでした。")

# --------------------------------------------------------------------------
# --- Page Implementations: PLデータ解析 (元のコードを維持) ---
# --------------------------------------------------------------------------

def page_pl_analysis():
    # ... (元のコードを維持) ...
    st.header("🔬 PL Data Analysis (PLデータ解析)")
    st.markdown("PLデータファイルを選択し、グラフ描画とデータのエクスポートを行います。")
    
    # Session Stateの初期化 (波長校正係数)
    if 'pl_calib_a' not in st.session_state:
        st.session_state.pl_calib_a = 0.0
    if 'pl_calib_b' not in st.session_state:
        st.session_state.pl_calib_b = 0.0

    uploaded_files = st.file_uploader(
        "PLデータファイル（.txt または .csv）を選択してください",
        type=['txt', 'csv'],
        accept_multiple_files=True,
        key='pl_files'
    )

    if uploaded_files:
        st.subheader("波長校正 (Wavelength Calibration)")
        
        col_calib_a, col_calib_b = st.columns(2)
        with col_calib_a:
            st.session_state.pl_calib_a = st.number_input(
                "校正係数 a (Wavelength = a * pixel + b)",
                value=st.session_state.pl_calib_a,
                format="%.6f",
                key='pl_input_a'
            )
        with col_calib_b:
            st.session_state.pl_calib_b = st.number_input(
                "校正係数 b (Wavelength = a * pixel + b)",
                value=st.session_state.pl_calib_b,
                format="%.6f",
                key='pl_input_b'
            )
        
        a = st.session_state.pl_calib_a
        b = st.session_state.pl_calib_b

        st.subheader("📊 PL Characteristic Plot")
        
        fig, ax = plt.subplots(figsize=(12, 7))
        all_data_for_export = []
        
        # 1. データの読み込みとグラフ描画
        for uploaded_file in uploaded_files:
            df = load_pl_data(uploaded_file)
            file_name = uploaded_file.name
            
            if df is not None and not df.empty:
                # ピクセルを波長に変換
                df['wavelength_nm'] = df['pixel'] * a + b
                
                # グラフにプロット
                ax.plot(df['wavelength_nm'], df['intensity'], label=file_name)
                
                # エクスポート用に列名を整形
                df_export = df.rename(columns={'wavelength_nm': 'Wavelength_nm', 'intensity': f'Intensity_{file_name}'})
                all_data_for_export.append({'name': file_name, 'df': df_export[['Wavelength_nm', f'Intensity_{file_name}']]})

        
        # グラフ設定 (文字化け対策: すべて英語)
        ax.set_title('PL Spectrum Plot', fontsize=16)
        ax.set_xlabel('Wavelength (nm)', fontsize=14)
        ax.set_ylabel('PL Intensity (a.u.)', fontsize=14)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend(title='File Name', loc='best')
        # 軸のスケール調整はユーザーに任せる
        
        st.pyplot(fig, use_container_width=True)
        plt.close(fig) # メモリ解放

        # ------------------------------------------------------------------
        # 2. Excelエクスポート (IV解析と同様の結合ロジックをPLに適用)
        # ------------------------------------------------------------------
        if all_data_for_export:
            st.subheader("📝 データのエクスポート")
            
            output = BytesIO()
            file_count = len(all_data_for_export)
            
            # PLデータも10個以下の場合は結合フラグをTrueに (IV解析から流用)
            SHOULD_COMBINE = file_count <= MAX_COMBINED_FILES
            
            if SHOULD_COMBINE:
                st.info(f"✅ ファイル数が{file_count}個のため、個別シートに加えて**結合データシート**を作成します。")
            else:
                st.warning(f"⚠️ ファイル数が{file_count}個と多いため、クラッシュ防止のため**個別データシートのみ**を作成します。（結合シートはスキップされます）")
            
            with st.spinner("データをExcelに書き込んでいます..."):
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    
                    # --- (A) 各ファイルを別シートに出力 (共通処理) ---
                    for data_item in all_data_for_export:
                        file_name = data_item['name']
                        df_export = data_item['df']
                        
                        sheet_name = file_name.replace('.txt', '').replace('.csv', '')
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:28] 
                        
                        df_export.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 個別DFのメモリを直後に解放
                        del df_export

                    # --- (B) 結合データを出力 (10個以下の場合のみ) ---
                    if SHOULD_COMBINE:
                        
                        # 最初のデータフレームを基準に結合を開始
                        start_df = all_data_for_export[0]['df']
                        combined_df = start_df.copy() 
                        
                        # 2つ目以降のデータフレームを 'Wavelength_nm' をキーに結合
                        for i in range(1, len(all_data_for_export)):
                            item = all_data_for_export[i]
                            df_current = item['df']
                            # 'Wavelength_nm'をキーに、2つ目の列（強度データ）のみを結合
                            combined_df = pd.merge(combined_df, df_current[['Wavelength_nm', df_current.columns[1]]], on='Wavelength_nm', how='outer')
                        
                        # 波長順にソート (昇順)
                        combined_df.sort_values(by='Wavelength_nm', inplace=True)
                        
                        # 結合DFのプレビュー
                        st.dataframe(combined_df.head())
                        
                        # 結合DFを最終シートに出力
                        combined_df.to_excel(writer, sheet_name='__COMBINED_DATA__', index=False)
                        
                        # 処理落ち対策: 結合DFのメモリを直後に解放
                        del combined_df
                        
            
            processed_data = output.getvalue()
            
            download_label = "📈 結合/個別データを含むExcelファイルとしてダウンロード" if SHOULD_COMBINE else "📁 全データを個別シートに保存してダウンロード"
            
            # 出力ファイル名に中心波長を付ける (元のコードから流用)
            center_wavelength_input = st.number_input("出力ファイル名に使用する中心波長 (nm)", value=800, key='pl_center_wavelength_input')
            
            st.download_button(
                label=download_label,
                data=processed_data,
                file_name=f"pl_analysis_export_{center_wavelength_input}nm_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        else:
            st.warning("有効なデータファイルが見つかりませんでした。")
    else:
        st.info("測定データファイルをアップロードしてください。")

# --- Dummy Pages (未実装のページ) ---
# ... (元のコードを維持) ...
def page_calendar():
    st.header("🗓️ スケジュール・装置予約")
    st.info("このページは未実装です。")
    # ... (元のコードを維持) ...

def page_qa_box():
    # ... (元のコードを維持) ...
    st.header("💡 知恵袋・質問箱")
    st.markdown("質問を投稿し、過去の質問・回答を閲覧します。")
    
    # ... (元のロジックを維持) ...
    QA_FOLDER_NAME = 'qa_files'

    # --- 質問投稿フォーム ---
    with st.expander("❓ 質問を投稿する"):
        title = st.text_input("質問タイトル", key='qa_title')
        content = st.text_area("質問内容", height=200, key='qa_content')
        contact = st.text_input("連絡先メールアドレス (任意)", key='qa_contact')
        file_attachments = st.file_uploader("添付ファイル", accept_multiple_files=False, key='qa_attachments')
        
        if st.button("質問を投稿", key='post_qa_button'):
            if not title or not content:
                st.error("質問タイトルと質問内容は必須です。")
                return
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            filename = ""
            file_url = ""
            
            if file_attachments:
                with st.spinner("ファイルをCloud Storageにアップロード中..."):
                    filename, file_url = upload_file_to_gcs(storage_client, file_attachments, QA_FOLDER_NAME)

            row_data = [
                timestamp, title, content, contact, filename, file_url, "未解決"
            ]
            
            # スプレッドシートへの書き込み
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_QA_DATA).append_row(row_data)
                st.success("質問をアーカイブしました。")
                st.cache_data.clear(); st.rerun()
            except Exception as e:
                st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_QA_DATA}' が存在するか確認してください。")
                st.exception(e)

    # --- 質問と回答の一覧表示 ---
    st.subheader("📋 過去の質問と回答")
    df_questions = get_sheet_as_df(SPREADSHEET_NAME, SHEET_QA_DATA)
    df_answers = get_sheet_as_df(SPREADSHEET_NAME, SHEET_QA_ANSWER)

    if not df_questions.empty:
        # 質問IDをキーに結合
        df_merged = pd.merge(
            df_questions, 
            df_answers[['質問タイムスタンプ (質問ID)', '解答内容', '解答者 (任意)']],
            left_on=QA_COL_TIMESTAMP,
            right_on='質問タイムスタンプ (質問ID)',
            how='left'
        )
        
        # 表示用の列を選択し、新しい列名で整理
        df_display = df_merged[[
            QA_COL_TIMESTAMP, QA_COL_TITLE, QA_COL_CONTENT, QA_COL_STATUS, QA_COL_FILE_URL,
            '解答内容', '解答者 (任意)'
        ]].rename(columns={
            QA_COL_TIMESTAMP: '質問ID',
            QA_COL_TITLE: 'タイトル',
            QA_COL_CONTENT: '質問内容',
            QA_COL_STATUS: 'ステータス',
            QA_COL_FILE_URL: '添付URL',
            '解答者 (任意)': '解答者'
        })
        
        # 絞り込み
        status_filter = st.selectbox("ステータスで絞り込み", ["すべて"] + list(df_display['ステータス'].unique()), key='qa_status_filter')
        if status_filter != "すべて":
            df_display = df_display[df_display['ステータス'] == status_filter]
            
        search_query = st.text_input("キーワード検索 (タイトル/内容)", key='qa_search_query')
        if search_query:
             df_display = df_display[
                df_display['タイトル'].astype(str).str.contains(search_query, case=False, na=False) |
                df_display['質問内容'].astype(str).str.contains(search_query, case=False, na=False)
            ]
        
        st.dataframe(df_display.sort_values(by='質問ID', ascending=False).reset_index(drop=True), use_container_width=True)
    else:
        st.info("まだ質問は投稿されていません。")

def page_handoff_notes():
    # ... (元のコードを維持) ...
    st.header("🤝 装置引き継ぎメモ")
    st.markdown("装置のマニュアルや引き継ぎ情報をアーカイブし、一覧表示します。")
    
    # ... (元のロジックを維持) ...
    HANDOVER_FOLDER_NAME = 'handoff_files'
    
    # --- 記録フォーム ---
    with st.expander("📝 引き継ぎ情報を記録する"):
        ho_type = st.selectbox("種類", ["マニュアル", "手順書", "その他メモ"], key='ho_type')
        ho_title = st.text_input("タイトル/装置名", key='ho_title')
        ho_url = st.text_input("関連ファイル/ドキュメントのURL (G Driveなど)", key='ho_url')
        ho_memo = st.text_area("メモ", height=150, key='ho_memo')
        
        if st.button("記録を保存 (引き継ぎ)", key='save_handoff_button'):
            if not ho_title or not ho_url:
                st.error("タイトルとURLは必須です。")
                return
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # 内容1,2,3の列は使わず、メモ列に統合
            row_data = [
                timestamp, ho_type, ho_title, ho_url, "", "", ho_memo 
            ]
            
            # スプレッドシートへの書き込み
            try:
                # 元のコードのデータ構造を維持するため、内容1-3も空で渡す
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_HANDOVER_DATA).append_row(row_data)
                st.success("引き継ぎ情報をアーカイブしました。")
                st.cache_data.clear(); st.rerun()
            except Exception as e:
                st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_HANDOVER_DATA}' が存在するか確認してください。")
                st.exception(e)

    # --- 一覧表示 ---
    page_data_list(SHEET_HANDOVER_DATA, "装置引き継ぎメモ", HANDOVER_COL_TIMESTAMP, col_filter=HANDOVER_COL_TYPE, detail_cols=[HANDOVER_COL_TITLE, HANDOVER_COL_MEMO])


def page_trouble_report():
    # ... (元のコードを維持) ...
    st.header("🚨 トラブル報告")
    st.markdown("装置のトラブル内容を報告・記録し、Google SpreadSheetとCloud Storageにアーカイブします。")
    
    # ... (元のロジックを維持) ...
    TROUBLE_FOLDER_NAME = 'trouble_files'
    
    # --- 報告フォーム ---
    with st.expander("📝 トラブルを報告する"):
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            device = st.selectbox("機器/場所", ["MBE", "RTA", "ALD", "D1", "D2", "その他"], key='trouble_device')
        with col_t2:
            report_date = st.date_input("発生日", key='trouble_date')
            
        report_title = st.text_input("件名/タイトル", key='trouble_title')
        t_occur = st.text_area("トラブル発生時の状況 (発生時間含む)", height=150, key='trouble_occur')
        t_cause = st.text_area("原因/究明", height=150, key='trouble_cause')
        t_solution = st.text_area("対策/復旧内容", height=150, key='trouble_solution')
        t_prevention = st.text_area("再発防止策", height=150, key='trouble_prevention')
        reporter_name = st.text_input("報告者名", key='trouble_reporter')
        file_attachments = st.file_uploader("関連写真/ファイル添付", accept_multiple_files=True, key='trouble_attachments')
        
        if st.button("報告をアーカイブ", key='archive_trouble_button'):
            if not report_title or not t_occur or not reporter_name:
                st.error("件名/タイトル、発生時の状況、報告者名は必須です。")
                return
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            filenames_json = json.dumps([f.name for f in file_attachments])
            urls_list = []
            
            # ファイルアップロード処理
            with st.spinner("ファイルをCloud Storageにアップロード中..."):
                for uploaded_file in file_attachments:
                    original_filename, public_url = upload_file_to_gcs(storage_client, uploaded_file, TROUBLE_FOLDER_NAME)
                    if public_url:
                        urls_list.append(public_url)

            urls_json = json.dumps(urls_list)
            
            # タイムスタンプ, 機器/場所, 発生日, トラブル発生時, 原因/究明, 対策/復旧, 再発防止策, 報告者, ファイル名, ファイルURL, 件名/タイトル
            row_data = [
                timestamp, device, report_date.isoformat(), t_occur,
                t_cause, t_solution, t_prevention,
                reporter_name, filenames_json, urls_json, report_title # JSON文字列として保存
            ]
            
            # スプレッドシートへの書き込み
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_TROUBLE_DATA).append_row(row_data)
                st.success("トラブル報告をアーカイブしました。")
                st.cache_data.clear(); st.rerun()
            except Exception as e:
                st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_TROUBLE_DATA}' が存在するか確認してください。")
                st.exception(e)
                
    # --- 一覧表示 ---
    page_data_list(SHEET_TROUBLE_DATA, "トラブル報告", TROUBLE_COL_TIMESTAMP, col_filter=TROUBLE_COL_DEVICE, detail_cols=[TROUBLE_COL_TITLE, TROUBLE_COL_OCCUR_DATE, TROUBLE_COL_CAUSE, TROUBLE_COL_SOLUTION, TROUBLE_COL_REPORTER, TROUBLE_COL_FILE_URL])


def page_contact():
    # ... (元のコードを維持) ...
    st.header("✉️ 連絡・問い合わせ")
    st.markdown("アプリ管理者への連絡やバグ報告を行います。")
    
    # ... (元のロジックを維持) ...
    
    contact = st.text_input("連絡先メールアドレス", key='contact_email')
    contact_type = st.selectbox("お問い合わせの種類", ["バグ報告", "機能要望", "その他"], key='contact_type')
    detail = st.text_area("詳細内容", height=150, key='contact_detail')
    
    if st.button("送信", key='send_contact_button'):
        if not contact_type or not detail:
            st.error("お問い合わせの種類と詳細内容は必須です。")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        row_data = [
            timestamp, contact_type, detail, contact
        ]
        
        # スプレッドシートへの書き込み
        try:
            gc.open(SPREADSHEET_NAME).worksheet(SHEET_CONTACT_DATA).append_row(row_data)
            st.success("お問い合わせを送信しました。")
            st.cache_data.clear(); st.rerun()
        except Exception as e:
            st.error(f"データの書き込み中にエラーが発生しました。シート名 '{SHEET_CONTACT_DATA}' が存在するか確認してください。")
            st.exception(e)


# --------------------------------------------------------------------------
# --- Main App Execution ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    # メニューを記録・一覧で統合
    menu_selection = st.sidebar.radio("機能選択", [
        "エピノート", "メンテノート", "議事録", "知恵袋・質問箱", "装置引き継ぎメモ", "トラブル報告", "連絡・問い合わせ",
        "⚡ IVデータ解析", "🔬 PLデータ解析", "🗓️ スケジュール・装置予約"
    ])
    
    # ページルーティング
    if menu_selection == "エピノート": page_epi_note()
    elif menu_selection == "メンテノート": page_mainte_note()
    elif menu_selection == "議事録": page_meeting_note()
    elif menu_selection == "知恵袋・質問箱": page_qa_box()
    elif menu_selection == "装置引き継ぎメモ": page_handoff_notes()
    elif menu_selection == "トラブル報告": page_trouble_report()
    elif menu_selection == "連絡・問い合わせ": page_contact()
    elif menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()

if __name__ == "__main__":
    main()

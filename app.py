# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v20.6.1 (PL波長校正対応版)
# - NEW: load_pl_data(uploaded_file) 関数を追加し、データカラム名を 'pixel', 'intensity' に固定。
# - CHG: page_pl_analysis() をユーザー提供の波長校正ロジックに置き換え、校正係数をセッションステートで保持。
# - FIX: 全てのリストでデフォルト開始日を2025/4/1に設定 (v20.6.0から変更なし)。
# --------------------------------------------------------------------------

import streamlit as st
import gspread
import pandas as pd
import os # NEW: Added for os.path.splitext in page_pl_analysis
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
    
# app.py (修正箇所)
# ...
import calendar
import matplotlib.font_manager as fm # <--- fmのインポートは既にあり

# app.py (import文の直後あたりに追記)

# ... (前略)

# --- Matplotlib 日本語フォント設定（packages.txt利用時） ---
try:
    # Noto Sans CJK JPがインストールされていることを期待
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Noto Sans CJK JP', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False
    st.info("✅ Matplotlib: 'Noto Sans CJK JP' を設定しました。")

except Exception as e:
    st.error(f"❌ フォント設定中に予期せぬエラーが発生しました: {e}")
    
# --- Global Configuration & Setup ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
# ↓↓↓↓↓↓ 【重要】ご自身の「バケット名」に書き換えてください ↓↓↓↓↓↓
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" 
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

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

# --- IV/PLデータ解析用コアユーティリティ ---
# (既存のload_data_fileのロジックを流用)
def _load_two_column_data_core(uploaded_file_bytes, column_names):
    """IV/PLデータファイルから2列のデータを読み込み、指定されたカラム名を付けてDataFrameを返す"""
    try:
        # ロバストな読み込みロジック 
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
        # st.error(f"データファイルの読み込み中にエラーが発生しました: {e}") # エラー表示は上位関数で行う
        return None

# --- IVデータ解析用 (既存関数をコアユーティリティで置き換え) ---
@st.cache_data(show_spinner="IVデータを解析中...", max_entries=50)
def load_data_file(uploaded_file_bytes, uploaded_file_name):
    """IVファイル (Axis_X vs Filename) を読み込み、DataFrameを返す (IV/PL共通ロジック)"""
    return _load_two_column_data_core(uploaded_file_bytes, ['Axis_X', uploaded_file_name])

# --- PLデータ解析用 (新規追加 R4) ---
@st.cache_data(show_spinner="PLデータを解析中...", max_entries=50)
def load_pl_data(uploaded_file):
    """PLファイル (pixel vs intensity) を読み込み、DataFrame (pixel, intensity) を返す"""
    df = _load_two_column_data_core(uploaded_file.getvalue(), ['pixel', 'intensity'])
    # load_pl_dataは、uploaded_fileオブジェクトを直接受け取るため、getvalue()を使用
    if df is not None and not df.empty:
        return df[['pixel', 'intensity']]
    return None

@st.cache_data(show_spinner="データを結合中...")
def combine_dataframes(dataframes, filenames):
    """複数のDataFrameを共通のX軸をキーに外部結合する"""
    if not dataframes: return None
    
    # 結合キーは 'X_Value' (load_data_fileの出力に合わせる)
    combined_df = dataframes[0].rename(columns={'Axis_X': 'X_Value'})
    
    for i in range(1, len(dataframes)):
        df_to_merge = dataframes[i].rename(columns={'Axis_X': 'X_Value'})
        combined_df = pd.merge(combined_df, df_to_merge, on='X_Value', how='outer')
        
    combined_df = combined_df.sort_values(by='X_Value', ascending=False).reset_index(drop=True)
    
    for col in combined_df.columns:
        if col != 'X_Value':
            combined_df[col] = combined_df[col].round(4)
            
    # X軸の列名を結合前に戻す
    combined_df = combined_df.rename(columns={'X_Value': 'X_Axis'})
    
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

# --- 汎用的な一覧表示関数 ---
def page_data_list(sheet_name, title, col_time, col_filter=None, col_memo=None, col_url=None, detail_cols=None):
    """汎用的なデータ一覧ページ (R2, R3, R1対応)"""
    
    st.header(f"📚 {title}一覧")
    
    df = get_sheet_as_df(SPREADSHEET_NAME, sheet_name) 

    if df.empty: st.info("データがありません。"); return
        
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
                errors='coerce', 
                format='%Y%m%d'
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
            except ValueError:
                default_start_date = date.today() - timedelta(days=365) # 安全策
                
            # 実際の日付の最小値と、指定されたデフォルト開始日のうち、新しい方を選択
            initial_start_date = max(min_date, default_start_date) if isinstance(min_date, date) else default_start_date
            
            col_date1, col_date2 = st.columns(2)
            with col_date1:
                start_date = st.date_input("開始日", value=initial_start_date)
            with col_date2:
                end_date = st.date_input("終了日", value=max_date)
            
            df = df_valid_date[(df_valid_date['date_only'] >= start_date) & (df_valid_date['date_only'] <= end_date)].drop(columns=['date_only'])
        else:
            if 'date_only' in df.columns:
                 df = df.drop(columns=['date_only'])

    if df.empty: st.info("絞り込み条件に一致するデータがありません。"); return

    df = df.sort_values(by=col_time, ascending=False).reset_index(drop=True)
    
    st.markdown("---")
    st.subheader(f"検索結果 ({len(df)}件)")

    def format_func(idx):
        row = df.loc[idx]
        time_str = str(row[col_time])
        filter_str = row[col_filter] if col_filter and pd.notna(row[col_filter]) else ""
        memo_str = row[col_memo] if col_memo and pd.notna(row[col_memo]) else "メモなし"
        # メモは最初の1行または50文字で表示
        display_memo = memo_str.split('\n')[0] if '\n' in memo_str else memo_str
        return f"[{time_str.split('_')[0]}] {filter_str} - {display_memo[:50].replace('\n', ' ')}..."

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
                if col in row and pd.notna(row[col]):
                    if col_memo == col or '内容' in col: # メモや内容が多い場合はテキストエリアで表示
                        st.markdown(f"**{col}:**"); st.text(row[col])
                    elif 'URL' in col: # URLは添付ファイルセクションで処理
                         continue
                    else:
                        st.write(f"**{col}:** {row[col]}")
        
        # 添付ファイル (R3: インライン画像表示対応)
        if col_url and col_url in row:
            st.markdown("##### 添付ファイル")
            
            try:
                # JSONデコードを試みる (最新の保存形式)
                urls = json.loads(row[col_url])
                filenames = json.loads(row[EPI_COL_FILENAME]) if EPI_COL_FILENAME in row and row[EPI_COL_FILENAME] else ['ファイル'] * len(urls)
                
                if urls:
                    for filename, url in zip(filenames, urls):
                        if url:
                            is_image = url.lower().endswith(('.png', '.jpg', '.jpeg'))
                            
                            if is_image and ("storage.googleapis.com" in url or "drive.google.com" in url):
                                # 画像の場合はインライン表示 (R3)
                                st.markdown(f"**画像ファイル:** {filename}")
                                st.image(url, caption=filename, use_column_width=True)
                            elif "drive.google.com" in url:
                                # Google Driveのリンク
                                st.markdown(f"🔗 **Google Drive:** [{filename}](<{url}>)")
                            else:
                                # その他のURL（GCSなど）
                                st.markdown(f"🔗 [添付ファイル]({url}) ({filename})")
                        
                else:
                    st.info("添付ファイルはありません。")

            except Exception:
                # JSON形式ではない場合（古いデータや手動入力、単一URLの直接保存）
                if pd.notna(row[col_url]) and row[col_url]:
                    url_list = row[col_url].split(',')
                    for url in url_list:
                         url = url.strip().strip('"')
                         if url:
                            is_image = url.lower().endswith(('.png', '.jpg', '.jpeg'))
                            if is_image and ("storage.googleapis.com" in url or "drive.google.com" in url):
                                st.image(url, caption="添付画像", use_column_width=True) # R3
                            else:
                                st.markdown(f"🔗 [添付ファイルURL]({url})")
                else:
                    st.info("添付ファイルはありません。")


# 1. エピノート機能
def page_epi_note_recording():
    st.markdown("#### 📝 新しいエピノートを記録")
    
    with st.form(key='epi_note_form'):
        col1, col2 = st.columns(2)
        with col1:
            # R6: カテゴリをD1/D2の選択式に
            ep_category = st.selectbox(f"{EPI_COL_CATEGORY} (装置種別)", ["D1", "D2", "その他"], key='ep_category_input')
        with col2:
            # R8: タイトル/要約を「番号(例：791)」に変更 (必須チェック)
            ep_title = st.text_input("番号 (例: 791) (必須)", key='ep_title_input')
        
        # R9: 詳細メモを「構造（空白でも可）」に変更
        ep_memo = st.text_area("構造 (例: 10nm GaAs/AlGaAs/GaAs) (空白でも可)", height=100, key='ep_memo_input')
        
        uploaded_files = st.file_uploader("添付ファイル (画像、グラフなど)", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], accept_multiple_files=True)
        
        st.markdown("---")
        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')

    if submit_button:
        if not ep_title:
            st.warning("番号 (例: 791) は必須項目です。")
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
        
        # R8, R9: メモ欄には「番号\n構造」の形式で保存
        memo_content = f"{ep_title}\n{ep_memo}"
        row_data = [
            timestamp, EPI_COL_NOTE_TYPE, ep_category, 
            memo_content, filenames_json, urls_json
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_EPI_DATA)
            worksheet.append_row(row_data)
            st.success("✅ エピノートをアップロードしました！"); st.cache_data.clear(); st.rerun() # R7: 成功メッセージ
        except Exception:
            st.error(f"❌ データの書き込み中にエラーが発生しました。シート名 '{SHEET_EPI_DATA}' が存在するか確認してください。")

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
    
    with st.form(key='mainte_note_form'):
        
        # R10: 選択式から記入式へ変更 & ラベル変更
        mainte_title = st.text_input("メンテタイトル (例: D1 ドライポンプ交換) (必須)", key='mainte_title_input')
        memo_content = st.text_area("詳細メモ", height=150, key='mainte_memo_input')
        uploaded_files = st.file_uploader("添付ファイル (画像、グラフなど)", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], accept_multiple_files=True)
        
        st.markdown("---")
        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')

    if submit_button:
        if not mainte_title:
            st.warning("メンテタイトルを入力してください。")
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

        # R10: 結合方法を変更 (タイトルとメモを結合)
        memo_to_save = f"[{mainte_title}]\n{memo_content}"
        row_data = [
            timestamp, MAINT_COL_NOTE_TYPE, memo_to_save, 
            filenames_json, urls_json
        ]
        
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_MAINTE_DATA)
            worksheet.append_row(row_data)
            st.success("✅ メンテノートをアップロードしました！"); st.cache_data.clear(); st.rerun() # R7: 成功メッセージ
        except Exception:
            st.error(f"❌ データの書き込み中にエラーが発生しました。シート名 '{SHEET_MAINTE_DATA}' が存在するか確認してください。")

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
            st.success("✅ 議事録をアップロードしました！"); st.cache_data.clear(); st.rerun() # R7: 成功メッセージ
        except Exception:
            st.error(f"❌ データの書き込み中にエラーが発生しました。シート名 '{SHEET_MEETING_DATA}' が存在するか確認してください。")

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
            st.success("✅ 質問をアップロードしました！回答があるまでお待ちください。"); st.cache_data.clear(); st.rerun() # R7: 成功メッセージ
        except Exception:
            st.error(f"❌ データの書き込み中にエラーが発生しました。シート名 '{SHEET_QA_DATA}' が存在するか確認してください。")

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
            st.success("✅ 引き継ぎメモをアップロードしました！"); st.cache_data.clear(); st.rerun() # R7: 成功メッセージ
        except Exception:
            st.error(f"❌ データの書き込み中にエラーが発生しました。シート名 '{SHEET_HANDOVER_DATA}' が存在するか確認してください。")

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
    
    # R5: 機器/場所の選択肢
    DEVICE_OPTIONS = ["MBE", "XRD", "PL", "IV", "TEM・SEM", "抵抗加熱蒸着", "RTA", "フォトリソ", "ドラフター", "その他"]

    with st.form(key='trouble_form'):
        
        st.subheader("基本情報")
        col1, col2 = st.columns(2)
        with col1:
            report_date = st.date_input(f"{TROUBLE_COL_OCCUR_DATE} (発生日)", datetime.now().date())
        with col2:
            # R5: 選択肢に変更
            device_to_save = st.selectbox(f"{TROUBLE_COL_DEVICE} (機器/場所)", DEVICE_OPTIONS, key='device_input')
            
        report_title = st.text_input(f"{TROUBLE_COL_TITLE} (件名/タイトル) (必須)", key='trouble_title_input')
        occur_time = st.text_area(f"{TROUBLE_COL_OCCUR_TIME} (状況詳細)", height=100)
        
        st.subheader("対応と考察")
        cause = st.text_area(f"{TROUBLE_COL_CAUSE} (原因/究明)", height=100)
        solution = st.text_area(f"{TROUBLE_COL_SOLUTION} (対策/復旧)", height=100)
        prevention = st.text_area(f"{TROUBLE_COL_PREVENTION} (再発防止策)", height=100)

        col3, col4 = st.columns(2)
        with col3:
            reporter_name = st.text_input(f"{TROUBLE_COL_REPORTER} (報告者) (必須)", key='reporter_input')
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
            st.success("✅ トラブル報告をアップロードしました！"); st.cache_data.clear(); st.rerun() # R7: 成功メッセージ
        except Exception:
            st.error(f"❌ データの書き込み中にエラーが発生しました。シート名 '{SHEET_TROUBLE_DATA}' が存在するか確認してください。")

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
            st.success("✅ お問い合わせを送信しました。"); st.cache_data.clear(); st.rerun() # R7: 成功メッセージ
        except Exception:
            st.error(f"❌ データの書き込み中にエラーが発生しました。シート名 '{SHEET_CONTACT_DATA}' が存在するか確認してください。")

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
            
            ax.set_xlabel("Voltage (V)") # 日本語ラベル
            ax.set_ylabel("Current (A)") # 日本語ラベル
            ax.grid(True)
            ax.legend(title="ファイル名", loc='best')
            ax.set_title("IV Characteristic Plot") # 日本語タイトル
            
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


# 9. PLデータ解析 (ユーザー提供の高度な波長校正ロジックに置き換え R4)
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
                # load_pl_data (pixel, intensity) を使用
                df1 = load_pl_data(cal1_file)
                df2 = load_pl_data(cal2_file)
                
                if df1 is not None and df2 is not None:
                    # ピーク位置の計算
                    peak_pixel1 = df1['pixel'].iloc[df1['intensity'].idxmax()]
                    peak_pixel2 = df2['pixel'].iloc[df2['intensity'].idxmax()]
                    
                    st.write("---"); st.subheader("校正結果")
                    col_res1, col_res2, col_res3 = st.columns(3)
                    col_res1.metric(f"{cal1_wavelength}nmのピーク位置", f"{int(peak_pixel1)} pixel")
                    col_res2.metric(f"{cal2_wavelength}nmのピーク位置", f"{int(peak_pixel2)} pixel")
                    
                    try:
                        delta_wave = float(cal2_wavelength - cal1_wavelength)
                        # ピクセル値の差分計算 (ユーザーロジックの方向を維持)
                        delta_pixel = float(peak_pixel1 - peak_pixel2) 
                        
                        if delta_pixel == 0:
                            st.error("2つのピーク位置が同じです。異なる校正ファイルを選択するか、データを確認してください。")
                        else:
                            slope = delta_wave / delta_pixel
                            col_res3.metric("校正係数 (nm/pixel)", f"{slope:.4f}")
                            st.session_state['pl_calibrated'] = True
                            st.session_state['pl_slope'] = slope
                            st.session_state['pl_center_wl_cal'] = cal1_wavelength
                            st.session_state['pl_center_pixel_cal'] = peak_pixel1
                            st.success("✅ 校正係数を保存しました。ステップ2に進んでください。")
                    except Exception as e:
                        st.error(f"校正パラメータの計算中にエラーが発生しました: {e}")
                else:
                    st.error("校正ファイルのデータ読み込みに失敗しました。ファイル形式を確認してください。")
            else:
                st.warning("両方の校正ファイルをアップロードしてください。")

    st.write("---")
    st.subheader("ステップ2：測定データ解析")
    if 'pl_calibrated' not in st.session_state or not st.session_state['pl_calibrated']:
        st.info("💡 まず、ステップ1の波長校正を完了させてください。")
    else:
        st.success(f"✅ 波長校正済みです。（校正係数: **{st.session_state['pl_slope']:.4f} nm/pixel**）")
        
        with st.container(border=True):
            center_wavelength_input = st.number_input(
                "測定時の中心波長 (nm)", min_value=0, value=1700, step=10,
                help="この測定で装置に設定した中心波長を入力してください。"
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
                        # センターピクセルは 256.5 を使用 (ユーザーロジックを維持)
                        center_pixel = 256.5 
                        
                        # 波長変換の実行: Wavelength = (Pixel - Center_Pixel) * Slope + Center_Wavelength
                        df['wavelength_nm'] = (df['pixel'] - center_pixel) * slope + center_wavelength_input
                        
                        base_name = os.path.splitext(uploaded_file.name)[0]
                        # 凡例の自動整形 (中心波長部分を削除)
                        cleaned_label = base_name.replace(str(int(center_wavelength_input)), "").strip(' _-')
                        label = cleaned_label if cleaned_label else base_name
                        
                        ax.plot(df['wavelength_nm'], df['intensity'], label=label, linewidth=2.5)
                        
                        export_df = df[['wavelength_nm', 'intensity']].copy()
                        export_df.rename(columns={'intensity': base_name}, inplace=True)
                        all_dataframes.append(export_df)

                if all_dataframes:
                    # 波長をキーにデータフレームを結合
                    final_df = all_dataframes[0].rename(columns={'wavelength_nm': 'wavelength_nm'})
                    for i in range(1, len(all_dataframes)):
                        final_df = pd.merge(final_df, all_dataframes[i], on='wavelength_nm', how='outer')
                        
                    final_df = final_df.sort_values(by='wavelength_nm').reset_index(drop=True)

                    # グラフ設定
                    ax.set_title(f"PL spectrum (Center: {center_wavelength_input} nm)")
                    ax.set_xlabel("Wavelength [nm]"); ax.set_ylabel("PL intensity [a.u.]")
                    ax.legend(loc='upper left', frameon=False, fontsize=10)
                    ax.grid(axis='y', linestyle='-', color='lightgray', zorder=0)
                    ax.tick_params(direction='in', top=True, right=True, which='both')
                    
                    min_wl = final_df['wavelength_nm'].min()
                    max_wl = final_df['wavelength_nm'].max()
                    padding = (max_wl - min_wl) * 0.05
                    ax.set_xlim(min_wl - padding, max_wl + padding)
                    
                    st.pyplot(fig)
                    
                    # ダウンロード機能
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: # openpyxlをxlsxwriterに変更
                        final_df.to_excel(writer, index=False, sheet_name='Combined PL Data')

                    processed_data = output.getvalue()
                    st.download_button(label="📈 Excelデータとしてダウンロード", data=processed_data, file_name=f"pl_analysis_combined_{center_wavelength_input}nm.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.warning("有効なデータファイルが見つかりませんでした。")
            else:
                 st.info("測定データファイルをアップロードしてください。")

# --- Dummy Pages (未実装のページ) ---
def page_calendar(): st.header("🗓️ スケジュール・装置予約"); st.info("このページは未実装です。")

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
    elif menu_selection == "装置引き継ぎメモ": page_handover_note()
    elif menu_selection == "トラブル報告": page_trouble_report()
    elif menu_selection == "連絡・問い合わせ": page_contact_form()
    elif menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()


if __name__ == "__main__":
    main()




# -*- coding: utf-8 -*-
"""
bennriyasann3_fixed_v2_part1.py
Yamane Lab Convenience Tool - 修正版パート1（共通ユーティリティ・認証・データ読み込み等）
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

# --- Streamlit ページ設定 ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# ---------------------------
# --- Global constants ------
# ---------------------------
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files"  # 必要に応じて置き換えてください
SPREADSHEET_NAME = "エピノート"

# --- シート名 & カラム名（既存のスプレッドシート構成に合わせています） ---
SHEET_EPI_DATA = 'エピノート_データ'
EPI_COL_TIMESTAMP = 'タイムスタンプ'
EPI_COL_NOTE_TYPE = 'ノート種別'
EPI_COL_CATEGORY = 'カテゴリ'
EPI_COL_MEMO = 'メモ'
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
HANDOVER_COL_MEMO = 'メモ'

SHEET_QA_DATA = '知恵袋_データ'
QA_COL_TIMESTAMP = 'タイムスタンプ'
QA_COL_TITLE = '質問タイトル'
QA_COL_CONTENT = '質問内容'
QA_COL_CONTACT = '連絡先メールアドレス'
QA_COL_FILENAME = '添付ファイル名'
QA_COL_FILE_URL = '添付ファイルURL'
QA_COL_STATUS = 'ステータス'
SHEET_QA_ANSWER = '知恵袋_解答'

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

# --- 研究室スケジュールデータ（新しいシートが必要） ---
SHEET_SCHEDULE_DATA = "Schedule"
SCH_COL_TIMESTAMP = "登録日時"
SCH_COL_TITLE = "タイトル"
SCH_COL_DETAIL = "詳細"
SCH_COL_START_DATETIME = "開始日時"
SCH_COL_END_DATETIME = "終了日時"
SCH_COL_USER = "登録者"

try:
    CALENDAR_ID
except NameError:
    # 暫定的な定義（エラー防止）
    CALENDAR_ID = "yamane.lab.6747@gmail.com"

# --- 予約/作業のカテゴリ（タイトル生成用） ---
CATEGORY_OPTIONS = [
    "D1エピ", "D2エピ", "MBEメンテ", "XRD", "PL", "AFM", "フォトリソ", "アニール", "蒸着", "その他入力"
]

# --- Google Calendar API連携用定数 ---
# 鍵ファイルは st.secrets から読み込むため、ファイル名は不要です
SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = "yamane.lab.6747@gmail.com" # ターゲットカレンダーID

# ---------------------------
# --- Google Service Stubs ---
# ---------------------------
class DummyGSClient:
    """認証失敗時用ダミー gspread クライアント"""
    def open(self, name): return self
    def worksheet(self, name): return self
    def get_all_records(self): return []
    def get_all_values(self): return []
    def append_row(self, values): pass

class DummyStorageClient:
    """認証失敗時用ダミー GCS クライアント"""
    def bucket(self, name): return self
    def blob(self, name): return self
    def upload_from_file(self, file_obj, content_type): pass
    def list_blobs(self, **kwargs): return []

# グローバル初期値（認証されていない状態でもUIは起動する）
gc = DummyGSClient()
storage_client = DummyStorageClient()

# ---------------------------
# --- Google 認証初期化 ---
# ---------------------------
@st.cache_resource(ttl=3600)
def initialize_google_services():
    global storage
    if storage is None:
        # google.cloud.storage が import できない環境
        st.sidebar.warning("⚠️ `google-cloud-storage` が利用できません。ファイルアップロード機能は制限されます。")
        return DummyGSClient(), DummyStorageClient()

    if "gcs_credentials" not in st.secrets:
        st.sidebar.info("Streamlit secrets に `gcs_credentials` を設定してください（オフラインでも一部機能は動きます）。")
        return DummyGSClient(), DummyStorageClient()

    try:
        raw = st.secrets["gcs_credentials"]
        # クレンジング
        cleaned = raw.strip().replace('\t', '').replace('\r', '').replace('\n', '')
        info = json.loads(cleaned)
        gc_real = gspread.service_account_from_dict(info)
        storage_real = storage.Client.from_service_account_info(info)
        st.sidebar.success("✅ Googleサービス認証 成功")
        return gc_real, storage_real
    except json.JSONDecodeError as e:
        st.sidebar.error(f"認証情報のJSONが不正です: {e}")
        return DummyGSClient(), DummyStorageClient()
    except Exception as e:
        st.sidebar.error(f"Googleサービスの初期化に失敗しました: {e}")
        return DummyGSClient(), DummyStorageClient()

# 実際に初期化してグローバルを書き換え
gc, storage_client = initialize_google_services()

# ---------------------------
# --- Spreadsheet 関連 ---
# ---------------------------
@st.cache_data(ttl=600, show_spinner="スプレッドシートからデータを読み込み中...")
def get_sheet_as_df(spreadsheet_name, sheet_name):
    global gc
    try:
        if isinstance(gc, DummyGSClient):
            # 認証されていない場合は空DFを返す（UIテスト用）
            return pd.DataFrame()
        ws = gc.open(spreadsheet_name).worksheet(sheet_name)
        data = ws.get_all_values()
        if not data or len(data) <= 1:
            return pd.DataFrame(columns=data[0] if data else [])
        df = pd.DataFrame(data[1:], columns=data[0])
        return df
    except Exception:
        return pd.DataFrame()

# ---------------------------
# --- データ読み込みコア ---
# ---------------------------
def _load_two_column_data_core(uploaded_bytes, column_names):
    try:
        text = uploaded_bytes.decode('utf-8', errors='ignore').splitlines()
        # コメント/空行を除く
        data_lines = []
        is_first_data_line = True # 【追加】ヘッダー行をスキップするためのフラグ
        for line in text:
            s = line.strip()
            if not s:
                continue
            if s.startswith(('#', '!', '/')):  # コメント行
                continue
            
            # 【修正】最初に見つかった非コメント/非空行（ヘッダー行）をスキップ
            if is_first_data_line:
                is_first_data_line = False
                continue
                
            data_lines.append(line)
            
        if not data_lines:
            return None
        # pandas に渡す
        df = pd.read_csv(io.StringIO("\n".join(data_lines)),
                         sep=r'\s+|,|\t', engine='python', header=None)
        if df.shape[1] < 2:
            return None
        df = df.iloc[:, :2]
        df.columns = column_names
        # 数値変換
        df[column_names[0]] = pd.to_numeric(df[column_names[0]], errors='coerce')
        df[column_names[1]] = pd.to_numeric(df[column_names[1]], errors='coerce')
        
        # 【修正】往路/復路解析のためにソート（.sort_values）を削除し、dropnaのみを行う
        df = df.dropna().reset_index(drop=True)
        
        if df.empty:
            return None
        return df
    except Exception:
        # 読み込み失敗時は None を返す（元の挙動を維持）
        return None
# ---------------------------
# --- IV / PL 専用読み込み ---
# ---------------------------
@st.cache_data(show_spinner="IVデータを解析中...", max_entries=128)
def load_data_file(uploaded_bytes, uploaded_filename):
    # 修正済み: 全角括弧を半角括弧に修正
    """IVファイルを読み込み Axis_X と filename 列を返す (uploaded_bytes: bytes)"""
    return _load_two_column_data_core(uploaded_bytes, ['Axis_X', uploaded_filename])

@st.cache_data(show_spinner="PLデータを解析中...", max_entries=128)
def load_pl_data(uploaded_file):
    """
    PLデータ読み込み関数（最終安定版）。
    コメント行(#,!,/)をスキップし、カンマ・スペース・タブ区切りすべてに対応。
    例: '1, 303' / '1 303' / '1\t303'
    """
    try:
        # 読み込み
        content = uploaded_file.getvalue().decode('utf-8', errors='ignore').splitlines()

        # コメント行・空行スキップ
        data_lines = []
        for line in content:
            s = line.strip()
            if not s or s.startswith(('#', '!', '/')):
                continue
            data_lines.append(s)

        if not data_lines:
            st.warning(f"'{uploaded_file.name}' に有効なデータ行が見つかりません。")
            return None

        # --- データを統一形式に整形 ---
        # 「, 」や「 ,」などを統一してカンマまたは空白に変換
        normalized = []
        for line in data_lines:
            # カンマ→スペースに統一
            line = line.replace(',', ' ')
            # タブをスペースに変換
            line = line.replace('\t', ' ')
            # 余分なスペースを1つに
            line = re.sub(r'\s+', ' ', line.strip())
            normalized.append(line)

        df = pd.read_csv(io.StringIO("\n".join(normalized)),
                         sep=' ', header=None, names=['pixel', 'intensity'])

        # 数値変換
        df['pixel'] = pd.to_numeric(df['pixel'], errors='coerce')
        df['intensity'] = pd.to_numeric(df['intensity'], errors='coerce')
        df.dropna(inplace=True)

        if df.empty:
            st.warning(f"'{uploaded_file.name}' に有効な数値データが見つかりません。")
            return None

        return df

    except Exception as e:
        st.error(f"エラー：'{uploaded_file.name}' の読み込みに失敗しました。({e})")
        return None


# ---------------------------
# --- IV データ結合（補間） ---
# ---------------------------
@st.cache_data(show_spinner="IVデータを結合中...", max_entries=64)
def combine_dataframes(dataframes, filenames, num_points=500):
    """
    複数のIVデータを共通電圧軸で線形補間して結合（欠損を作らない）。
    - dataframes: list of DataFrame (each has 'Axis_X' and a second column)
    - filenames: list of str (列名に使用)
    """
    if not dataframes:
        return None

    # 各DFの Axis_X を集める
    try:
        all_x = np.concatenate([df['Axis_X'].values for df in dataframes if 'Axis_X' in df.columns])
    except Exception:
        return None

    if all_x.size == 0:
        return None

    x_common = np.linspace(all_x.min(), all_x.max(), num_points)
    combined_df = pd.DataFrame({'X_Axis': x_common})

    for df, name in zip(dataframes, filenames):
        # df は Axis_X, <value> の2列構成を仮定
        df_sorted = df.sort_values('Axis_X')
        y_vals = df_sorted.iloc[:, 1].values
        x_vals = df_sorted['Axis_X'].values
        # 線形補間（境界外は最外端の値を使用）
        y_interp = np.interp(x_common, x_vals, y_vals, left=y_vals[0], right=y_vals[-1])
        combined_df[name] = y_interp

    return combined_df

# ---------------------------
# --- 署名付きURL生成ユーティリティ ---
# ---------------------------
def generate_signed_url(gcs_path, expiration_minutes=60):
    """
    非公開GCSオブジェクトに対して、有効期限付きの署名付きURLを生成する。
    """
    if isinstance(storage_client, DummyStorageClient) or storage is None:
        return None

    try:
        bucket_name = CLOUD_STORAGE_BUCKET_NAME
        blob_name = gcs_path
        
        # 署名付きURLを生成
        bucket = storage_client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        
        # サービスアカウントを使用して署名
        signed_url = blob.generate_signed_url(
            version="v4",
            # 有効期限: 60分
            expiration=timedelta(minutes=expiration_minutes), 
            method="GET"
        )
        return signed_url
    except Exception as e:
        st.error(f"署名付きURLの生成中にエラーが発生しました: {e}")
        return None
        
# ---------------------------
# --- GCS へのファイルアップロード ---
# ---------------------------
def upload_file_to_gcs(uploaded_file, folder_name=""):
    """StreamlitのアップロードファイルをGCSのルートに保存し、公開URLを返す。"""
    if isinstance(storage_client, DummyStorageClient) or storage is None:
        st.warning("GCSクライアントが認証されていません。ファイル保存スキップ。")
        return None, None

    try:
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_filename = uploaded_file.name.replace("/", "_").replace("\\", "_")

        # 【修正】フォルダ分けをせず、常にルート ('') に保存する
        # gcs_filenameにはフォルダ名を含めない
        gcs_filename = f"{current_time}_{safe_filename}"
        
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)

        # ファイルをメモリからアップロード
        blob.upload_from_file(uploaded_file, rewind=True)
        
        # make_public() は不要。非公開のまま、既存の認証済みロジックでアクセスする
        
        # 公開URLを生成
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        
        return gcs_filename, public_url

    except Exception as e:
        st.error(f"GCSへのアップロード中にエラーが発生しました: {e}")
        return None, None

# ---------------------------
# --- GCS ファイルリスト取得ユーティリティ ---
# ---------------------------
def get_note_files_from_gcs(folder_type):
    """
    指定されたフォルダタイプに対応するGCSパスからファイルリストを取得する。
    ルート ('')、ep_notes/、mainte_notes/ の複数パスを検索する。
    Returns: list of (file_name, full_gcs_path, public_url)
    """
    if isinstance(storage_client, DummyStorageClient) or storage is None:
        st.info("GCSクライアントが認証されていません。")
        return []

    search_prefixes = []
    if folder_type == "エピノート":
        # ルート ('') と ep_notes/ の両方を検索
        search_prefixes = ["", "ep_notes/"] 
    elif folder_type == "メンテノート":
        # ルート ('') と mainte_notes/ の両方を検索
        search_prefixes = ["", "mainte_notes/"] 
    # その他のタイプも必要に応じて追加

    all_files = {} # {表示名: (フルパス, URL)}

    try:
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        
        for prefix in search_prefixes:
            # GCS Blobオブジェクトのリストを取得
            blobs = bucket.list_blobs(prefix=prefix)
            
            for blob in blobs:
                # フォルダ自体（例: 'ep_notes/'）は除外
                if blob.name.endswith('/'):
                    continue
                
                # ファイル名を取得 (prefix部分を削除)
                file_name_display = blob.name.replace(prefix, '')
                
                # ルート直下のファイルの場合
                if not file_name_display and prefix == "":
                     file_name_display = blob.name

                # public_urlを作成 (url_quoteを使用)
                public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(blob.name)}"
                
                # 同じファイル名が見つかった場合、新しいパス（通常はサブフォルダ）を優先
                if file_name_display not in all_files:
                    all_files[file_name_display] = (blob.name, public_url)

        # ファイル名でソートして返す (list of (file_name, full_gcs_path, public_url))
        # ファイル名順の降順で、新しいファイル（通常は新しいタイムスタンプ）が上に来るようにする
        return sorted([
            (name, path_url[0], path_url[1]) 
            for name, path_url in all_files.items()
        ], key=lambda x: x[0], reverse=True) 

    except Exception as e:
        st.error(f"GCSファイルリストの取得中にエラーが発生しました: {e}")
        return []

# ---------------------------
# --- GCSファイル閲覧機能 ---
# ---------------------------
# ---------------------------
# --- GCSファイル閲覧機能 ---
# ---------------------------
def display_gcs_files(folder_type):
    st.markdown("#### 📂 GCS ファイルブラウザ (直接ファイル一覧)")
    st.caption(f"GCSバケット ({CLOUD_STORAGE_BUCKET_NAME}) のルートと /{folder_type}/ フォルダからファイルを検索しています。")
    
    # get_note_files_from_gcs は以前の修正で定義済みと仮定
    file_list = get_note_files_from_gcs(folder_type)

    if not file_list:
        st.info("GCSからファイルが見つかりませんでした。")
        return
        
    # ファイル名のみのリストで選択ボックスを作成
    selected_file_name = st.selectbox(
        f"GCS内のファイルを選択 (合計 {len(file_list)} 件)", 
        options=[item[0] for item in file_list],
        key=f"{folder_type}_gcs_browser"
    )
    
    if selected_file_name:
        # 選択されたファイルの情報を検索 (name, gcs_path, url)
        selected_info = next((item for item in file_list if item[0] == selected_file_name), None)
        
        if selected_info:
            file_name, gcs_path, public_url = selected_info
            
            # **【重要修正】非公開ファイルアクセスのため、署名付きURLを生成**
            signed_url = generate_signed_url(gcs_path, expiration_minutes=5) # 有効期限5分
            
            if signed_url is None:
                st.error("認証済みアクセスリンクの生成に失敗しました。GCS認証設定を確認してください。")
                return
            
            # --- ファイル表示 ---
            st.markdown("---")
            st.markdown(f"**ファイル名:** `{file_name}`")
            st.markdown(f"**GCSパス:** `{gcs_path}`")
            st.warning("※このリンクは署名付きURLであり、5分間の有効期限があります。")
            
            # display_attached_files が期待する形式に合わせる
            pseudo_row = {
                'url': json.dumps([signed_url]), # 署名付きURLを使用
                'filename': json.dumps([file_name])
            }
            
            # display_attached_files(row_dict, col_url_key, col_filename_key)
            display_attached_files(pseudo_row, 'url', 'filename')
        else:
            st.error("ファイル情報が見つかりません。")
# ... (後略: page_mainte_list など、他のリスト表示関数もすべて page_data_list を呼び出しており、page_data_list が display_attached_files を呼び出しているため、自動的に新しい表示方法が適用されます。) ...

# ---------------------------
# --- ユーティリティ参照 ---
# ---------------------------
# 前半部を同一ファイルにまとめない場合は import で呼ぶ（例: from bennriyasann3_fixed_v2_part1 import *）
# ここでは「同一実行環境にpart1がロード済み」と仮定します。

# ---------------------------
# --- 汎用的な一覧表示関数 ---
# ---------------------------
def page_data_list(sheet_name, title, col_time, col_filter=None, col_memo=None, col_url=None, detail_cols=None, col_filename=None):
    """汎用的なデータ一覧ページ"""
    st.header(f"📚 {title} 一覧")
    df = get_sheet_as_df(SPREADSHEET_NAME, sheet_name)

    if df.empty:
        st.info("データがありません。")
        return

    st.subheader("絞り込みと検索")

    # フィルタ列があれば選択肢を表示
    if col_filter and col_filter in df.columns:
        df[col_filter] = df[col_filter].fillna('なし')
        options = ["すべて"] + sorted(list(df[col_filter].unique()))
        sel = st.selectbox(f"「{col_filter}」で絞り込み", options)
        if sel != "すべて":
            df = df[df[col_filter] == sel]

    # 日付フィルタ（タイムスタンプ列がある場合）
    if col_time and col_time in df.columns:
        try:
            df['date_only'] = pd.to_datetime(
                df[col_time].astype(str).str.replace(r'[^0-9]', '', regex=True).str[:8],
                errors='coerce', format='%Y%m%d'
            ).dt.date
        except Exception:
            df['date_only'] = pd.NaT

        df_valid = df.dropna(subset=['date_only'])
        if not df_valid.empty:
            min_date = df_valid['date_only'].min()
            max_date = df_valid['date_only'].max()
            # 存在しない日付をデフォルトにするとエラーになるため、適切なデフォルトを設定
            default_start = min(date(2025, 4, 1), max_date) if isinstance(max_date, date) else date(2025, 4, 1)
            start_date = st.date_input("開始日", value=max(min_date, default_start) if isinstance(min_date, date) else default_start)
            end_date = st.date_input("終了日", value=max_date)
            df = df_valid[(df_valid['date_only'] >= start_date) & (df_valid['date_only'] <= end_date)].drop(columns=['date_only'])

    if df.empty:
        st.info("絞り込み条件に一致するデータがありません。")
        return

    df = df.sort_values(by=col_time, ascending=False).reset_index(drop=True)

    st.markdown("---")
    st.subheader(f"検索結果 ({len(df)} 件)")

    # 表示用の選択ボックス（行を選ぶと詳細表示）
    df['display_index'] = df.index
    def fmt(i):
        row = df.loc[i]
        t = str(row[col_time]) if col_time in row and pd.notna(row[col_time]) else ""
        filt = row[col_filter] if col_filter and col_filter in row and pd.notna(row[col_filter]) else ""
        memo_preview = row[col_memo].split('\n')[0] if col_memo and col_memo in row and pd.notna(row[col_memo]) else ""
        return f"[{t.split('_')[0]}] {filt} - {memo_preview[:50]}"

    sel_idx = st.selectbox("詳細を表示する記録を選択", options=df['display_index'], format_func=fmt)

    if sel_idx is not None:
        row = df.loc[sel_idx]
        st.markdown(f"#### 選択された記録 (ID: {sel_idx+1})")
        
        # 👇 NameErrorを解消するため、ここで定義します
        cols_to_skip = [col_url, col_filename] 
        
        if detail_cols:
            for c in detail_cols:
                # 添付ファイルの列であればスキップ
                if c in cols_to_skip:
                    continue
                    
                if c in row and pd.notna(row[c]):
                    # メモや長文はテキスト表示
                    if col_memo == c or '内容' in c or len(str(row[c])) > 200:
                        st.markdown(f"**{c}:**")
                        st.text(row[c])
                    else:
                        st.write(f"**{c}:** {row[c]}")

        # 添付ファイルの表示
        if col_url and col_url in row:
            st.markdown("##### 添付ファイル")
            display_attached_files(row, col_url, col_filename)
# ---------------------------
# --- エピノートページ ---
# ---------------------------
# app (4).py: 約842行目から
def page_epi_note_recording():
    st.markdown("#### 📝 新しいエピノートを記録")
    
    # フォーム全体
    with st.form(key='epi_note_form'):
        
        # ユーザー入力フィールド
        # (ここでは既存のコードの内容を正確に再現する必要があります。不足しているフィールドは既存のコードから補完してください)
        
        ep_title = st.text_input("タイトル/番号 (例: 791)", key="epi_title")
        ep_category = st.selectbox("カテゴリ", ["測定", "作製", "データ整理", "その他"], key="epi_category")
        ep_memo = st.text_area("詳細メモ", height=200, key="epi_memo")
        
        uploaded_files = st.file_uploader(
            "添付ファイル (画像, PDF, データファイルなど)", 
            type=None, 
            accept_multiple_files=True,
            key="epi_uploader"
        )
        
        st.markdown("---")
        
        # データのインポート expaneder (ご指摘のあったインデントエラーの発生源)
        with st.expander("データのインポート"):
            # インデントエラー回避のため、ダミーの処理として pass を挿入
            pass  
            
        # フォーム送信ボタン
        submit_button = st.form_submit_button("記録を保存") 
        
    # フォームの処理はフォームの外側ではなく、submit_button の戻り値で制御
    if submit_button:
        if not ep_title:
            st.warning("番号 (例: 791) は必須項目です。")
            return
            
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    # 【修正済み】フォルダ名引数を削除し、GCSルートに保存
                    # storage_clientはグローバルに定義されていると仮定
                    filename, url = upload_file_to_gcs(storage_client, file_obj) 
                    
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)
                    else:
                        st.error(f"ファイル {file_obj.name} のアップロードに失敗しました。")

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        memo_content = f"{ep_title}\n{ep_memo}"
        
        # 定数 (EPI_COL_NOTE_TYPE, SHEET_EPI_DATA, SPREADSHEET_NAME) はグローバルに定義されていると仮定
        EPI_COL_NOTE_TYPE = "エピノート" # 仮の定数
        SHEET_EPI_DATA = "エピノート"   # 仮の定数
        
        row_data = [timestamp, EPI_COL_NOTE_TYPE, ep_category, memo_content, filenames_json, urls_json]
        
        try:
            # gc は authenticate_gspread() で取得したグローバル変数と仮定
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_EPI_DATA)
            worksheet.append_row(row_data)
            st.success("✅ エピノートをアップロードしました！")
            
            # キャッシュクリアとリランは、 Streamlit の仕様に従い、適切に配置
            if 'st.cache_data' in st.__dict__:
                st.cache_data.clear()
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

def page_epi_note_list():
    detail_cols = [EPI_COL_TIMESTAMP, EPI_COL_CATEGORY, EPI_COL_NOTE_TYPE, EPI_COL_MEMO, EPI_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_EPI_DATA,
        title="エピノート",
        col_time=EPI_COL_TIMESTAMP,
        col_filter=EPI_COL_CATEGORY,
        col_memo=EPI_COL_MEMO,
        col_url=EPI_COL_FILE_URL,
        detail_cols=detail_cols,
        col_filename=EPI_COL_FILENAME
    )

def page_epi_note():
    st.header("エピノート機能")
    st.markdown("---")
    # 既存のtabを修正し、「GCS閲覧」を追加
    tab_titles = ["📝 記録", "📚 一覧 (シート)", "📂 GCSファイル閲覧"]
    tab = st.radio("表示切替", tab_titles, key="epi_tab", horizontal=True)
    
    if tab == "📝 記録":
        page_epi_note_recording()
    elif tab == "📚 一覧 (シート)":
        page_epi_note_list()
    elif tab == "📂 GCSファイル閲覧":
        display_gcs_files("エピノート") # 新規追加

# ---------------------------
# --- メンテノートページ ---
# ---------------------------
# app (4).py: 約911行目から
def page_mainte_recording():
    st.markdown("#### 🛠️ 新しいメンテノートを記録")
    
    # フォーム全体
    with st.form(key='mainte_note_form'):
        
        # ユーザー入力フィールド (既存のコードに基づき再現)
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
        
        # データのインポート expaneder (インデントエラーの発生源)
        with st.expander("データのインポート"):
            # インデントエラー回避のため pass を挿入
            pass
            
        # フォーム送信ボタン
        submit_button = st.form_submit_button("記録を保存")
        
    # フォームの処理
    if submit_button:
        if not mainte_title:
            st.warning("メンテタイトルを入力してください。")
            return
            
        filenames_list = []; urls_list = []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for file_obj in uploaded_files:
                    # 【修正済み】フォルダ名引数を削除し、GCSルートに保存
                    filename, url = upload_file_to_gcs(storage_client, file_obj)
                    
                    if url:
                        filenames_list.append(filename)
                        urls_list.append(url)
                    else:
                        st.error(f"ファイル {file_obj.name} のアップロードに失敗しました。")
        
        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        memo_to_save = f"[{mainte_title}]\n{memo_content}"
        
        # 定数 (MAINTE_COL_NOTE_TYPE, SHEET_MAINTE_DATA, SPREADSHEET_NAME) はグローバルに定義されていると仮定
        MAINTE_COL_NOTE_TYPE = "メンテノート" # 仮の定数
        SHEET_MAINTE_DATA = "メンテノート"   # 仮の定数
        
        row_data = [timestamp, MAINTE_COL_NOTE_TYPE, mainte_device, memo_to_save, filenames_json, urls_json]
        
        try:
            # gc は authenticate_gspread() で取得したグローバル変数と仮定
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_MAINTE_DATA)
            worksheet.append_row(row_data)
            st.success("✅ メンテノートをアップロードしました！")
            
            # キャッシュクリアとリラン
            if 'st.cache_data' in st.__dict__:
                st.cache_data.clear()
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

def page_mainte_list():
    detail_cols = [MAINT_COL_TIMESTAMP, MAINT_COL_NOTE_TYPE, MAINT_COL_MEMO, MAINT_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_MAINTE_DATA,
        title="メンテノート",
        col_time=MAINT_COL_TIMESTAMP,
        col_filter=MAINT_COL_NOTE_TYPE,
        col_memo=MAINT_COL_MEMO,
        col_url=MAINT_COL_FILE_URL,
        detail_cols=detail_cols,
        col_filename=MAINT_COL_FILENAME
    )

def page_mainte_note():
    st.header("メンテノート機能")
    st.markdown("---")
    # 既存のtabを修正し、「GCS閲覧」を追加
    tab_titles = ["📝 記録", "📚 一覧 (シート)", "📂 GCSファイル閲覧"]
    tab = st.radio("表示切替", tab_titles, key="mainte_tab", horizontal=True)
    
    if tab == "📝 記録":
        page_mainte_recording()
    elif tab == "📚 一覧 (シート)":
        page_mainte_list()
    elif tab == "📂 GCSファイル閲覧":
        display_gcs_files("メンテノート") # 新規追加

# ---------------------------
# --- 議事録ページ ---
# ---------------------------
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
        row_data = [timestamp, meeting_title, audio_name, audio_url, meeting_content]
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_MEETING_DATA)
            worksheet.append_row(row_data)
            st.success("✅ 議事録をアップロードしました！")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

def page_meeting_list():
    detail_cols = [MEETING_COL_TIMESTAMP, MEETING_COL_TITLE, MEETING_COL_CONTENT, MEETING_COL_AUDIO_NAME, MEETING_COL_AUDIO_URL]
    page_data_list(
        sheet_name=SHEET_MEETING_DATA,
        title="議事録",
        col_time=MEETING_COL_TIMESTAMP,
        col_filter=MEETING_COL_TITLE,
        col_memo=MEETING_COL_CONTENT,
        col_url=MEETING_COL_AUDIO_URL,
        detail_cols=detail_cols,
        col_filename=MEETING_COL_AUDIO_NAME
    )

def page_meeting_note():
    st.header("議事録・ミーティングメモ機能")
    st.markdown("---")
    tab = st.radio("表示切替", ["📝 記録", "📚 一覧"], key="meeting_tab", horizontal=True)
    if tab == "📝 記録":
        page_meeting_recording()
    else:
        page_meeting_list()

# ---------------------------
# --- 知恵袋ページ ---
# ---------------------------
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
        row_data = [timestamp, qa_title, qa_content, qa_contact, filenames_json, urls_json, "未解決"]
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_QA_DATA)
            worksheet.append_row(row_data)
            st.success("✅ 質問をアップロードしました！")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

def page_qa_list():
    detail_cols = [QA_COL_TIMESTAMP, QA_COL_TITLE, QA_COL_CONTENT, QA_COL_CONTACT, QA_COL_STATUS, QA_COL_FILENAME]
    page_data_list(
        sheet_name=SHEET_QA_DATA,
        title="知恵袋・質問箱",
        col_time=QA_COL_TIMESTAMP,
        col_filter=QA_COL_STATUS,
        col_memo=QA_COL_CONTENT,
        col_url=QA_COL_FILE_URL,
        detail_cols=detail_cols,
        col_filename=QA_COL_FILENAME
    )

def page_qa_box():
    st.header("知恵袋・質問箱機能")
    st.markdown("---")
    tab = st.radio("表示切替", ["💡 質問投稿", "📚 質問一覧"], key="qa_tab", horizontal=True)
    if tab == "💡 質問投稿":
        page_qa_recording()
    else:
        page_qa_list()

# ---------------------------
# --- 引き継ぎページ ---
# ---------------------------
def page_handover_recording():
    st.markdown("#### 🤝 新しい引き継ぎメモを記録")
    with st.form(key='handover_form'):
        handover_type = st.selectbox(f"{HANDOVER_COL_TYPE} (カテゴリ)", ["マニュアル", "装置設定", "その他メモ"])
        handover_title = st.text_input(f"{HANDOVER_COL_TITLE} (例: D1 MBE起動手順)", key='handover_title_input')
        handover_memo = st.text_area(f"{HANDOVER_COL_MEMO}", height=150, key='handover_memo_input')
        st.markdown("---")
        submit_button = st.form_submit_button(label='記録をスプレッドシートに保存')
    if submit_button:
        if not handover_title:
            st.warning("タイトルを入力してください。")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        row_data = [timestamp, handover_type, handover_title, handover_memo, "", "", ""]
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_HANDOVER_DATA)
            worksheet.append_row(row_data)
            st.success("✅ 引き継ぎメモをアップロードしました！")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

def page_handover_list():
    detail_cols = [HANDOVER_COL_TIMESTAMP, HANDOVER_COL_TYPE, HANDOVER_COL_TITLE, '内容1', '内容2', '内容3', HANDOVER_COL_MEMO]
    page_data_list(
        sheet_name=SHEET_HANDOVER_DATA,
        title="装置引き継ぎメモ",
        col_time=HANDOVER_COL_TIMESTAMP,
        col_filter=HANDOVER_COL_TYPE,
        col_memo=HANDOVER_COL_TITLE,
        col_url='内容1',
        detail_cols=detail_cols,
        col_filename=None
    )

def page_handover_note():
    st.header("装置引き継ぎメモ機能")
    st.markdown("---")
    tab = st.radio("表示切替", ["📝 記録", "📚 一覧"], key="handover_tab", horizontal=True)
    if tab == "📝 記録":
        page_handover_recording()
    else:
        page_handover_list()

# ---------------------------
# --- トラブル報告ページ ---
# ---------------------------
def page_trouble_recording():
    st.markdown("#### 🚨 新しいトラブルを報告")
    DEVICE_OPTIONS = ["MBE", "XRD", "PL", "IV", "TEM・SEM", "抵抗加熱蒸着", "RTA", "フォトリソ", "ドラフター", "その他"]
    with st.form(key='trouble_form'):
        st.subheader("基本情報")
        col1, col2 = st.columns(2)
        with col1:
            report_date = st.date_input(f"{TROUBLE_COL_OCCUR_DATE} (発生日)", datetime.now().date())
        with col2:
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
        row_data = [timestamp, device_to_save, report_date.isoformat(), occur_time, cause, solution, prevention, reporter_name, filenames_json, urls_json, report_title]
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_TROUBLE_DATA)
            worksheet.append_row(row_data)
            st.success("✅ トラブル報告をアップロードしました！")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

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
        detail_cols=detail_cols,
        col_filename=TROUBLE_COL_FILENAME
    )

def page_trouble_report():
    st.header("トラブル報告機能")
    st.markdown("---")
    tab = st.radio("表示切替", ["📝 記録", "📚 一覧"], key="trouble_tab", horizontal=True)
    if tab == "📝 記録":
        page_trouble_recording()
    else:
        page_trouble_list()

# ---------------------------
# --- 連絡・問い合わせページ ---
# ---------------------------
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
        row_data = [timestamp, contact_type, contact_detail, contact_info]
        try:
            worksheet = gc.open(SPREADSHEET_NAME).worksheet(SHEET_CONTACT_DATA)
            worksheet.append_row(row_data)
            st.success("✅ お問い合わせを送信しました。")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

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
    tab = st.radio("表示切替", ["📝 記録", "📚 一覧"], key="contact_tab", horizontal=True)
    if tab == "📝 記録":
        page_contact_recording()
    else:
        page_contact_list()

# ---------------------------
# --- IVデータ解析ページ ---
# ---------------------------
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    st.markdown("IV測定データファイル（2列データ：X軸/Y軸）をアップロードし、ソートせずに往路/復路の特性をプロットします。")
    
    uploaded_files = st.file_uploader("IV測定データファイル (.txt) をアップロード", type=['txt'], accept_multiple_files=True)
    if not uploaded_files:
        st.info("ファイルをアップロードしてください。")
        return

    data_frames = {}
    file_names = []
    
    st.subheader("ステップ1: ファイル読み込みと解析")
    
    for uploaded_file in uploaded_files:
        # load_data_file は bytes を受け取る 
        df = load_data_file(uploaded_file.getvalue(), uploaded_file.name)
        if df is not None and not df.empty:
            data_frames[uploaded_file.name] = df
            file_names.append(uploaded_file.name)
        else:
            # 読み込みに失敗したファイル名を表示
            st.warning(f"{uploaded_file.name} の読み込みに失敗したか、データが不正です。") 
            
    if not data_frames:
        st.error("有効なデータファイルが見つかりませんでした。")
        return

    st.success(f"{len(data_frames)} 個の有効なファイルを読み込みました。")

    st.subheader("ステップ2: グラフ表示 (往路/復路)")
    
    fig, ax = plt.subplots(figsize=(12, 7))
    
    # 軸ラベルは最初のDFから取得 (データが2列であることを前提)
    x_col = data_frames[file_names[0]].columns[0] # Axis_X
    y_col = data_frames[file_names[0]].columns[1] # <ファイル名>
    
    for name, df in data_frames.items():
        # X軸データ
        x_data = df.iloc[:, 0] 
        
        # 往路/復路の分割点を特定: X軸の最大値のインデックス
        if x_data.empty:
             continue
             
        max_index = x_data.idxmax()

        # 往路 (Forward: データの最初から最大値まで)
        df_forward = df.iloc[:max_index + 1]
        if not df_forward.empty:
            ax.plot(df_forward.iloc[:, 0], df_forward.iloc[:, 1], 
                    label=f"{name} (往路)", linestyle='-', marker='o', markersize=3, alpha=0.7)

        # 復路 (Reverse: 最大値の次から最後まで)
        df_reverse = df.iloc[max_index + 1:]
        if not df_reverse.empty:
            ax.plot(df_reverse.iloc[:, 0], df_reverse.iloc[:, 1], 
                    label=f"{name} (復路)", linestyle='--', marker='x', markersize=3, alpha=0.7)

    # グラフの装飾
    ax.set_xlabel(x_col)
    ax.set_ylabel(y_col)
    ax.set_title("IV特性比較 (ソートなし、往路/復路表示)")
    ax.legend(title="ファイル名", loc='best')
    ax.grid(True)
    st.pyplot(fig, use_container_width=True)
    plt.close(fig) # メモリ解放
    
    st.info("※ ユーザーの要望に基づき、データ結合（補間）およびX軸でのソート処理は行っていません。")

# ---------------------------
# --- PLデータ解析ページ ---
# ---------------------------
def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    st.write("ステップ1：波長校正（2点） → ステップ2：測定データ解析 の順に実行してください。")

    # --- ステップ1: 校正 ---
    with st.expander("ステップ1：波長校正", expanded=True):
        st.write("2つの基準波長の反射光データをアップロードして、分光器の傾き（nm/pixel）を校正します。")
        col1, col2 = st.columns(2)
        with col1:
            cal1_wavelength = st.number_input("基準波長1 (nm)", value=1500)
            cal1_file = st.file_uploader(f"{cal1_wavelength} nm の校正ファイル (.txt)", type=['txt'], key="pl_cal1")
        with col2:
            cal2_wavelength = st.number_input("基準波長2 (nm)", value=1570)
            cal2_file = st.file_uploader(f"{cal2_wavelength} nm の校正ファイル (.txt)", type=['txt'], key="pl_cal2")

        if st.button("校正を実行", key="run_pl_cal"):
            if not (cal1_file and cal2_file):
                st.warning("両方の校正ファイルをアップロードしてください。")
            else:
                df1 = load_pl_data(cal1_file)
                df2 = load_pl_data(cal2_file)
                if df1 is None or df2 is None:
                    st.error("校正ファイルのデータ読み込みに失敗しました。ファイル内容・形式を確認してください。")
                else:
                    try:
                        peak_pixel1 = df1['pixel'].iloc[df1['intensity'].idxmax()]
                        peak_pixel2 = df2['pixel'].iloc[df2['intensity'].idxmax()]

                        st.write("---")
                        st.subheader("校正結果")
                        c1, c2, c3 = st.columns(3)
                        c1.metric(f"{cal1_wavelength} nm のピーク位置", f"{int(peak_pixel1)} pixel")
                        c2.metric(f"{cal2_wavelength} nm のピーク位置", f"{int(peak_pixel2)} pixel")

                        delta_wave = float(cal2_wavelength - cal1_wavelength)
                        delta_pixel = float(peak_pixel1 - peak_pixel2)
                        if delta_pixel == 0:
                            st.error("2つのピーク位置が同じです。異なる校正ファイルを選んでください。")
                        else:
                            slope = delta_wave / delta_pixel
                            c3.metric("校正係数 (nm/pixel)", f"{slope:.6f}")
                            st.session_state['pl_calibrated'] = True
                            st.session_state['pl_slope'] = slope
                            st.session_state['pl_center_wl_cal'] = cal1_wavelength
                            st.session_state['pl_center_pixel_cal'] = peak_pixel1
                            st.success("校正係数を保存しました。ステップ2に進んでください。")
                    except Exception as e:
                        st.error(f"校正計算中にエラーが発生しました: {e}")

    st.write("---")
    st.subheader("ステップ2：測定データ解析")

    if not st.session_state.get('pl_calibrated', False):
        st.info("まずステップ1の波長校正を完了してください。")
        return

    # --- ステップ2: 測定データの解析 ---
    center_wavelength_input = st.number_input(
        "測定時の中心波長 (nm)", min_value=0, value=1700, step=10,
        help="この測定で装置に設定した中心波長を入力してください（凡例整形にも利用）。"
    )
    uploaded_files = st.file_uploader("測定データファイル（複数選択可）をアップロード", type=['txt'], accept_multiple_files=True, key="pl_measure_files")

    if not uploaded_files:
        st.info("測定データファイルをアップロードしてください。")
        return

    st.subheader("解析結果")
    fig, ax = plt.subplots(figsize=(10, 6))
    all_dataframes = []
    slope = st.session_state['pl_slope']
    center_pixel = 256.5  # あなたの既存ロジックをそのまま使用

    for uploaded_file in uploaded_files:
        df = load_pl_data(uploaded_file)
        if df is None:
            st.warning(f"{uploaded_file.name} の読み込みに失敗したためスキップします。")
            continue

        # 波長変換
        df['wavelength_nm'] = (df['pixel'] - center_pixel) * slope + center_wavelength_input

        base_name = os.path.splitext(uploaded_file.name)[0]
        cleaned_label = base_name.replace(str(int(center_wavelength_input)), "").strip(' _-')
        label = cleaned_label if cleaned_label else base_name

        ax.plot(df['wavelength_nm'], df['intensity'], label=label, linewidth=2.5)

        export_df = df[['wavelength_nm', 'intensity']].copy()
        export_df.rename(columns={'intensity': base_name}, inplace=True)
        all_dataframes.append(export_df)

    if not all_dataframes:
        st.warning("有効な測定データがありません。")
        return

    # 結合（波長をキーに外部結合）
    final_df = all_dataframes[0]
    for i in range(1, len(all_dataframes)):
        final_df = pd.merge(final_df, all_dataframes[i], on='wavelength_nm', how='outer')

    final_df = final_df.sort_values(by='wavelength_nm').reset_index(drop=True)

    # グラフ整形
    ax.set_title(f"PL spectrum (Center: {center_wavelength_input} nm)")
    ax.set_xlabel("Wavelength [nm]")
    ax.set_ylabel("PL intensity [a.u.]")
    ax.legend(loc='upper left', frameon=False, fontsize=10)
    ax.grid(axis='y', linestyle='-', color='lightgray', zorder=0)
    ax.tick_params(direction='in', top=True, right=True, which='both')

    # x 範囲パディング
    min_wl = final_df['wavelength_nm'].min()
    max_wl = final_df['wavelength_nm'].max()
    if pd.notna(min_wl) and pd.notna(max_wl) and max_wl > min_wl:
        padding = (max_wl - min_wl) * 0.05
        ax.set_xlim(min_wl - padding, max_wl + padding)

    st.pyplot(fig, use_container_width=True)

    # Excel 出力（openpyxl を使用）
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Combined PL Data')
        processed_data = output.getvalue()
        st.download_button(
            label="📈 Excelデータとしてダウンロード",
            data=processed_data,
            file_name=f"pl_analysis_combined_{center_wavelength_input}nm.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Excel出力に失敗しました: {e}")

# ---------------------------
# --- Google Calendar API接続ユーティリティを追加 ---
# ---------------------------
@st.cache_resource(ttl=3600)
def get_calendar_service():
    """Google Calendar API サービスオブジェクトを取得する (page_calendar()で使用)"""
    if "gcs_credentials" not in st.secrets:
        st.error("Google Calendar API認証情報 (`gcs_credentials`) が Streamlit secrets に設定されていません。")
        return None
    try:
        raw = st.secrets["gcs_credentials"]
        # クレンジング
        cleaned = raw.strip().replace('\t', '').replace('\r', '').replace('\n', '')
        info = json.loads(cleaned)

        # サービスアカウント認証情報を使用してCredentialsオブジェクトを作成
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        # Calendar API サービスオブジェクトを構築
        service = build('calendar', 'v3', credentials=creds)
        return service
    except json.JSONDecodeError:
        st.error("カレンダー認証情報JSONの解析に失敗しました。")
        return None
    except Exception as e:
        st.error(f"Google Calendar Serviceの初期化に失敗しました: {e}")
        return None
        
# --------------------------
# --- 予約・カレンダーページ（最終調整版） ---
# --------------------------
def page_calendar():
# --- 0. セッションステートに保存されたメッセージを表示 ---
    if 'calendar_success_message' in st.session_state:
        st.success(st.session_state['calendar_success_message'])
        # メッセージを一度表示したら消去
        del st.session_state['calendar_success_message']
        
    st.header("🗓️ スケジュール・装置予約")
    
    # --- 1. 外部予約サイトへのリンク ---
    st.subheader("外部予約サイト")
    
    col_evers, col_rac = st.columns(2)
    
    # Evers 予約サイト
    evers_url = "https://www.eiiris.tut.ac.jp/evers/Web/dashboard.php"
    col_evers.markdown(
        f'<a href="{evers_url}" target="_blank">'
        f'<button style="width:100%; height:40px; background-color:#4CAF50; color:white; border:none; border-radius:5px; cursor:pointer;">'
        f'Evers 予約サイトへアクセス</button></a>',
        unsafe_allow_html=True
    )
    col_evers.caption("（学内共用装置予約システム）")

    # 教育研究基盤センター 予約ポータル
    rac_url = "https://tech.rac.tut.ac.jp/regist/potal_0.php"
    col_rac.markdown(
        f'<a href="{rac_url}" target="_blank">'
        f'<button style="width:100%; height:40px; background-color:#2196F3; color:white; border:none; border-radius:5px; cursor:pointer;">'
        f'教育研究基盤センター ポータルへ</button></a>',
        unsafe_allow_html=True
    )
    col_rac.caption("（共用施設利用登録）")

    st.markdown("---")
    
    # --- 2. Googleカレンダーの埋め込み ---
    st.subheader("予約カレンダー（Googleカレンダー）")

    calendar_html = f"""
    <iframe src="https://calendar.google.com/calendar/embed?height=600&wkst=1&bgcolor=%23ffffff&ctz=Asia%2FTokyo&src={CALENDAR_ID}&color=%237986CB&showTitle=0&showPrint=0&showCalendars=0&showTz=0" style="border-width:0" width="100%" height="600" frameborder="0" scrolling="no"></iframe>
    """
    
    st.markdown(calendar_html, unsafe_allow_html=True)
    
    st.caption("カレンダーの予約状況を確認し、以下のフォームから予定を登録してください。")
    st.markdown("---") 
    
    # ------------------------------------------------------------------
    # --- 3. 予約登録の制御部分（フォーム外: 即時応答が必要な要素） ---
    # ------------------------------------------------------------------
    st.subheader("🗓️ 新規予定の登録")
    
    initial_user_name = st.session_state.get('user_name', '')
    
    # 1. 登録者名をフォームの外に配置 (即時表示のため)
    user_name = st.text_input("登録者名 / グループ名", value=initial_user_name, key="user_name_outside")
    
    # 2. カテゴリ選択とカスタム入力欄をフォームの外に配置 (即時表示のため)
    col_cat, col_other = st.columns([1, 2])
    
    with col_cat:
        category = st.selectbox("作業/装置カテゴリ", CATEGORY_OPTIONS, key="category_select_outside")
        
    custom_category = ""
    with col_other:
        # 💡 即時表示が働く
        if category == "その他入力":
            custom_category = st.text_input(
                "カスタムカテゴリを直接入力", 
                placeholder="例: 学会発表準備", 
                key="custom_category_input_cal_outside"
            ) 
    
    # 3. 最終カテゴリ名を決定 (submitボタンが押される前に確定)
    final_category = custom_category if category == "その他入力" else category
    
    st.markdown("---") # フォームとの区切り線
    
    # -----------------------------------------------------
    # --- 4. フォーム本体 ---
    # -----------------------------------------------------
    with st.form(key='schedule_form'):
        
        # 💡 フォーム外で決定した情報を、フォームの最初で表示
        st.markdown(f"**📚 選択されたカテゴリ:** `{final_category}`") 
        
        # 💡 予定タイトルを計算し表示
        final_title_preview = f"{user_name} ({final_category})" if user_name and final_category else ""
        st.markdown(f"**💡 予定のタイトル:** `{final_title_preview}`")

        st.markdown("---")
        
        # 5. 開始日時と終了日時
        st.markdown("##### 予定日時")
        
        cols_start_date, cols_start_time = st.columns(2)
        # st.date_input はそのまま
        start_date = cols_start_date.date_input("開始日", value=date.today())
        
        # ⏰ 開始時刻を st.time_input に変更 
        # time_input は datetime.time オブジェクトを返す
        start_time_obj = cols_start_time.time_input("開始時刻", value=datetime.strptime("09:00", '%H:%M').time())

        cols_end_date, cols_end_time = st.columns(2)
        # st.date_input はそのまま
        end_date = cols_end_date.date_input("終了日", value=date.today())
        
        # ⏰ 終了時刻を st.time_input に変更
        end_time_obj = cols_end_time.time_input("終了時刻", value=datetime.strptime("11:00", '%H:%M').time())
        
        # 💡 st.time_input を使った場合、時刻のオブジェクト (time_obj) は既に正しい形式なので、
        # 後の API 処理で strftime/strptime を使わずに直接使えます。
        
        # 6. 詳細（メモ）
        detail = st.text_area("詳細（予定の内容）", height=100)
        
        submit_button = st.form_submit_button(label='⬆️ Googleカレンダーに自動登録')

        if submit_button:
            # フォーム外の user_name と final_category を使用
            if not user_name or not final_category:
                st.error("「登録者名」と「作業カテゴリ」は必須です。")
                return 
            
            # 最終タイトルを確定
            final_title = f"{user_name} ({final_category})"

            # ----------------------------------------
            # API経由で直接カレンダーに書き込み 
            # ----------------------------------------
            service = get_calendar_service() 
            if service is None:
                return 

            try:
                # ⏰ time_input で取得した time_obj を直接使用するため、strptime は不要
                start_dt_obj = datetime.combine(start_date, start_time_obj) 
                end_dt_obj = datetime.combine(end_date, end_time_obj)
                
                if end_dt_obj <= start_dt_obj:
                    st.error("終了日時は開始日時より後に設定してください。")
                    return

                # 予定のボディを作成
                event_body = {
                    'summary': final_title,
                    'description': detail,
                    'start': {
                        'dateTime': start_dt_obj.isoformat(),
                        'timeZone': 'Asia/Tokyo',
                    },
                    'end': {
                        'dateTime': end_dt_obj.isoformat(),
                        'timeZone': 'Asia/Tokyo',
                    },
                }

                # API経由で予定を挿入
                event = service.events().insert(calendarId=CALENDAR_ID, body=event_body).execute()
                
                # 成功メッセージとユーザー名をセッションに保存
                st.session_state['user_name'] = user_name 
                st.session_state['calendar_success_message'] = f"✅ 予定 `{final_title}` がカレンダーに自動登録されました！"
                
                # ページ全体を再実行し、カレンダー埋め込みを再ロードして更新
                # st.rerun() は st.success の前に実行されるため、メッセージが残るようになります
                st.rerun()
                    
            except ValueError:
                st.error("時刻のフォーマットが無効です。「HH:MM」の形式で入力してください。")
            except HttpError as e:
                st.error(f"カレンダー登録に失敗しました。権限とカレンダーIDを確認してください。詳細: {e.content.decode()}")
            except Exception as e:
                st.error(f"予定の登録中に予期せぬエラーが発生しました: {e}")
# ---------------------------
# --- メインルーティング ---
# ---------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    menu_selection = st.sidebar.radio("機能選択", [
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
    ])

    if menu_selection == "エピノート":
        page_epi_note()
    elif menu_selection == "メンテノート":
        page_mainte_note()
    elif menu_selection == "IVデータ解析":
        page_iv_analysis()
    elif menu_selection == "PLデータ解析":
        page_pl_analysis()
    elif menu_selection == "議事録":
        page_meeting_note()
    elif menu_selection == "知恵袋・質問箱":
        page_qa_box()
    elif menu_selection == "装置引き継ぎメモ":
        page_handover_note()
    elif menu_selection == "トラブル報告":
        page_trouble_report()
    elif menu_selection == "連絡・問い合わせ":
        page_contact_form()
    elif menu_selection == "🗓️ スケジュール・装置予約":
        page_calendar()
    else:
        st.info("選択した機能は未実装です。")

if __name__ == "__main__":
    main()













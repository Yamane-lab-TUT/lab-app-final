# -*- coding: utf-8 -*-
"""
Yamane Lab Convenience Tool - Complete Refactored Version
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

# Google Services
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Optional GCS
try:
    from google.cloud import storage
except ImportError:
    storage = None

# --- Streamlit ページ設定 ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide", page_icon="🧪")

# --- Matplotlib 日本語フォント設定 ---
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
# --- Constants ---
# ---------------------------
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" # 必要に応じて変更
SPREADSHEET_NAME = "エピノート"

# シート定義
SHEET_EPI_DATA = 'エピノート_データ'
SHEET_MAINTE_DATA = 'メンテノート_データ'
SHEET_MEETING_DATA = '議事録_データ'
SHEET_HANDOVER_DATA = '引き継ぎ_データ'
SHEET_QA_DATA = '知恵袋_データ'
SHEET_CONTACT_DATA = 'お問い合わせ_データ'
SHEET_TROUBLE_DATA = 'トラブル報告_データ'

# Google Calendar Config
CALENDAR_ID = "yamane.lab.6747@gmail.com" # ターゲットカレンダーID
SCOPES = ['https://www.googleapis.com/auth/calendar']

# ---------------------------
# --- Dummy Classes for Offline/Error Mode ---
# ---------------------------
class DummyGSClient:
    def open(self, name): return self
    def worksheet(self, name): return self
    def get_all_records(self): return []
    def get_all_values(self): return []
    def append_row(self, values): pass

class DummyStorageClient:
    def bucket(self, name): return self
    def blob(self, name): return self
    def list_blobs(self, **kwargs): return []

# ---------------------------
# --- Google Services Initialization ---
# ---------------------------
@st.cache_resource
def initialize_google_services():
    """Google Sheets, Drive, GCS, Calendarの認証を行う"""
    # デフォルト（失敗時）
    gc_client = DummyGSClient()
    storage_client_obj = DummyStorageClient()
    calendar_service = None
    
    if "gcs_credentials" not in st.secrets:
        st.sidebar.warning("⚠️ Secretsに `gcs_credentials` が設定されていません。")
        return gc_client, storage_client_obj, calendar_service

    try:
        # SecretsからJSON文字列を取得してパース
        raw = st.secrets["gcs_credentials"]
        # 制御文字の削除などクレンジング
        cleaned = raw.strip().replace('\t', '').replace('\r', '').replace('\n', '')
        info = json.loads(cleaned)
        
        # 1. Gspread (Sheets)
        gc_client = gspread.service_account_from_dict(info)
        
        # 2. GCS
        if storage:
            storage_client_obj = storage.Client.from_service_account_info(info)
        
        # 3. Calendar API
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        calendar_service = build('calendar', 'v3', credentials=creds)
        
        return gc_client, storage_client_obj, calendar_service

    except Exception as e:
        st.sidebar.error(f"Googleサービス認証エラー: {e}")
        return gc_client, storage_client_obj, calendar_service

# グローバル変数として初期化
gc, storage_client, calendar_service = initialize_google_services()

# ---------------------------
# --- Utils: GCS Upload & File Handling ---
# ---------------------------
def upload_file_to_gcs(file_obj):
    """ファイルをGCSにアップロードし、ファイル名と公開URLを返す"""
    if isinstance(storage_client, DummyStorageClient) or storage is None:
        return None, None
        
    try:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        original_filename = file_obj.name
        safe_filename = re.sub(r'[^a-zA-Z0-9_.]', '_', original_filename)
        gcs_filename = f"{timestamp}_{safe_filename}"
        
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        
        blob.upload_from_string(
            file_obj.getvalue(), 
            content_type=file_obj.type if hasattr(file_obj, 'type') else 'application/octet-stream'
        )
        
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        return original_filename, public_url
    except Exception as e:
        st.error(f"アップロードエラー: {e}")
        return None, None

def generate_signed_url(blob_name, expiration_minutes=15):
    """署名付きURLを生成（非公開バケット用）"""
    if isinstance(storage_client, DummyStorageClient): return None
    try:
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(blob_name)
        return blob.generate_signed_url(version="v4", expiration=timedelta(minutes=expiration_minutes), method="GET")
    except Exception:
        return None

def get_note_files_from_gcs(folder_prefix=""):
    """GCS内のファイル一覧を取得"""
    if isinstance(storage_client, DummyStorageClient): return []
    try:
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        # ルートと特定のフォルダプレフィックスを検索
        blobs = list(bucket.list_blobs(prefix=folder_prefix))
        if folder_prefix != "":
            # ルートも検索対象に加える
            blobs += list(bucket.list_blobs(prefix=""))
            
        file_list = []
        seen = set()
        for blob in blobs:
            if blob.name.endswith('/'): continue
            if blob.name in seen: continue
            seen.add(blob.name)
            
            # URL生成
            public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(blob.name)}"
            file_list.append((blob.name, blob.name, public_url))
            
        # 新しい順にソート
        return sorted(file_list, key=lambda x: x[0], reverse=True)
    except Exception:
        return []

# ---------------------------
# --- Utils: Spreadsheet & Data ---
# ---------------------------
@st.cache_data(ttl=600)
def get_sheet_as_df(spreadsheet_name, sheet_name):
    """スプレッドシートをDataFrameとして読み込む"""
    try:
        if isinstance(gc, DummyGSClient): return pd.DataFrame()
        ws = gc.open(spreadsheet_name).worksheet(sheet_name)
        data = ws.get_all_values()
        if not data or len(data) <= 1:
            return pd.DataFrame()
        return pd.DataFrame(data[1:], columns=data[0])
    except Exception:
        return pd.DataFrame()

def display_attached_files(row, col_url, col_filename):
    """JSON形式または文字列形式の添付ファイルリンクを表示"""
    raw_urls = row.get(col_url, '')
    raw_names = row.get(col_filename, '')
    
    urls = []
    names = []
    
    # URL解析
    try:
        urls = json.loads(raw_urls)
        if not isinstance(urls, list): urls = [raw_urls]
    except:
        # 古い形式：単一URLまたはカンマ区切りと仮定
        if raw_urls.startswith('http'): urls = [raw_urls]

    # 名前解析
    try:
        names = json.loads(raw_names)
        if not isinstance(names, list): names = [raw_names]
    except:
        names = [f"File {i+1}" for i in range(len(urls))]

    # 長さ合わせ
    while len(names) < len(urls): names.append(f"File {len(names)+1}")
    
    # 表示
    if urls:
        st.markdown("**📎 添付ファイル:**")
        for u, n in zip(urls, names):
            if u and isinstance(u, str) and u.startswith('http'):
                st.markdown(f"- [{n}]({u})")

# ---------------------------
# --- Utils: Analysis Loaders ---
# ---------------------------
@st.cache_data
def load_iv_data(uploaded_file):
    """IVデータ（2列）の読み込み"""
    try:
        content = uploaded_file.getvalue().decode('utf-8', errors='ignore')
        df = pd.read_csv(io.StringIO(content), sep=r'[\t, ]+', engine='python', header=None)
        if df.shape[1] < 2: return None
        df = df.iloc[:, :2]
        df.columns = ['Axis_X', uploaded_file.name]
        df = df.apply(pd.to_numeric, errors='coerce').dropna()
        return df
    except:
        return None

@st.cache_data
def load_pl_data(uploaded_file):
    """PLデータ（Pixel, Intensity）の読み込み"""
    try:
        content = uploaded_file.getvalue().decode('utf-8', errors='ignore').splitlines()
        # コメント行スキップ
        data_lines = [line.strip() for line in content if line.strip() and not line.strip().startswith(('#','!','/'))]
        
        # 正規化（カンマ、タブをスペースに）
        normalized = [re.sub(r'[\t,]+', ' ', line) for line in data_lines]
        
        df = pd.read_csv(io.StringIO("\n".join(normalized)), sep=' ', header=None, names=['pixel', 'intensity'])
        df = df.apply(pd.to_numeric, errors='coerce').dropna()
        return df
    except:
        return None

# ---------------------------
# --- Components: Generic List & GCS Browser ---
# ---------------------------
def page_data_list_view(sheet_name, title, col_time, col_filter, col_memo, detail_cols):
    st.subheader(f"📚 {title} 一覧")
    df = get_sheet_as_df(SPREADSHEET_NAME, sheet_name)
    
    if df.empty:
        st.info("データがありません。")
        return

    # フィルタ
    if col_filter and col_filter in df.columns:
        options = ["すべて"] + sorted(list(df[col_filter].unique()))
        sel = st.selectbox(f"{col_filter}で絞り込み", options)
        if sel != "すべて":
            df = df[df[col_filter] == sel]

    # ソート
    if col_time in df.columns:
        df = df.sort_values(col_time, ascending=False)

    # リスト表示
    st.markdown("---")
    for i, row in df.iterrows():
        with st.expander(f"{row.get(col_time,'')} - {str(row.get(col_memo,''))[:30]}..."):
            for col in detail_cols:
                if col in row:
                    st.write(f"**{col}:** {row[col]}")
            # 添付ファイル列の自動検出
            url_col = next((c for c in row.index if 'URL' in c), None)
            name_col = next((c for c in row.index if 'ファイル名' in c), None)
            if url_col:
                display_attached_files(row, url_col, name_col)

def display_gcs_browser(folder_type):
    st.subheader("📂 GCS ファイルブラウザ")
    files = get_note_files_from_gcs(folder_type)
    if not files:
        st.info("ファイルが見つかりません。")
        return
        
    sel_name = st.selectbox("ファイルを選択", [f[0] for f in files])
    if sel_name:
        sel_file = next(f for f in files if f[0] == sel_name)
        # 署名付きURL生成
        signed_url = generate_signed_url(sel_file[1])
        if signed_url:
            st.success(f"ファイル名: {sel_file[0]}")
            st.markdown(f"[ダウンロード/表示]({signed_url}) (リンクは一時的に有効です)")

# ---------------------------
# --- Page: Epi Note ---
# ---------------------------
def page_epi_note():
    st.header("エピノート")
    tab1, tab2, tab3 = st.tabs(["📝 記録", "📚 一覧", "📂 ファイル閲覧"])
    
    with tab1:
        with st.form("epi_form"):
            category = st.selectbox("カテゴリ", ["D1", "D2", "その他"])
            title = st.text_input("タイトル/番号 (例: 791)")
            memo = st.text_area("メモ")
            files = st.file_uploader("添付", accept_multiple_files=True)
            if st.form_submit_button("保存"):
                if not title:
                    st.error("タイトルは必須です")
                else:
                    file_names, file_urls = [], []
                    if files:
                        for f in files:
                            n, u = upload_file_to_gcs(f)
                            if u: file_names.append(n); file_urls.append(u)
                    
                    row = [
                        datetime.now().strftime("%Y%m%d_%H%M%S"),
                        "エピノート", category, f"{title}\n{memo}",
                        json.dumps(file_names), json.dumps(file_urls)
                    ]
                    try:
                        gc.open(SPREADSHEET_NAME).worksheet(SHEET_EPI_DATA).append_row(row)
                        st.success("保存しました")
                        get_sheet_as_df.clear() # キャッシュクリア
                    except Exception as e:
                        st.error(f"保存エラー: {e}")

    with tab2:
        page_data_list_view(SHEET_EPI_DATA, "エピノート", 'タイムスタンプ', 'カテゴリ', 'メモ', 
                            ['タイムスタンプ', 'カテゴリ', 'メモ', 'ファイル名'])
    
    with tab3:
        display_gcs_browser("ep_notes")

# ---------------------------
# --- Page: Mainte Note ---
# ---------------------------
def page_mainte_note():
    st.header("メンテノート")
    tab1, tab2, tab3 = st.tabs(["📝 記録", "📚 一覧", "📂 ファイル閲覧"])
    
    with tab1:
        with st.form("mainte_form"):
            device = st.selectbox("装置", ["MBE", "XRD", "PL", "AFM", "その他"])
            title = st.text_input("作業タイトル")
            memo = st.text_area("詳細")
            files = st.file_uploader("添付", accept_multiple_files=True)
            if st.form_submit_button("保存"):
                if not title:
                    st.error("タイトルは必須です")
                else:
                    file_names, file_urls = [], []
                    if files:
                        for f in files:
                            n, u = upload_file_to_gcs(f)
                            if u: file_names.append(n); file_urls.append(u)
                    
                    row = [
                        datetime.now().strftime("%Y%m%d_%H%M%S"),
                        "メンテノート", f"[{title}] {device}\n{memo}",
                        json.dumps(file_names), json.dumps(file_urls)
                    ]
                    try:
                        gc.open(SPREADSHEET_NAME).worksheet(SHEET_MAINTE_DATA).append_row(row)
                        st.success("保存しました")
                        get_sheet_as_df.clear()
                    except Exception as e:
                        st.error(f"保存エラー: {e}")

    with tab2:
        page_data_list_view(SHEET_MAINTE_DATA, "メンテノート", 'タイムスタンプ', None, 'メモ', 
                            ['タイムスタンプ', 'メモ', 'ファイル名'])
    
    with tab3:
        display_gcs_browser("mainte_notes")

# ---------------------------
# --- Page: Meeting Note ---
# ---------------------------
def page_meeting_note():
    st.header("議事録")
    tab1, tab2 = st.tabs(["📝 記録", "📚 一覧"])
    
    with tab1:
        with st.form("meeting_form"):
            title = st.text_input("会議タイトル (例: 2025-10-28 定例)")
            content = st.text_area("内容")
            audio_url = st.text_input("音声ファイルURL (Google Drive等)")
            if st.form_submit_button("保存"):
                if not title:
                    st.error("タイトルは必須です")
                else:
                    row = [
                        datetime.now().strftime("%Y%m%d_%H%M%S"),
                        title, "", audio_url, content
                    ]
                    try:
                        gc.open(SPREADSHEET_NAME).worksheet(SHEET_MEETING_DATA).append_row(row)
                        st.success("保存しました")
                        get_sheet_as_df.clear()
                    except Exception as e:
                        st.error(f"保存エラー: {e}")

    with tab2:
        page_data_list_view(SHEET_MEETING_DATA, "議事録", 'タイムスタンプ', None, '会議タイトル', 
                            ['タイムスタンプ', '会議タイトル', '議事録内容', '音声ファイルURL'])

# ---------------------------
# --- Page: QA Box ---
# ---------------------------
def page_qa_box():
    st.header("知恵袋・質問箱")
    tab1, tab2 = st.tabs(["💡 質問投稿", "📚 質問一覧"])
    
    with tab1:
        with st.form("qa_form"):
            title = st.text_input("質問タイトル")
            content = st.text_area("内容")
            contact = st.text_input("連絡先 (任意)")
            files = st.file_uploader("添付", accept_multiple_files=True)
            if st.form_submit_button("送信"):
                if not title:
                    st.error("タイトルは必須です")
                else:
                    file_names, file_urls = [], []
                    if files:
                        for f in files:
                            n, u = upload_file_to_gcs(f)
                            if u: file_names.append(n); file_urls.append(u)
                    
                    row = [
                        datetime.now().strftime("%Y%m%d_%H%M%S"),
                        title, content, contact,
                        json.dumps(file_names), json.dumps(file_urls), "未解決"
                    ]
                    try:
                        gc.open(SPREADSHEET_NAME).worksheet(SHEET_QA_DATA).append_row(row)
                        st.success("送信しました")
                        get_sheet_as_df.clear()
                    except Exception as e:
                        st.error(f"送信エラー: {e}")
    with tab2:
        page_data_list_view(SHEET_QA_DATA, "質問リスト", 'タイムスタンプ', 'ステータス', '質問タイトル',
                            ['タイムスタンプ', '質問タイトル', '質問内容', 'ステータス', '連絡先'])

# ---------------------------
# --- Page: Handover & Trouble & Contact ---
# ---------------------------
def page_handover_note():
    st.header("引き継ぎメモ")
    tab1, tab2 = st.tabs(["📝 記録", "📚 一覧"])
    with tab1:
        with st.form("handover_form"):
            htype = st.selectbox("種類", ["マニュアル", "装置設定", "その他"])
            title = st.text_input("タイトル")
            memo = st.text_area("内容")
            if st.form_submit_button("保存"):
                if not title: st.error("タイトル必須"); return
                try:
                    gc.open(SPREADSHEET_NAME).worksheet(SHEET_HANDOVER_DATA).append_row([
                        datetime.now().strftime("%Y%m%d_%H%M%S"), htype, title, memo
                    ])
                    st.success("保存しました"); get_sheet_as_df.clear()
                except Exception as e: st.error(f"エラー: {e}")
    with tab2:
        page_data_list_view(SHEET_HANDOVER_DATA, "引き継ぎ", 'タイムスタンプ', '種類', 'タイトル', 
                            ['タイムスタンプ', '種類', 'タイトル', 'メモ'])

def page_trouble_report():
    st.header("トラブル報告")
    tab1, tab2 = st.tabs(["🚨 報告", "📚 履歴"])
    with tab1:
        with st.form("trouble_form"):
            device = st.selectbox("機器", ["MBE", "XRD", "PL", "その他"])
            title = st.text_input("件名")
            cause = st.text_area("原因/現象")
            solution = st.text_area("対策/復旧")
            reporter = st.text_input("報告者")
            if st.form_submit_button("保存"):
                try:
                    gc.open(SPREADSHEET_NAME).worksheet(SHEET_TROUBLE_DATA).append_row([
                        datetime.now().strftime("%Y%m%d_%H%M%S"), device, "", "",
                        cause, solution, "", reporter, "", "", title
                    ])
                    st.success("保存しました"); get_sheet_as_df.clear()
                except Exception as e: st.error(f"エラー: {e}")
    with tab2:
        page_data_list_view(SHEET_TROUBLE_DATA, "トラブル", 'タイムスタンプ', '機器/場所', '件名/タイトル',
                            ['タイムスタンプ', '機器/場所', '件名/タイトル', '原因/究明', '対策/復旧'])

def page_contact_form():
    st.header("お問い合わせ")
    with st.form("contact_form"):
        ctype = st.selectbox("種類", ["バグ報告", "要望", "その他"])
        detail = st.text_area("詳細")
        contact = st.text_input("連絡先")
        if st.form_submit_button("送信"):
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_CONTACT_DATA).append_row([
                    datetime.now().strftime("%Y%m%d_%H%M%S"), ctype, detail, contact
                ])
                st.success("送信しました")
            except Exception as e: st.error(f"エラー: {e}")

# ---------------------------
# --- Page: Analysis (IV / PL) ---
# ---------------------------
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    files = st.file_uploader("IVデータファイル (txt)", accept_multiple_files=True)
    if files:
        dfs = []
        names = []
        for f in files:
            df = load_iv_data(f)
            if df is not None:
                dfs.append(df)
                names.append(f.name)
        
        if dfs:
            fig, ax = plt.subplots()
            for df, name in zip(dfs, names):
                ax.plot(df['Axis_X'], df.iloc[:,1], label=name)
            ax.set_xlabel("Voltage (V)")
            ax.set_ylabel("Current (A)")
            ax.legend()
            st.pyplot(fig)

def page_pl_analysis():
    st.header("PLデータ解析")
    
    # Session Stateの初期化
    if 'pl_slope' not in st.session_state: st.session_state['pl_slope'] = None

    tab1, tab2 = st.tabs(["Step 1: 波長校正", "Step 2: データプロット"])

    # --- Step 1: Calibration ---
    with tab1:
        st.info("2つの既知の波長のピーク位置から校正係数(nm/pixel)を算出します。")
        col1, col2 = st.columns(2)
        wl1 = col1.number_input("波長1 (nm)", value=546.1)
        wl2 = col2.number_input("波長2 (nm)", value=577.0)
        
        f1 = col1.file_uploader("波長1のデータ", key="cal1")
        f2 = col2.file_uploader("波長2のデータ", key="cal2")

        if f1 and f2:
            df1 = load_pl_data(f1)
            df2 = load_pl_data(f2)
            
            if df1 is not None and df2 is not None:
                p1 = df1.loc[df1['intensity'].idxmax(), 'pixel']
                p2 = df2.loc[df2['intensity'].idxmax(), 'pixel']
                
                if p1 != p2:
                    slope = (wl2 - wl1) / (p2 - p1)
                    st.success(f"校正係数: {slope:.4f} nm/pixel")
                    if st.button("この係数を保存して次へ"):
                        st.session_state['pl_slope'] = slope
                        st.session_state['pl_cal_base_wl'] = wl1
                        st.session_state['pl_cal_base_px'] = p1
                else:
                    st.error("ピーク位置が同じです。")

    # --- Step 2: Analysis ---
    with tab2:
        if st.session_state['pl_slope'] is None:
            st.warning("Step 1 で校正を行ってください。")
        else:
            slope = st.session_state['pl_slope']
            base_wl = st.session_state.get('pl_cal_base_wl', 546.1)
            base_px = st.session_state.get('pl_cal_base_px', 0)
            
            st.write(f"現在の校正係数: `{slope:.4f}` nm/pixel")
            
            center_wl = st.number_input("測定中心波長 (nm)", value=1700)
            # 中心ピクセル（通常はCCDの中央、例: 256 or 512）
            # ここでは簡易的に、校正時の基準を用いるか、固定値(256.5など)を使用するか選択
            # 既存コードに合わせて補正ロジックを適用
            
            files = st.file_uploader("測定データ", accept_multiple_files=True, key="pl_meas")
            if files:
                fig, ax = plt.subplots()
                for f in files:
                    df = load_pl_data(f)
                    if df is not None:
                        # 波長変換: (pixel - center_pixel_of_detector) * slope + center_wavelength
                        # ただし、簡易校正の場合は (pixel - base_px) * slope + base_wl のオフセットを使うこともある
                        # ここでは元のコードのロジック「(df['pixel'] - 256.5) * slope + center_wavelength」を採用
                        center_pixel_const = 256.5 
                        df['wavelength'] = (df['pixel'] - center_pixel_const) * slope + center_wl
                        
                        ax.plot(df['wavelength'], df['intensity'], label=f.name)
                
                ax.set_xlabel("Wavelength (nm)")
                ax.set_ylabel("Intensity")
                ax.legend()
                st.pyplot(fig)

# ---------------------------
# --- Page: Calendar ---
# ---------------------------
def page_calendar():
    st.header("🗓️ スケジュール・装置予約")
    
    # Embed Calendar
    src = CALENDAR_ID.replace("@", "%40")
    st.markdown(f"""
    <iframe src="https://calendar.google.com/calendar/embed?height=600&wkst=1&bgcolor=%23ffffff&ctz=Asia%2FTokyo&src={src}&color=%237986CB" style="border:solid 1px #777" width="100%" height="600" frameborder="0" scrolling="no"></iframe>
    """, unsafe_allow_html=True)
    
    # Reservation Form
    with st.expander("➕ 新しい予定を追加"):
        with st.form("cal_form"):
            summary = st.text_input("予定タイトル")
            start_d = st.date_input("開始日")
            start_t = st.time_input("開始時刻")
            end_t = st.time_input("終了時刻")
            desc = st.text_area("詳細")
            
            if st.form_submit_button("予約登録"):
                if calendar_service:
                    start_dt = datetime.combine(start_d, start_t).isoformat()
                    end_dt = datetime.combine(start_d, end_t).isoformat()
                    
                    event = {
                        'summary': summary,
                        'description': desc,
                        'start': {'dateTime': start_dt, 'timeZone': 'Asia/Tokyo'},
                        'end': {'dateTime': end_dt, 'timeZone': 'Asia/Tokyo'},
                    }
                    try:
                        calendar_service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
                        st.success("予約を追加しました！")
                        st.rerun()
                    except Exception as e:
                        st.error(f"登録エラー: {e}")
                else:
                    st.error("カレンダー機能は無効です（Secrets設定を確認してください）。")

# ---------------------------
# --- Main App & Router ---
# ---------------------------
def main():
    st.sidebar.title("Yamane Lab Tools")
    
    menu = st.sidebar.radio("メニュー", [
        "エピノート",
        "メンテノート",
        "🗓️ スケジュール・装置予約",
        "IVデータ解析",
        "PLデータ解析",
        "議事録",
        "知恵袋・質問箱",
        "引き継ぎメモ",
        "トラブル報告",
        "お問い合わせ"
    ])

    # セッション状態によるキャッシュクリア制御
    if 'current_menu' not in st.session_state:
        st.session_state['current_menu'] = menu
    
    if st.session_state['current_menu'] != menu:
        # メニュー切り替え時にデータをリフレッシュしたい場合
        get_sheet_as_df.clear()
        st.session_state['current_menu'] = menu

    # ルーティング
    if menu == "エピノート":
        page_epi_note()
    elif menu == "メンテノート":
        page_mainte_note()
    elif menu == "🗓️ スケジュール・装置予約":
        page_calendar()
    elif menu == "IVデータ解析":
        page_iv_analysis()
    elif menu == "PLデータ解析":
        page_pl_analysis()
    elif menu == "議事録":
        page_meeting_note()
    elif menu == "知恵袋・質問箱":
        page_qa_box()
    elif menu == "引き継ぎメモ":
        page_handover_note()
    elif menu == "トラブル報告":
        page_trouble_report()
    elif menu == "お問い合わせ":
        page_contact_form()

if __name__ == "__main__":
    main()
    

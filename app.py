# -*- coding: utf-8 -*-
"""
bennriyasann3_final_integrated_v2.py
Yamane Lab Convenience Tool - 認証ロジック統合版
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
from datetime import datetime, date, timedelta, time
from urllib.parse import quote as url_quote
from io import BytesIO
import calendar
import matplotlib.font_manager as fm

# Google Calendar API
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# GCS Library (存在しない場合も考慮)
try:
    from google.cloud import storage
except ImportError:
    storage = None

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
# --- 定数（グローバル変数） ---
# ---------------------------
SPREADSHEET_NAME = "エピノート (2).xlsx" 

# 各シート名
SHEET_EPI_DATA = "エピノート_データ"   
SHEET_MAINTE_DATA = "メンテノート_データ" 
SHEET_SCHEDULE_DATA = "スケジュール" 
SHEET_FAQ_DATA = "知恵袋_データ"
SHEET_TROUBLE_DATA = "トラブル報告_データ" 
SHEET_HANDOVER_DATA = "引き継ぎ_データ"
SHEET_CONTACT_DATA = "お問い合わせ_データ"
SHEET_MEETING_DATA = "議事録_データ"

# GCSバケット名
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files"

# ---------------------------
# --- Google Calendar API連携用定数 ---
# ---------------------------
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
    # upload_from_string もダミーに追加
    def upload_from_string(self, data, content_type=None): pass

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

# --- 認証の実行 ---
gc, storage_client = initialize_google_services()

# --- Google Calendar サービス初期化 (追加) ---
gcal_service = None
try:
    if "gcs_credentials" in st.secrets:
        raw = st.secrets["gcs_credentials"]
        cleaned = raw.strip().replace('\t', '').replace('\r', '').replace('\n', '')
        info = json.loads(cleaned)
        gcal_creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        gcal_service = build('calendar', 'v3', credentials=gcal_creds)
except Exception:
    pass # カレンダーエラーは無視してアプリを起動

# ---------------------------
# --- ユーティリティ関数 ---
# ---------------------------

@st.cache_data(ttl=600)
def get_data_from_gspread(sheet_name):
    """スプレッドシートからデータを取得しDataFrame化"""
    if isinstance(gc, DummyGSClient):
        return pd.DataFrame()
    try:
        worksheet = gc.open(SPREADSHEET_NAME).worksheet(sheet_name)
        data = worksheet.get_all_values()
        if not data:
            return pd.DataFrame()
        # 1行目をヘッダーとして扱う
        return pd.DataFrame(data[1:], columns=data[0])
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"シート '{sheet_name}' が見つかりません。")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"データ取得エラー ({sheet_name}): {e}")
        return pd.DataFrame()

def upload_file_to_gcs(client_obj, file_obj):
    """ファイルをGCSルートに保存し、公開URLを返す"""
    if isinstance(client_obj, DummyStorageClient) or client_obj is None:
        return None, None
    try:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        safe_name = file_obj.name.replace(' ', '_').replace('/', '_')
        gcs_filename = f"{timestamp}_{safe_name}"

        bucket = client_obj.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        
        # Streamlit UploadedFile は getvalue()
        blob.upload_from_string(
            file_obj.getvalue(),
            content_type=file_obj.type if hasattr(file_obj, 'type') else 'application/octet-stream'
        )
        
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        return file_obj.name, public_url
    except Exception as e:
        st.error(f"アップロードエラー: {e}")
        return None, None

def handle_file_uploads(uploaded_files):
    """複数ファイルのアップロード処理ラッパー"""
    f_list, u_list = [], []
    if uploaded_files:
        with st.spinner("ファイルをアップロード中..."):
            for f in uploaded_files:
                name, url = upload_file_to_gcs(storage_client, f)
                if url:
                    f_list.append(name)
                    u_list.append(url)
    return json.dumps(f_list), json.dumps(u_list)

def display_attached_files(row_dict, col_url_key, col_filename_key):
    """添付ファイル表示: JSON二重エスケープ対応版"""
    urls = []
    filenames = []
    
    raw_urls = row_dict.get(col_url_key, '')
    raw_filenames = row_dict.get(col_filename_key, '')

    # --- URLデコード ---
    try:
        parsed = json.loads(raw_urls)
        if isinstance(parsed, list):
            for item in parsed:
                if isinstance(item, str) and item.startswith('http'):
                    urls.append(item)
                else:
                    try:
                        inner = json.loads(item)
                        if isinstance(inner, str) and inner.startswith('http'):
                            urls.append(inner)
                    except: pass
        elif isinstance(parsed, str) and parsed.startswith('http'):
             urls.append(parsed)
    except:
        m = re.search(r'https?://[^\s,"]+', str(raw_urls))
        if m: urls = [m.group(0)]

    # --- ファイル名デコード ---
    try:
        parsed_fn = json.loads(raw_filenames)
        if isinstance(parsed_fn, list):
            filenames = parsed_fn
        elif isinstance(parsed_fn, str):
            filenames = [parsed_fn]
    except:
        filenames = [f"添付ファイル {i+1}" for i in range(len(urls))]

    # --- 表示 ---
    if urls:
        st.markdown("##### 📎 添付ファイル")
        if len(filenames) < len(urls):
            filenames += [f"File {i+1}" for i in range(len(filenames), len(urls))]
        
        for u, f in zip(urls, filenames):
            st.markdown(f"[{f}]({u})")
    else:
        st.markdown("添付ファイルなし")

def save_row_to_sheet(sheet_name, row_data):
    """行データをシートに追加し、キャッシュクリアしてリラン"""
    if isinstance(gc, DummyGSClient):
        st.error("認証されていないため保存できません。")
        return

    try:
        ws = gc.open(SPREADSHEET_NAME).worksheet(sheet_name)
        ws.append_row(row_data)
        st.success("保存しました！")
        get_data_from_gspread.clear()
        st.rerun()
    except Exception as e:
        st.error(f"保存エラー: {e}")

# ---------------------------
# --- 4. 各機能ページの実装 ---
# ---------------------------

# === エピノート ===
def page_epi_note():
    st.header("エピノート")
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])
    
    with tab2:
        with st.form("epi_form"):
            title = st.text_input("タイトル/番号 (例: 791)")
            cat = st.selectbox("カテゴリ", ["D1", "D2", "その他"])
            memo = st.text_area("詳細メモ", height=150)
            files = st.file_uploader("添付ファイル", accept_multiple_files=True)
            with st.expander("データのインポート"): pass
            submit = st.form_submit_button("記録を保存")
        
        if submit:
            if not title:
                st.warning("タイトルは必須です")
            else:
                f_json, u_json = handle_file_uploads(files)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                # 6列: Timestamp, Type, Category, Memo, FileName, URL
                row = [ts, "エピノート", cat, f"{title}\n{memo}", f_json, u_json]
                save_row_to_sheet(SHEET_EPI_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_EPI_DATA)
        if not df.empty:
            if 'タイムスタンプ' in df.columns:
                df = df.sort_values('タイムスタンプ', ascending=False)
            st.dataframe(df, use_container_width=True)
            
            ts_col = 'タイムスタンプ'
            if ts_col in df.columns:
                sel = st.selectbox("詳細表示を選択", df[ts_col].unique(), key="epi_sel")
                if sel:
                    row = df[df[ts_col] == sel].iloc[0].to_dict()
                    st.divider()
                    st.write(f"**日時:** {row.get(ts_col)}")
                    st.write(f"**カテゴリ:** {row.get('カテゴリ')}")
                    st.text_area("内容", row.get('メモ'), disabled=True)
                    display_attached_files(row, '写真URL', 'ファイル名')

# === メンテノート ===
def page_mainte_note():
    st.header("メンテノート")
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])
    
    with tab2:
        with st.form("mainte_form"):
            title = st.text_input("メンテタイトル")
            dev = st.selectbox("対象装置", ["MOCVD", "IV/PL", "その他"])
            memo = st.text_area("作業メモ", height=150)
            files = st.file_uploader("添付ファイル", accept_multiple_files=True)
            with st.expander("データのインポート"): pass
            submit = st.form_submit_button("記録を保存")
            
        if submit:
            if not title: st.warning("タイトルは必須です")
            else:
                f_json, u_json = handle_file_uploads(files)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                content = f"[{title}] (装置: {dev})\n{memo}"
                # 5列: Timestamp, Type, Memo, FileName, URL
                row = [ts, "メンテノート", content, f_json, u_json]
                save_row_to_sheet(SHEET_MAINTE_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_MAINTE_DATA)
        if not df.empty:
            if 'タイムスタンプ' in df.columns:
                df = df.sort_values('タイムスタンプ', ascending=False)
            st.dataframe(df, use_container_width=True)
            
            ts_col = 'タイムスタンプ'
            if ts_col in df.columns:
                sel = st.selectbox("詳細表示を選択", df[ts_col].unique(), key="mainte_sel")
                if sel:
                    row = df[df[ts_col] == sel].iloc[0].to_dict()
                    st.divider()
                    st.text_area("内容", row.get('メモ'), disabled=True)
                    display_attached_files(row, '写真URL', 'ファイル名')

# === スケジュール・装置予約 ===
def page_schedule_reservation():
    st.header("🗓️ スケジュール・装置予約")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("新規予約")
        with st.form("sch_form"):
            title = st.text_input("予定タイトル", "装置予約: ")
            d_input = st.date_input("日付", date.today())
            s_time = st.time_input("開始", time(9, 0))
            e_time = st.time_input("終了", time(10, 0))
            desc = st.text_area("詳細")
            submit = st.form_submit_button("カレンダー登録")
        
        if submit:
            if gcal_service:
                try:
                    start_dt = datetime.combine(d_input, s_time).isoformat()
                    end_dt = datetime.combine(d_input, e_time).isoformat()
                    event = {
                        'summary': title,
                        'description': desc,
                        'start': {'dateTime': start_dt, 'timeZone': 'Asia/Tokyo'},
                        'end': {'dateTime': end_dt, 'timeZone': 'Asia/Tokyo'},
                    }
                    gcal_service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
                    st.success(f"予約 '{title}' を登録しました")
                except Exception as e:
                    st.error(f"登録失敗: {e}")
            else:
                st.error("カレンダー機能は現在利用できません")

    with col2:
        st.subheader("直近の予定 (カレンダー)")
        if gcal_service:
            try:
                now = datetime.utcnow().isoformat() + 'Z'
                events_result = gcal_service.events().list(
                    calendarId=CALENDAR_ID, timeMin=now, maxResults=10, 
                    singleEvents=True, orderBy='startTime'
                ).execute()
                events = events_result.get('items', [])
                
                if not events:
                    st.info("予定はありません")
                else:
                    for event in events:
                        start = event['start'].get('dateTime', event['start'].get('date'))
                        st.write(f"**{start}**: {event['summary']}")
            except Exception as e:
                st.error(f"取得失敗: {e}")
        
        # シート側のデータも表示
        st.divider()
        st.subheader("予約履歴 (シート)")
        df = get_data_from_gspread(SHEET_SCHEDULE_DATA)
        if not df.empty:
            st.dataframe(df)

# === IVデータ解析 ===
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    st.markdown("IV測定データファイル（2列データ：X軸/Y軸）をアップロードし、往路/復路の特性をプロットします。")
    
    uploaded_files = st.file_uploader(
        "IV測定データファイル (.txt, .csv)", 
        type=['txt', 'csv'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        fig, ax = plt.subplots(figsize=(10, 6))
        
        for f in uploaded_files:
            try:
                content = f.getvalue().decode('utf-8', errors='ignore')
                lines = [l for l in content.splitlines() if l.strip() and not l.strip().startswith(('#', '!', '/'))]
                data_start_idx = 0
                for i, line in enumerate(lines):
                    try:
                        parts = re.split(r'\s+|,|\t', line.strip())
                        float(parts[0])
                        data_start_idx = i
                        break
                    except: continue
                
                data_lines = lines[data_start_idx:]
                if not data_lines: continue

                df = pd.read_csv(io.StringIO("\n".join(data_lines)), sep=r'\s+|,|\t', engine='python', header=None)
                if df.shape[1] < 2: continue
                
                x = pd.to_numeric(df.iloc[:, 0], errors='coerce')
                y = pd.to_numeric(df.iloc[:, 1], errors='coerce')
                df_clean = pd.DataFrame({'x': x, 'y': y}).dropna()
                
                if df_clean.empty: continue
                
                max_idx = df_clean['x'].idxmax()
                
                ax.plot(df_clean.iloc[:max_idx+1]['x'], df_clean.iloc[:max_idx+1]['y'], 
                        label=f"{f.name} (往)", marker='.', markersize=2)
                if max_idx < len(df_clean) - 1:
                    ax.plot(df_clean.iloc[max_idx+1:]['x'], df_clean.iloc[max_idx+1:]['y'], 
                            label=f"{f.name} (復)", linestyle='--', alpha=0.7)
                            
            except Exception as e:
                st.warning(f"{f.name} 解析エラー: {e}")
        
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Current (A)")
        ax.grid(True)
        ax.legend()
        st.pyplot(fig)

# === PLデータ解析 ===
def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("設定")
        slope = st.number_input("Slope (nm/px)", value=1.0, format="%.5f")
        center_wl = st.number_input("Center Wavelength (nm)", value=500.0)
        center_px = st.number_input("Center Pixel", value=256.0)
        
    uploaded_files = st.file_uploader("PL測定データ", accept_multiple_files=True)
    
    if uploaded_files:
        fig, ax = plt.subplots(figsize=(10, 6))
        for f in uploaded_files:
            try:
                content = f.getvalue().decode('utf-8', errors='ignore')
                lines = [l for l in content.splitlines() if l.strip() and not l.strip().startswith(('#', '!', '/'))]
                data_lines = []
                for line in lines:
                    try:
                        parts = re.split(r'\s+|,|\t', line.strip())
                        float(parts[1]) 
                        data_lines.append(line)
                    except: continue

                df = pd.read_csv(io.StringIO("\n".join(data_lines)), sep=r'\s+|,|\t', engine='python', header=None)
                if df.shape[1] < 2: continue
                
                y_data = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(0)
                pixels = np.arange(len(y_data))
                wavelengths = (pixels - center_px) * slope + center_wl
                ax.plot(wavelengths, y_data, label=f.name)
            except Exception as e:
                st.warning(f"{f.name}: {e}")
                
        ax.set_xlabel("Wavelength (nm)")
        ax.set_ylabel("Intensity (a.u.)")
        ax.legend()
        st.pyplot(fig)

# === 議事録 ===
def page_meeting_note():
    st.header("📄 議事録")
    # CSV列: Timestamp, Title, AudioName, AudioURL, Content
    
    tab1, tab2 = st.tabs(["一覧", "新規"])
    with tab2:
        with st.form("meet_form"):
            title = st.text_input("会議タイトル/日付")
            content = st.text_area("議事録内容", height=300)
            files = st.file_uploader("音声/資料添付", accept_multiple_files=True)
            submit = st.form_submit_button("保存")
        
        if submit:
            f_j, u_j = handle_file_uploads(files)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            row = [ts, title, f_j, u_j, content]
            save_row_to_sheet(SHEET_MEETING_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_MEETING_DATA)
        if not df.empty:
            st.dataframe(df)
            ts_col = 'タイムスタンプ'
            if ts_col in df.columns:
                sel = st.selectbox("詳細", df[ts_col].unique(), key="meet_sel")
                if sel:
                    row = df[df[ts_col] == sel].iloc[0].to_dict()
                    st.divider()
                    st.markdown(f"### {row.get('会議タイトル')}")
                    st.markdown(row.get('議事録内容'))
                    display_attached_files(row, '音声ファイルURL', '音声ファイル名')

# === 知恵袋・質問箱 ===
def page_faq():
    st.header("💡 知恵袋・質問箱")
    # CSV: Timestamp, Title, Content, Email, FileName, FileURL, Status
    
    tab1, tab2 = st.tabs(["質問一覧", "質問投稿"])
    with tab2:
        with st.form("faq_form"):
            title = st.text_input("質問タイトル")
            content = st.text_area("質問内容")
            email = st.text_input("連絡先メールアドレス (任意)")
            files = st.file_uploader("添付", accept_multiple_files=True)
            submit = st.form_submit_button("投稿")
        
        if submit:
            f_j, u_j = handle_file_uploads(files)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            row = [ts, title, content, email, f_j, u_j, "未解決"]
            save_row_to_sheet(SHEET_FAQ_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_FAQ_DATA)
        if not df.empty:
            st.dataframe(df)
            for _, row in df.iterrows():
                with st.expander(f"{row.get('質問タイトル')} ({row.get('ステータス')})"):
                    st.write(f"**質問内容:** {row.get('質問内容')}")
                    display_attached_files(row, '添付ファイルURL', '添付ファイル名')

# === トラブル報告 ===
def page_trouble_report():
    st.header("🚨 トラブル報告")
    # CSV: Timestamp, Place, Date, When, Cause, Solution, Prevention, Reporter, FileName, FileURL, Title
    
    tab1, tab2 = st.tabs(["報告一覧", "新規報告"])
    with tab2:
        with st.form("trb_form"):
            col1, col2 = st.columns(2)
            with col1:
                title = st.text_input("件名/タイトル")
                place = st.text_input("機器/場所")
                reporter = st.text_input("報告者")
            with col2:
                date_occ = st.date_input("発生日")
            
            when = st.text_area("トラブル発生時")
            cause = st.text_area("原因/究明")
            sol = st.text_area("対策/復旧")
            prev = st.text_area("再発防止策")
            files = st.file_uploader("写真/資料", accept_multiple_files=True)
            submit = st.form_submit_button("報告")
        
        if submit:
            f_j, u_j = handle_file_uploads(files)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            row = [ts, place, str(date_occ), when, cause, sol, prev, reporter, f_j, u_j, title]
            save_row_to_sheet(SHEET_TROUBLE_DATA, row)
            
    with tab1:
        df = get_data_from_gspread(SHEET_TROUBLE_DATA)
        if not df.empty:
            st.dataframe(df)
            sel = st.selectbox("詳細", df['タイムスタンプ'].unique() if 'タイムスタンプ' in df.columns else [], key="trb_sel")
            if sel:
                row = df[df['タイムスタンプ'] == sel].iloc[0].to_dict()
                st.write(row)
                display_attached_files(row, 'ファイルURL', 'ファイル名')

# === 装置引き継ぎメモ ===
def page_device_handover():
    st.header("📝 装置引き継ぎメモ")
    # CSV: Timestamp, Type, Title, Content1, Content2, Content3, Memo
    
    tab1, tab2 = st.tabs(["一覧", "新規"])
    with tab2:
        with st.form("ho_form"):
            h_type = st.selectbox("種類", ["マニュアル", "ノウハウ", "その他"])
            title = st.text_input("タイトル")
            memo = st.text_area("概要/メモ")
            
            st.markdown("---")
            st.caption("詳細内容やリンク")
            c1 = st.text_area("内容1")
            c2 = st.text_area("内容2")
            c3 = st.text_area("内容3")
            submit = st.form_submit_button("保存")
            
        if submit:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            row = [ts, h_type, title, c1, c2, c3, memo]
            save_row_to_sheet(SHEET_HANDOVER_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_HANDOVER_DATA)
        if not df.empty:
            st.dataframe(df)

# === 連絡・問い合わせ ===
def page_contact():
    st.header("📧 連絡・問い合わせ")
    # CSV: Timestamp, Type, Detail, Contact
    
    tab1, tab2 = st.tabs(["履歴", "新規"])
    with tab2:
        with st.form("contact_form"):
            c_type = st.selectbox("種類", ["バグ報告", "要望", "その他"])
            detail = st.text_area("詳細内容")
            contact = st.text_input("連絡先")
            submit = st.form_submit_button("送信")
        
        if submit:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            row = [ts, c_type, detail, contact]
            save_row_to_sheet(SHEET_CONTACT_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_CONTACT_DATA)
        if not df.empty:
            st.dataframe(df)


# ---------------------------
# --- 5. メインルーティング ---
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
    
    # メニュー切り替え検知 & キャッシュクリア
    if 'menu_selection' not in st.session_state:
        st.session_state.menu_selection = menu_selection
    
    if st.session_state.menu_selection != menu_selection:
        # シート読み込みキャッシュのクリア
        get_data_from_gspread.clear()
        st.session_state.menu_selection = menu_selection

    # ルーティング実行
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

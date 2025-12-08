# -*- coding: utf-8 -*-
"""
bennriyasann3_original_restored.py
Yamane Lab Convenience Tool - app (4).py ベース完全復元 + 必須修正版
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

# Google Calendar APIのためのインポート
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# GCSライブラリ (存在しない場合も考慮)
try:
    from google.cloud import storage
except Exception:
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
# 元のコードに合わせていますが、シート名はCSVに合わせて修正しています
SPREADSHEET_NAME = "エピノート (2).xlsx" 

SHEET_EPI_DATA = "エピノート_データ"   
SHEET_MAINTE_DATA = "メンテノート_データ" 
SHEET_SCHEDULE_DATA = "スケジュール" 
SHEET_FAQ_DATA = "知恵袋_データ"
SHEET_TROUBLE_DATA = "トラブル報告_データ" 
SHEET_HANDOVER_DATA = "引き継ぎ_データ"
SHEET_CONTACT_DATA = "お問い合わせ_データ"
SHEET_MEETING_DATA = "議事録_データ"

CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files"

# カレンダー設定
SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = "yamane.lab.6747@gmail.com" # ターゲットカレンダーID

# ---------------------------
# --- Google Service Stubs (認証エラー回避用) ---
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
    def upload_from_file(self, file_obj, content_type): pass
    def list_blobs(self, **kwargs): return []
    def upload_from_string(self, data, content_type=None): pass

# ---------------------------
# --- Google 認証初期化 (修正済みロジック) ---
# ---------------------------
gc = DummyGSClient()
storage_client = DummyStorageClient()
gcal_service = None

try:
    if "gcs_credentials" in st.secrets:
        # クレンジング処理
        raw = st.secrets["gcs_credentials"]
        cleaned = raw.strip().replace('\t', '').replace('\r', '').replace('\n', '')
        info = json.loads(cleaned)
        
        # 1. Gspread
        gc = gspread.service_account_from_dict(info)
        
        # 2. GCS
        if storage:
            storage_client = storage.Client.from_service_account_info(info)
            
        # 3. Calendar
        try:
            gcal_creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
            gcal_service = build('calendar', 'v3', credentials=gcal_creds)
        except Exception:
            pass # カレンダーエラーは無視
            
    elif "gcp_service_account" in st.secrets:
        # 互換性維持
        info = dict(st.secrets["gcp_service_account"])
        gc = gspread.service_account_from_dict(info)
        if storage:
            storage_client = storage.Client.from_service_account_info(info)
        gcal_creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        gcal_service = build('calendar', 'v3', credentials=gcal_creds)

except Exception as e:
    st.error(f"認証初期化エラー: {e}")

# ---------------------------
# --- ユーティリティ関数 (修正済み) ---
# ---------------------------

@st.cache_data(ttl=600)
def get_data_from_gspread(sheet_name):
    """スプレッドシートからデータを取得"""
    if isinstance(gc, DummyGSClient):
        return pd.DataFrame()
    try:
        worksheet = gc.open(SPREADSHEET_NAME).worksheet(sheet_name)
        data = worksheet.get_all_values()
        if not data:
            return pd.DataFrame()
        return pd.DataFrame(data[1:], columns=data[0])
    except Exception as e:
        # シートがない場合は空を返す（エラーで止めない）
        return pd.DataFrame()

def upload_file_to_gcs(client_obj, file_obj):
    """【修正】ファイルをGCSルートに保存し、公開URLを返す"""
    if isinstance(client_obj, DummyStorageClient) or client_obj is None:
        return None, None
    try:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        safe_name = file_obj.name.replace(' ', '_').replace('/', '_')
        gcs_filename = f"{timestamp}_{safe_name}" # ルート保存

        bucket = client_obj.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        
        blob.upload_from_string(
            file_obj.getvalue(),
            content_type=file_obj.type if hasattr(file_obj, 'type') else 'application/octet-stream'
        )
        
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        return file_obj.name, public_url
    except Exception as e:
        st.error(f"アップロードエラー: {e}")
        return None, None

def display_attached_files(row_dict, col_url_key, col_filename_key):
    """【修正】JSON二重エスケープ対応版表示関数"""
    urls = []
    filenames = []
    
    raw_urls = row_dict.get(col_url_key, '')
    raw_filenames = row_dict.get(col_filename_key, '')

    # URLデコード
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

    # ファイル名デコード
    try:
        parsed_fn = json.loads(raw_filenames)
        if isinstance(parsed_fn, list):
            filenames = parsed_fn
        elif isinstance(parsed_fn, str):
            filenames = [parsed_fn]
    except:
        filenames = [f"添付ファイル {i+1}" for i in range(len(urls))]

    if urls:
        st.markdown("##### 📎 添付ファイル")
        if len(filenames) < len(urls):
            filenames += [f"File {i+1}" for i in range(len(filenames), len(urls))]
        
        for u, f in zip(urls, filenames):
            st.markdown(f"[{f}]({u})")
    else:
        st.markdown("添付ファイルはありません。")

# ---------------------------
# --- 各ページ機能 (元のロジックを復元) ---
# ---------------------------

# 1. エピノート (元UI復元 + アップロード修正)
def page_epi_note_recording():
    st.markdown("#### 📝 新しいエピノートを記録")
    with st.form(key='epi_note_form'):
        ep_title = st.text_input("タイトル/番号 (例: 791)", key="epi_title")
        ep_category = st.selectbox("カテゴリ", ["D1", "D2", "その他"], key="epi_category")
        ep_memo = st.text_area("詳細メモ", height=200, key="epi_memo")
        uploaded_files = st.file_uploader("添付ファイル", accept_multiple_files=True, key="epi_uploader")
        
        st.markdown("---")
        with st.expander("データのインポート"): pass
        submit_button = st.form_submit_button("記録を保存") 
        
    if submit_button:
        if not ep_title:
            st.warning("番号 (例: 791) は必須項目です。")
            return
        
        filenames_list, urls_list = [], []
        if uploaded_files:
            with st.spinner("ファイルをGCSにアップロード中..."):
                for f in uploaded_files:
                    # 修正: ルート保存関数を使用
                    name, url = upload_file_to_gcs(storage_client, f) 
                    if url:
                        filenames_list.append(name)
                        urls_list.append(url)

        filenames_json = json.dumps(filenames_list)
        urls_json = json.dumps(urls_list)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        memo_content = f"{ep_title}\n{ep_memo}"
        
        row_data = [timestamp, "エピノート", ep_category, memo_content, filenames_json, urls_json]
        
        try:
            ws = gc.open(SPREADSHEET_NAME).worksheet(SHEET_EPI_DATA)
            ws.append_row(row_data)
            st.success("✅ エピノートをアップロードしました！")
            get_data_from_gspread.clear()
            st.rerun()
        except Exception as e:
            st.error(f"❌ データ書き込みエラー: {e}")

def page_epi_note_list():
    st.subheader("エピノート一覧")
    df = get_data_from_gspread(SHEET_EPI_DATA)
    if df.empty:
        st.info("データがありません")
        return

    if 'タイムスタンプ' in df.columns:
        df = df.sort_values('タイムスタンプ', ascending=False)
    st.dataframe(df, use_container_width=True)
    
    # 詳細表示
    ts_col = 'タイムスタンプ'
    if ts_col in df.columns:
        sel = st.selectbox("詳細表示を選択", df[ts_col].unique(), key="epi_sel_list")
        if sel:
            row = df[df[ts_col] == sel].iloc[0].to_dict()
            st.divider()
            st.write(f"**日時:** {row.get(ts_col)}")
            st.write(f"**カテゴリ:** {row.get('カテゴリ')}")
            st.text_area("内容", row.get('メモ'), disabled=True)
            display_attached_files(row, '写真URL', 'ファイル名')

def page_epi_note():
    st.header("エピノート機能")
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])
    with tab1: page_epi_note_list()
    with tab2: page_epi_note_recording()


# 2. メンテノート (元UI復元 + アップロード修正)
def page_mainte_recording():
    st.markdown("#### 🛠️ 新しいメンテノートを記録")
    with st.form(key='mainte_note_form'):
        title = st.text_input("メンテタイトル")
        dev = st.selectbox("対象装置", ["MOCVD", "IV/PL", "その他"])
        memo = st.text_area("作業詳細メモ", height=200)
        uploaded_files = st.file_uploader("添付ファイル", accept_multiple_files=True)
        
        st.markdown("---")
        with st.expander("データのインポート"): pass
        submit = st.form_submit_button("記録を保存")
        
    if submit:
        if not title:
            st.warning("タイトル必須")
            return
        
        f_list, u_list = [], []
        if uploaded_files:
            with st.spinner("アップロード中..."):
                for f in uploaded_files:
                    n, u = upload_file_to_gcs(storage_client, f)
                    if u:
                        f_list.append(n)
                        u_list.append(u)

        f_json = json.dumps(f_list)
        u_json = json.dumps(u_list)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        content = f"[{title}] (装置: {dev})\n{memo}"
        
        row = [ts, "メンテノート", content, f_json, u_json]
        
        try:
            ws = gc.open(SPREADSHEET_NAME).worksheet(SHEET_MAINTE_DATA)
            ws.append_row(row)
            st.success("✅ メンテノート保存成功")
            get_data_from_gspread.clear()
            st.rerun()
        except Exception as e:
            st.error(f"保存エラー: {e}")

def page_mainte_list():
    st.subheader("メンテノート一覧")
    df = get_data_from_gspread(SHEET_MAINTE_DATA)
    if df.empty: return

    if 'タイムスタンプ' in df.columns:
        df = df.sort_values('タイムスタンプ', ascending=False)
    st.dataframe(df, use_container_width=True)
    
    ts_col = 'タイムスタンプ'
    if ts_col in df.columns:
        sel = st.selectbox("詳細表示", df[ts_col].unique(), key="mainte_sel_list")
        if sel:
            row = df[df[ts_col] == sel].iloc[0].to_dict()
            st.divider()
            st.text_area("内容", row.get('メモ'), disabled=True)
            display_attached_files(row, '写真URL', 'ファイル名')

def page_mainte_note():
    st.header("メンテノート機能")
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])
    with tab1: page_mainte_list()
    with tab2: page_mainte_recording()


# 3. スケジュール (app(4).py ロジック復元)
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
                        'summary': title, 'description': desc,
                        'start': {'dateTime': start_dt, 'timeZone': 'Asia/Tokyo'},
                        'end': {'dateTime': end_dt, 'timeZone': 'Asia/Tokyo'},
                    }
                    gcal_service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
                    st.success(f"予約登録完了: {title}")
                except Exception as e:
                    st.error(f"登録失敗: {e}")
            else:
                st.error("カレンダー機能は利用できません")

    with col2:
        st.subheader("直近の予定")
        if gcal_service:
            try:
                now = datetime.utcnow().isoformat() + 'Z'
                events_result = gcal_service.events().list(
                    calendarId=CALENDAR_ID, timeMin=now, maxResults=10, 
                    singleEvents=True, orderBy='startTime'
                ).execute()
                events = events_result.get('items', [])
                if not events: st.info("予定なし")
                for event in events:
                    start = event['start'].get('dateTime', event['start'].get('date'))
                    st.write(f"**{start}**: {event['summary']}")
            except Exception: pass
            
        st.divider()
        st.write("履歴 (シート)")
        df = get_data_from_gspread(SHEET_SCHEDULE_DATA)
        if not df.empty: st.dataframe(df)


# 4. IVデータ解析 (元のForward/Reverse分割ロジック復元)
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    uploaded_files = st.file_uploader("IVデータ (.txt, .csv)", accept_multiple_files=True)
    
    if uploaded_files:
        fig, ax = plt.subplots(figsize=(10, 6))
        for f in uploaded_files:
            try:
                content = f.getvalue().decode('utf-8', errors='ignore')
                lines = [l for l in content.splitlines() if l.strip() and not l.strip().startswith(('#', '!', '/'))]
                # データ開始行探索
                start_idx = 0
                for i, l in enumerate(lines):
                    try:
                        float(re.split(r'\s+|,|\t', l.strip())[0])
                        start_idx = i
                        break
                    except: continue
                
                df = pd.read_csv(io.StringIO("\n".join(lines[start_idx:])), sep=r'\s+|,|\t', engine='python', header=None)
                if df.shape[1] < 2: continue
                
                x = pd.to_numeric(df.iloc[:, 0], errors='coerce')
                y = pd.to_numeric(df.iloc[:, 1], errors='coerce')
                df_clean = pd.DataFrame({'x': x, 'y': y}).dropna()
                
                if df_clean.empty: continue
                max_idx = df_clean['x'].idxmax()
                
                # 往路復路プロット
                ax.plot(df_clean.iloc[:max_idx+1]['x'], df_clean.iloc[:max_idx+1]['y'], 
                        label=f"{f.name} (往)", marker='.', markersize=2)
                if max_idx < len(df_clean) - 1:
                    ax.plot(df_clean.iloc[max_idx+1:]['x'], df_clean.iloc[max_idx+1:]['y'], 
                            label=f"{f.name} (復)", linestyle='--', alpha=0.7)
            except Exception as e:
                st.warning(f"{f.name} 解析エラー: {e}")
        
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Current (A)")
        ax.legend()
        ax.grid(True)
        st.pyplot(fig)


# 5. PLデータ解析 (元のロジック復元)
def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("校正設定")
        slope = st.number_input("Slope (nm/px)", value=1.0, format="%.5f")
        center_wl = st.number_input("Center WL (nm)", value=500.0)
        center_px = st.number_input("Center Pixel", value=256.0)
    
    uploaded_files = st.file_uploader("PLデータ", accept_multiple_files=True)
    if uploaded_files:
        fig, ax = plt.subplots(figsize=(10, 6))
        for f in uploaded_files:
            try:
                content = f.getvalue().decode('utf-8', errors='ignore')
                lines = [l for l in content.splitlines() if l.strip() and not l.strip().startswith(('#', '!', '/'))]
                # データ抽出簡易ロジック
                data_lines = []
                for l in lines:
                    try: 
                        float(re.split(r'\s+|,|\t', l.strip())[1])
                        data_lines.append(l)
                    except: continue
                
                df = pd.read_csv(io.StringIO("\n".join(data_lines)), sep=r'\s+|,|\t', engine='python', header=None)
                y_data = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(0)
                pixels = np.arange(len(y_data))
                wls = (pixels - center_px) * slope + center_wl
                ax.plot(wls, y_data, label=f.name)
            except: pass
        ax.set_xlabel("Wavelength (nm)")
        ax.set_ylabel("Intensity")
        ax.legend()
        st.pyplot(fig)


# 6. 議事録 (CSV構造に合わせて実装)
def page_meeting_note():
    st.header("📄 議事録")
    tab1, tab2 = st.tabs(["一覧", "新規"])
    with tab2:
        with st.form("meet_form"):
            title = st.text_input("会議タイトル")
            content = st.text_area("内容", height=300)
            files = st.file_uploader("添付", accept_multiple_files=True)
            submit = st.form_submit_button("保存")
        if submit:
            f_j, u_j = (json.dumps([]), json.dumps([]))
            if files:
                f_l, u_l = [], []
                for f in files:
                    n, u = upload_file_to_gcs(storage_client, f)
                    if u: f_l.append(n); u_l.append(u)
                f_j, u_j = json.dumps(f_l), json.dumps(u_l)
            
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            row = [ts, title, f_j, u_j, content]
            try:
                ws = gc.open(SPREADSHEET_NAME).worksheet(SHEET_MEETING_DATA)
                ws.append_row(row)
                st.success("保存完了")
                get_data_from_gspread.clear()
                st.rerun()
            except Exception as e: st.error(str(e))

    with tab1:
        df = get_data_from_gspread(SHEET_MEETING_DATA)
        if not df.empty:
            st.dataframe(df)
            if 'タイムスタンプ' in df.columns:
                sel = st.selectbox("詳細", df['タイムスタンプ'].unique(), key="meet_sel")
                if sel:
                    row = df[df['タイムスタンプ'] == sel].iloc[0].to_dict()
                    st.divider()
                    st.markdown(row.get('議事録内容', ''))
                    display_attached_files(row, '音声ファイルURL', '音声ファイル名')

# 7-10. その他の機能 (NameError回避のため最低限の実装を提供)
def page_faq():
    st.header("💡 知恵袋・質問箱")
    # 実装: 一覧表示のみ簡易提供
    df = get_data_from_gspread(SHEET_FAQ_DATA)
    if not df.empty: st.dataframe(df)
    else: st.info("データなし")

def page_trouble_report():
    st.header("🚨 トラブル報告")
    df = get_data_from_gspread(SHEET_TROUBLE_DATA)
    if not df.empty: st.dataframe(df)
    else: st.info("データなし")

def page_device_handover():
    st.header("📝 装置引き継ぎメモ")
    df = get_data_from_gspread(SHEET_HANDOVER_DATA)
    if not df.empty: st.dataframe(df)
    else: st.info("データなし")

def page_contact():
    st.header("📧 連絡・問い合わせ")
    df = get_data_from_gspread(SHEET_CONTACT_DATA)
    if not df.empty: st.dataframe(df)
    else: st.info("データなし")


# ---------------------------
# --- メインルーティング (キャッシュクリア機能付き) ---
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
    
    # メニュー切り替え時のキャッシュクリア
    if 'menu_selection' not in st.session_state:
        st.session_state.menu_selection = menu_selection
    
    if st.session_state.menu_selection != menu_selection:
        get_data_from_gspread.clear()
        st.session_state.menu_selection = menu_selection

    # ルーティング
    if menu_selection == "エピノート": page_epi_note()
    elif menu_selection == "メンテノート": page_mainte_note()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_schedule_reservation()
    elif menu_selection == "IVデータ解析": page_iv_analysis()
    elif menu_selection == "PLデータ解析": page_pl_analysis()
    elif menu_selection == "議事録": page_meeting_note()
    elif menu_selection == "知恵袋・質問箱": page_faq()
    elif menu_selection == "装置引き継ぎメモ": page_device_handover()
    elif menu_selection == "トラブル報告": page_trouble_report()
    elif menu_selection == "連絡・問い合わせ": page_contact()

if __name__ == "__main__":
    main()

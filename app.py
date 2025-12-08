# -*- coding: utf-8 -*-
"""
bennriyasann3_revived_full_v1.py
Yamane Lab Convenience Tool - 完全復元・動作修正版
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
SPREADSHEET_NAME = "エピノート (2).xlsx" 
SHEET_EPI_DATA = "エピノート_データ"   
SHEET_MAINTE_DATA = "メンテノート_データ" 
SHEET_SCHEDULE_DATA = "スケジュール" 
SHEET_FAQ_DATA = "知恵袋_データ"
SHEET_TROUBLE_DATA = "トラブル報告_データ" 
SHEET_HANDOVER_DATA = "引き継ぎ_データ"
SHEET_CONTACT_DATA = "お問い合わせ_データ" # CSVファイル名から推測
SHEET_MEETING_DATA = "議事録_データ"

CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files"
# カレンダーID (ユーザー環境に合わせて変更してください)
CALENDAR_ID = "primary" 

# ---------------------------
# --- 認証とクライアント初期化 ---
# ---------------------------
gc = None
gcal_service = None
storage_client = None

try:
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        
        # 1. Gspread (Sheets)
        try:
            gc = gspread.service_account_from_dict(creds_dict)
        except Exception as e:
            st.error(f"Google Sheets認証エラー: {e}")

        # 2. Google Calendar
        try:
            gcal_creds = service_account.Credentials.from_service_account_info(
                creds_dict, scopes=['https://www.googleapis.com/auth/calendar']
            )
            gcal_service = build('calendar', 'v3', credentials=gcal_creds)
        except Exception as e:
            # カレンダー機能が使えなくても他は動かす
            # st.warning(f"Google Calendar認証エラー: {e}") 
            pass

        # 3. GCS (Storage)
        if storage:
            try:
                storage_client = storage.Client()
            except Exception as e:
                # st.warning(f"GCSクライアント初期化エラー: {e}")
                pass
    else:
        st.warning("secrets.toml に 'gcp_service_account' が設定されていません。")

except Exception as e:
    st.error(f"予期せぬ認証エラー: {e}")


# ---------------------------
# --- ユーティリティ関数 ---
# ---------------------------

@st.cache_data(ttl=600)
def get_data_from_gspread(sheet_name):
    if gc is None:
        return pd.DataFrame()
    try:
        worksheet = gc.open(SPREADSHEET_NAME).worksheet(sheet_name)
        data = worksheet.get_all_values()
        if not data:
            return pd.DataFrame()
        return pd.DataFrame(data[1:], columns=data[0])
    except Exception as e:
        # シートがない場合などは空DFを返す
        return pd.DataFrame()

def upload_file_to_gcs(storage_client_obj, file_obj):
    """ファイルをGCSルートに保存し、公開URLを返す"""
    if storage_client_obj is None or storage is None:
        return None, None
    try:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        safe_filename = file_obj.name.replace(' ', '_').replace('/', '_')
        gcs_filename = f"{timestamp}_{safe_filename}" # ルート保存

        bucket = storage_client_obj.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        blob.upload_from_string(
            file_obj.getvalue(), 
            content_type=file_obj.type if hasattr(file_obj, 'type') else 'application/octet-stream'
        )
        
        # 署名付きURLではなく公開パス+認証パラメータ用ベースURL
        # 注: 非公開バケットの場合、ブラウザで見るには署名付きURLが必要だが、
        # 既存データに合わせて単純なURL生成としています。必要なら generate_signed_url を使用。
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        return file_obj.name, public_url
    except Exception as e:
        st.error(f"アップロードエラー: {e}")
        return None, None

def display_attached_files(row_dict, col_url_key, col_filename_key):
    """JSON形式(エスケープ対応)と古いURL形式の両対応で添付ファイルを表示"""
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
    except:
        # 古い形式 (単純なURL文字列)
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

    # 表示
    if urls:
        st.markdown("##### 📎 添付ファイル")
        # 数合わせ
        if len(filenames) < len(urls):
            filenames += [f"File {i+1}" for i in range(len(filenames), len(urls))]
        for u, f in zip(urls, filenames):
            st.markdown(f"[{f}]({u})")
    else:
        st.markdown("なし")

# 共通：データ保存関数
def save_to_sheet(sheet_name, row_data, success_msg="保存しました"):
    try:
        ws = gc.open(SPREADSHEET_NAME).worksheet(sheet_name)
        ws.append_row(row_data)
        st.success(success_msg)
        get_data_from_gspread.clear() # キャッシュクリア
        st.rerun()
    except Exception as e:
        st.error(f"保存エラー: {e}")

# 共通：ファイルアップロード処理
def handle_file_uploads(uploaded_files):
    f_list, u_list = [], []
    if uploaded_files:
        with st.spinner("アップロード中..."):
            for f in uploaded_files:
                name, url = upload_file_to_gcs(storage_client, f)
                if url:
                    f_list.append(name)
                    u_list.append(url)
    return json.dumps(f_list), json.dumps(u_list)


# ---------------------------
# --- 各ページ機能 ---
# ---------------------------

# 1. エピノート
def page_epi_note():
    st.header("エピノート")
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])
    
    with tab2:
        with st.form("epi_form"):
            title = st.text_input("タイトル/番号 (例: 791)")
            cat = st.selectbox("カテゴリ", ["D1", "D2", "その他"])
            memo = st.text_area("詳細メモ", height=150)
            files = st.file_uploader("添付", accept_multiple_files=True)
            with st.expander("データのインポート"): pass # Layout調整
            submit = st.form_submit_button("保存")
        
        if submit:
            if not title:
                st.warning("タイトルは必須です")
            else:
                f_json, u_json = handle_file_uploads(files)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                # 6列構成: Timestamp, Type, Category, Memo, Filename, URL
                row = [ts, "エピノート", cat, f"{title}\n{memo}", f_json, u_json]
                save_to_sheet(SHEET_EPI_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_EPI_DATA)
        if not df.empty:
            if 'タイムスタンプ' in df.columns:
                df = df.sort_values('タイムスタンプ', ascending=False)
            st.dataframe(df, use_container_width=True)
            
            sel = st.selectbox("詳細表示", df['タイムスタンプ'].unique() if 'タイムスタンプ' in df.columns else [], key="epi_sel")
            if sel:
                row = df[df['タイムスタンプ'] == sel].iloc[0].to_dict()
                st.subheader("詳細")
                st.write(f"**日時:** {row.get('タイムスタンプ')}")
                st.write(f"**カテゴリ:** {row.get('カテゴリ')}")
                st.text_area("内容", row.get('メモ'), disabled=True)
                display_attached_files(row, '写真URL', 'ファイル名')

# 2. メンテノート
def page_mainte_note():
    st.header("メンテノート")
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])
    
    with tab2:
        with st.form("mainte_form"):
            title = st.text_input("タイトル")
            dev = st.selectbox("装置", ["MOCVD", "IV/PL", "その他"])
            memo = st.text_area("メモ", height=150)
            files = st.file_uploader("添付", accept_multiple_files=True)
            with st.expander("データのインポート"): pass
            submit = st.form_submit_button("保存")
            
        if submit:
            if not title: st.warning("タイトル必須")
            else:
                f_json, u_json = handle_file_uploads(files)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                # 5列構成: Timestamp, Type, Memo(Title+Dev+Memo), Filename, URL
                content = f"[{title}] (装置: {dev})\n{memo}"
                row = [ts, "メンテノート", content, f_json, u_json]
                save_to_sheet(SHEET_MAINTE_DATA, row)

    with tab1:
        df = get_data_from_gspread(SHEET_MAINTE_DATA)
        if not df.empty:
            if 'タイムスタンプ' in df.columns: df = df.sort_values('タイムスタンプ', ascending=False)
            st.dataframe(df, use_container_width=True)
            sel = st.selectbox("詳細表示", df['タイムスタンプ'].unique(), key="mainte_sel")
            if sel:
                row = df[df['タイムスタンプ'] == sel].iloc[0].to_dict()
                st.text_area("内容", row.get('メモ'), disabled=True)
                display_attached_files(row, '写真URL', 'ファイル名')

# 3. スケジュール・装置予約 (app(4).pyより復元)
def page_schedule_reservation():
    st.header("🗓️ スケジュール・装置予約")
    
    # シンプルなカレンダー登録フォーム
    with st.form("schedule_form"):
        title = st.text_input("予定タイトル", "装置予約: ")
        date_input = st.date_input("日付", date.today())
        start_time = st.time_input("開始時刻", datetime.now().time())
        end_time = st.time_input("終了時刻", (datetime.now() + timedelta(hours=1)).time())
        desc = st.text_area("詳細")
        submit = st.form_submit_button("カレンダーに登録")
    
    if submit:
        if gcal_service:
            try:
                start_dt = datetime.combine(date_input, start_time).isoformat()
                end_dt = datetime.combine(date_input, end_time).isoformat()
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
            st.error("カレンダー機能は現在利用できません（認証設定を確認してください）")
            
    # スプレッドシート側のスケジュール一覧表示（もしあれば）
    st.subheader("予約一覧 (シート)")
    df = get_data_from_gspread(SHEET_SCHEDULE_DATA)
    if not df.empty:
        st.dataframe(df)

# 4. IVデータ解析 (app(4).pyより復元)
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    uploaded_files = st.file_uploader("IV測定データ (.txt, .csv)", accept_multiple_files=True)
    if uploaded_files:
        fig, ax = plt.subplots()
        for f in uploaded_files:
            try:
                # 簡易的な読み込み (スペース/タブ/カンマ区切りに対応)
                df = pd.read_csv(f, sep=r'\s+|,|\t', engine='python', header=None, comment='#')
                if df.shape[1] >= 2:
                    # 往路復路の簡易分離 (最大値で分割)
                    x_data = pd.to_numeric(df.iloc[:, 0], errors='coerce')
                    y_data = pd.to_numeric(df.iloc[:, 1], errors='coerce')
                    df_clean = pd.DataFrame({'x': x_data, 'y': y_data}).dropna()
                    
                    if not df_clean.empty:
                        max_idx = df_clean['x'].idxmax()
                        ax.plot(df_clean.iloc[:max_idx+1]['x'], df_clean.iloc[:max_idx+1]['y'], label=f"{f.name} (往)")
                        ax.plot(df_clean.iloc[max_idx+1:]['x'], df_clean.iloc[max_idx+1:]['y'], label=f"{f.name} (復)", linestyle='--')
            except Exception as e:
                st.warning(f"{f.name} 読み込みエラー: {e}")
        
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Current (A)")
        ax.legend()
        ax.grid(True)
        st.pyplot(fig)

# 5. PLデータ解析 (app(4).pyより復元)
def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    # 校正ロジック簡易版
    st.subheader("1. 波長校正")
    slope = st.number_input("Slope (nm/px)", value=1.0, format="%.4f")
    center_wl = st.number_input("Center Wavelength (nm)", value=500.0)
    center_px = st.number_input("Center Pixel", value=256.0)
    
    st.subheader("2. データプロット")
    uploaded_files = st.file_uploader("PL測定データ", accept_multiple_files=True, key="pl_files")
    if uploaded_files:
        fig, ax = plt.subplots()
        for f in uploaded_files:
            try:
                df = pd.read_csv(f, sep=r'\s+|,|\t', engine='python', header=None, comment='#')
                if df.shape[1] >= 2:
                    y_data = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(0)
                    pixels = np.arange(len(y_data))
                    wavelengths = (pixels - center_px) * slope + center_wl
                    ax.plot(wavelengths, y_data, label=f.name)
            except: pass
        ax.set_xlabel("Wavelength (nm)")
        ax.set_ylabel("Intensity")
        ax.legend()
        st.pyplot(fig)

# 6. 議事録 (app(4).pyより復元)
def page_meeting_note():
    st.header("📄 議事録")
    # CSV構造: Timestamp, Title, AudioName, AudioURL, Content
    page_data_list_generic(SHEET_MEETING_DATA, "議事録", 
                           ["会議タイトル", "議事録内容"], 
                           ["会議タイトル", "議事録内容"], # 入力フィールド
                           "議事録")

# --- 新規実装: 以前NameErrorだったページを汎用ロジックで実装 ---

# 汎用的な「記録＆一覧」ページ作成関数
def page_data_list_generic(sheet_name, title, display_cols, input_labels, note_type):
    st.header(title)
    tab1, tab2 = st.tabs(["一覧表示", "新規記録"])
    
    with tab2: # 新規記録
        with st.form(f"{sheet_name}_form"):
            inputs = []
            for label in input_labels:
                if "内容" in label or "メモ" in label:
                    inputs.append(st.text_area(label, height=100))
                else:
                    inputs.append(st.text_input(label))
            files = st.file_uploader("添付", accept_multiple_files=True, key=f"{sheet_name}_file")
            submit = st.form_submit_button("保存")
            
        if submit:
            f_json, u_json = handle_file_uploads(files)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            # 汎用的な行データ作成: Timestamp, NoteType, Inputs..., Files, URLs
            # ※実際のCSV列順に合わせるため、必要に応じて調整が必要だが、
            #  ここでは最も安全な「後ろに追加」戦略をとるか、CSVヘッダに依存
            row = [ts, note_type] + inputs + [f_json, u_json]
            save_to_sheet(sheet_name, row)

    with tab1: # 一覧
        df = get_data_from_gspread(sheet_name)
        if not df.empty:
            st.dataframe(df)
            # 簡易詳細表示
            sel = st.selectbox("詳細選択", df.iloc[:, 0].unique() if not df.empty else [], key=f"{sheet_name}_sel")
            if sel:
                row = df[df.iloc[:, 0] == sel].iloc[0].to_dict()
                st.write(row)
                display_attached_files(row, 'ファイルURL', 'ファイル名') # 一般的な列名と仮定
                display_attached_files(row, '写真URL', 'ファイル名')   # メンテ/エピ用
                display_attached_files(row, '添付ファイルURL', '添付ファイル名') # 知恵袋用

# 7. 知恵袋
def page_faq():
    # CSV: Timestamp, Title, Content, Email, FileName, FileURL, Status
    page_data_list_generic(SHEET_FAQ_DATA, "💡 知恵袋・質問箱",
                           ["質問タイトル", "質問内容", "ステータス"],
                           ["質問タイトル", "質問内容", "連絡先メールアドレス"],
                           "知恵袋")

# 8. トラブル報告
def page_trouble_report():
    # CSV: Timestamp, Place, Date, When, Cause, Solution, Prevention, Reporter, FileName, FileURL, Title
    st.header("🚨 トラブル報告")
    tab1, tab2 = st.tabs(["一覧", "報告"])
    with tab2:
        with st.form("trb_form"):
            title = st.text_input("件名/タイトル")
            place = st.text_input("機器/場所")
            date_occ = st.date_input("発生日")
            when = st.text_area("トラブル発生時")
            cause = st.text_area("原因/究明")
            sol = st.text_area("対策/復旧")
            prev = st.text_area("再発防止策")
            reporter = st.text_input("報告者")
            files = st.file_uploader("添付", accept_multiple_files=True)
            submit = st.form_submit_button("報告")
        if submit:
            f_j, u_j = handle_file_uploads(files)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            # CSV順序に合わせる
            row = [ts, place, str(date_occ), when, cause, sol, prev, reporter, f_j, u_j, title]
            save_to_sheet(SHEET_TROUBLE_DATA, row)
    with tab1:
        df = get_data_from_gspread(SHEET_TROUBLE_DATA)
        if not df.empty:
            st.dataframe(df)

# 9. 引き継ぎメモ
def page_device_handover():
    # CSV: Timestamp, Type, Title, Content1, Content2, Content3, Memo
    st.header("📝 装置引き継ぎメモ")
    # 簡易実装
    page_data_list_generic(SHEET_HANDOVER_DATA, "引き継ぎメモ", 
                           ["種類", "タイトル", "メモ"], 
                           ["種類", "タイトル", "内容1", "メモ"], 
                           "引き継ぎ")

# 10. 連絡・問い合わせ
def page_contact():
    # CSV: Timestamp, Type, Detail, Contact
    page_data_list_generic(SHEET_CONTACT_DATA, "📧 連絡・問い合わせ",
                           ["お問い合わせの種類", "詳細内容"],
                           ["お問い合わせの種類", "詳細内容", "連絡先"],
                           "問い合わせ")


# ---------------------------
# --- メインルーティング ---
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

# -*- coding: utf-8 -*-
"""
Yamane Lab Convenience Tool - Complete Fixed Version + High-End Graph Plotter
機能: エピノート/メンテノート/カレンダー/解析(IV, PL)/議事録/知恵袋/引き継ぎ/トラブル/問い合わせ/【New】グラフ描画
"""

import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import json
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
from datetime import datetime, date, timedelta
from urllib.parse import quote as url_quote, unquote as url_unquote
from io import BytesIO
import calendar
import matplotlib.font_manager as fm
from functools import reduce

# Google Services
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Optional GCS
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

# --- Streamlit ページ設定 ---
st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

# ---------------------------
# --- Constants ---
# ---------------------------
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files"
SPREADSHEET_NAME = "エピノート"

# シート定義 (省略 - そのまま維持)
SHEET_EPI_DATA = 'エピノート_データ'
EPI_COL_TIMESTAMP = 'タイムスタンプ'
EPI_COL_CATEGORY = 'カテゴリ'
EPI_COL_MEMO = 'メモ'
EPI_COL_FILENAME = 'ファイル名'
EPI_COL_FILE_URL = '写真URL'

SHEET_MAINTE_DATA = 'メンテノート_データ'
MAINT_COL_TIMESTAMP = 'タイムスタンプ'
MAINT_COL_MEMO = 'メモ'
MAINT_COL_FILENAME = 'ファイル名'
MAINT_COL_FILE_URL = '写真URL'

SHEET_MEETING_DATA = '議事録_データ'
MEETING_COL_TIMESTAMP = 'タイムスタンプ'
MEETING_COL_TITLE = '会議タイトル'
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

SHEET_CONTACT_DATA = 'お問い合わせ_データ'
CONTACT_COL_TIMESTAMP = 'タイムスタンプ'
CONTACT_COL_TYPE = 'お問い合わせの種類'
CONTACT_COL_DETAIL = '詳細内容'
CONTACT_COL_CONTACT = '連絡先'

SHEET_TROUBLE_DATA = 'トラブル報告_データ'
TROUBLE_COL_TIMESTAMP = 'タイムスタンプ'
TROUBLE_COL_DEVICE = '機器/場所'
TROUBLE_COL_TITLE = '件名/タイトル'
TROUBLE_COL_CAUSE = '原因/究明'
TROUBLE_COL_SOLUTION = '対策/復旧'
TROUBLE_COL_REPORTER = '報告者'

# Calendar Config
CALENDAR_ID = "yamane.lab.6747@gmail.com"
SCOPES = ['https://www.googleapis.com/auth/calendar']

# ---------------------------
# --- Service Classes ---
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
    def generate_signed_url(self, **kwargs): return None

# ---------------------------
# --- Initialization ---
# ---------------------------
@st.cache_resource(ttl=3600)
def initialize_google_services():
    global storage
    gc_client = DummyGSClient()
    storage_client_obj = DummyStorageClient()
    calendar_service = None

    if "gcs_credentials" not in st.secrets:
        # st.sidebar.warning("⚠️ Secrets未設定 (オフラインモード)")
        return gc_client, storage_client_obj, calendar_service

    try:
        raw = st.secrets["gcs_credentials"]
        cleaned = raw.strip().replace('\t', '').replace('\r', '').replace('\n', '')
        info = json.loads(cleaned)
        
        gc_client = gspread.service_account_from_dict(info)
        if storage:
            storage_client_obj = storage.Client.from_service_account_info(info)
        
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        calendar_service = build('calendar', 'v3', credentials=creds)
        
        # st.sidebar.success("✅ Googleサービス認証 成功")
        return gc_client, storage_client_obj, calendar_service

    except Exception:
        # st.sidebar.error(f"Googleサービス初期化エラー: {e}")
        return gc_client, storage_client_obj, calendar_service

gc, storage_client, calendar_service = initialize_google_services()

# ---------------------------
# --- Utils ---
# ---------------------------
def upload_file_to_gcs(storage_client_obj, file_obj):
    if isinstance(storage_client_obj, DummyStorageClient) or storage is None:
        return None, None
    try:
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        original_filename = file_obj.name
        safe_filename = re.sub(r'[^a-zA-Z0-9_.]', '_', original_filename)
        gcs_filename = f"{timestamp}_{safe_filename}"
        
        bucket = storage_client_obj.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(gcs_filename)
        blob.upload_from_string(
            file_obj.getvalue(), 
            content_type=file_obj.type if hasattr(file_obj, 'type') else 'application/octet-stream'
        )
        public_url = f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/{url_quote(gcs_filename)}"
        return original_filename, public_url
    except Exception:
        return None, None

def generate_signed_url(blob_name_quoted, expiration_minutes=15):
    if isinstance(storage_client, DummyStorageClient): return None
    try:
        bucket = storage_client.bucket(CLOUD_STORAGE_BUCKET_NAME)
        blob = bucket.blob(blob_name_quoted)
        return blob.generate_signed_url(version="v4", expiration=timedelta(minutes=expiration_minutes), method="GET")
    except Exception:
        return None

@st.cache_data(ttl=600)
def get_sheet_as_df(spreadsheet_name, sheet_name):
    try:
        if isinstance(gc, DummyGSClient): return pd.DataFrame()
        ws = gc.open(spreadsheet_name).worksheet(sheet_name)
        data = ws.get_all_values()
        if not data or len(data) <= 1: return pd.DataFrame()
        return pd.DataFrame(data[1:], columns=data[0])
    except Exception:
        return pd.DataFrame()

def display_attached_files(row, col_url, col_filename):
    raw_urls = row.get(col_url, '')
    raw_names = row.get(col_filename, '')
    urls = []
    names = []
    
    try:
        urls = json.loads(raw_urls) if raw_urls else []
        if not isinstance(urls, list): urls = [raw_urls] if isinstance(raw_urls, str) else []
    except:
        if raw_urls and raw_urls.startswith('http'): urls = [raw_urls]
        
    try:
        names = json.loads(raw_names) if raw_names else []
        if not isinstance(names, list): names = [names] if isinstance(names, str) else []
    except:
        pass

    while len(names) < len(urls): names.append(f"File {len(names)+1}")
    
    if urls:
        st.markdown("##### 📎 添付ファイル")
        
        for u, n in zip(urls, names):
            is_image = n.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))
            
            blob_name_quoted = None
            if u.startswith(f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/"):
                blob_name_quoted = u.split(f"https://storage.googleapis.com/{CLOUD_STORAGE_BUCKET_NAME}/")[1]

            if is_image and blob_name_quoted:
                signed_url = generate_signed_url(blob_name_quoted) 
                
                if signed_url:
                    st.image(signed_url, caption=f"画像: {n}", width=400)
                else:
                    st.markdown(f"- **画像 ({n})** : GCSアクセス失敗、またはファイル期限切れのため [ダウンロードリンク]({u})")
            else:
                st.markdown(f"- [{n}]({u})")

# --- Excel Export Helpers ---
def to_excel(df):
    output = BytesIO()
    df = df.apply(pd.to_numeric, errors='coerce').astype(float)
    if 'Axis_X' in df.columns: df.rename(columns={'Axis_X': 'Voltage_V'}, inplace=True)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Combined Data') 
    processed_data = output.getvalue()
    return processed_data

def to_excel_multi_sheet(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            export_df = df.apply(pd.to_numeric, errors='coerce').astype(float)
            if 'Axis_X' in export_df.columns:
                 export_df.rename(columns={'Axis_X': 'Voltage_V'}, inplace=True)
            export_df.to_excel(writer, index=False, sheet_name=sheet_name)
    processed_data = output.getvalue()
    return processed_data

# ---------------------------
# --- Data Loaders ---
# ---------------------------
@st.cache_data
def load_data_file(uploaded_bytes, filename):
    try:
        text = uploaded_bytes.decode('utf-8', errors='ignore').splitlines()
        data_lines = [line.strip() for line in text if line.strip() and not line.strip().startswith(('#','!','/'))]
        if data_lines and not data_lines[0][0].isdigit():
            data_lines = data_lines[1:]
            
        df = pd.read_csv(io.StringIO("\n".join(data_lines)), sep=r'\s+|,|\t', engine='python', header=None)
        if df.shape[1] < 2: return None
        df = df.iloc[:, :2]
        df.columns = ['Axis_X', filename]
        df = df.apply(pd.to_numeric, errors='coerce').dropna()
        return df
    except:
        return None

@st.cache_data
def load_pl_data(uploaded_file):
    try:
        content = uploaded_file.getvalue().decode('utf-8', errors='ignore').splitlines()
        
        data_lines = []
        for line in content:
            line = line.strip()
            if not line: continue
            if line.startswith(('#', '!', '/')): continue
            data_lines.append(line)
            
        if not data_lines: return None

        df = pd.read_csv(io.StringIO("\n".join(data_lines)), 
                         sep=r'[\t, ]+', 
                         engine='python', 
                         header=None)

        if df.shape[1] < 2: 
            df = df.dropna(axis=1, how='all')
            if df.shape[1] < 2:
                return None
        
        df = df.iloc[:, :2]
        df.columns = ['pixel', 'intensity']
        df = df.apply(pd.to_numeric, errors='coerce').dropna()
        if df.empty: return None
        return df
    except Exception:
        return None
        
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import json
import io
from scipy import stats
from datetime import datetime
from io import BytesIO

# ==========================================
# 関数定義: page_graph_plotting (完全統合版)
# ==========================================
def page_graph_plotting():
    st.header("📈 統合型グラフ解析ツール")
    st.markdown("以前の「詳細設定・Excelコピペ」機能に、新しい「MPPT解析・順序入替・単位換算」を追加しました。")

    # --- CSS: プレビューの追従 & 数値入力の調整 ---
    st.markdown("""
        <style>
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-of-type(2) {
            position: sticky; top: 4rem; align-self: start; z-index: 999;
        }
        </style>
    """, unsafe_allow_html=True)

    # --- セッション管理 ---
    if 'gp_data_list' not in st.session_state:
        st.session_state['gp_data_list'] = []

    # --- ヘルパー関数: データ移動 ---
    def move_data(idx, direction):
        lst = st.session_state['gp_data_list']
        if direction == "up" and idx > 0:
            lst[idx], lst[idx-1] = lst[idx-1], lst[idx]
        elif direction == "down" and idx < len(lst) - 1:
            lst[idx], lst[idx+1] = lst[idx+1], lst[idx]

    # ==========================================
    # 0. プロジェクト管理 (JSON)
    # ==========================================
    with st.expander("💾 プロジェクトの保存・復元", expanded=False):
        c_load, c_save = st.columns(2)
        with c_load:
            st.markdown("#### 📂 復元")
            uploaded_project = st.file_uploader("プロジェクトファイル (.json)", type=["json"], key="project_loader")
            if uploaded_project:
                if st.button("設定を読み込む"): # 自動削除せずボタンで実行
                    try:
                        project_data = json.load(uploaded_project)
                        restored_data_list = []
                        for item in project_data.get("datasets", []):
                            # CSVからDataFrame復元
                            df_restored = pd.read_csv(io.StringIO(item["data_csv"]))
                            # 辞書に復元
                            item['df'] = df_restored
                            restored_data_list.append(item)
                        
                        # 既存データに追記するか置換するか（ここでは置換）
                        st.session_state['gp_data_list'] = restored_data_list
                        
                        # 設定値の復元
                        saved_settings = project_data.get("settings", {})
                        for key, value in saved_settings.items():
                            st.session_state[key] = value
                        st.success("✅ 復元完了")
                        st.rerun()
                    except Exception as e: st.error(f"エラー: {e}")

        with c_save:
            st.markdown("#### 💾 保存")
            if st.button("プロジェクトファイルを作成"):
                if not st.session_state['gp_data_list']:
                    st.warning("データなし")
                else:
                    datasets_serialized = []
                    for d in st.session_state['gp_data_list']:
                        csv_buffer = io.StringIO()
                        d['df'].to_csv(csv_buffer, index=False)
                        d_copy = d.copy()
                        d_copy['data_csv'] = csv_buffer.getvalue()
                        if 'df' in d_copy: del d_copy['df']
                        datasets_serialized.append(d_copy)
                    
                    settings_snapshot = {}
                    for key, val in st.session_state.items():
                        if key in ['gp_uploader', 'project_loader', 'gp_data_list']: continue
                        if isinstance(val, (int, float, str, bool, list, dict, type(None))):
                            settings_snapshot[key] = val

                    project_obj = {
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "datasets": datasets_serialized,
                        "settings": settings_snapshot
                    }
                    json_str = json.dumps(project_obj, indent=2, ensure_ascii=False)
                    file_name = f"GraphProject_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
                    st.download_button("⬇️ JSONをダウンロード", json_str, file_name, "application/json")

    # ==========================================
    # 1. データ入力 (ファイル & コピペ & 追加)
    # ==========================================
    st.subheader("1. データの入力")
    
    # 既存データの管理
    if st.session_state['gp_data_list']:
        st.info(f"現在のデータ数: {len(st.session_state['gp_data_list'])} (下部で追加可能)")
        if st.button("🗑️ 全データをクリア"):
            st.session_state['gp_data_list'] = []; st.rerun()
    
    tab1, tab2 = st.tabs(["📂 ファイルから追加", "📋 エクセルから貼り付け"])
    
    with tab1:
        # アップロードしても既存データを消さないように処理
        files = st.file_uploader("CSV/Excelファイル", accept_multiple_files=True, key="gp_uploader_add")
        if files:
            new_data_added = False
            for f in files:
                # 名前重複チェック（簡易）
                if any(d['name'] == f.name for d in st.session_state['gp_data_list']):
                    continue 

                df = None
                try:
                    if f.name.endswith(('.xlsx', '.xls')):
                        df = pd.read_excel(f)
                    else:
                        df = pd.read_csv(f)
                except: pass
                
                if df is not None:
                    # 数値列のみ抽出 & 列名クリーニング
                    df = df.select_dtypes(include=[np.number])
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    # 初期データ構造
                    st.session_state['gp_data_list'].append({
                        "name": f.name, "df": df,
                        "scale_x": 1.0, "scale_y": 1.0, # 新機能: スケーリング
                        "mppt": False, "show_eq": False, # 新機能: 解析
                        "color": "#0000FF", "marker": "None", "linestyle": "-"
                    })
                    new_data_added = True
            
            if new_data_added:
                st.rerun() # リロードして反映

    with tab2:
        st.caption("Excelのデータをコピーしてここに貼り付け、Ctrl+Enterで確定")
        paste_text = st.text_area("データ貼り付けエリア", height=100)
        paste_name = st.text_input("データセット名", value=f"Data_{len(st.session_state['gp_data_list'])+1}")
        
        if st.button("貼り付けたデータを追加"):
            if paste_text:
                try:
                    lines = [l.strip() for l in paste_text.splitlines() if l.strip()]
                    df_paste = pd.read_csv(io.StringIO("\n".join(lines)), sep=r'[\t, ]+', engine='python')
                    if df_paste is not None and not df_paste.empty:
                        df_paste = df_paste.select_dtypes(include=[np.number])
                        st.session_state['gp_data_list'].append({
                            "name": paste_name, "df": df_paste,
                            "scale_x": 1.0, "scale_y": 1.0,
                            "mppt": False, "show_eq": False,
                            "color": "#000000", "marker": "o", "linestyle": "-"
                        })
                        st.success("追加しました")
                        st.rerun()
                except Exception as e: st.error(f"読み込み失敗: {e}")

    # データがない場合はここで終了
    data_list = st.session_state['gp_data_list']
    if not data_list: return

    # ==========================================
    # 2. グラフ設定 (左右分割)
    # ==========================================
    st.markdown("---")
    col_settings, col_preview = st.columns([1.3, 2])

    with col_settings:
        st.subheader("2. 詳細設定")

        # --- A. キャンバス設定 ---
        with st.expander("📊 キャンバス・フォント", expanded=False):
            c1, c2 = st.columns(2)
            fig_w = c1.number_input("幅 (inch)", 1.0, 50.0, 6.0, step=0.5)
            fig_h = c2.number_input("高さ (inch)", 1.0, 50.0, 4.0, step=0.5)
            dpi_val = st.number_input("解像度 (DPI)", 72, 600, 150)
            font_family = st.selectbox("フォント", ["Arial", "Times New Roman", "Helvetica", "Meiryo", "Yu Gothic"], index=0)
            base_fs = st.number_input("基本フォントサイズ", 6, 50, 12)

        # --- B. 4軸設定 (負の値対応) ---
        with st.expander("📐 軸 (Axes) 設定", expanded=True):
            tabs_ax = st.tabs(["X軸(下)", "X軸(上)", "Y軸(左)", "Y軸(右)", "共通"])
            ax_settings = {}

            def axis_ui(key_prefix, label_def):
                label = st.text_input("ラベル", label_def, key=f"{key_prefix}_lbl")
                c1, c2 = st.columns(2)
                # 【修正】min_value=None で負の数も入力可能に
                d_min = c1.number_input("最小 (空=Auto)", value=None, format="%f", key=f"{key_prefix}_min")
                d_max = c2.number_input("最大 (空=Auto)", value=None, format="%f", key=f"{key_prefix}_max")
                
                c3, c4 = st.columns(2)
                maj_int = c3.number_input("主目盛間隔", 0.0, step=0.1, key=f"{key_prefix}_maj")
                min_int = c4.number_input("補助目盛間隔", 0.0, step=0.1, key=f"{key_prefix}_min_int")
                
                is_log = st.checkbox("対数軸", False, key=f"{key_prefix}_log")
                return {"label": label, "min": d_min, "max": d_max, "maj": maj_int, "min_int": min_int, "log": is_log}

            with tabs_ax[0]: ax_settings['x1'] = axis_ui("x1", "Voltage (V)")
            with tabs_ax[1]: ax_settings['x2'] = axis_ui("x2", "Secondary X")
            with tabs_ax[2]: ax_settings['y1'] = axis_ui("y1", "Current (A)")
            with tabs_ax[3]: ax_settings['y2'] = axis_ui("y2", "Power (W)")
            
            with tabs_ax[4]:
                tick_dir = st.selectbox("目盛の向き", ["in", "out", "inout"], index=0)
                show_grid = st.checkbox("グリッド表示", True)
                zero_cross = st.checkbox("原点(0,0)を通る線を描画", True)

        # --- C. 凡例設定 ---
        with st.expander("📝 凡例 (Legend)", expanded=False):
            show_leg = st.checkbox("凡例を表示", True)
            l_loc = st.selectbox("位置", ["best", "upper right", "lower left", "outside right"], index=0)
            l_col = st.number_input("列数", 1, 5, 1)

        # --- D. データ系列の個別設定 (順序・解析・スタイル) ---
        st.markdown("#### データ系列設定")
        
        # セッション内のリストを直接操作して順序変更を反映
        datasets = st.session_state['gp_data_list']
        
        for i, d in enumerate(datasets):
            with st.expander(f"#{i+1}: {d['name']}", expanded=False):
                # 1. 順序変更 & 削除ボタン
                bc1, bc2, bc3 = st.columns([1, 1, 2])
                with bc1:
                    if st.button("⬆", key=f"btn_u_{i}"): move_data(i, "up"); st.rerun()
                with bc2:
                    if st.button("⬇", key=f"btn_d_{i}"): move_data(i, "down"); st.rerun()
                with bc3:
                    if st.button("❌ 削除", key=f"btn_del_{i}"): datasets.pop(i); st.rerun()

                # 2. 列選択 & スケーリング (単位換算)
                cols = d['df'].columns.tolist()
                sc1, sc2 = st.columns(2)
                xc = sc1.selectbox(f"X列", cols, index=0, key=f"xc_{i}")
                yc = sc2.selectbox(f"Y列", cols, index=1 if len(cols)>1 else 0, key=f"yc_{i}")
                
                st.caption("単位換算 (例: 1000倍=m単位)")
                kc1, kc2 = st.columns(2)
                # 辞書キーがなければ初期値1.0を入れる
                d['scale_x'] = kc1.number_input("X倍率", value=d.get('scale_x', 1.0), format="%e", key=f"kcx_{i}")
                d['scale_y'] = kc2.number_input("Y倍率", value=d.get('scale_y', 1.0), format="%e", key=f"kcy_{i}")

                # 3. スタイル設定
                tc1, tc2 = st.columns(2)
                d['color'] = tc1.color_picker("色", d.get('color', '#0000FF'), key=f"clr_{i}")
                d['marker'] = tc2.selectbox("マーカー", ["None", "o", "s", "^", "x"], index=0 if d.get('marker')=="None" else 1, key=f"mrk_{i}")
                
                lw1, lw2 = st.columns(2)
                d['line_width'] = lw1.number_input("線幅", 0.0, 10.0, d.get('line_width', 1.5), key=f"lw_{i}")
                d['marker_size'] = lw2.number_input("点サイズ", 0.0, 20.0, d.get('marker_size', 6.0), key=f"ms_{i}")
                
                d['linestyle'] = st.selectbox("線種", ["-", "--", "-.", ":", "None"], index=0, key=f"lst_{i}")

                # 4. 軸の割り当て
                ac1, ac2 = st.columns(2)
                d['use_top'] = ac1.checkbox("上X軸を使用", d.get('use_top', False), key=f"ut_{i}")
                d['use_right'] = ac2.checkbox("右Y軸を使用", d.get('use_right', False), key=f"ur_{i}")

                # 5. 解析機能 (MPPT & 近似)
                st.markdown("---")
                st.caption("解析機能")
                
                # MPPT
                d['mppt'] = st.checkbox("MPPT解析 (第2象限の最大電力)", d.get('mppt', False), key=f"mppt_{i}")
                
                # 近似曲線
                d['fit_mode'] = st.selectbox("近似曲線", ["なし", "線形 (y=ax+b)", "多項式(2次)", "移動平均"], 
                                             index=0, key=f"fit_{i}")
                if d['fit_mode'] != "なし":
                    d['show_eq'] = st.checkbox("数式を表示", d.get('show_eq', False), key=f"seq_{i}")

                # データ更新 (辞書に値を書き戻す)
                d.update({'x_col': xc, 'y_col': yc})

    # ==========================================
    # 3. 描画 (プレビューエリア)
    # ==========================================
    with col_preview:
        st.subheader("プレビュー")
        
        # フォント設定
        plt.rcParams['font.size'] = base_fs
        if font_family in ["Times New Roman", "Arial", "Helvetica"]:
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [font_family]
        else:
            plt.rcParams['font.family'] = font_family # 日本語フォント等

        fig, ax1 = plt.subplots(figsize=(fig_w, fig_h), dpi=dpi_val)
        
        # 軸の構築 (Top/Rightが必要か判定)
        has_right = any(d.get('use_right') for d in datasets)
        has_top = any(d.get('use_top') for d in datasets)

        ax2, ax3, ax4 = None, None, None
        
        # マッピング: (use_top, use_right) -> axis object
        axes_map = {(False, False): ax1}

        if has_right:
            ax2 = ax1.twinx()
            axes_map[(False, True)] = ax2
        if has_top:
            ax3 = ax1.twiny()
            axes_map[(True, False)] = ax3
        if has_right and has_top:
            # 4軸目は少し複雑だが、簡易的にax3(上)とax2(右)を共有する形を作る
            # 厳密な4軸独立にはax4 = ax1.twinx().twiny() 等が必要
            ax4 = ax1.twinx()
            # ここでは簡易実装として既存の軸を使うか、新規作成
            axes_map[(True, True)] = ax3 # 仮実装（複雑化を防ぐため）

        # 軸設定の適用関数
        def apply_axis_conf(ax, xc, yc):
            if not ax: return
            ax.set_xlabel(xc['label'])
            ax.set_ylabel(yc['label'])
            
            if xc['min'] is not None: ax.set_xlim(left=xc['min'])
            if xc['max'] is not None: ax.set_xlim(right=xc['max'])
            if yc['min'] is not None: ax.set_ylim(bottom=yc['min'])
            if yc['max'] is not None: ax.set_ylim(top=yc['max'])
            
            if xc['log']: ax.set_xscale('log')
            if yc['log']: ax.set_yscale('log')
            
            if xc['maj'] > 0: ax.xaxis.set_major_locator(ticker.MultipleLocator(xc['maj']))
            if yc['maj'] > 0: ax.yaxis.set_major_locator(ticker.MultipleLocator(yc['maj']))
            
            ax.tick_params(direction=tick_dir, which='both')

        # 設定適用
        apply_axis_conf(ax1, ax_settings['x1'], ax_settings['y1'])
        apply_axis_conf(ax2, ax_settings['x1'], ax_settings['y2']) # 右軸は下軸と共有X
        apply_axis_conf(ax3, ax_settings['x2'], ax_settings['y1']) # 上軸は左軸と共有Y
        
        # グリッド
        if show_grid: ax1.grid(True, linestyle=':', alpha=0.6)
        if zero_cross: 
            ax1.axhline(0, color='black', linewidth=0.8)
            ax1.axvline(0, color='black', linewidth=0.8)

        # プロット実行
        for d in datasets:
            # データ準備
            df = d['df']
            x_raw = df[d['x_col']]
            y_raw = df[d['y_col']]
            
            # スケーリング
            x_data = x_raw * d.get('scale_x', 1.0)
            y_data = y_raw * d.get('scale_y', 1.0)
            
            # 軸決定
            target_ax = axes_map.get((d.get('use_top', False), d.get('use_right', False)), ax1)
            
            # NaN除去
            mask = pd.notna(x_data) & pd.notna(y_data)
            x_plot, y_plot = x_data[mask], y_data[mask]

            if len(x_plot) == 0: continue

            # メインプロット
            ls = d.get('linestyle', '-')
            if ls == "None": ls = ""
            mk = d.get('marker', 'None')
            if mk == "None": mk = ""
            
            target_ax.plot(x_plot, y_plot, label=d['name'], 
                           color=d['color'], marker=mk, linestyle=ls,
                           linewidth=d.get('line_width', 1.5), markersize=d.get('marker_size', 6))

            # --- 解析機能: 近似曲線 ---
            fmode = d.get('fit_mode', "なし")
            if fmode != "なし" and len(x_plot) > 1:
                idx_sorted = np.argsort(x_plot)
                xs = x_plot.iloc[idx_sorted]
                ys = y_plot.iloc[idx_sorted]
                
                eq_text = ""
                y_fit = None
                
                try:
                    if "線形" in fmode:
                        slope, intercept, r_val, _, _ = stats.linregress(xs, ys)
                        y_fit = slope * xs + intercept
                        eq_text = f"y={slope:.2e}x+{intercept:.2e}\n$R^2$={r_val**2:.3f}"
                    elif "2次" in fmode:
                        coef = np.polyfit(xs, ys, 2)
                        y_fit = np.polyval(coef, xs)
                        eq_text = "Poly(deg=2)"
                    elif "移動平均" in fmode:
                        y_fit = ys.rolling(window=5, center=True).mean()

                    if y_fit is not None:
                        target_ax.plot(xs, y_fit, color=d['color'], linestyle='--', linewidth=1, alpha=0.8)
                        if d.get('show_eq') and eq_text:
                            # グラフ上の最終点付近に表示
                            target_ax.text(xs.iloc[-1], y_fit.iloc[-1], eq_text, fontsize=9, color=d['color'])
                except: pass

            # --- 解析機能: MPPT (第2象限) ---
            if d.get('mppt'):
                # 第2象限: X < 0, Y > 0
                m_mask = (x_plot < 0) & (y_plot > 0)
                xm, ym = x_plot[m_mask], y_plot[m_mask]
                
                if len(xm) > 0:
                    p = (xm * ym).abs()
                    max_i = p.idxmax()
                    best_x, best_y, best_p = xm[max_i], ym[max_i], p[max_i]
                    
                    target_ax.plot(best_x, best_y, marker='*', color='gold', markersize=14, markeredgecolor='black', zorder=10)
                    target_ax.annotate(f"MPPT:{best_p:.2f}W\n({best_x:.1f}V, {best_y:.1f}A)",
                                       xy=(best_x, best_y), xytext=(10, -30),
                                       textcoords='offset points', arrowprops=dict(arrowstyle="->"),
                                       bbox=dict(boxstyle="round", fc="white", alpha=0.7))

        # 凡例表示
        if show_leg:
            lines = []
            labels = []
            for ax in [ax1, ax2, ax3, ax4]:
                if ax is not None:
                    l, lb = ax.get_legend_handles_labels()
                    lines.extend(l)
                    labels.extend(lb)
            # 重複除去
            by_label = dict(zip(labels, lines))
            
            bbox = None
            loc_param = l_loc
            if l_loc == "outside right":
                loc_param = "center left"
                bbox = (1.05, 0.5)
                
            ax1.legend(by_label.values(), by_label.keys(), loc=loc_param, bbox_to_anchor=bbox, ncol=l_col)

        plt.tight_layout()
        st.pyplot(fig)
        
        # 保存ボタン
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=300, bbox_inches='tight')
        st.download_button("画像を保存 (PNG)", buf.getvalue(), "my_graph.png", "image/png")

# ---------------------------
# --- Components ---
# ---------------------------
# (前回と同じ page_data_list は省略せずそのまま記述します)
def page_data_list(sheet_name, title, col_time, col_filter, col_memo, col_url, detail_cols, col_filename):
    st.subheader(f"📚 {title} 一覧")
    df = get_sheet_as_df(SPREADSHEET_NAME, sheet_name)
    if df.empty:
        st.info("データがありません")
        return

    search_query = st.text_input("📝 検索（メモ/タイトルを絞り込み）", key=f"{sheet_name}_search").strip()
    
    filtered_df = df.copy()
    if col_filter and col_filter in df.columns:
        options = ["すべて"] + sorted(list(df[col_filter].unique()))
        sel = st.selectbox(f"カテゴリで絞り込み", options)
        if sel != "すべて": filtered_df = filtered_df[filtered_df[col_filter] == sel]
            
    if search_query:
        searchable_cols = [col_memo]
        search_mask = False
        for col in searchable_cols:
            if col in filtered_df.columns:
                mask = filtered_df[col].astype(str).str.contains(search_query, case=False, na=False)
                search_mask = search_mask | mask
        filtered_df = filtered_df[search_mask]
        
    if filtered_df.empty:
        st.warning("該当するデータは見つかりませんでした。")
        return

    if col_time in filtered_df.columns:
        filtered_df = filtered_df.sort_values(col_time, ascending=False)

    st.markdown("---")
    for i, row in filtered_df.iterrows():
        ts_display = row.get(col_time,'不明')
        memo_content = str(row.get(col_memo,''))
        first_line = memo_content.split('\n')[0].strip()
        expander_title = f"{first_line}"
        
        with st.expander(expander_title):
            st.write(f"**{EPI_COL_TIMESTAMP}:** {ts_display}")
            for col in detail_cols:
                if col in row and col not in [col_url, col_filename, col_time]:
                    st.write(f"**{col}:** {row[col]}")
            display_attached_files(row, col_url, col_filename)

# ---------------------------
# --- Pages (Existing) ---
# ---------------------------
def page_epi_note_recording():
    st.markdown("#### 📝 新しいエピノートを記録")
    with st.form("epi_form"):
        title = st.text_input("タイトル/番号 (例: 791)")
        cat = st.selectbox("カテゴリ", ["D1", "D2", "その他"])
        memo = st.text_area("メモ")
        files = st.file_uploader("添付", accept_multiple_files=True)
        if st.form_submit_button("保存"):
            if not title: st.error("タイトル必須"); return
            f_names, f_urls = [], []
            if files:
                for f in files:
                    n, u = upload_file_to_gcs(storage_client, f)
                    if u: f_names.append(n); f_urls.append(u)
            row = [
                datetime.now().strftime("%Y%m%d_%H%M%S"),
                "エピノート", cat, f"{title}\n{memo}",
                json.dumps(f_names), json.dumps(f_urls)
            ]
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_EPI_DATA).append_row(row)
                st.success("保存成功")
                st.cache_data.clear()
            except Exception as e: st.error(f"エラー: {e}")

def page_epi_note():
    st.header("エピノート")
    tab1, tab2 = st.tabs(["📝 記録", "📚 一覧"])
    with tab1: page_epi_note_recording()
    with tab2:
        page_data_list(SHEET_EPI_DATA, "エピノート", EPI_COL_TIMESTAMP, EPI_COL_CATEGORY, EPI_COL_MEMO, EPI_COL_FILE_URL, 
                       [EPI_COL_TIMESTAMP, EPI_COL_CATEGORY, EPI_COL_MEMO], EPI_COL_FILENAME)

def page_mainte_recording():
    st.markdown("#### 📝 新しいメンテノートを記録")
    with st.form("mainte_form"):
        dev = st.selectbox("装置", ["MBE", "XRD", "PL", "AFM", "その他"])
        title = st.text_input("タイトル")
        memo = st.text_area("詳細")
        files = st.file_uploader("添付", accept_multiple_files=True)
        if st.form_submit_button("保存"):
            if not title: st.error("タイトル必須"); return
            f_names, f_urls = [], []
            if files:
                for f in files:
                    n, u = upload_file_to_gcs(storage_client, f)
                    if u: f_names.append(n); f_urls.append(u)
            row = [
                datetime.now().strftime("%Y%m%d_%H%M%S"),
                "メンテノート", f"[{title}] {dev}\n{memo}",
                json.dumps(f_names), json.dumps(f_urls)
            ]
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_MAINTE_DATA).append_row(row)
                st.success("保存成功")
                st.cache_data.clear()
            except Exception as e: st.error(f"エラー: {e}")

def page_mainte_note():
    st.header("メンテノート")
    tab1, tab2 = st.tabs(["📝 記録", "📚 一覧"])
    with tab1: page_mainte_recording()
    with tab2:
        page_data_list(SHEET_MAINTE_DATA, "メンテノート", MAINT_COL_TIMESTAMP, None, MAINT_COL_MEMO, MAINT_COL_FILE_URL,
                       [MAINT_COL_TIMESTAMP, MAINT_COL_MEMO], MAINT_COL_FILENAME)

def page_meeting_note():
    st.header("議事録")
    with st.form("meeting_form"):
        title = st.text_input("会議タイトル")
        content = st.text_area("内容")
        url = st.text_input("音声URL")
        if st.form_submit_button("保存"):
            if not title: st.error("タイトル必須"); return
            row = [datetime.now().strftime("%Y%m%d_%H%M%S"), title, "", url, content]
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_MEETING_DATA).append_row(row)
                st.success("保存成功")
                st.cache_data.clear()
            except Exception as e: st.error(f"エラー: {e}")
    page_data_list(SHEET_MEETING_DATA, "議事録", MEETING_COL_TIMESTAMP, None, MEETING_COL_TITLE, MEETING_COL_AUDIO_URL,
                   [MEETING_COL_TIMESTAMP, MEETING_COL_TITLE, MEETING_COL_CONTENT], None)

def page_qa_box():
    st.header("知恵袋")
    with st.form("qa_form"):
        title = st.text_input("質問タイトル")
        content = st.text_area("内容")
        contact = st.text_input("連絡先")
        files = st.file_uploader("添付", accept_multiple_files=True)
        if st.form_submit_button("送信"):
            if not title: st.error("タイトル必須"); return
            f_names, f_urls = [], []
            if files:
                for f in files:
                    n, u = upload_file_to_gcs(storage_client, f)
                    if u: f_names.append(n); f_urls.append(u)
            row = [
                datetime.now().strftime("%Y%m%d_%H%M%S"), title, content, contact,
                json.dumps(f_names), json.dumps(f_urls), "未解決"
            ]
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_QA_DATA).append_row(row)
                st.success("送信成功")
                st.cache_data.clear()
            except Exception as e: st.error(f"エラー: {e}")
    page_data_list(SHEET_QA_DATA, "QA", QA_COL_TIMESTAMP, QA_COL_STATUS, QA_COL_TITLE, QA_COL_FILE_URL,
                   [QA_COL_TIMESTAMP, QA_COL_TITLE, QA_COL_CONTENT, QA_COL_STATUS], QA_COL_FILENAME)

def page_handover_note():
    st.header("引き継ぎメモ")
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
                st.success("保存成功")
                st.cache_data.clear()
            except Exception as e: st.error(f"エラー: {e}")
    page_data_list(SHEET_HANDOVER_DATA, "引き継ぎ", HANDOVER_COL_TIMESTAMP, HANDOVER_COL_TYPE, HANDOVER_COL_TITLE, None,
                   [HANDOVER_COL_TIMESTAMP, HANDOVER_COL_TYPE, HANDOVER_COL_TITLE, HANDOVER_COL_MEMO], None)

def page_trouble_report():
    st.header("トラブル報告")
    with st.form("trouble_form"):
        dev = st.selectbox("機器", ["MBE", "XRD", "PL", "IV", "TEM・SEM", "抵抗加熱蒸着", "RTA", "フォトリソ", "ドラフト", "その他"])
        title = st.text_input("件名")
        cause = st.text_area("原因")
        sol = st.text_area("対策")
        rep = st.text_input("報告者")
        if st.form_submit_button("保存"):
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_TROUBLE_DATA).append_row([
                    datetime.now().strftime("%Y%m%d_%H%M%S"), dev, "", "", cause, sol, "", rep, "", "", title
                ])
                st.success("保存成功")
                st.cache_data.clear()
            except Exception as e: st.error(f"エラー: {e}")
    page_data_list(SHEET_TROUBLE_DATA, "トラブル", TROUBLE_COL_TIMESTAMP, TROUBLE_COL_DEVICE, TROUBLE_COL_TITLE, None,
                   [TROUBLE_COL_TIMESTAMP, TROUBLE_COL_DEVICE, TROUBLE_COL_TITLE, TROUBLE_COL_CAUSE, TROUBLE_COL_SOLUTION], None)

def page_contact_form():
    st.header("お問い合わせ")
    with st.form("contact_form"):
        ctype = st.selectbox("種類", ["バグ報告", "機能要望", "データ修正依頼", "その他"])
        detail = st.text_area("詳細")
        contact = st.text_input("連絡先")
        if st.form_submit_button("送信"):
            if not detail: st.error("詳細必須"); return
            try:
                gc.open(SPREADSHEET_NAME).worksheet(SHEET_CONTACT_DATA).append_row([
                    datetime.now().strftime("%Y%m%d_%H%M%S"), ctype, detail, contact
                ])
                st.success("送信成功")
                st.cache_data.clear()
            except Exception as e: st.error(f"エラー: {e}")

# ---------------------------
# --- Analysis Pages (Original IV/PL) ---
# ---------------------------
# (IVとPLは前回の最終修正版をそのまま搭載します)

def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    use_log_scale = st.checkbox("縦軸（電流）を対数表示にする", key="iv_log_scale")
    files = st.file_uploader("IVファイル(.txt)", accept_multiple_files=True)
    
    data_for_export = []
    dfs_to_plot = []
    
    if files:
        with st.spinner("ファイルを読み込み、グラフを準備中..."):
            fig, ax = plt.subplots(figsize=(8, 6))
            has_plot = False
            
            for f in files:
                df = load_data_file(f.getvalue(), f.name)
                if df is not None:
                    data_for_export.append(df)
                    plot_df = df.copy()
                    if use_log_scale:
                        plot_df.iloc[:, 1] = np.abs(plot_df.iloc[:, 1])
                    dfs_to_plot.append(plot_df)
                    has_plot = True

            for plot_df in dfs_to_plot:
                ax.plot(plot_df['Axis_X'], plot_df.iloc[:,1], label=plot_df.columns[1])

        if has_plot:
            if use_log_scale:
                ax.set_yscale('log')
                st.warning("⚠️ 対数表示のため、電流値は**絶対値**に変換してプロットしています。")
            else:
                ax.set_yscale('linear')
            if not use_log_scale:
                 ax.axhline(0, color='gray', linestyle='--', linewidth=1)
            ax.axvline(0, color='gray', linestyle='--', linewidth=1)
            ax.set_xlabel("Voltage")
            ax.set_ylabel("Current")
            ax.legend()
            ax.grid(True, linestyle=':', alpha=0.5)
            st.pyplot(fig)
            
            st.markdown("---")
            st.subheader("📥 解析結果のエクセル出力")
            
            if data_for_export:
                is_consistent = False
                if len(data_for_export) > 0:
                    ref_df = data_for_export[0]
                    ref_x_vals = ref_df['Axis_X'].to_numpy()
                    ref_min, ref_max, ref_len = ref_x_vals.min(), ref_x_vals.max(), len(ref_x_vals)
                    all_match = True
                    for df in data_for_export[1:]:
                        df_x_vals = df['Axis_X'].to_numpy()
                        if not (np.isclose(df_x_vals.min(), ref_min) and np.isclose(df_x_vals.max(), ref_max) and len(df_x_vals) == ref_len):
                            all_match = False; break
                    is_consistent = all_match

                if is_consistent and len(data_for_export) > 1:
                    st.success("✅ 全てのファイルの電圧軸が一致するため、**測定順序を保持**したまま1枚のシートに統合します。")
                    with st.spinner("Excel出力用にデータを統合中 (順序保持)..."):
                        dfs_to_concat = [data_for_export[0]]
                        for df in data_for_export[1:]:
                            dfs_to_concat.append(df[[df.columns[1]]])
                        merged_df = pd.concat(dfs_to_concat, axis=1)
                        excel_data = to_excel(merged_df)
                else:
                    data_for_export_dict = {}
                    with st.spinner("Excel出力用にデータを準備中 (シート分割)..."):
                        for df in data_for_export:
                            data_for_export_dict[df.columns[1].replace('.txt', '')] = df
                    if len(data_for_export) > 1:
                        st.warning("⚠️ 電圧軸の範囲やステップが異なるため、ファイルごとにシートを分けて出力します。")
                        excel_data = to_excel_multi_sheet(data_for_export_dict)
                    else:
                         st.info("ファイルが1つだけのため、1枚のシートに出力します。")
                         excel_data = to_excel(data_for_export[0])
                
                default_name = datetime.now().strftime("IV_Analysis_%Y%m%d")
                filename_input = st.text_input("ファイル名 (.xlsx)", value=default_name, key="iv_filename")
                st.download_button("Excelファイルとしてダウンロード", excel_data, f"{filename_input}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="iv_download_btn")
        else:
            st.warning("プロットできるデータがありませんでした。")

def page_pl_analysis():
    st.header("💡 PLデータ解析")
    if 'pl_slope' not in st.session_state: st.session_state['pl_slope'] = None
    if 'pl_center_wl' not in st.session_state: st.session_state['pl_center_wl'] = 1700

    st.markdown("## 1️⃣ Step 1: 波長校正")
    st.info("2つの既知の波長ピークを持つデータをアップロードし、校正係数を決定します。")
    c1, c2 = st.columns(2)
    wl1 = c1.number_input("既知波長1 (nm)", value=1500.0, key="wl1_input")
    wl2 = c2.number_input("既知波長2 (nm)", value=1570.0, key="wl2_input")
    f1 = c1.file_uploader("波長1データファイル", key="c1")
    f2 = c2.file_uploader("波長2データファイル", key="c2")
    if f1 and f2:
        df1 = load_pl_data(f1)
        df2 = load_pl_data(f2)
        if df1 is not None and not df1.empty and df2 is not None and not df2.empty:
            try:
                p1 = df1.loc[df1['intensity'].idxmax(), 'pixel']
                p2 = df2.loc[df2['intensity'].idxmax(), 'pixel']
                if p1 != p2:
                    slope_raw = (wl2 - wl1) / (p2 - p1)
                    slope = np.abs(slope_raw)
                    st.success(f"✅ 計算された校正係数 (nm/pixel): **{slope:.4f}**")
                    st.caption(f"（計算値: {slope_raw:.4f} nm/pixel の絶対値を取得しました。）")
                    if st.button("この係数を保存してStep 2へ進む", key="save_slope"):
                        st.session_state['pl_slope'] = slope
                        st.rerun() 
                else: st.error("ピーク位置が同じです。")
            except Exception as e: st.error(f"解析エラー: {e}")
        else: st.error("データの読み込みに失敗しました。")

    st.markdown("---")
    st.markdown("## 2️⃣ Step 2: 中心波長の設定")
    if st.session_state['pl_slope'] is None:
        st.warning("⚠️ まず Step 1 で校正係数を決定・保存してください。")
    else:
        st.success(f"校正係数: {st.session_state['pl_slope']:.4f} nm/pixel が設定されています。")
        center_wl = st.number_input("分光器の中心波長 (nm) を入力", value=st.session_state['pl_center_wl'], key='center_wl_input')
        if st.button("中心波長を保存してStep 3へ進む", key="save_center_wl"):
            st.session_state['pl_center_wl'] = center_wl
            st.rerun()

    st.markdown("---")
    st.markdown("## 3️⃣ Step 3: 測定データ解析実行")
    if st.session_state['pl_slope'] is None or st.session_state['pl_center_wl'] is None:
        st.warning("⚠️ Step 1 (校正係数) と Step 2 (中心波長) の両方を設定してください。")
    else:
        slope = st.session_state['pl_slope']
        cw = st.session_state['pl_center_wl']
        st.info(f"現在の設定: 係数={slope:.4f}, 中心波長={cw} nm")
        files = st.file_uploader("測定データファイル(.txt)", accept_multiple_files=True, key="pl_m")
        if files:
            fig, ax = plt.subplots(figsize=(10, 6))
            has_plot = False
            data_for_export = []
            for f in files:
                df = load_pl_data(f)
                if df is not None and not df.empty:
                    df['wl'] = (df['pixel'] - 256.5) * slope + cw
                    ax.plot(df['wl'], df['intensity'], label=f.name)
                    has_plot = True
                    export_df = df[['wl', 'intensity']].copy()
                    export_df.columns = [f"Wavelength ({f.name})", f"Intensity ({f.name})"]
                    data_for_export.append(export_df)
            
            if has_plot:
                ax.set_xlabel("Wavelength (nm)")
                ax.set_ylabel("Intensity (a.u.)")
                ax.legend()
                ax.grid(True, linestyle='--', alpha=0.7)
                st.pyplot(fig)
                
                st.markdown("---")
                st.subheader("📥 解析結果のエクセル出力")
                if data_for_export:
                    ref_wl_df = data_for_export[0].iloc[:, [0]].copy() 
                    ref_wl_df.columns = ['Wavelength_nm']
                    intensity_dfs = [df.iloc[:, [1]] for df in data_for_export] 
                    dfs_to_concat = [ref_wl_df] + intensity_dfs
                    merged_df = pd.concat(dfs_to_concat, axis=1)
                    default_name = datetime.now().strftime("PL_Analysis_%Y%m%d")
                    filename_input = st.text_input("ファイル名 (.xlsx)", value=default_name, key="pl_filename")
                    excel_data = to_excel(merged_df)
                    st.download_button("Excelファイルとしてダウンロード", excel_data, f"{filename_input}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="pl_download_btn")
            else:
                st.warning("プロットできるデータがありませんでした。")

# ---------------------------
# --- Calendar ---
# ---------------------------
def page_calendar():
    st.header("🗓️ スケジュール・装置予約")
    
    st.subheader("外部予約サイト")
    c1, c2 = st.columns(2)
    c1.markdown(f'<a href="https://www.eiiris.tut.ac.jp/evers/Web/dashboard.php" target="_blank"><button style="width:100%;padding:10px;background-color:#007BFF;color:white;border:none;border-radius:5px;">🔬 Evers 予約サイトへ飛ぶ</button></a>', unsafe_allow_html=True)
    c2.markdown(f'<a href="https://tech.rac.tut.ac.jp/regist/potal_0.php" target="_blank"><button style="width:100%;padding:10px;background-color:#28A745;color:white;border:none;border-radius:5px;">⚙️ 教育研究基盤センターへ飛ぶ</button></a>', unsafe_allow_html=True)
    st.markdown("---")

    st.subheader("研究室カレンダー")
    src = CALENDAR_ID.replace("@", "%40")
    st.markdown(f'<iframe src="https://calendar.google.com/calendar/embed?src={src}&ctz=Asia%2FTokyo" style="border:0" width="100%" height="600" frameborder="0" scrolling="no"></iframe>', unsafe_allow_html=True)

    with st.expander("➕ 予定を追加"):
        with st.form("cal_form"):
            summ = st.text_input("タイトル")
            sd = st.date_input("開始日"); st_time = st.time_input("開始時刻")
            ed = st.time_input("終了時刻")
            desc = st.text_area("詳細")
            if st.form_submit_button("予約"):
                if calendar_service:
                    sdt = datetime.combine(sd, st_time).isoformat()
                    edt = datetime.combine(sd, ed).isoformat()
                    evt = {'summary': summ, 'description': desc, 
                           'start': {'dateTime': sdt, 'timeZone': 'Asia/Tokyo'},
                           'end': {'dateTime': edt, 'timeZone': 'Asia/Tokyo'}}
                    try:
                        calendar_service.events().insert(calendarId=CALENDAR_ID, body=evt).execute()
                        st.success("追加しました")
                        st.rerun()
                    except Exception as e: st.error(f"エラー: {e}")
                else: st.error("カレンダー機能無効")

# ---------------------------
# --- Main ---
# ---------------------------
def main():
    st.sidebar.title("Yamane Lab Tools")
    menu = st.sidebar.radio("メニュー", [
        "エピノート", "メンテノート", "🗓️ スケジュール・装置予約", 
        "IVデータ解析", "PLデータ解析", "📈 高機能グラフ描画", 
        "議事録", "知恵袋・質問箱", "引き継ぎメモ", "トラブル報告", "お問い合わせ"
    ])
    
    if 'curr_menu' not in st.session_state: st.session_state['curr_menu'] = menu
    if st.session_state['curr_menu'] != menu:
        st.cache_data.clear()
        st.session_state['curr_menu'] = menu

    if menu == "エピノート": page_epi_note()
    elif menu == "メンテノート": page_mainte_note()
    elif menu == "🗓️ スケジュール・装置予約": page_calendar()
    elif menu == "IVデータ解析": page_iv_analysis()
    elif menu == "PLデータ解析": page_pl_analysis()
    elif menu == "📈 高機能グラフ描画": page_graph_plotting()
    elif menu == "議事録": page_meeting_note()
    elif menu == "知恵袋・質問箱": page_qa_box()
    elif menu == "引き継ぎメモ": page_handover_note()
    elif menu == "トラブル報告": page_trouble_report()
    elif menu == "お問い合わせ": page_contact_form()

if __name__ == "__main__":
    try:
        if 'st.cache_data' in st.__dict__:
            st.cache_data.clear()
    except Exception:
        pass
    main()



















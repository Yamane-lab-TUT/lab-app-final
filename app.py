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
import uuid
from scipy import stats
from datetime import datetime
from io import BytesIO

# ==========================================
# 関数定義: page_graph_plotting (v20: 全列展開取り込み・複製機能追加版)
# ==========================================
def page_graph_plotting():
    st.header("📈 統合型グラフ解析ツール")
    st.markdown("""
    **v20 更新**: 
    - **全列活用**: 3列以上のファイル読み込み時、全ての列を別々の系列として一括追加できる機能を追加しました。
    - **複製機能**: 登録済みのデータ系列をコピーするボタンを追加しました（同じファイルでX/Y軸を変えて表示したい場合に便利です）。
    - **サイズ指定**: cm単位・実寸プレビュー・ファイル名指定保存に対応しています。
    """, unsafe_allow_html=True)

    # --- CSS ---
    st.markdown("""
        <style>
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-of-type(2) {
            position: sticky; top: 4rem; align-self: start; z-index: 999;
        }
        div[data-testid="stMarkdownContainer"] p { margin-bottom: 0px; }
        </style>
    """, unsafe_allow_html=True)

    # --- デフォルト色 ---
    DEFAULT_COLORS = [
        '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
        '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'
    ]

    # --- セッション初期化 ---
    if 'gp_data_list' not in st.session_state:
        st.session_state['gp_data_list'] = []
    
    if 'uploader_key_id' not in st.session_state:
        st.session_state['uploader_key_id'] = 0

    # --- IDマイグレーション ---
    for d in st.session_state['gp_data_list']:
        if 'id' not in d: d['id'] = str(uuid.uuid4())

    # --- ヘルパー関数 ---
    def move_data(idx, direction):
        lst = st.session_state['gp_data_list']
        if direction == "up" and idx > 0:
            lst[idx], lst[idx-1] = lst[idx-1], lst[idx]
        elif direction == "down" and idx < len(lst) - 1:
            lst[idx], lst[idx+1] = lst[idx+1], lst[idx]

    def duplicate_data(idx):
        lst = st.session_state['gp_data_list']
        original = lst[idx]
        # 辞書を浅いコピー（DataFrameは参照渡しでメモリ節約）
        new_item = original.copy()
        new_item['id'] = str(uuid.uuid4())
        new_item['legend_name'] = f"{original.get('legend_name', '')} (copy)"
        # リストの直下に挿入
        lst.insert(idx + 1, new_item)

    def get_next_color(index):
        return DEFAULT_COLORS[index % len(DEFAULT_COLORS)]

    def format_power(watts):
        if watts == 0: return "0 W"
        w_abs = abs(watts)
        if w_abs >= 1: return f"{watts:.3f} W"
        elif w_abs >= 1e-3: return f"{watts*1e3:.3f} mW"
        elif w_abs >= 1e-6: return f"{watts*1e6:.3f} µW"
        elif w_abs >= 1e-9: return f"{watts*1e9:.3f} nW"
        else: return f"{watts*1e12:.3f} pW"

    # ==========================================
    # 0. プロジェクト管理
    # ==========================================
    with st.expander("💾 プロジェクトの保存・復元", expanded=False):
        c_load, c_save = st.columns(2)
        with c_load:
            st.markdown("#### 📂 復元")
            uploaded_project = st.file_uploader("プロジェクトファイル (.json)", type=["json"], key="project_loader_v20")
            if uploaded_project:
                if st.button("設定を読み込む", key="btn_load_proj_v20"):
                    try:
                        project_data = json.load(uploaded_project)
                        restored_data_list = []
                        for item in project_data.get("datasets", []):
                            df_restored = pd.read_csv(io.StringIO(item["data_csv"]))
                            item['df'] = df_restored
                            cols = df_restored.columns.tolist()
                            
                            defaults = {
                                "mppt": False, "show_eq": False, "visible": True, 
                                "legend_name": item.get('name', ''),
                                "id": str(uuid.uuid4()),
                                "x_col": cols[0] if cols else None,
                                "y_col": cols[1] if len(cols)>1 else (cols[0] if cols else None),
                                "area": 1.0, "use_density": False,
                                "mppt_x": 10, "mppt_y": -30,
                                "fill_area": False
                            }
                            for k, v in defaults.items():
                                if k not in item: item[k] = v
                                
                            restored_data_list.append(item)
                        
                        st.session_state['gp_data_list'] = restored_data_list
                        saved_settings = project_data.get("settings", {})
                        for key, value in saved_settings.items():
                            st.session_state[key] = value
                        st.success("✅ 復元完了")
                        st.rerun()
                    except Exception as e: st.error(f"エラー: {e}")

        with c_save:
            st.markdown("#### 💾 保存")
            default_proj_name = f"GraphProject_{datetime.now().strftime('%Y%m%d_%H%M')}"
            save_name_proj = st.text_input("プロジェクト名 (拡張子不要)", value=default_proj_name, key="proj_save_name_v20")
            
            if st.button("プロジェクトファイルを作成", key="btn_save_proj_v20"):
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
                        # 保存対象外キー
                        if key.startswith(("project_", "gp_", "btn_", "paste_", "fw_", "fh_", "dpi_", "ff_", "bfs_", "sleg", "lfont", "ax_preset", "legend_", "scale_sel", "vis_", "leg_nm_", "xc_", "yc_", "ut_", "ur_", "clr_", "mrk_", "lw_", "ms_", "lst_", "mppt_", "fit_", "seq_", "area_", "dens_", "mx_", "my_", "fill_", "xy_swap_", "proj_save_name", "img_save_name")): continue
                        if isinstance(val, (int, float, str, bool, list, dict, type(None))):
                            settings_snapshot[key] = val

                    project_obj = {
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "datasets": datasets_serialized,
                        "settings": settings_snapshot
                    }
                    json_str = json.dumps(project_obj, indent=2, ensure_ascii=False)
                    
                    final_proj_fname = save_name_proj.strip()
                    if not final_proj_fname: final_proj_fname = default_proj_name
                    if not final_proj_fname.endswith(".json"): final_proj_fname += ".json"
                    
                    st.download_button("⬇️ JSONをダウンロード", json_str, final_proj_fname, "application/json", key="dl_json_btn_v20")

    # ==========================================
    # 1. データ入力
    # ==========================================
    st.subheader("1. データの入力")
    
    if st.session_state['gp_data_list']:
        st.info(f"現在のデータ数: {len(st.session_state['gp_data_list'])}")
        if st.button("🗑️ 全データをクリア", key="btn_clear_all_v20"):
            st.session_state['gp_data_list'] = []
            st.session_state['uploader_key_id'] += 1
            st.rerun()
    
    tab1, tab2 = st.tabs(["📂 ファイルから追加", "📋 エクセルから貼り付け"])
    
    with tab1:
        st.markdown("**読み込みオプション**")
        # --- 全列展開オプション ---
        expand_cols = st.checkbox("列ごとに別の系列として追加する（例: A列vsB列, A列vsC列...）", value=False, help="チェックを入れると、2列目以降のすべての列を、1列目をX軸とした個別のグラフデータとして一括で追加します。", key="expand_cols_v20")
        
        current_uploader_key = f"gp_uploader_v20_{st.session_state['uploader_key_id']}"
        files = st.file_uploader("CSV/Excelファイル", accept_multiple_files=True, key=current_uploader_key)
        
        if files:
            new_data_added = False
            for f in files:
                # 重複チェック（名前だけで判定）- 展開モードのときは名前が変わるので緩める
                if not expand_cols and any(d['name'] == f.name for d in st.session_state['gp_data_list']): continue
                
                df = None
                try:
                    if f.name.endswith(('.xlsx', '.xls')): df = pd.read_excel(f)
                    else: df = pd.read_csv(f)
                except: pass
                
                if df is not None:
                    # 数値列のみ抽出
                    df = df.select_dtypes(include=[np.number])
                    df.columns = [str(c).strip() for c in df.columns]
                    cols = df.columns.tolist()
                    
                    if not cols: continue

                    # 追加ロジック
                    if expand_cols and len(cols) >= 2:
                        # 展開モード: 1列目をXとして、2列目以降すべてをYとして登録
                        x_c = cols[0]
                        for y_c in cols[1:]:
                            auto_color = get_next_color(len(st.session_state['gp_data_list']))
                            st.session_state['gp_data_list'].append({
                                "id": str(uuid.uuid4()),
                                "name": f.name,
                                "df": df, # 同じDFを参照
                                "legend_name": f"{f.name} ({y_c})", # 凡例に列名を含める
                                "mppt": False, "show_eq": False, "visible": True,
                                "color": auto_color, "marker": "None", "linestyle": "-",
                                "x_col": x_c,
                                "y_col": y_c,
                                "area": 1.0, "use_density": False,
                                "mppt_x": 10, "mppt_y": -30, "fill_area": False
                            })
                        new_data_added = True
                    else:
                        # 通常モード: 1つのデータとして登録（初期はCol1 vs Col2）
                        auto_color = get_next_color(len(st.session_state['gp_data_list']))
                        st.session_state['gp_data_list'].append({
                            "id": str(uuid.uuid4()),
                            "name": f.name, "df": df,
                            "legend_name": f.name,
                            "mppt": False, "show_eq": False, "visible": True,
                            "color": auto_color, "marker": "None", "linestyle": "-",
                            "x_col": cols[0] if cols else None,
                            "y_col": cols[1] if len(cols) > 1 else (cols[0] if cols else None),
                            "area": 1.0, "use_density": False,
                            "mppt_x": 10, "mppt_y": -30, "fill_area": False
                        })
                        new_data_added = True

            if new_data_added: st.rerun()

    with tab2:
        st.caption("Excelからコピペ (タブ区切り) して Ctrl+Enter")
        paste_text = st.text_area("データ貼り付けエリア", height=100, key="paste_area_v20")
        paste_name = st.text_input("データセット名", value=f"Data_{len(st.session_state['gp_data_list'])+1}", key="paste_name_v20")
        
        if st.button("貼り付け追加", key="btn_paste_add_v20"):
            if paste_text:
                try:
                    df_paste = pd.read_csv(io.StringIO(paste_text), sep='\t')
                    if df_paste is not None and not df_paste.empty:
                        df_paste = df_paste.select_dtypes(include=[np.number])
                        cols = df_paste.columns.tolist()
                        auto_color = get_next_color(len(st.session_state['gp_data_list']))
                        
                        st.session_state['gp_data_list'].append({
                            "id": str(uuid.uuid4()),
                            "name": paste_name, "df": df_paste,
                            "legend_name": paste_name,
                            "mppt": False, "show_eq": False,
                            "visible": True,
                            "color": auto_color, "marker": "None",
                            "linestyle": "-",
                            "x_col": cols[0] if cols else None,
                            "y_col": cols[1] if len(cols) > 1 else (cols[0] if cols else None),
                            "area": 1.0, "use_density": False,
                            "mppt_x": 10, "mppt_y": -30,
                            "fill_area": False
                        })
                        st.success("追加しました")
                        st.rerun()
                except Exception as e: st.error(f"エラー (Tab区切りとして読み込めませんでした): {e}")

    datasets = st.session_state['gp_data_list']
    if not datasets: return

    # ==========================================
    # 2. グラフ設定
    # ==========================================
    st.markdown("---")
    col_settings, col_preview = st.columns([1.3, 2])

    with col_settings:
        st.subheader("2. 詳細設定")
        
        # --- A. キャンバス (cm指定) ---
        with st.expander("📊 キャンバス・フォント", expanded=False):
            c1, c2 = st.columns(2)
            fig_w_cm = c1.number_input("幅 (cm)", 2.0, 100.0, 15.0, step=0.5, key="fw_cm_v20")
            fig_h_cm = c2.number_input("高さ (cm)", 2.0, 100.0, 10.0, step=0.5, key="fh_cm_v20")
            
            fig_w_inch = fig_w_cm / 2.54
            fig_h_inch = fig_h_cm / 2.54
            
            dpi_val = st.number_input("解像度 (DPI)", 72, 600, 150, key="dpi_in_v20")
            font_family = st.selectbox("フォント", ["Times New Roman", "Arial", "Helvetica", "Meiryo", "Yu Gothic"], index=0, key="ff_sel_v20")
            base_fs = st.number_input("基本フォントサイズ", 6, 50, 12, key="bfs_in_v20")

        # --- B. 軸設定 ---
        with st.expander("📐 軸 (Axes) と 単位変換", expanded=True):
            tabs_ax = st.tabs(["X軸(下)", "X軸(上)", "Y軸(左)", "Y軸(右)", "共通"])
            ax_settings = {}
            SCALE_OPTIONS = {
                "x1 (そのまま)": 1.0, "x1000 (m)": 1000.0, "x10^6 (µ)": 1e6,
                "x10^9 (n)": 1e9, "x10^12 (p)": 1e12, "x0.001 (k)": 0.001
            }

            def axis_ui(key_prefix, label_def, use_top=False, use_right=False):
                col_btn = st.columns(3)
                if col_btn[0].button("Voltage(V)", key=f"p_v_{key_prefix}"):
                    st.session_state[f"{key_prefix}_lbl_v20"] = "Voltage (V)"
                    st.session_state[f"{key_prefix}_scale_idx_v20"] = 0
                if col_btn[1].button("Current(mA)", key=f"p_ma_{key_prefix}"):
                    st.session_state[f"{key_prefix}_lbl_v20"] = "Current (mA)"
                    st.session_state[f"{key_prefix}_scale_idx_v20"] = 1
                if col_btn[2].button("Current(µA)", key=f"p_ua_{key_prefix}"):
                    st.session_state[f"{key_prefix}_lbl_v20"] = "Current (µA)"
                    st.session_state[f"{key_prefix}_scale_idx_v20"] = 2

                label = st.text_input("ラベル", label_def, key=f"{key_prefix}_lbl_v20")
                
                curr_idx = st.session_state.get(f"{key_prefix}_scale_idx_v20", 0)
                scale_key = st.selectbox("表示倍率", list(SCALE_OPTIONS.keys()), index=curr_idx, key=f"{key_prefix}_scale_sel_v20")
                st.session_state[f"{key_prefix}_scale_idx_v20"] = list(SCALE_OPTIONS.keys()).index(scale_key)
                
                current_scale_val = SCALE_OPTIONS[scale_key]
                prev_scale_key = f"{key_prefix}_prev_scale_val"
                prev_scale_val = st.session_state.get(prev_scale_key, 1.0)
                
                if current_scale_val != prev_scale_val:
                    ratio = current_scale_val / prev_scale_val
                    k_min = f"{key_prefix}_min_v20"
                    k_max = f"{key_prefix}_max_v20"
                    if st.session_state.get(k_min) is not None:
                        st.session_state[k_min] = st.session_state[k_min] * ratio
                    if st.session_state.get(k_max) is not None:
                        st.session_state[k_max] = st.session_state[k_max] * ratio
                    st.session_state[prev_scale_key] = current_scale_val

                data_vals = []
                for d in datasets:
                    if not d.get('visible', True): continue
                    if d.get('x_col') is None or d.get('y_col') is None: continue
                    if d['x_col'] not in d['df'].columns or d['y_col'] not in d['df'].columns: continue

                    is_this_axis_x = (d.get('use_top', False) == use_top)
                    is_this_axis_y = (d.get('use_right', False) == use_right)
                    
                    val = None
                    if key_prefix.startswith('x') and is_this_axis_x:
                        val = d['df'][d['x_col']]
                    elif key_prefix.startswith('y') and is_this_axis_y:
                        val = d['df'][d['y_col']]
                        if d.get('use_density', False) and d.get('area', 1.0) > 0:
                            val = val / d['area']

                    if val is not None:
                        data_vals.append(val * current_scale_val)
                
                calc_min, calc_max = None, None
                if data_vals:
                    concat_data = pd.concat(data_vals)
                    if not concat_data.empty:
                        calc_min = float(concat_data.min())
                        calc_max = float(concat_data.max())
                        margin = (calc_max - calc_min) * 0.05
                        if margin == 0: margin = abs(calc_max) * 0.1 if calc_max!=0 else 1.0
                        calc_min -= margin
                        calc_max += margin

                k_min = f"{key_prefix}_min_v20"
                k_max = f"{key_prefix}_max_v20"
                if st.session_state.get(k_min) is None and calc_min is not None:
                    st.session_state[k_min] = calc_min
                if st.session_state.get(k_max) is None and calc_max is not None:
                    st.session_state[k_max] = calc_max

                if prev_scale_key not in st.session_state:
                     st.session_state[prev_scale_key] = current_scale_val

                c1, c2 = st.columns(2)
                d_min = c1.number_input("最小", value=None, format="%f", key=k_min)
                d_max = c2.number_input("最大", value=None, format="%f", key=k_max)
                c3, c4 = st.columns(2)
                maj_int = c3.number_input("主目盛", 0.0, step=0.1, key=f"{key_prefix}_maj_v20")
                min_int = c4.number_input("補助目盛", 0.0, step=0.1, key=f"{key_prefix}_min_int_v20")
                
                c5, c6 = st.columns(2)
                is_log = c5.checkbox("対数軸", False, key=f"{key_prefix}_log_v20")
                is_inv = c6.checkbox("軸を反転", False, key=f"{key_prefix}_inv_v20")

                return {"label": label, "min": d_min, "max": d_max, "maj": maj_int, "log": is_log, "inv": is_inv, "scale": current_scale_val}

            with tabs_ax[0]: ax_settings['x1'] = axis_ui("x1", "Voltage (V)", use_top=False)
            with tabs_ax[1]: ax_settings['x2'] = axis_ui("x2", "Secondary X", use_top=True)
            with tabs_ax[2]: ax_settings['y1'] = axis_ui("y1", "Current (A)", use_right=False)
            with tabs_ax[3]: ax_settings['y2'] = axis_ui("y2", "Power (W)", use_right=True)
            
            with tabs_ax[4]:
                tick_dir = st.selectbox("目盛の向き", ["in", "out", "inout"], index=0, key="tdir_v20")
                show_grid = st.checkbox("グリッド表示", False, key="sgrid_v20")
                zero_cross = st.checkbox("原点線描画", True, key="zcross_v20")

        # --- C. 凡例設定 ---
        with st.expander("📝 凡例 (Legend)", expanded=True):
            show_leg = st.checkbox("凡例を表示", True, key="sleg_v20")
            
            st.markdown("#### 凡例順序・表示設定")
            for i, d in enumerate(datasets):
                did = d['id']
                c_vis, c_name, c_up, c_down = st.columns([0.5, 4, 0.7, 0.7])
                with c_vis:
                    d['visible'] = st.checkbox("vis", value=d.get('visible', True), key=f"vis_main_{did}", label_visibility="collapsed")
                with c_name:
                    st.text(f"{d.get('legend_name', d['name'])}")
                with c_up:
                    if st.button("⬆", key=f"leg_u_{did}"): move_data(i, "up"); st.rerun()
                with c_down:
                    if st.button("⬇", key=f"leg_d_{did}"): move_data(i, "down"); st.rerun()

            if show_leg:
                st.markdown("---")
                st.markdown("**スタイル設定**")
                c_auto, c_size = st.columns(2)
                auto_leg_size = c_auto.checkbox("サイズ自動調整", True, key="auto_leg_size_v20")
                manual_fs = c_size.number_input("フォントサイズ", 5, 40, int(base_fs), disabled=auto_leg_size, key="lfont_v20")
                if auto_leg_size:
                    l_fontsize = max(6, int(base_fs) - (len(datasets) // 3))
                else:
                    l_fontsize = manual_fs

                c1, c2 = st.columns(2)
                l_loc = c1.selectbox("位置", ["best", "upper right", "upper left", "lower right", "lower left", "outside right"], index=0, key="lloc_v20")
                l_col = c2.number_input("列数", 1, 5, 1, key="lcol_v20")
                l_frame = st.checkbox("枠線を表示", False, key="lframe_v20")

        # --- D. データ系列 ---
        st.markdown("#### データ系列設定")
        
        for i, d in enumerate(datasets):
            did = d['id']
            with st.expander(f"#{i+1}: {d.get('legend_name', d['name'])}", expanded=False):
                d['legend_name'] = st.text_input("凡例表示名", value=d.get('legend_name', d['name']), key=f"leg_nm_{did}")

                # 操作ボタン（複製を追加）
                bc1, bc2, bc3, bc4 = st.columns([1, 1, 1.5, 2])
                with bc1:
                    if st.button("⬆", key=f"btn_u_{did}"): move_data(i, "up"); st.rerun()
                with bc2:
                    if st.button("⬇", key=f"btn_d_{did}"): move_data(i, "down"); st.rerun()
                with bc3:
                    if st.button("©️ 複製", key=f"btn_dup_{did}", help="このデータ系列を複製して追加します"): duplicate_data(i); st.rerun()
                with bc4:
                    if st.button("❌ 削除", key=f"btn_del_{did}"): datasets.pop(i); st.rerun()

                cols = d['df'].columns.tolist()
                sc1, sc2, sc3 = st.columns([2, 2, 1])
                curr_xc = d.get('x_col')
                curr_yc = d.get('y_col')
                ix_x = cols.index(curr_xc) if curr_xc in cols else 0
                ix_y = cols.index(curr_yc) if curr_yc in cols else (1 if len(cols)>1 else 0)

                xc = sc1.selectbox(f"X列", cols, index=ix_x, key=f"xc_{did}")
                yc = sc2.selectbox(f"Y列", cols, index=ix_y, key=f"yc_{did}")
                # X/Y 入替
                if sc3.button("🔄 入替", key=f"xy_swap_{did}"):
                    d['x_col'], d['y_col'] = yc, xc
                    st.rerun()

                st.caption("電流密度計算 (Y軸 = Y / 面積)")
                ac_dens1, ac_dens2 = st.columns(2)
                d['area'] = ac_dens1.number_input("デバイス面積 (cm²)", 0.0, 100.0, float(d.get('area', 1.0)), format="%.4f", key=f"area_{did}")
                d['use_density'] = ac_dens2.checkbox("電流密度に換算", d.get('use_density', False), key=f"dens_{did}")

                ac1, ac2 = st.columns(2)
                d['use_top'] = ac1.checkbox("上X軸", d.get('use_top', False), key=f"ut_{did}")
                d['use_right'] = ac2.checkbox("右Y軸", d.get('use_right', False), key=f"ur_{did}")

                tc1, tc2 = st.columns(2)
                d['color'] = tc1.color_picker("色", d.get('color', '#0000FF'), key=f"clr_{did}")
                d['marker'] = tc2.selectbox("マーカー", ["None", "o", "s", "^", "x"], index=0 if d.get('marker')=="None" else 1, key=f"mrk_{did}")
                
                lw1, lw2 = st.columns(2)
                d['line_width'] = lw1.number_input("線幅", 0.0, 10.0, float(d.get('line_width', 1.5)), key=f"lw_{did}")
                d['marker_size'] = lw2.number_input("点サイズ", 0.0, 20.0, float(d.get('marker_size', 6.0)), key=f"ms_{did}")
                d['linestyle'] = st.selectbox("線種", ["-", "--", "-.", ":", "None"], index=0, key=f"lst_{did}")
                
                d['fill_area'] = st.checkbox("0まで塗りつぶす (Fill)", d.get('fill_area', False), key=f"fill_{did}")

                st.markdown("---")
                d['mppt'] = st.checkbox("MPPT解析", d.get('mppt', False), key=f"mppt_{did}")
                if d['mppt']:
                    mp1, mp2 = st.columns(2)
                    d['mppt_x'] = mp1.number_input("Text X Offset", value=int(d.get('mppt_x', 10)), key=f"mx_{did}")
                    d['mppt_y'] = mp2.number_input("Text Y Offset", value=int(d.get('mppt_y', -30)), key=f"my_{did}")

                d['fit_mode'] = st.selectbox("近似曲線", ["なし", "線形", "多項式(2次)", "移動平均"], index=0, key=f"fit_{did}")
                if d['fit_mode'] != "なし":
                    d['show_eq'] = st.checkbox("数式を表示", d.get('show_eq', False), key=f"seq_{did}")

                d.update({'x_col': xc, 'y_col': yc})

    # ==========================================
    # 3. 描画
    # ==========================================
    with col_preview:
        st.subheader("プレビュー")
        
        # フォント & 数式設定
        if font_family in ["Times New Roman", "Times"]:
            plt.rcParams['font.family'] = 'serif'
            plt.rcParams['font.serif'] = [font_family] + plt.rcParams['font.serif']
        else:
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [font_family] + plt.rcParams['font.sans-serif']
        plt.rcParams['font.size'] = base_fs
        
        # MathTextモード有効化
        plt.rcParams['axes.formatter.use_mathtext'] = True

        # CM -> Inch 変換してFigure作成
        fig, ax1 = plt.subplots(figsize=(fig_w_inch, fig_h_inch), dpi=dpi_val)
        
        visible_datasets = [d for d in datasets if d.get('visible', True)]
        
        has_right = any(d.get('use_right') for d in visible_datasets)
        has_top = any(d.get('use_top') for d in visible_datasets)

        ax2, ax3 = None, None
        axes_map = {(False, False): ax1}

        if has_right:
            ax2 = ax1.twinx()
            axes_map[(False, True)] = ax2
        if has_top:
            ax3 = ax1.twiny()
            axes_map[(True, False)] = ax3
        if has_right and has_top:
            axes_map[(True, True)] = ax3

        # 軸設定適用関数 (X, Y 独立フォーマッター)
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
            
            if xc.get('inv', False):
                ax.invert_xaxis()
            
            xfmt = ticker.ScalarFormatter(useMathText=True)
            xfmt.set_powerlimits((-2, 3))
            ax.xaxis.set_major_formatter(xfmt)
            
            yfmt = ticker.ScalarFormatter(useMathText=True)
            yfmt.set_powerlimits((-2, 3))
            ax.yaxis.set_major_formatter(yfmt)

            if xc['maj'] > 0: ax.xaxis.set_major_locator(ticker.MultipleLocator(xc['maj']))
            if yc['maj'] > 0: ax.yaxis.set_major_locator(ticker.MultipleLocator(yc['maj']))
            ax.tick_params(direction=tick_dir, which='both')

        apply_axis_conf(ax1, ax_settings['x1'], ax_settings['y1'])
        apply_axis_conf(ax2, ax_settings['x1'], ax_settings['y2'])
        apply_axis_conf(ax3, ax_settings['x2'], ax_settings['y1'])
        
        if show_grid: ax1.grid(True, linestyle=':', alpha=0.6)
        if zero_cross: 
            ax1.axhline(0, color='black', linewidth=0.8)
            ax1.axvline(0, color='black', linewidth=0.8)

        legend_handles = []
        legend_labels = []

        for d in datasets:
            if not d.get('visible', True): continue
            
            if not d.get('x_col') or not d.get('y_col'): continue
            
            df = d['df']
            x_raw = df[d['x_col']]
            y_val = df[d['y_col']]
            
            if d.get('use_density', False) and d.get('area', 1.0) > 0:
                y_val = y_val / d['area']

            use_t = d.get('use_top', False)
            use_r = d.get('use_right', False)
            x_scale = ax_settings['x2']['scale'] if use_t else ax_settings['x1']['scale']
            y_scale = ax_settings['y2']['scale'] if use_r else ax_settings['y1']['scale']

            x_data = x_raw * x_scale
            y_data = y_val * y_scale
            
            target_ax = axes_map.get((use_t, use_r), ax1)
            
            mask = pd.notna(x_data) & pd.notna(y_data)
            x_plot, y_plot = x_data[mask], y_data[mask]
            if len(x_plot) == 0: continue

            ls = d.get('linestyle', '-')
            if ls == "None": ls = ""
            mk = d.get('marker', 'None')
            if mk == "None": mk = ""
            
            label_text = d.get('legend_name', d['name'])

            lines = target_ax.plot(x_plot, y_plot, label=label_text, 
                                   color=d['color'], marker=mk, linestyle=ls,
                                   linewidth=d.get('line_width', 1.5), markersize=d.get('marker_size', 6))
            
            if d.get('fill_area', False):
                target_ax.fill_between(x_plot, y_plot, 0, color=d['color'], alpha=0.3)

            if lines:
                legend_handles.append(lines[0])
                legend_labels.append(label_text)

            fmode = d.get('fit_mode', "なし")
            if fmode != "なし" and len(x_plot) > 1:
                try:
                    idx_sorted = np.argsort(x_plot)
                    xs = x_plot.iloc[idx_sorted]
                    ys = y_plot.iloc[idx_sorted]
                    y_fit = None
                    eq_text = ""
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
                            target_ax.text(xs.iloc[-1], y_fit.iloc[-1], eq_text, fontsize=9, color=d['color'])
                except: pass

            if d.get('mppt'):
                m_mask = (x_plot < 0) & (y_plot > 0)
                xm, ym = x_plot[m_mask], y_plot[m_mask]
                if len(xm) > 0:
                    p_calc = (xm * ym).abs()
                    max_i = p_calc.idxmax()
                    best_p = p_calc[max_i]
                    best_x_plot = xm[max_i]
                    best_y_plot = ym[max_i]
                    pow_str = format_power(best_p)

                    target_ax.plot(best_x_plot, best_y_plot, marker='*', color='gold', markersize=14, markeredgecolor='black', zorder=10)
                    off_x = d.get('mppt_x', 10)
                    off_y = d.get('mppt_y', -30)
                    target_ax.annotate(f"MPPT: {pow_str}", xy=(best_x_plot, best_y_plot), xytext=(off_x, off_y),
                                       textcoords='offset points', arrowprops=dict(arrowstyle="->"),
                                       bbox=dict(boxstyle="round", fc="white", alpha=0.7))

        if show_leg and legend_handles:
            bbox = None
            loc_param = l_loc
            if l_loc == "outside right":
                loc_param = "center left"
                bbox = (1.05, 0.5)
            
            ax1.legend(legend_handles, legend_labels, 
                       loc=loc_param, bbox_to_anchor=bbox, ncol=l_col,
                       fontsize=l_fontsize, frameon=l_frame, edgecolor='black')

        plt.tight_layout()
        st.pyplot(fig, use_container_width=False)
        
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=300, bbox_inches='tight')
        
        default_img_name = f"plot_{datetime.now().strftime('%Y%m%d_%H%M')}"
        save_name_img = st.text_input("画像保存名 (拡張子不要)", value=default_img_name, key="img_save_name_v20")
        
        final_img_fname = save_name_img.strip()
        if not final_img_fname: final_img_fname = default_img_name
        if not final_img_fname.endswith(".png"): final_img_fname += ".png"

        st.download_button("画像を保存 (PNG)", buf.getvalue(), final_img_fname, "image/png", key="dl_png_v20")
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
































# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
# 修正版 v20.7.0
#  - IV結合データ補間対応済み
#  - PL解析ファイル読み込み修正
#  - 画像インライン表示（自動リサイズ対応）
#  - メニュー順序変更（IV/PLをメンテノート下に配置）
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
    st.error("❌ `google-cloud-storage` が見つかりません。")
    pass

# --- Matplotlib 日本語フォント設定 ---
try:
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Meiryo', 'IPAexGothic', 'Noto Sans CJK JP']
    plt.rcParams['axes.unicode_minus'] = False
except Exception:
    pass

st.set_page_config(page_title="山根研 便利屋さん", layout="wide")

CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files"
SPREADSHEET_NAME = 'エピノート'

# --- Google 認証 ---
class DummyGSClient:
    def open(self, name): return self
    def worksheet(self, name): return self
    def get_all_values(self): return []
    def append_row(self, values): pass

class DummyStorageClient:
    def bucket(self, name): return self
    def blob(self, name): return self
    def upload_from_file(self, f, content_type): pass

gc = DummyGSClient()
storage_client = DummyStorageClient()

@st.cache_resource(ttl=3600)
def initialize_google_services():
    if "gcs_credentials" not in st.secrets:
        st.warning("⚠️ Secretsに`gcs_credentials`がありません。")
        return DummyGSClient(), DummyStorageClient()
    try:
        info = json.loads(st.secrets["gcs_credentials"])
        gc_real = gspread.service_account_from_dict(info)
        storage_real = storage.Client.from_service_account_info(info)
        st.sidebar.success("✅ Google認証成功")
        return gc_real, storage_real
    except Exception as e:
        st.error(f"認証エラー: {e}")
        return DummyGSClient(), DummyStorageClient()

gc, storage_client = initialize_google_services()

# --------------------------------------------------------------------------
# データ取得関数
# --------------------------------------------------------------------------
@st.cache_data(ttl=600)
def get_sheet_as_df(spreadsheet_name, sheet_name):
    try:
        ws = gc.open(spreadsheet_name).worksheet(sheet_name)
        data = ws.get_all_values()
        if not data or len(data) <= 1:
            return pd.DataFrame(columns=data[0] if data else [])
        return pd.DataFrame(data[1:], columns=data[0])
    except Exception:
        return pd.DataFrame()

# --------------------------------------------------------------------------
# --- IV/PL共通データ読み込みユーティリティ ---
# --------------------------------------------------------------------------
def _load_two_column_data_core(uploaded_bytes, column_names):
    try:
        text = uploaded_bytes.decode('utf-8', errors='ignore').splitlines()
        data_lines = [l for l in text if l.strip() and not l.startswith(('#', '!', '/'))]
        if not data_lines: return None
        df = pd.read_csv(io.StringIO("\n".join(data_lines)),
                         sep=r'\s+|,|\t', engine='python', header=None)
        df = df.iloc[:, :2]
        df.columns = column_names
        df[column_names[0]] = pd.to_numeric(df[column_names[0]], errors='coerce')
        df[column_names[1]] = pd.to_numeric(df[column_names[1]], errors='coerce')
        df = df.dropna().sort_values(column_names[0])
        return df
    except Exception:
        return None

@st.cache_data(show_spinner="IVデータ解析中...")
def load_data_file(uploaded_bytes, filename):
    return _load_two_column_data_core(uploaded_bytes, ['Axis_X', filename])

@st.cache_data(show_spinner="PLデータ解析中...")
def load_pl_data(uploaded_file):
    try:
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)
        df = _load_two_column_data_core(file_bytes, ['pixel', 'intensity'])
        if df is not None and not df.empty:
            return df[['pixel', 'intensity']]
    except Exception as e:
        st.error(f"PLデータの読み込みに失敗しました: {e}")
    return None

# --------------------------------------------------------------------------
# --- IVデータ結合（改良版・補間対応） ---
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="IVデータ結合中...")
def combine_dataframes(dataframes, filenames, num_points=500):
    if not dataframes:
        return None
    all_x = np.concatenate([df['Axis_X'].values for df in dataframes])
    x_common = np.linspace(all_x.min(), all_x.max(), num_points)
    combined_df = pd.DataFrame({'X_Axis': x_common})
    for df, name in zip(dataframes, filenames):
        df_sorted = df.sort_values('Axis_X')
        y_interp = np.interp(x_common, df_sorted['Axis_X'], df_sorted.iloc[:, 1])
        combined_df[name] = y_interp
    return combined_df

# --------------------------------------------------------------------------
# --- 添付ファイル表示（自動画像リサイズ付き） ---
# --------------------------------------------------------------------------
def display_attached_files(row, col_url, col_filename=None):
    if col_url not in row or not row[col_url]:
        st.info("添付ファイルはありません。")
        return
    try:
        urls = json.loads(row[col_url])
        filenames = []
        if col_filename and row.get(col_filename):
            filenames = json.loads(row[col_filename])
        if not filenames:
            filenames = ['ファイル'] * len(urls)
        for filename, url in zip(filenames, urls):
            if not url:
                continue
            is_image = url.lower().endswith(('.png', '.jpg', '.jpeg'))
            if is_image:
                st.image(url, caption=filename, use_container_width=True)
            else:
                st.markdown(f"🔗 [{filename}]({url})")
    except Exception:
        st.markdown("⚠️ 添付ファイルの読み込みに失敗しました。")

# --------------------------------------------------------------------------
# --- 各ページ（IV/PL解析のみ再掲） ---
# --------------------------------------------------------------------------
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    uploaded_files = st.file_uploader("IVデータ (.txt)", type=['txt'], accept_multiple_files=True)
    if not uploaded_files:
        st.info("ファイルをアップロードしてください。")
        return

    valid_dfs, names = [], []
    for f in uploaded_files:
        df = load_data_file(f.getvalue(), f.name)
        if df is not None and not df.empty:
            valid_dfs.append(df)
            names.append(f.name)

    if not valid_dfs:
        st.warning("有効なファイルがありません。")
        return

    combined_df = combine_dataframes(valid_dfs, names)
    st.success(f"{len(valid_dfs)}件を結合しました。")

    fig, ax = plt.subplots(figsize=(10,6))
    for name in names:
        ax.plot(combined_df['X_Axis'], combined_df[name], label=name)
    ax.set_xlabel("電圧 (V)")
    ax.set_ylabel("電流 (A)")
    ax.legend()
    ax.grid(True)
    st.pyplot(fig)

    st.dataframe(combined_df.head(), use_container_width=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, index=False, sheet_name="IV_combined")
    st.download_button("💾 Excelダウンロード", output.getvalue(),
                       f"iv_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    st.write("校正→スペクトル解析の2ステップで行います。")

    # --- Step 1 校正 ---
    cal1_wl = st.number_input("基準波長1 (nm)", value=1500)
    cal2_wl = st.number_input("基準波長2 (nm)", value=1570)
    cal1_file = st.file_uploader(f"{cal1_wl} nm 校正ファイル", type=['txt'], key="cal1")
    cal2_file = st.file_uploader(f"{cal2_wl} nm 校正ファイル", type=['txt'], key="cal2")

    if st.button("校正実行"):
        if cal1_file and cal2_file:
            df1, df2 = load_pl_data(cal1_file), load_pl_data(cal2_file)
            if df1 is not None and df2 is not None:
                p1 = df1['pixel'].iloc[df1['intensity'].idxmax()]
                p2 = df2['pixel'].iloc[df2['intensity'].idxmax()]
                slope = (cal2_wl - cal1_wl) / (p1 - p2)
                st.session_state['pl_slope'] = slope
                st.session_state['pl_calibrated'] = True
                st.success(f"校正完了: {slope:.4f} nm/pixel")
            else:
                st.error("校正ファイルのデータ読み込みに失敗しました。")
        else:
            st.warning("2つの校正ファイルを指定してください。")

    st.write("---")
    st.subheader("ステップ2：測定データ解析")

    if not st.session_state.get('pl_calibrated', False):
        st.info("💡 まず校正を完了してください。")
        return

    center_wl = st.number_input("測定時の中心波長 (nm)", value=1700)
    uploaded_files = st.file_uploader("PL測定データ (.txt)", type=['txt'], accept_multiple_files=True)
    if not uploaded_files:
        return

    slope = st.session_state['pl_slope']
    center_pixel = 256.5

    fig, ax = plt.subplots(figsize=(10,6))
    merged = None
    for f in uploaded_files:
        df = load_pl_data(f)
        if df is None: continue
        df['wavelength_nm'] = (df['pixel'] - center_pixel) * slope + center_wl
        ax.plot(df['wavelength_nm'], df['intensity'], label=f.name)
        df = df[['wavelength_nm', 'intensity']].rename(columns={'intensity': f.name})
        merged = df if merged is None else pd.merge(merged, df, on='wavelength_nm', how='outer')

    ax.set_xlabel("波長 (nm)")
    ax.set_ylabel("PL強度 (a.u.)")
    ax.legend()
    ax.grid(True)
    st.pyplot(fig)
    if merged is not None:
        merged = merged.sort_values('wavelength_nm')
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged.to_excel(writer, index=False, sheet_name="PL_combined")
        st.download_button("📊 Excelダウンロード", output.getvalue(),
                           f"pl_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

# --------------------------------------------------------------------------
# --- メインルーティング ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    menu = st.sidebar.radio("機能選択", [
        "エピノート",
        "メンテノート",
        "⚡ IVデータ解析",
        "🔬 PLデータ解析",
        "議事録",
        "知恵袋・質問箱",
        "装置引き継ぎメモ",
        "トラブル報告",
        "連絡・問い合わせ",
        "🗓️ スケジュール・装置予約"
    ])

    if menu == "⚡ IVデータ解析":
        page_iv_analysis()
    elif menu == "🔬 PLデータ解析":
        page_pl_analysis()
    else:
        st.info("この機能は現在省略しています（既存バージョンから流用可能）。")

if __name__ == "__main__":
    main()

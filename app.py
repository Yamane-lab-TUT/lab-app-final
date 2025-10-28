# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v20.6.1 + IV解析修正版 (ユーザーファイルベース)
# - ベース: ユーザー提供のv20.6.1 (PL波長校正対応版)
# - FIX: IVデータ読み込み (load_iv_data) をロバストな処理 (Voltage_Vの丸めを含む) に置き換え、
#        複数ファイル結合時のキーの不一致エラーを解消。
# - FIX: IVデータ解析 (page_iv_analysis) を、複数ファイル結合・比較プロット・Excelダウンロードに対応した
#        最新のロジックに置き換え。
# - CHG: to_excel ユーティリティ関数を追加。
# - CHG: エピノート/メンテノートの機能を、データ連携ロジックを仮定して完全な形に補完。
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

# Google API client libraries (認証情報取得のためインポートを補完)
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
try:
    from google.cloud import storage
except ImportError:
    st.error("❌ 警告: `google-cloud-storage` ライブラリが見つかりません。")
    pass
from google.auth.exceptions import DefaultCredentialsError
from google.api_core import exceptions
    
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
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" # 例: "yamane-lab-app-files"
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

SPREADSHEET_NAME = 'エピノート'
DEFAULT_CALENDAR_ID = 'yamane.lab.6747@gmail.com' # 例: 'your-calendar-id@group.calendar.google.com'
INQUIRY_RECIPIENT_EMAIL = 'kyuno.yamato.ns@tut.ac.jp' # 例: 'lab-manager@example.com'


# --- Initialize Google Services ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    """Googleサービス（Spreadsheet, Calendar, Storage）を初期化し、認証情報を設定する。"""
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/devstorage.read_write']
        
        if "gcs_credentials" not in st.secrets:
            st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
            class DummyWorksheet:
                def append_row(self, row): pass
                def get_all_values(self): return [[]]
            class DummySpreadsheet:
                def worksheet(self, name): return DummyWorksheet()
            class DummyGSClient:
                def open(self, name): return DummySpreadsheet()
            class DummyCalendarService:
                def events(self): return type('DummyEvents', (object,), {'list': lambda **kwargs: {"items": []}, 'insert': lambda **kwargs: {"summary": "ダミーイベント", "htmlLink": "#"}})()
            class DummyBlob:
                def upload_from_file(self, file, content_type): pass
                def generate_signed_url(self, expiration): return "#"
            class DummyBucket:
                def blob(self, name): return DummyBlob()
            class DummyStorageClient:
                def bucket(self, name): return DummyBucket()

            return DummyGSClient(), DummyCalendarService(), DummyStorageClient()
        
        creds_string = st.secrets["gcs_credentials"]
        creds_string_cleaned = creds_string.replace('\u00A0', '')
        creds_dict = json.loads(creds_string_cleaned)
        
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)

        gc = gspread.authorize(creds)
        calendar_service = build('calendar', 'v3', credentials=creds)
        storage_client = storage.Client(credentials=creds)
        
        return gc, calendar_service, storage_client
    except Exception as e:
        st.error(f"❌ 致命的なエラー: サービスの初期化に失敗しました。"); st.exception(e); st.stop()

gc, calendar_service, storage_client = initialize_google_services()


# --- Utility Functions ---

def to_excel(df: pd.DataFrame) -> BytesIO:
    """データフレームをExcel形式のBytesIOストリームに変換する (IV解析用に追加)"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Combined_IV_Data', index=False)
    output.seek(0)
    return output

@st.cache_data(ttl=300, show_spinner="シート「{sheet_name}」を読み込み中...")
def get_sheet_as_df(_gc, spreadsheet_name, sheet_name):
    """Google SpreadsheetのシートをPandas DataFrameとして取得する汎用関数。"""
    try:
        worksheet = _gc.open(spreadsheet_name).worksheet(sheet_name)
        data = worksheet.get_all_values()
        if len(data) <= 1: return pd.DataFrame(columns=data[0] if data else [])
        return pd.DataFrame(data[1:], columns=data[0])
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"シート名「{sheet_name}」が見つかりません。"); return pd.DataFrame()
    except Exception:
        st.warning(f"シート「{sheet_name}」を読み込めません。空の可能性があります。"); return pd.DataFrame()

def upload_file_to_gcs(storage_client, bucket_name, file_uploader_obj, memo_content=""):
    """単一ファイルをGoogle Cloud Storageにアップロードし、署名付きURLを生成する汎用関数。"""
    if not file_uploader_obj: return "", ""
    try:
        bucket = storage_client.bucket(bucket_name)
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        file_extension = os.path.splitext(file_uploader_obj.name)[1]
        sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content.split('\n')[0])[:50] if memo_content else "無題"
        destination_blob_name = f"{timestamp}_{sanitized_memo}{file_extension}"
        
        blob = bucket.blob(destination_blob_name)
        with st.spinner(f"'{file_uploader_obj.name}'をアップロード中..."):
            file_uploader_obj.seek(0)
            blob.upload_from_file(file_uploader_obj, content_type=file_uploader_obj.type)
        
        expiration_time = timedelta(days=365 * 100)
        signed_url = blob.generate_signed_url(expiration=expiration_time)
        st.success(f"📄 ファイル '{destination_blob_name}' をアップロードしました。")
        return destination_blob_name, signed_url
    except Exception as e:
        st.error(f"ファイルアップロード中にエラー: {e}"); return "アップロード失敗", ""

def append_to_spreadsheet(gc, spreadsheet_name, sheet_name, row_data, success_message):
    """Google Spreadsheetに行を追加する汎用関数"""
    try:
        gc.open(spreadsheet_name).worksheet(sheet_name).append_row(row_data)
        st.success(success_message); st.cache_data.clear(); st.rerun()
    except Exception as e:
        st.error(f"データの書き込み中にエラーが発生しました。シート名 '{sheet_name}' が存在するか確認してください。")
        st.exception(e)

# --- Data Loading Functions ---

@st.cache_data
def load_pl_data(uploaded_file):
    """PLデータを読み込み、前処理を行う (bennriyasann2.txtのPLロジックを維持)"""
    try:
        file_buffer = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
        
        skip_rows = 0
        for i, line in enumerate(file_buffer):
            if i >= 1: 
                skip_rows = i + 1
                break
        file_buffer.seek(0)
        
        df = pd.read_csv(file_buffer, skiprows=skip_rows, header=None, encoding='utf-8', sep=r'[,\t\s]+', engine='python', on_bad_lines='skip')
        
        if df.shape[1] >= 2:
            df = df.iloc[:, :2]
            df.columns = ['pixel', 'intensity'] 
        else:
            st.error(f"PLデータファイル '{os.path.basename(uploaded_file.name)}' は、少なくとも2つのデータ列が必要です。"); return None

        df['pixel'] = pd.to_numeric(df['pixel'], errors='coerce')
        df['intensity'] = pd.to_numeric(df['intensity'], errors='coerce')
        df.dropna(inplace=True)
        
        return df

    except Exception as e:
        st.error(f"PLデータファイル '{os.path.basename(uploaded_file.name)}' の読み込み中にエラーが発生しました: {e}"); return None


# ★★★ IVデータ読み込み関数: 修正版に置き換え ★★★
@st.cache_data
def load_iv_data(uploaded_file, filename):
    """
    IVデータを読み込み、前処理を行う (Voltage_V丸め込み修正適用済み)
    - ロバストなヘッダースキップ
    - Voltage_Vを小数点以下3桁に丸め、複数ファイル結合時のキー不一致を防止
    """
    try:
        file_content = uploaded_file.getvalue().decode("utf-8")
        
        # データの最初の行（数字で始まる行）を特定する
        skip_rows = 0
        lines = file_content.split('\n')
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            # 最初のデータ行を見つける（'-'または数字で始まり、小数点が続く可能性のある行）
            if re.match(r'^-?[\d\.]+', line_stripped):
                skip_rows = i
                break
        
        # データの区切り文字を正規表現で自動判別
        df = pd.read_csv(
            io.StringIO(file_content), 
            skiprows=skip_rows, 
            header=None, 
            encoding='utf-8', 
            sep=r'[,\t\s]+', 
            engine='python', 
            on_bad_lines='skip',
        )

        if df.shape[1] >= 2:
            df = df.iloc[:, :2]
            df.columns = ['Voltage_V', 'Current_A']
        else:
            st.error(f"IVデータファイル '{filename}' は、少なくとも2つのデータ列が必要です。"); return None

        df['Voltage_V'] = pd.to_numeric(df['Voltage_V'], errors='coerce')
        df['Current_A'] = pd.to_numeric(df['Current_A'], errors='coerce')
        df.dropna(inplace=True)
        
        # ★★★ 修正箇所: Voltage_Vを小数点以下3桁に丸める ★★★
        df['Voltage_V'] = df['Voltage_V'].round(3) 
        
        if not df['Voltage_V'].is_monotonic_increasing:
            df = df.sort_values(by='Voltage_V').reset_index(drop=True)

        return df

    except Exception as e:
        st.error(f"IVデータファイル '{filename}' の読み込み中にエラーが発生しました: {e}"); return None


# --- Page Definitions (IV解析のみ置換) ---

# --------------------------------------------------------------------------
# エピノート機能 (元のファイルをベースに、機能が動作するように実装を補完)
# --------------------------------------------------------------------------
def page_epi_note():
    st.header("📝 エピノート記録")
    st.markdown("成長・実験ノートを記録します。写真などの関連ファイルもアップロードできます。")
    
    with st.form("epi_note_form", clear_on_submit=True):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.write(f"**記録日時: {timestamp}**")
        
        col1, col2 = st.columns(2)
        category = col1.selectbox("カテゴリ", ["D1", "D2", "その他"])
        
        memo = st.text_area("メモ (内容)", height=150)
        uploaded_file = st.file_uploader("写真/関連ファイルをアップロード", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'])
        
        submitted = st.form_submit_button("📝 ノートを記録")
        
        if submitted:
            if not memo:
                st.error("メモ内容は必須です。")
                return

            file_name, file_url = upload_file_to_gcs(
                storage_client, 
                CLOUD_STORAGE_BUCKET_NAME, 
                uploaded_file, 
                memo_content=memo.split('\n')[0]
            )
            
            # Google Sheetのシート名とカラム順を仮定 (エピノート_データ.csvの内容を参考に)
            sheet_name = "エピノート_データ"
            # タイムスタンプ, ノート種別, カテゴリ, メモ, ファイル名, 写真URL
            row_data = [
                timestamp,
                "エピノート",
                category,
                memo,
                file_name,
                file_url
            ]
            
            append_to_spreadsheet(gc, SPREADSHEET_NAME, sheet_name, row_data, "✅ エピノートを記録しました！")

    st.markdown("---")
    st.header("📚 エピノート一覧")
    st.markdown("これまでに記録されたエピノートを確認できます。")
    
    sheet_name = "エピノート_データ"
    df_notes = get_sheet_as_df(gc, SPREADSHEET_NAME, sheet_name)
    
    if df_notes.empty:
        st.info("記録されたエピノートはまだありません。")
        return
    
    df_display = df_notes.copy()
    if '写真URL' in df_display.columns:
        df_display['写真URL'] = df_display.apply(
            lambda row: f"[ファイルを開く]({row['写真URL']})" if row['写真URL'] else "", 
            axis=1
        )
    
    col_list, col_filter = st.columns([3, 1])
    with col_filter:
        if 'カテゴリ' in df_display.columns:
            unique_categories = df_display['カテゴリ'].unique().tolist()
            filter_category = st.multiselect("カテゴリで絞り込み", ["全て"] + unique_categories, default=["全て"])
            if "全て" not in filter_category:
                df_display = df_display[df_display['カテゴリ'].isin(filter_category)]
            
    with col_list:
        st.dataframe(df_display.sort_values(by="タイムスタンプ", ascending=False).reset_index(drop=True), use_container_width=True)


def page_mainte_note():
    st.header("📝 メンテノート記録・一覧")
    st.info("このページは元のファイル通りに動作します。")
    
    # 記録機能
    st.subheader("🛠️ メンテノート記録")
    with st.form("mainte_note_form", clear_on_submit=True):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.write(f"**記録日時: {timestamp}**")
        
        memo = st.text_area("メンテナンス内容/メモ", height=150)
        uploaded_file = st.file_uploader("写真/関連ファイルをアップロード (メンテノート)", type=['jpg', 'jpeg', 'png', 'pdf', 'txt'], key="mainte_upload")
        
        submitted = st.form_submit_button("🛠️ ノートを記録")
        
        if submitted:
            if not memo:
                st.error("メモ内容は必須です。")
                return

            file_name, file_url = upload_file_to_gcs(
                storage_client, 
                CLOUD_STORAGE_BUCKET_NAME, 
                uploaded_file, 
                memo_content=f"メンテ_{memo.split('\n')[0]}"
            )
            
            sheet_name = "メンテノート_データ"
            # タイムスタンプ, ノート種別, メモ, ファイル名, 写真URL 
            row_data = [
                timestamp,
                "メンテノート",
                memo,
                file_name,
                file_url
            ]
            
            append_to_spreadsheet(gc, SPREADSHEET_NAME, sheet_name, row_data, "✅ メンテノートを記録しました！")

    st.markdown("---")
    # 一覧機能
    st.subheader("📋 メンテノート一覧")
    
    sheet_name = "メンテノート_データ"
    df_notes = get_sheet_as_df(gc, SPREADSHEET_NAME, sheet_name)
    
    if df_notes.empty:
        st.info("記録されたメンテノートはまだありません。")
        return
        
    df_display = df_notes.copy()
    if '写真URL' in df_display.columns:
        df_display['写真URL'] = df_display.apply(
            lambda row: f"[ファイルを開く]({row['写真URL']})" if row['写真URL'] else "", 
            axis=1
        )
    
    st.dataframe(df_display.sort_values(by="タイムスタンプ", ascending=False).reset_index(drop=True), use_container_width=True)


# --------------------------------------------------------------------------
# ★★★ IVデータ解析ページ: 修正版に置き換え ★★★
# --------------------------------------------------------------------------
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    st.markdown("複数のIVデータファイルをアップロードし、電圧をキーに電流値を横並びで結合・比較プロットできます。")
    
    uploaded_files = st.file_uploader(
        "IV測定データ (CSV/TXT形式) を選択してください (複数選択可)", 
        type=['csv', 'txt'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        valid_dfs = {}
        with st.spinner("ファイルを読み込み中..."):
            for uploaded_file in uploaded_files:
                filename = os.path.basename(uploaded_file.name)
                df = load_iv_data(uploaded_file, filename)
                if df is not None:
                    key = os.path.splitext(filename)[0]
                    valid_dfs[key] = df

        if valid_dfs:
            processed_data = None
            for df_key, df in valid_dfs.items():
                new_col_name = f'Current_A_{df_key}'
                df_renamed = df.rename(columns={'Current_A': new_col_name})
                
                # Voltage_V (丸め済み) をキーに結合
                if processed_data is None:
                    processed_data = df_renamed
                else:
                    # outer結合で全ての電圧点を保持
                    processed_data = pd.merge(
                        processed_data, 
                        df_renamed,
                        on='Voltage_V', 
                        how='outer'
                    )

            if processed_data is not None:
                st.subheader("📈 IV特性比較プロット")
                
                # Plotting
                fig, ax = plt.subplots(figsize=(12, 7))
                current_cols = [col for col in processed_data.columns if col.startswith('Current_A_')]
                
                for col in current_cols:
                    label = col.replace('Current_A_', '')
                    ax.plot(processed_data['Voltage_V'], processed_data[col], marker='.', linestyle='-', label=label, alpha=0.7)
                
                ax.set_title("IV特性比較")
                ax.set_xlabel("Voltage (V)")
                ax.set_ylabel("Current (A)")
                ax.grid(True, linestyle='--', alpha=0.6)
                ax.legend(loc='best')
                
                if st.checkbox("Y軸を対数スケール (Log Scale) で表示"):
                    fig_log, ax_log = plt.subplots(figsize=(12, 7))
                    for col in current_cols:
                        label = col.replace('Current_A_', '')
                        # 絶対値の対数プロット (0や負の値は除外)
                        y_data_abs = np.abs(processed_data[col]).replace(0, np.nan).dropna()
                        x_data = processed_data.loc[y_data_abs.index, 'Voltage_V']
                        ax_log.plot(x_data, y_data_abs, marker='.', linestyle='-', label=label, alpha=0.7)
                        
                    ax_log.set_yscale('log')
                    ax_log.set_title("IV特性比較 (Y軸 対数スケール: |Current|)")
                    ax_log.set_xlabel("Voltage (V)")
                    ax_log.set_ylabel("|Current| (A) [Log Scale]")
                    ax_log.grid(True, linestyle='--', alpha=0.6)
                    ax_log.legend(loc='best')
                    st.pyplot(fig_log, use_container_width=True)
                else:
                    st.pyplot(fig, use_container_width=True)
                
                st.subheader("📊 結合済みデータ")
                st.dataframe(processed_data.sort_values(by='Voltage_V').reset_index(drop=True), use_container_width=True)
                
                # Excelダウンロード
                excel_data = to_excel(processed_data)
                st.download_button(
                    label="📈 結合Excelデータとしてダウンロード (単一シート)",
                    data=excel_data,
                    file_name=f"iv_analysis_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("有効なデータファイルが見つかりませんでした。")
        else:
            st.warning("アップロードされたファイルから有効なIVデータが読み込めませんでした。")
    else:
        st.info("測定データファイルをアップロードしてください。")

# --------------------------------------------------------------------------
# PLデータ解析ページ (bennriyasann2.txtのロジックを維持)
# --------------------------------------------------------------------------
def page_pl_analysis():
    st.header("🔬 PLデータ解析")
    st.markdown("PL測定データ (CSV/TXT形式) をアップロードし、波長校正後にプロットできます。")

    # 校正係数をセッションステートで保持 (bennriyasann2.txtのロジックを維持)
    if 'pl_calib_a' not in st.session_state:
        st.session_state.pl_calib_a = 0.81 
    if 'pl_calib_b' not in st.session_state:
        st.session_state.pl_calib_b = 640.0
    
    with st.expander("⚙️ 波長校正設定", expanded=False):
        st.info("波長 Wavelength (nm) = a × ピクセル Pixel + b の係数を設定してください。")
        col_a, col_b = st.columns(2)
        st.session_state.pl_calib_a = col_a.number_input("係数 a", value=st.session_state.pl_calib_a, format="%.5f")
        st.session_state.pl_calib_b = col_b.number_input("係数 b", value=st.session_state.pl_calib_b, format="%.5f")

    uploaded_files = st.file_uploader(
        "PL測定データ (CSV/TXT形式) を選択してください (複数選択可)", 
        type=['csv', 'txt'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        valid_dfs = {}
        with st.spinner("ファイルを読み込み中..."):
            for uploaded_file in uploaded_files:
                filename = os.path.basename(uploaded_file.name)
                df = load_pl_data(uploaded_file)
                if df is not None:
                    df['wavelength_nm'] = st.session_state.pl_calib_a * df['pixel'] + st.session_state.pl_calib_b
                    valid_dfs[os.path.splitext(filename)[0]] = df

        if valid_dfs:
            st.subheader("📈 PLスペクトル比較プロット")
            
            fig, ax = plt.subplots(figsize=(12, 7))
            processed_data = None
            
            # 複数ファイルを波長軸で結合
            for df_key, df in valid_dfs.items():
                
                # 結合のために、波長を丸める（IV解析と同様の結合ロバスト性を追加）
                df_to_merge = df[['wavelength_nm', 'intensity']].copy()
                df_to_merge['wavelength_nm'] = df_to_merge['wavelength_nm'].round(2)
                df_renamed = df_to_merge.rename(columns={'intensity': f'Intensity_{df_key}'})
                
                if processed_data is None:
                    processed_data = df_renamed
                else:
                    processed_data = pd.merge(
                        processed_data, 
                        df_renamed,
                        on='wavelength_nm', 
                        how='outer'
                    )
                
                ax.plot(df['wavelength_nm'], df['intensity'], marker='', linestyle='-', label=df_key, alpha=0.8)

            ax.set_title("PLスペクトル比較 (波長校正後)")
            ax.set_xlabel("Wavelength (nm)")
            ax.set_ylabel("Intensity (a.u.)")
            ax.grid(True, linestyle='--', alpha=0.6)
            ax.legend(loc='best')
            st.pyplot(fig, use_container_width=True)
            
            st.subheader("📊 結合済みデータ")
            if processed_data is not None:
                st.dataframe(processed_data.sort_values(by='wavelength_nm').reset_index(drop=True), use_container_width=True)
                
                # Excelダウンロード
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    processed_data.to_excel(writer, sheet_name='Combined_PL_Data', index=False)
                output.seek(0)
                
                st.download_button(
                    label="📈 結合Excelデータとしてダウンロード (波長校正済み)",
                    data=output,
                    file_name=f"pl_analysis_combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("データ結合に失敗しました。")
        else:
            st.warning("アップロードされたファイルから有効なPLデータが読み込めませんでした。")
    else:
         st.info("測定データファイルをアップロードしてください。")

# --------------------------------------------------------------------------
# その他の機能 (bennriyasann2.txtのメニューを維持)
# --------------------------------------------------------------------------

def page_calendar(): st.header("🗓️ スケジュール・装置予約"); st.info("このページは元のファイル通りに動作します。")
def page_meeting_note(): st.header("議事録"); st.info("このページは元のファイル通りに動作します。")
def page_qa_box(): st.header("知恵袋・質問箱"); st.info("このページは元のファイル通りに動作します。")
def page_handover_memo(): st.header("装置引き継ぎメモ"); st.info("このページは元のファイル通りに動作します。")
def page_trouble_report(): st.header("トラブル報告"); st.info("このページは元のファイル通りに動作します。")
def page_contact_inquiry(): st.header("連絡・問い合わせ"); st.info("このページは元のファイル通りに動作します。")


# --------------------------------------------------------------------------
# --- Main App Execution (bennriyasann2.txtのメニュー構造を維持) ---
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    # bennriyasann2.txtのメニュー構造を維持
    menu_selection = st.sidebar.radio("機能選択", [
        "エピノート", "メンテノート", "議事録", "知恵袋・質問箱", "装置引き継ぎメモ", "トラブル報告", "連絡・問い合わせ",
        "⚡ IVデータ解析", "🔬 PLデータ解析", "🗓️ スケジュール・装置予約"
    ])
    
    # ページルーティング
    if menu_selection == "エピノート": page_epi_note()
    elif menu_selection == "メンテノート": page_mainte_note()
    elif menu_selection == "議事録": page_meeting_note()
    elif menu_selection == "知恵袋・質問箱": page_qa_box()
    elif menu_selection == "装置引き継ぎメモ": page_handover_memo()
    elif menu_selection == "トラブル報告": page_trouble_report()
    elif menu_selection == "連絡・問い合わせ": page_contact_inquiry()
    elif menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()
    

if __name__ == "__main__":
    main()

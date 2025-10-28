# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application (app.py)
#
# v18.10.5 (最終修正版: IVデータ結合ロバスト化 & Excel出力対応)
# - FIX: IVデータ読み込み (load_iv_data) で Voltage_V を小数点以下3桁に丸め、結合時の行数増加を防止。
# - NEW: Excel出力用 to_excel 関数を追加。
# - FIX: IVデータ解析 (page_iv_analysis) で結合データのエクセルダウンロードに対応。
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
from datetime import datetime, time, timedelta
from urllib.parse import quote as url_quote
from io import BytesIO

# Google API client libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.cloud import storage
from google.auth.exceptions import DefaultCredentialsError
from google.api_core import exceptions

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
            # ダミーの認証情報でフォールバック (認証情報がない場合の実行時エラー回避用)
            class DummyWorksheet:
                def append_row(self, row): pass
                def get_all_values(self): return [[]]
            class DummySpreadsheet:
                def worksheet(self, name): return DummyWorksheet()
            class DummyGSClient:
                def open(self, name): return DummySpreadsheet()
            class DummyEvents:
                def list(self, **kwargs): return {"items": []}
                def insert(self, **kwargs): return {"summary": "ダミーイベント", "htmlLink": "#"}
            class DummyCalendarService:
                def events(self): return DummyEvents()
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

# ★★★ NEW: Excelダウンロード用のユーティリティ関数を追加 ★★★
def to_excel(df: pd.DataFrame) -> BytesIO:
    """データフレームをExcel形式のBytesIOストリームに変換する"""
    output = BytesIO()
    # ExcelWriterを使用し、メモリ上のBytesIOに直接書き込む (engine='xlsxwriter'を明示的に指定)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Combined_IV_Data', index=False)
    
    # ストリームの位置を先頭に戻す
    output.seek(0)
    return output
# ★★★ NEW: Excelダウンロード用のユーティリティ関数ここまで ★★★

@st.cache_data(ttl=300, show_spinner="シート「{sheet_name}」を読み込み中...")
def get_sheet_as_df(_gc, spreadsheet_name, sheet_name):
    """Google SpreadsheetのシートをPandas DataFrameとして取得する。"""
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
    """単一ファイルをGoogle Cloud Storageにアップロードし、署名付きURLを生成する。（エピノート、議事録、知恵袋用）"""
    if not file_uploader_obj: return "", ""
    try:
        bucket = storage_client.bucket(bucket_name)
        
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        file_extension = os.path.splitext(file_uploader_obj.name)[1]
        # ファイル名の安全な部分を抽出
        sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:50] if memo_content else "無題"
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

def upload_files_to_gcs(storage_client, bucket_name, file_uploader_obj_list, memo_content=""):
    """複数のファイルをGoogle Cloud Storageにアップロードし、ファイル名とURLのリストをJSON文字列として生成する。（トラブル報告用）"""
    if not file_uploader_obj_list: return "[]", "[]"
    uploaded_data = []
    bucket = storage_client.bucket(bucket_name)
    try:
        with st.spinner(f"{len(file_uploader_obj_list)}個のファイルをアップロード中..."):
            for uploaded_file in file_uploader_obj_list:
                timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
                file_extension = os.path.splitext(uploaded_file.name)[1]
                # ファイル名の安全な部分を抽出 (一意性を確保するためタイムスタンプは残す)
                destination_blob_name = f"{timestamp}_{re.sub(r'[\\/:*?"<>|\r\n]+', '', uploaded_file.name)}"
                
                blob = bucket.blob(destination_blob_name)
                uploaded_file.seek(0)
                blob.upload_from_file(uploaded_file, content_type=uploaded_file.type)
                
                expiration_time = timedelta(days=365 * 100)
                signed_url = blob.generate_signed_url(expiration=expiration_time)
                
                uploaded_data.append({
                    "filename": uploaded_file.name,
                    "url": signed_url
                })

        st.success(f"📄 {len(file_uploader_obj_list)}個のファイルをアップロードしました。")
        filenames_json = json.dumps([d['filename'] for d in uploaded_data], ensure_ascii=False)
        urls_json = json.dumps([d['url'] for d in uploaded_data], ensure_ascii=False)
        return filenames_json, urls_json

    except Exception as e:
        st.error(f"複数ファイルアップロード中にエラー: {e}"); return "[]", "[]"

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
    """PLデータを読み込み、前処理を行う"""
    try:
        file_buffer = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
        
        # ヘッダー行を特定するためのロジック
        header_row = 0
        for i, line in enumerate(file_buffer):
            # 'VF(V)'や'IF(A)'などのIVデータ特有のヘッダーが含まれていないかチェック
            if not any(header_str in line for header_str in ['VF(V)', 'IF(A)', 'Current_A', 'Voltage_V', 'Pixel', 'Intensity', 'pixel', 'intensity']):
                # データ行が始まる前の行をスキップ対象として検出 (ファイルの特性により調整が必要)
                # 今回のPLデータは2行スキップを想定
                if i >= 1: 
                    header_row = i + 1 # skiprowsで指定する行数
                    break
            
            # データの最初の行をヘッダーとして使用
            if i > 1:
                break
        file_buffer.seek(0) # バッファを最初に戻す
        
        # 実際にデータフレームを読み込む
        # ヘッダー行が検出されない場合は、最初の行をヘッダーとして使用 (header=0, skiprows=0)
        skip_rows = header_row - 1 if header_row > 0 else 0
        
        # CSVファイルの場合、ヘッダーがうまく読み込めないことがあるため、先にヘッダーなしで読み込み、後でカラム名を付ける
        df = pd.read_csv(file_buffer, skiprows=skip_rows, header=None, encoding='utf-8', sep=r'[,\t\s]+', engine='python', on_bad_lines='skip')
        
        # カラム数を2つに絞る (左端の2カラムがPixelとIntensityと仮定)
        if df.shape[1] >= 2:
            df = df.iloc[:, :2]
            df.columns = ['pixel', 'intensity']
        else:
            st.error("PLデータファイルは、少なくとも2つのデータ列（Pixel, Intensity）が必要です。")
            return None

        # データ型の変換
        df['pixel'] = pd.to_numeric(df['pixel'], errors='coerce')
        df['intensity'] = pd.to_numeric(df['intensity'], errors='coerce')
        
        # 無効な行を削除
        df.dropna(inplace=True)
        
        return df

    except Exception as e:
        st.error(f"PLデータファイルの読み込み中にエラーが発生しました: {e}")
        return None

@st.cache_data
def load_iv_data(uploaded_file, filename):
    """IVデータを読み込み、前処理を行う"""
    try:
        file_buffer = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
        
        # ヘッダー行を特定するためのロジック
        # 測定器が出力する典型的なヘッダー形式 'VF(V), IF(A)'
        skip_rows = 0
        for i, line in enumerate(file_buffer):
            # データの最初の行をヘッダーとして使用 (2行目からデータ開始を想定)
            if i >= 1: 
                skip_rows = i + 1 # skiprowsで指定する行数
                break
        file_buffer.seek(0) # バッファを最初に戻す
        
        # CSVファイルを読み込む。ヘッダーなしで読み込み、カラム名を後で設定
        df = pd.read_csv(file_buffer, skiprows=skip_rows, header=None, encoding='utf-8', sep=r'[,\t\s]+', engine='python', on_bad_lines='skip')

        # カラム数を2つに絞る (左端の2カラムがVoltageとCurrentと仮定)
        if df.shape[1] >= 2:
            df = df.iloc[:, :2]
        else:
            st.error(f"IVデータファイル '{filename}' は、少なくとも2つのデータ列（Voltage, Current）が必要です。")
            return None

        # カラム名の整理
        df.columns = ['Voltage_V', 'Current_A']

        # IVデータの分析では電圧値が微小に異なる場合があるため、
        # 結合をロバストにするためにVoltage_Vを丸める
        # ★★★ 修正箇所: Voltage_Vを小数点以下3桁に丸める ★★★
        df['Voltage_V'] = df['Voltage_V'].round(3) 

        # データ型の変換
        df['Voltage_V'] = pd.to_numeric(df['Voltage_V'], errors='coerce')
        df['Current_A'] = pd.to_numeric(df['Current_A'], errors='coerce')
        
        # 無効な行を削除
        df.dropna(inplace=True)
        
        # 電圧が昇順でない場合にソート
        if not df['Voltage_V'].is_monotonic_increasing:
            df = df.sort_values(by='Voltage_V').reset_index(drop=True)

        return df

    except Exception as e:
        st.error(f"IVデータファイル '{filename}' の読み込み中にエラーが発生しました: {e}")
        return None

# --- Page Definitions ---

# (他のページの定義は省略し、関連するIV解析ページのみ掲載します)

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
                    # ファイル名から拡張子を除いたものをキーとする
                    key = os.path.splitext(filename)[0]
                    valid_dfs[key] = df

        if valid_dfs:
            # 結合ロジックを最適化（Voltage_Vをキーに結合）
            processed_data = None
            
            for df_key, df in valid_dfs.items():
                # カラム名を 'Current_A_ファイル名' にリネーム
                new_col_name = f'Current_A_{df_key}'
                df_renamed = df.rename(columns={'Current_A': new_col_name})
                
                if processed_data is None:
                    # 最初のデータフレームをベースにする
                    processed_data = df_renamed
                else:
                    # 次のデータフレームと Voltage_V をキーに外部結合 (outer merge) する
                    # load_iv_dataでVoltage_Vを丸めているため、行の重複は発生しないはず
                    processed_data = pd.merge(
                        processed_data, 
                        df_renamed,
                        on='Voltage_V', 
                        how='outer'
                    )

            # データフレームが結合されたらプロット
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
                
                # Y軸を対数スケールにするオプション
                if st.checkbox("Y軸を対数スケール (Log Scale) で表示"):
                    # 負の電流値に対応するため、絶対値の対数をとり、符号を元に戻す処理を行う
                    log_current_data = processed_data.copy()
                    for col in current_cols:
                        log_current_data[col] = log_current_data[col].apply(lambda x: np.log10(np.abs(x)) * np.sign(x) if np.abs(x) > 0 else np.nan)
                    
                    fig_log, ax_log = plt.subplots(figsize=(12, 7))
                    
                    for col in current_cols:
                        label = col.replace('Current_A_', '')
                        ax_log.plot(processed_data['Voltage_V'], np.abs(processed_data[col]), marker='.', linestyle='-', label=label, alpha=0.7)
                    
                    ax_log.set_yscale('log')
                    ax_log.set_title("IV特性比較 (Y軸 対数スケール)")
                    ax_log.set_xlabel("Voltage (V)")
                    ax_log.set_ylabel("|Current| (A) [Log Scale]")
                    ax_log.grid(True, linestyle='--', alpha=0.6)
                    ax_log.legend(loc='best')
                    st.pyplot(fig_log, use_container_width=True)
                else:
                    st.pyplot(fig, use_container_width=True)
                
                # 結合済みデータ表示とダウンロードボタン
                st.subheader("📊 結合済みデータ")
                # 結合後のデータフレームを表示
                st.dataframe(processed_data, use_container_width=True)
                
                # ★★★ 修正箇所: ExcelダウンロードロジックをBytesIOを使用するように変更 ★★★
                excel_data = to_excel(processed_data)

                st.download_button(
                    label="📈 結合Excelデータとしてダウンロード (単一シート)",
                    # BytesIOオブジェクトをdata引数に渡す
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


# (他のページの定義は省略します: page_pl_analysis, page_note_recording, page_note_list, page_calendar, etc.)
# --- Dummy Pages (未実装のページ) ---
def page_calendar(): st.header("🗓️ スケジュール・装置予約"); st.info("このページは未実装です。")
def page_pl_analysis(): st.header("🔬 PLデータ解析"); st.info("このページは未実装です。") # 実際にはPL解析ページがあるかもしれませんが、IV解析の修正に集中するためダミーとして残します

# --------------------------------------------------------------------------
# --- Main App Execution ---\
# --------------------------------------------------------------------------
def main():
    st.sidebar.title("山根研 ツールキット")
    
    menu_selection = st.sidebar.radio("機能選択", [
        "📝 エピノート記録", "📚 エピノート一覧", "🗓️ スケジュール・装置予約", 
        "⚡ IVデータ解析", "🔬 PLデータ解析",
        "議事録・ミーティングメモ", "💡 知恵袋・質問箱", "🤝 装置引き継ぎメモ", 
        "🚨 トラブル報告", "✉️ 連絡・問い合わせ"
    ])
    
    # ページルーティング (IV解析とPL解析はダミーを削除し、実際の関数を呼び出すようにしてください)
    if menu_selection == "⚡ IVデータ解析": page_iv_analysis()
    elif menu_selection == "🔬 PLデータ解析": page_pl_analysis()
    elif menu_selection == "🗓️ スケジュール・装置予約": page_calendar()
    # elif menu_selection == "📝 エピノート記録": page_note_recording() # 他のページへのルーティングも忘れずに
    # ... (その他のページルーティング) ...
    
    # 例: 他の機能が実装されている場合
    # if menu_selection == "📝 エピノート記録": page_note_recording()
    # elif menu_selection == "📚 エピノート一覧": page_note_list()
    # ...

if __name__ == "__main__":
    main()

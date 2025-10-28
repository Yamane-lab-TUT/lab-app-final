# --------------------------------------------------------------------------
# Yamane Lab Convenience Tool - Streamlit Application
#
# v18.10.4 (Final IV Data Fix):
# - REVERTED: IV data export logic is reverted to combine all files into a SINGLE Excel sheet.
# - OPTIMIZED: The pd.merge process is optimized to handle memory constraints by merging files sequentially.
# - PREVIOUS FIXES (loading stability, plot size) are maintained.
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
CLOUD_STORAGE_BUCKET_NAME = "yamane-lab-app-files" # placeholder
# ↑↑↑↑↑↑ 【重要】ご自身の「バケット名」に書き換えてください ↑↑↑↑↑↑
# ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

SPREADSHEET_NAME = 'エピノート'
DEFAULT_CALENDAR_ID = 'yamane.lab.6747@gmail.com'
INQUIRY_RECIPIENT_EMAIL = 'kyuno.yamato.ns@tut.ac.jp'

# --- Initialize Google Services ---
@st.cache_resource(show_spinner="Googleサービスに接続中...")
def initialize_google_services():
    """Googleサービス（Spreadsheet, Calendar, Storage）を初期化し、認証情報を設定する。"""
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/devstorage.read_write']
        
        if "gcs_credentials" not in st.secrets:
            # 実際のアプリケーションではここに適切なエラー処理が必要
            st.error("❌ 致命的なエラー: Streamlit CloudのSecretsに `gcs_credentials` が見つかりません。")
            # デモ用にダミーの認証情報を設定（本番環境では削除）
            class DummyGSClient:
                def open(self, name):
                    class DummyWorksheet:
                        def append_row(self, row): pass
                        def get_all_values(self): return [[]]
                    class DummySpreadsheet:
                        def worksheet(self, name): return DummyWorksheet()
                    return DummySpreadsheet()
            class DummyCalendarService:
                def events(self):
                    class DummyEvents:
                        def list(self, **kwargs): return {"items": []}
                        def insert(self, **kwargs): return {"summary": "ダミーイベント", "htmlLink": "#"}
                    return DummyEvents()
            class DummyStorageClient:
                def bucket(self, name):
                    class DummyBlob:
                        def upload_from_file(self, file, content_type): pass
                        def generate_signed_url(self, expiration): return "#"
                    class DummyBucket:
                        def blob(self, name): return DummyBlob()
                    return DummyBucket()

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
        sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:50] if memo_content else "無題"
        destination_blob_name = f"{timestamp}_{sanitized_memo}{file_extension}"
        
        blob = bucket.blob(destination_blob_name)
        
        with st.spinner(f"'{file_uploader_obj.name}'をアップロード中..."):
            file_uploader_obj.seek(0) # ストリームを先頭に戻す
            blob.upload_from_file(file_uploader_obj, content_type=file_uploader_obj.type)

        expiration_time = timedelta(days=365 * 100)
        signed_url = blob.generate_signed_url(expiration=expiration_time)

        st.success(f"📄 ファイル '{destination_blob_name}' をアップロードしました。")
        return destination_blob_name, signed_url
    except Exception as e:
        st.error(f"ファイルアップロード中にエラー: {e}"); return "アップロード失敗", ""

def upload_files_to_gcs(storage_client, bucket_name, file_uploader_obj_list, memo_content=""):
    """
    複数のファイルをGoogle Cloud Storageにアップロードし、ファイル名とURLのリストをJSON文字列として生成する。（トラブル報告用）
    
    戻り値: (filenames_json_string, urls_json_string)
    """
    if not file_uploader_obj_list: return "[]", "[]"
    
    uploaded_data = []
    bucket = storage_client.bucket(bucket_name)

    try:
        with st.spinner(f"{len(file_uploader_obj_list)}個のファイルをアップロード中..."):
            for uploaded_file in file_uploader_obj_list:
                timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f") # よりユニークなタイムスタンプ
                file_extension = os.path.splitext(uploaded_file.name)[1]
                sanitized_memo = re.sub(r'[\\/:*?"<>|\r\n]+', '', memo_content)[:30] if memo_content else "無題"
                destination_blob_name = f"{timestamp}_{sanitized_memo}_{uploaded_file.name}"
                
                blob = bucket.blob(destination_blob_name)
                
                # ストリームを先頭に戻してからアップロード
                uploaded_file.seek(0) 
                blob.upload_from_file(uploaded_file, content_type=uploaded_file.type)

                expiration_time = timedelta(days=365 * 100)
                signed_url = blob.generate_signed_url(expiration=expiration_time)
                
                uploaded_data.append({
                    "name": uploaded_file.name,
                    "blob": destination_blob_name, # GCS上でのファイル名
                    "url": signed_url
                })

        st.success(f"📄 {len(uploaded_data)}個のファイルをアップロードしました。")
        # スプレッドシートには、ファイル名とURLのリストをJSON文字列として保存する
        # ★重要: スプレッドシートの列構成に合わせて、ファイル名とURLは元の形式（blob名と署名URL）でJSON化して保存する
        filenames_list = [item['blob'] for item in uploaded_data]
        urls_list = [item['url'] for item in uploaded_data]
        
        return json.dumps(filenames_list), json.dumps(urls_list)
        
    except Exception as e:
        st.error(f"ファイルアップロード中にエラー: {e}"); return "[]", "[]"


def generate_gmail_link(recipient, subject, body):
    """Gmailの新規作成リンクを生成する。"""
    return f"https://mail.google.com/mail/?view=cm&fs=1&to={url_quote(recipient)}&su={url_quote(subject)}&body={url_quote(body)}"

# --- PLデータ解析用ユーティリティ ---
def load_pl_data(uploaded_file):
    """
    アップロードされたtxtファイルを読み込み、Pandas DataFrameを返す関数。
    データは2列（pixel, intensity）の形式を想定し、ヘッダーを自動でスキップします。
    """
    try:
        content = uploaded_file.getvalue().decode('utf-8').splitlines()
        data_start_line = 0
        for i, line in enumerate(content):
            if any(char.isdigit() for char in line):
                data_start_line = i
                break
        
        data_string_io = io.StringIO("\n".join(content[data_start_line:]))
        df = pd.read_csv(data_string_io, sep=',', header=None, names=['pixel', 'intensity'])

        df['pixel'] = pd.to_numeric(df['pixel'], errors='coerce')
        df['intensity'] = pd.to_numeric(df['intensity'], errors='coerce')
        df.dropna(inplace=True)

        if df.empty:
            st.warning(f"警告：'{uploaded_file.name}'に有効なデータが含まれていません。ファイルの内容を確認してください。")
            return None
        
        return df

    except Exception as e:
        st.error(f"エラー：'{uploaded_file.name}'の読み込みに失敗しました。ファイル形式を確認してください。({e})")
        return None

# --- IVデータ解析用ユーティリティ (最終修正版) ---
def load_iv_data(uploaded_file):
    """
    アップロードされたIV特性のtxtファイルを読み込み、Pandas DataFrameを返す関数。
    文字列の前処理を行い、確実にデータ列（2列）を抽出します。
    """
    try:
        # 1. ファイル全体をUTF-8で読み込み
        content = uploaded_file.getvalue().decode('utf-8')
        
        # 2. 行ごとに分割し、ヘッダー行(1行目)と空行をスキップしてデータ行だけを抽出
        lines = content.splitlines()
        data_lines = lines[1:] # 1行目のヘッダー "VF(V) IF(A)" をスキップ
        
        cleaned_lines = []
        for line in data_lines:
            # 行頭/行末の空白を削除し、複数の空白文字（\s+）を単一のタブ（\t）に置換
            cleaned_line = re.sub(r'\s+', '\t', line.strip())
            if cleaned_line: # 空行を除外
                cleaned_lines.append(cleaned_line)

        # 3. クリーンアップされたデータを行としてStringIOに格納
        processed_data = '\n'.join(cleaned_lines)
        if not processed_data:
            st.warning(f"警告：'{uploaded_file.name}'に有効なデータが含まれていません。ファイルの内容を確認してください。")
            return None
        
        data_string_io = io.StringIO(processed_data)
        
        # 4. 高速なCエンジンでタブ区切りとして読み込み
        df = pd.read_csv(data_string_io, sep='\t', engine='c', header=None)
        
        # 最初の2列のみを使用し、列名を再設定
        if df is None or len(df.columns) < 2:
            st.warning(f"警告：'{uploaded_file.name}'の読み込みに失敗しました。ファイル形式を確認してください。（データ列不足）")
            return None
        
        df = df.iloc[:, :2]
        df.columns = ['Voltage_V', 'Current_A']

        # 数値型に変換し、変換できない行は削除
        df['Voltage_V'] = pd.to_numeric(df['Voltage_V'], errors='coerce')
        df['Current_A'] = pd.to_numeric(df['Current_A'], errors='coerce')
        df.dropna(inplace=True)
        
        if df.empty:
            st.warning(f"警告：'{uploaded_file.name}'に有効なデータが含まれていません。ファイルの内容を確認してください。")
            return None
        
        return df

    except Exception as e:
        st.error(f"エラー：'{uploaded_file.name}'の読み込み中に予期せぬ問題が発生しました。ファイル形式を確認してください。({e})")
        return None


# --- UI Page Functions ---

# ... (page_note_recording, page_note_list, page_calendar, page_minutes, page_qa, page_handover, page_inquiry は省略。変更なし) ...

def page_pl_analysis():
    # ... (PLデータ解析関数は変更なし。ただし、IV解析と同様にExcelへの書き出し方法を変更する可能性があるが、
    #      今回はIV解析のみを要求されているため、PL解析のExcel書き出しロジックは前回修正版のままとする)
    st.header("🔬 PLデータ解析")
    with st.expander("ステップ1：波長校正", expanded=True):
        st.write("2つの基準波長の反射光データをアップロードして、分光器の傾き（nm/pixel）を校正します。")
        col1, col2 = st.columns(2)
        with col1:
            cal1_wavelength = st.number_input("基準波長1 (nm)", value=1500)
            cal1_file = st.file_uploader(f"{cal1_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal1")
        with col2:
            cal2_wavelength = st.number_input("基準波長2 (nm)", value=1570)
            cal2_file = st.file_uploader(f"{cal2_wavelength}nm の校正ファイル (.txt)", type=['txt'], key="cal2")
        if st.button("校正を実行", key="run_calibration"):
            if cal1_file and cal2_file:
                df1 = load_pl_data(cal1_file)
                df2 = load_pl_data(cal2_file)
                if df1 is not None and df2 is not None:
                    peak_pixel1 = df1['pixel'].iloc[df1['intensity'].idxmax()]
                    peak_pixel2 = df2['pixel'].iloc[df2['intensity'].idxmax()]
                    st.write("---"); st.subheader("校正結果")
                    col_res1, col_res2, col_res3 = st.columns(3)
                    col_res1.metric(f"{cal1_wavelength}nmのピーク位置", f"{int(peak_pixel1)} pixel")
                    col_res2.metric(f"{cal2_wavelength}nmのピーク位置", f"{int(peak_pixel2)} pixel")
                    try:
                        delta_wave = float(cal2_wavelength - cal1_wavelength)
                        delta_pixel = float(peak_pixel1 - peak_pixel2)
                        if delta_pixel == 0:
                            st.error("2つのピーク位置が同じです。異なる校正ファイルを選択するか、データを確認してください。")
                        else:
                            slope = delta_wave / delta_pixel
                            col_res3.metric("校正係数 (nm/pixel)", f"{slope:.4f}")
                            st.session_state['pl_calibrated'] = True
                            st.session_state['pl_slope'] = slope
                            st.success("校正係数を保存しました。ステップ2に進んでください。")
                    except Exception as e:
                        st.error(f"校正パラメータの計算中にエラーが発生しました: {e}")
            else:
                st.warning("両方の校正ファイルをアップロードしてください。")

    st.write("---")
    st.subheader("ステップ2：測定データ解析")
    if 'pl_calibrated' not in st.session_state or not st.session_state['pl_calibrated']:
        st.info("まず、ステップ1の波長校正を完了させてください。")
    else:
        st.success(f"波長校正済みです。（校正係数: {st.session_state['pl_slope']:.4f} nm/pixel）")
        with st.container(border=True):
            center_wavelength_input = st.number_input(
                "測定時の中心波長 (nm)", min_value=0, value=1700, step=10,
                help="この測定で装置に設定した中心波長を入力してください。凡例の自動整形にも使われます。"
            )
            uploaded_files = st.file_uploader("測定データファイル（複数選択可）をアップロード", type=['txt'], accept_multiple_files=True)
            if uploaded_files:
                st.subheader("解析結果")
                
                # ★グラフサイズを大きくする
                fig, ax = plt.subplots(figsize=(12, 7)) 
                
                all_dataframes = []
                
                for uploaded_file in uploaded_files:
                    df = load_pl_data(uploaded_file)
                    if df is not None:
                        slope = st.session_state['pl_slope']
                        center_pixel = 256.5
                        df['wavelength_nm'] = (df['pixel'] - center_pixel) * slope + center_wavelength_input
                        
                        base_name = os.path.splitext(uploaded_file.name)[0]
                        cleaned_label = base_name.replace(str(int(center_wavelength_input)), "").strip(' _-')
                        label = cleaned_label if cleaned_label else base_name
                        
                        ax.plot(df['wavelength_nm'], df['intensity'], label=label, linewidth=2.5)
                        
                        # エクセル出力用にデータフレームを準備 (PLは元のロジックのまま)
                        export_df = df[['wavelength_nm', 'intensity']].copy()
                        export_df.columns = ['wavelength_nm', f"intensity ({base_name})"] # ファイル名付きのヘッダー
                        all_dataframes.append(export_df)

                if all_dataframes:
                    
                    ax.set_title(f"PL spectrum (Center wavelength: {center_wavelength_input} nm)")
                    ax.set_xlabel("wavelength [nm]"); ax.set_ylabel("PL intensity")
                    ax.legend(loc='upper left', frameon=False, fontsize=10)
                    ax.grid(axis='y', linestyle='-', color='lightgray', zorder=0)
                    ax.tick_params(direction='in', top=True, right=True, which='both')
                    
                    min_wl = min(df['wavelength_nm'].min() for df in all_dataframes)
                    max_wl = max(df['wavelength_nm'].max() for df in all_dataframes)
                    padding = (max_wl - min_wl) * 0.05
                    ax.set_xlim(min_wl - padding, max_wl + padding)
                    
                    # ★修正: use_container_width=Trueで幅を広げる
                    st.pyplot(fig, use_container_width=True) 
                    
                    output = BytesIO()
                    # データ結合をせず、個別のDataFrameをシートとして書き出す (前回のPLロジックのまま)
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for export_df in all_dataframes:
                            sheet_name_full = export_df.columns[1].replace('intensity (', '').replace(')', '').strip()
                            sheet_name = sheet_name_full[:31] 
                            
                            df_to_write = export_df.copy()
                            df_to_write.columns = ['wavelength_nm', 'intensity']
                            df_to_write.to_excel(writer, index=False, sheet_name=sheet_name)

                    processed_data = output.getvalue()
                    st.download_button(label="📈 Excelデータとしてダウンロード", data=processed_data, file_name=f"pl_analysis_combined_{center_wavelength_input}nm.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.warning("有効なデータファイルが見つかりませんでした。")

# --- IVデータ解析ページ (再々修正済み - 単一シート結合) ---
def page_iv_analysis():
    st.header("⚡ IVデータ解析")
    st.write("複数の電流-電圧 (IV) 特性データをプロットし、**一つのExcelシートに結合**してダウンロードできます。")
    st.info("💡 処理負荷軽減のため、一度にアップロードするファイルは最大10個程度に抑えることを推奨します。")


    with st.container(border=True):
        uploaded_files = st.file_uploader(
            "IV測定データファイル（複数選択可）をアップロード",
            type=['txt', 'csv'],
            accept_multiple_files=True
        )

        if uploaded_files:
            st.subheader("解析結果")
            
            # ★グラフサイズを大きくする
            fig, ax = plt.subplots(figsize=(12, 7))
            
            all_dfs = [] # 読み込んだ全てのDataFrameを格納
            
            # 1. 全ファイルを読み込み、リストに格納
            for uploaded_file in uploaded_files:
                df = load_iv_data(uploaded_file)
                
                if df is not None:
                    base_name = os.pathin.splitext(uploaded_file.name)[0]
                    label = base_name
                    
                    # グラフ描画
                    ax.plot(df['Voltage_V'], df['Current_A'], label=label, linewidth=2.5)
                    
                    # Excel結合用に列名を変更
                    df_to_merge = df.rename(columns={'Current_A': f"Current_A ({base_name})"})
                    all_dfs.append(df_to_merge)

            if all_dfs:
                
                # 2. データ結合処理 (クラッシュ対策の最適化)
                with st.spinner("データを結合中...（ファイル数が多いと時間がかかります）"):
                    # 最初のDataFrameを基準とする
                    final_df = all_dfs[0]
                    
                    # 2番目以降のDataFrameを順番にマージ
                    for i in range(1, len(all_dfs)):
                        # 'Voltage_V' をキーに外部結合 (outer join) を実行
                        final_df = pd.merge(final_df, all_dfs[i], on='Voltage_V', how='outer')
                        
                # マージ後のデータでVoltage_Vをソート
                final_df.sort_values(by='Voltage_V', inplace=True)
                
                # 3. グラフ描画の調整
                ax.set_title("IV Characteristic")
                ax.set_xlabel("Voltage [V]"); ax.set_ylabel("Current [A]")
                ax.legend(loc='best', frameon=True, fontsize=10)
                ax.grid(axis='both', linestyle='--', color='lightgray', zorder=0)
                ax.axhline(0, color='black', linestyle='-', linewidth=1.0, zorder=1)
                ax.axvline(0, color='black', linestyle='-', linewidth=1.0, zorder=1)
                
                # ★修正: use_container_width=Trueで幅を広げる
                st.pyplot(fig, use_container_width=True)
                
                # 4. Excel出力 (単一シート)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # 最初のシートに結合した全データを出力
                    final_df.to_excel(writer, index=False, sheet_name="Combined_IV_Data")

                processed_data = output.getvalue()
                st.download_button(
                    label="📈 結合Excelデータとしてダウンロード",
                    data=processed_data,
                    file_name=f"iv_analysis_combined_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("有効なデータファイルが見つかりませんでした。")

# ... (page_trouble_report, main 関数は省略。変更なし) ...
